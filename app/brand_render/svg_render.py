#!/usr/bin/env python3
"""
Shared SVG helpers for the brand renderers.

Two concerns, both environment- and customer-agnostic (no customer literals):

  1. ``svg_to_png`` — render an SVG (wrapped in a minimal HTML shell) to a PNG.
     Environment-adaptive per KIWIAI-DECK-GENERATION-STANDARD-V1 §2.5.1:
       a. Local chrome/chromium binary present  -> use it directly (no Docker).
       b. No local binary, alpine-chrome already pulled locally -> use Docker.
       c. Neither available -> BLOCK (raise RuntimeError with probe results).
     Never spins on ``docker pull`` (per ADR-017 / §2 降级禁令). Failure always
     raises — no silent degradation.

  2. ``svg_layers`` / ``svg_flow`` / ``svg_bar`` — generate diagram SVG straight
     from IR data (architecture layers / process flow / bar chart). Lets both the
     pptx and docx brand renderers render diagrams from the IR alone, with no
     dependency on an external ``deck.html`` file.
"""
import os
import shutil
import subprocess

# Candidate local Chrome/Chromium binaries (checked in order; first existing wins).
# Plain binary paths only — no customer/environment literals.
_LOCAL_CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
]


def find_local_chrome():
    """Return path to a local chrome/chromium binary, or None if none found."""
    for c in _LOCAL_CHROME_CANDIDATES:
        if os.path.exists(c):
            return c
    for name in ("google-chrome", "chromium", "chromium-browser", "chrome"):
        found = shutil.which(name)
        if found:
            return found
    return None


def docker_image_present(image):
    """True iff the named docker image is already pulled locally."""
    try:
        result = subprocess.run(
            ["docker", "images", "-q", image],
            capture_output=True, timeout=10, text=True,
        )
        return result.returncode == 0 and result.stdout.strip() != ""
    except Exception:
        return False


def _cleanup_render(html_file, chrome_profile=None):
    """Remove render scratch artifacts (HTML wrapper + chrome user-data-dir).

    ``chrome_profile`` is None on the Docker branch, so this is a safe no-op for
    extra cleanup there. Always best-effort (never masks the real render error).
    """
    try:
        os.unlink(html_file)
    except Exception:
        pass
    if chrome_profile:
        shutil.rmtree(chrome_profile, ignore_errors=True)


def svg_to_png(svg_content, out_path, width=1000, height=520):
    """Render an SVG (wrapped in a minimal HTML shell) to a PNG.

    Raises RuntimeError on any failure (missing renderer, no output, tiny output,
    timeout). Never silently degrades.
    """
    html = (
        '<!DOCTYPE html><html><head><meta charset="UTF-8">'
        "<style>body{margin:0;width:" + str(width) + "px;height:" + str(height)
        + "px;display:flex;align-items:center;justify-content:center;"
        'background:white;font-family:"Microsoft YaHei","PingFang SC",'
        '"Noto Sans SC",sans-serif}svg{width:' + str(width) + "px;height:"
        + str(height) + "px}</style></head><body>" + svg_content + "</body></html>"
    )
    html_file = out_path + ".html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html)
    abs_html = os.path.abspath(html_file)
    abs_png = os.path.abspath(out_path)
    html_dir = os.path.dirname(abs_html)
    html_name = os.path.basename(abs_html)
    png_name = os.path.basename(abs_png)

    local_chrome = find_local_chrome()
    alpine_image = "zenika/alpine-chrome"
    docker_ready = docker_image_present(alpine_image)

    # Local chrome on a hardened host (e.g. systemd ProtectHome makes /home
    # read-only) dies with "mkdir '/home': Permission denied" because it tries
    # to seed its profile under /home/<user>/.local. Point BOTH --user-data-dir
    # and HOME at the render work dir (html_dir is guaranteed writable — we just
    # wrote the HTML wrapper there) so chrome is fully self-contained and never
    # touches /home, on any host. The Docker branch below is intentionally left
    # untouched (its container has its own writable /workspace); the
    # env-adaptive selection above is unchanged (LESSONS §F1-f — never hardcode
    # a single render path, keep probing).
    env = None
    chrome_profile = None
    if local_chrome:
        chrome_profile = os.path.join(html_dir, ".chrome-profile")
        try:
            os.makedirs(chrome_profile, exist_ok=True)
        except Exception:
            pass
        cmd = [
            local_chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars",
            "--no-first-run", "--no-default-browser-check",
            "--user-data-dir=%s" % chrome_profile,
            "--window-size=%d,%d" % (width, height),
            "--screenshot=%s" % abs_png, "file://%s" % abs_html,
        ]
        env = dict(os.environ)
        env["HOME"] = html_dir
    elif docker_ready:
        try:
            os.chmod(html_dir, 0o777)
        except Exception:
            pass
        cmd = [
            "docker", "run", "--rm", "-v", "%s:/workspace" % html_dir, alpine_image,
            "--no-sandbox", "--headless=new", "--disable-gpu", "--hide-scrollbars",
            "--window-size=%d,%d" % (width, height),
            "--screenshot=/workspace/%s" % png_name,
            "file:///workspace/%s" % html_name,
        ]
    else:
        _cleanup_render(html_file, chrome_profile)
        raise RuntimeError(
            "SVG render BLOCK: no local chrome/chromium binary found and Docker "
            "image '%s' not present locally. Environment probe: "
            "local_chrome_candidates=%s, "
            "which(google-chrome/chromium/chromium-browser/chrome)=None, "
            "docker_image_present(%s)=%s. "
            "Refusing to `docker pull` and spin (per ADR-017). Install local "
            "Chrome/Chromium, or pre-pull the alpine-chrome image, then retry."
            % (alpine_image, _LOCAL_CHROME_CANDIDATES, alpine_image, docker_ready)
        )

    try:
        result = subprocess.run(
            cmd, capture_output=True, timeout=60, text=True, env=env)
    except subprocess.TimeoutExpired:
        _cleanup_render(html_file, chrome_profile)
        raise RuntimeError("SVG render timeout (60s). Cmd: %s" % " ".join(cmd))
    except Exception as e:
        _cleanup_render(html_file, chrome_profile)
        raise RuntimeError("SVG render error: %s" % e)

    _cleanup_render(html_file, chrome_profile)
    if not os.path.exists(out_path):
        raise RuntimeError(
            "SVG render failed: no output. Cmd: %s\nstdout: %s\nstderr: %s\nrc: %s"
            % (" ".join(cmd), result.stdout, result.stderr, result.returncode)
        )
    if os.path.getsize(out_path) < 100:
        raise RuntimeError(
            "SVG render failed: file too small (%d bytes)"
            % os.path.getsize(out_path)
        )
    return True


# ---------------------------------------------------------------------------
# SVG diagram generators (from IR data — no external deck.html dependency)
# ---------------------------------------------------------------------------

def svg_layers(title, layers, primary="#0070C0", primary_pale="#F0F7FD"):
    """Render architecture layers diagram from IR ``layers`` data."""
    n = len(layers)
    layer_h = 420 // n if n else 420
    rows = []
    y = 60
    for layer in layers:
        name = layer.get("name", "")
        items = layer.get("items", [])
        items_str = " · ".join(items)
        rows.append(
            '<rect x="60" y="%d" width="880" height="%d" rx="6" fill="%s" stroke="%s" stroke-width="1.5"/>'
            % (y, layer_h - 8, primary_pale, primary)
            + '<text x="85" y="%d" font-family="Microsoft YaHei, sans-serif" font-size="15" font-weight="bold" fill="%s">%s</text>'
            % (y + 25, primary, _esc(name))
            + '<text x="85" y="%d" font-family="Microsoft YaHei, sans-serif" font-size="12" fill="#555">%s</text>'
            % (y + 48, _esc(items_str))
        )
        y += layer_h
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 520" width="1000" height="520">'
        '<rect width="1000" height="520" fill="white"/>'
        '<text x="500" y="35" font-family="Microsoft YaHei, sans-serif" font-size="18" font-weight="bold" fill="#1A1A2E" text-anchor="middle">%s</text>'
        '%s</svg>' % (_esc(title), "".join(rows))
    )


def svg_flow(title, steps, primary="#0070C0", primary_light="#E6F2FA"):
    """Render flow/process steps diagram from IR ``steps`` data."""
    n = len(steps)
    step_w = 880 // n if n else 880
    arrows = []
    for i in range(n - 1):
        ax = 100 + (i + 1) * step_w - 25
        arrows.append(
            '<polygon points="%d,230 %d,245 %d,260" fill="%s" opacity="0.5"/>'
            % (ax, ax + 20, ax, primary)
        )
    boxes = []
    for i, step in enumerate(steps):
        x = 60 + i * step_w
        num = step.get("number", str(i + 1))
        name = step.get("name", "")
        desc = step.get("description", "")
        boxes.append(
            '<rect x="%d" y="150" width="%d" height="100" rx="8" fill="%s" stroke="%s" stroke-width="1.5"/>'
            % (x + 5, step_w - 30, primary_light, primary)
            + '<circle cx="%d" cy="175" r="14" fill="%s"/>'
            % (x + step_w / 2 - 10, primary)
            + '<text x="%d" y="180" font-family="Microsoft YaHei, sans-serif" font-size="12" font-weight="bold" fill="white" text-anchor="middle">%s</text>'
            % (x + step_w / 2 - 10, _esc(num))
            + '<text x="%d" y="215" font-family="Microsoft YaHei, sans-serif" font-size="13" font-weight="bold" fill="#1A1A2E" text-anchor="middle">%s</text>'
            % (x + step_w / 2 - 10, _esc(name))
            + '<text x="%d" y="235" font-family="Microsoft YaHei, sans-serif" font-size="10" fill="#666" text-anchor="middle">%s</text>'
            % (x + step_w / 2 - 10, _esc(desc))
        )
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 520" width="1000" height="520">'
        '<rect width="1000" height="520" fill="white"/>'
        '<text x="500" y="35" font-family="Microsoft YaHei, sans-serif" font-size="18" font-weight="bold" fill="#1A1A2E" text-anchor="middle">%s</text>'
        '%s%s</svg>' % (_esc(title), "".join(arrows), "".join(boxes))
    )


def svg_bar(title, labels, values, unit, primary="#0070C0"):
    """Render bar chart SVG for reference (chart in docx uses tables;
    pptx uses native charts). Kept for parity / tooling."""
    n = len(labels)
    if not values:
        return None
    max_v = max(values) * 1.15 or 1
    bar_w = 800 // n - 20 if n else 0
    bars = []
    for i, (label, val) in enumerate(zip(labels, values)):
        x = 80 + i * (800 // n if n else 0)
        bar_h = (val / max_v) * 300
        y = 440 - bar_h
        bars.append(
            '<rect x="%d" y="%d" width="%d" height="%d" rx="4" fill="%s" opacity="0.85"/>'
            % (x, y, bar_w, bar_h, primary)
            + '<text x="%d" y="%d" font-family="Microsoft YaHei, sans-serif" font-size="13" font-weight="bold" fill="#1A1A2E" text-anchor="middle">%s%s</text>'
            % (x + bar_w / 2, y - 8, val, _esc(unit))
            + '<text x="%d" y="470" font-family="Microsoft YaHei, sans-serif" font-size="10" fill="#666" text-anchor="middle">%s</text>'
            % (x + bar_w / 2, _esc(label))
        )
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 520" width="1000" height="520">'
        '<rect width="1000" height="520" fill="white"/>'
        '<text x="500" y="30" font-family="Microsoft YaHei, sans-serif" font-size="18" font-weight="bold" fill="#1A1A2E" text-anchor="middle">%s</text>'
        '%s</svg>' % (_esc(title), "".join(bar for bar in bars))
    )


def _esc(s):
    """Minimal XML text escaping for SVG <text> content."""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
