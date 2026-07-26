"""Brand deck renderers — IR-adaptive brand rendering for artifact_generator.

Turns a deck IR (``brand_tokens`` + ``cover_reference`` + ``slides[]``) into
branded ``.pptx`` / ``.docx`` bytes. Consumed by ``app/server.py``'s
``generate_pptx`` / ``generate_docx`` ONLY when the request body carries a deck
IR (``brand_tokens`` + ``slides[]``); bodies with only ``sections[]`` keep using
the legacy synthetic renderer (no regression, API contract unchanged).

Design properties (per KPRY-CODER-ARTIFACT-GENERATOR-BRAND-UPGRADE):
  * Environment-adaptive SVG→PNG: local Chrome > Docker alpine-chrome > BLOCK
    (never spins on ``docker pull`` — ADR-017).
  * Customer-agnostic: NO customer name/slug/path literal lives here. Brand data
    comes from the IR (``brand_tokens`` / ``cover_reference``) and the logo from a
    caller-supplied ``brand_dir`` (multi-tenant). This makes the renderer reusable
    across every channel (portal / WeCom / future) and every tenant.

Exports:
    render_pptx_from_ir(ir, brand_dir=None, logo_path=None, deck_html=None, work_dir=None) -> pptx bytes
    render_docx_from_ir(ir, brand_dir=None, work_dir=None) -> docx bytes
    parse_deck(html_path) -> IR dict   (HTML→IR tool, for chatflow/portal use)
"""
from .render_pptx import render_pptx_from_ir
from .render_docx import render_docx_from_ir
from .emit_ir import parse_deck

__all__ = ["render_pptx_from_ir", "render_docx_from_ir", "parse_deck"]
