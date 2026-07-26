#!/usr/bin/env python3
"""
IR Emitter: parse deck HTML → structured IR JSON per KIWIAI-DECK-GENERATION-STANDARD-V1 §2.5
"""
import re, json, sys, os

def parse_deck(html_path):
    with open(html_path, 'r') as f:
        html = f.read()

    # Extract brand tokens from CSS variables
    tokens_match = re.search(r':root\s*\{([^}]+)\}', html)
    brand_tokens = {}
    if tokens_match:
        for line in tokens_match.group(1).split(';'):
            line = line.strip()
            if ':' in line:
                k, v = line.split(':', 1)
                brand_tokens[k.strip().lstrip('-')] = v.strip()

    # Find all slide divs
    pattern = r'<div class="slide[^"]*" id="(s\d+)">'
    matches = list(re.finditer(pattern, html))
    slides = []

    for m in matches:
        start = m.start()
        pos = m.end()
        depth = 1
        while depth > 0 and pos < len(html):
            next_open = html.find('<div', pos)
            next_close = html.find('</div>', pos)
            if next_open == -1: next_open = float('inf')
            if next_close == -1: next_close = float('inf')
            if next_close < next_open:
                depth -= 1
                pos = next_close + 6
            else:
                depth += 1
                pos = next_open + 4
        sh = html[start:pos]

        # Determine type
        slide_type = 'content'
        if 'slide-cover' in sh[:200]:
            slide_type = 'cover'
        elif 'slide-end' in sh[:200]:
            slide_type = 'end'

        slide_data = {'type': slide_type}

        # ---- Extract common elements ----
        # Section tag
        tag_match = re.search(r'class="header-tag"[^>]*>([^<]+)', sh)
        if tag_match:
            slide_data['tag'] = tag_match.group(1).strip()

        # Title
        title_match = re.search(r'class="header-title"[^>]*>([^<]+)', sh)

        # ---- Cover-specific ----
        cover_title = re.search(r'class="cover-title"[^>]*>([^<]+)', sh)
        cover_subtitle = re.search(r'class="cover-subtitle"[^>]*>\s*(.*?)\s*</div>', sh, re.DOTALL)
        # For cover: title = cover-title, subtitle = cover-subtitle
        if cover_title:
            slide_data['title'] = cover_title.group(1).strip()
        if cover_subtitle:
            sub = re.sub(r'<[^>]+>', '', cover_subtitle.group(1)).strip()
            # Split multi-line subtitle
            parts = [p.strip() for p in sub.split('\n') if p.strip()]
            slide_data['subtitle'] = ' | '.join(parts)

        # ---- Section-specific ----
        sec_title = re.search(r'class="section-title"[^>]*>([^<]+)', sh)
        if sec_title:
            slide_data['title'] = sec_title.group(1).strip()

        # ---- Content title ----
        if title_match and 'title' not in slide_data:
            slide_data['title'] = title_match.group(1).strip()

        # ---- Agenda items ----
        agenda_items = re.findall(r'class="agenda-text"[^>]*>([^<]+)', sh)
        if agenda_items:
            slide_data['agenda_items'] = [ai.strip() for ai in agenda_items]

        # ---- Bullets ----
        bullet_texts = re.findall(r'<li[^>]*>\s*(.*?)\s*</li>', sh, re.DOTALL)
        if bullet_texts:
            slide_data['bullets'] = [re.sub(r'<[^>]+>', '', b).strip() for b in bullet_texts]

        # ---- Cards ----
        card_blocks = re.findall(r'<div class="card">(.*?)</div>\s*</div>', sh, re.DOTALL)
        if card_blocks:
            cards = []
            for cb in card_blocks:
                card = {}
                ct = re.search(r'class="card-title"[^>]*>([^<]+)', cb)
                cd = re.search(r'class="card-desc"[^>]*>(.*?)</div>', cb, re.DOTALL)
                icon = re.search(r'class="card-icon"[^>]*>\s*(.*?)\s*</div>', cb, re.DOTALL)
                cstat = re.search(r'class="card-stat"[^>]*>([^<]+)', cb)
                cstat_label = re.search(r'class="card-stat-label"[^>]*>([^<]+)', cb)
                
                if ct: card['title'] = ct.group(1).strip()
                if cd: card['body'] = re.sub(r'<[^>]+>', '', cd.group(1)).strip()
                if icon: card['icon'] = icon.group(1).strip()
                if cstat: card['stat'] = cstat.group(1).strip()
                if cstat_label: card['stat_label'] = cstat_label.group(1).strip()
                if card:
                    cards.append(card)
            if cards:
                slide_data['cards'] = cards

        # ---- SVG processing ----
        svg_match = re.search(r'<svg[^>]*>(.*?)</svg>', sh, re.DOTALL)
        if svg_match:
            svg_full = svg_match.group(0)
            svg_texts = re.findall(r'<text[^>]*>([^<]+)</text>', svg_full)
            
            # Detect SVG kind
            has_rects = '<rect' in svg_full
            has_circles = '<circle' in svg_full
            num_rects = len(re.findall(r'<rect[^>]*/>', svg_full))
            num_circles = len(re.findall(r'<circle[^>]*/>', svg_full))

            # Heuristic: many similar-size rects with text = bar chart
            # circles + flow pattern = flow diagram
            # layered rects with distinct groups = layers

            # Check for layer structure (distinct y-groups of rects)
            rect_elements = re.findall(r'<rect[^>]*/>', svg_full)
            
            if num_circles >= 3 and num_rects >= 3:
                # Flow diagram
                slide_data['type'] = 'diagram'
                slide_data['svg_kind'] = 'flow'
                # Extract steps (pairs of number + name + description)
                steps = []
                # Pattern: number text → step name → description, repeated
                i = 0
                svg_texts_clean = svg_texts
                # Skip title (first text)
                if svg_texts_clean and not svg_texts_clean[0].isdigit():
                    svg_texts_clean = svg_texts_clean[1:]
                
                while i < len(svg_texts_clean):
                    step = {}
                    if i < len(svg_texts_clean) and svg_texts_clean[i].isdigit():
                        step['number'] = svg_texts_clean[i]
                        i += 1
                    if i < len(svg_texts_clean):
                        step['name'] = svg_texts_clean[i]
                        i += 1
                    if i < len(svg_texts_clean) and not svg_texts_clean[i].isdigit():
                        step['description'] = svg_texts_clean[i]
                        i += 1
                    if 'name' in step:
                        steps.append(step)
                if steps:
                    slide_data['steps'] = steps
                # Add slide-insight
                insight = re.search(r'class="slide-insight"[^>]*>(.*?)</div>', sh, re.DOTALL)
                if insight:
                    it = re.sub(r'<[^>]+>', ' ', insight.group(1)).strip()
                    it = re.sub(r'\s+', ' ', it)
                    slide_data['insight'] = it
                    
            elif num_rects > 10:
                # Layers diagram
                slide_data['type'] = 'diagram'
                slide_data['svg_kind'] = 'layers'
                # Parse layers
                layers = []
                current_layer = None
                layer_pattern = re.finditer(r'<text[^>]*font-size="16"[^>]*>([^<]+)</text>', svg_full)
                layer_names = [m.group(1) for m in layer_pattern]
                
                # Get all item texts (font-size="13")
                item_pattern = re.finditer(r'<text[^>]*font-size="13"[^>]*>([^<]+)</text>', svg_full)
                items = [m.group(1) for m in item_pattern]
                
                # Items are grouped by 4
                for li, lname in enumerate(layer_names):
                    start_idx = li * 4
                    layer_items = items[start_idx:start_idx+4] if start_idx < len(items) else []
                    layers.append({'name': lname, 'items': layer_items})
                
                if layers:
                    slide_data['layers'] = layers
                    
                # Add slide-insight
                insight = re.search(r'class="slide-insight"[^>]*>(.*?)</div>', sh, re.DOTALL)
                if insight:
                    it = re.sub(r'<[^>]+>', ' ', insight.group(1)).strip()
                    it = re.sub(r'\s+', ' ', it)
                    slide_data['insight'] = it
            else:
                # Bar chart
                slide_data['type'] = 'chart'
                slide_data['svg_kind'] = 'bar'
                slide_data['unit'] = '%'
                
                # Parse bar chart SVG: title, '％', axis ticks..., value1, label1, value2, label2...
                # The SVG text order is: title, %, y-axis-labels(0,17,35,52,70), value1, label1, value2, label2...
                title_text = svg_texts[0] if svg_texts else ''
                
                # Identify value positions: numbers > 0 that appear after y-axis ticks
                # y-axis ticks are: 0, 17, 35, 52, 70 - they are numbers too
                # values are the numbers that appear AFTER the axis tick group
                # Strategy: find numbers, skip the axis group (first 5 numbers after %), 
                # the rest are chart values
                all_numbers = []
                for vt in svg_texts:
                    if vt.lstrip('-').isdigit():
                        all_numbers.append(vt)
                
                # Values come in second group (after axis ticks 0,17,35,52,70)
                # The axis ticks are 0, 17, 35, 52, 70 → 5 numbers
                chart_values = all_numbers[5:] if len(all_numbers) > 5 else all_numbers
                
                # Labels are non-numeric texts that are not title/％/axis-ticks
                skip_set = {'%', title_text}
                labels = []
                for vt in svg_texts:
                    if vt in skip_set:
                        continue
                    if vt.lstrip('-').isdigit():
                        continue
                    if vt in labels:
                        continue
                    labels.append(vt)
                
                # Trim labels to match values count
                values = [int(v) for v in chart_values]
                if len(labels) != len(values):
                    labels = labels[:len(values)]
                
                if labels and values:
                    slide_data['labels'] = labels
                    slide_data['values'] = values
                
                # Add slide-insight
                insight = re.search(r'class="slide-insight"[^>]*>(.*?)</div>', sh, re.DOTALL)
                if insight:
                    it = re.sub(r'<[^>]+>', ' ', insight.group(1)).strip()
                    it = re.sub(r'\s+', ' ', it)
                    slide_data['insight'] = it

        # ---- Highlight box ----
        highlight = re.search(r'class="highlight-box"[^>]*>(.*?)</div>\s*</div>\s*</div>', sh, re.DOTALL)
        if highlight:
            ht = re.sub(r'<[^>]+>', ' ', highlight.group(1)).strip()
            ht = re.sub(r'\s+', ' ', ht)
            slide_data['highlight'] = ht

        # ---- Stat items ----
        stat_items = re.findall(r'class="stat-item"[^>]*>(.*?)</div>', sh, re.DOTALL)
        if stat_items:
            stats = []
            for si in stat_items:
                sn = re.search(r'class="stat-num"[^>]*>([^<]+)', si)
                sl = re.search(r'class="stat-label"[^>]*>([^<]+)', si)
                if sn and sl:
                    stats.append({'value': sn.group(1).strip(), 'label': sl.group(1).strip()})
            if stats:
                slide_data['stats'] = stats

        # ---- Timeline ----
        tl_items = re.findall(r'class="timeline-item"[^>]*>(.*?)</div>\s*</div>', sh, re.DOTALL)
        if tl_items:
            timeline = []
            for ti in tl_items:
                phase = re.search(r'class="timeline-phase"[^>]*>([^<]+)', ti)
                ttitle = re.search(r'class="timeline-title"[^>]*>([^<]+)', ti)
                tdesc = re.search(r'class="timeline-desc"[^>]*>([^<]+)', ti)
                item = {}
                if phase: item['phase'] = phase.group(1).strip()
                if ttitle: item['title'] = ttitle.group(1).strip()
                if tdesc: item['description'] = tdesc.group(1).strip()
                if item:
                    timeline.append(item)
            if timeline:
                slide_data['timeline'] = timeline

        # ---- Two-column detection (background slide has col-left/col-right) ----
        has_two_col = 'col-left' in sh and 'col-right' in sh
        if has_two_col:
            slide_data['layout'] = 'two-column'

        # ---- End slide ----
        end_title = re.search(r'class="end-title"[^>]*>([^<]+)', sh)
        end_subtitle = re.search(r'class="end-subtitle"[^>]*>([^<]+)', sh)
        end_contact = re.search(r'class="end-contact"[^>]*>(.*?)</div>', sh, re.DOTALL)
        if end_title:
            slide_data['title'] = end_title.group(1).strip()
        if end_subtitle:
            slide_data['subtitle'] = end_subtitle.group(1).strip()
        if end_contact:
            ec = re.sub(r'<br\s*/?>', ' | ', end_contact.group(1)).strip()
            ec = re.sub(r'<[^>]+>', '', ec).strip()
            slide_data['contact'] = ec

        slides.append(slide_data)

    # Build final IR
    ir = {
        'meta': {
            'source': os.path.basename(html_path),
            'slide_count': len(slides),
            'standard': 'KIWIAI-DECK-GENERATION-STANDARD-V1 §2.5'
        },
        'brand_tokens': {
            'palette': {
                'primary': brand_tokens.get('brand-primary', '#0070C0'),
                'primary_dark': brand_tokens.get('brand-primary-dark', '#005A9B'),
                'primary_light': brand_tokens.get('brand-primary-light', '#E6F2FA'),
                'primary_pale': brand_tokens.get('brand-primary-pale', '#F0F7FD'),
                'accent': brand_tokens.get('brand-accent', '#00A3E0'),
            },
            'typography': {
                'heading': 'Microsoft YaHei',
                'body': 'Microsoft YaHei',
            },
            'cover': {
                'bg_color': brand_tokens.get('cover-bg', '#F2F3F5'),
                'title_color': brand_tokens.get('brand-primary', '#0070C0'),
                'title_size': 34,
                'subtitle_color': brand_tokens.get('cover-text-muted', '#666'),
                'subtitle_size': 14,
            }
        },
        'cover_reference': {
            'layout': 'centered',
            # company_name is NOT extracted from the HTML cover slide yet; left empty
            # so NO customer name is hard-coded (redline). The chatflow side (card 2) /
            # portal supplies the real values, or a later emit_ir revision parses them
            # from the deck HTML cover/end slide.
            'company_name_cn': '',
            'company_name_en': '',
            'logo_path': 'logo.png'
        },
        'slides': slides
    }

    return ir


def main():
    html_path = sys.argv[1] if len(sys.argv) > 1 else 'deck.html'
    out_path = sys.argv[2] if len(sys.argv) > 2 else 'deck_ir.json'
    
    ir = parse_deck(html_path)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(ir, f, ensure_ascii=False, indent=2)
    
    print(f"IR emitted: {ir['meta']['slide_count']} slides → {out_path}")
    for i, s in enumerate(ir['slides']):
        print(f"  Slide {i+1}: type={s['type']}, title={s.get('title','?')[:50]}")

if __name__ == '__main__':
    main()
