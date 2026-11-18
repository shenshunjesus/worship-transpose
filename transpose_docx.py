#!/usr/bin/env python3
"""
Lossless .docx & Text Chord Sheet Transposer (敬拜歌譜轉調工具 CLI)
Supports:
- Lossless .docx XML transposition (preserves all fonts, sizes, grey shading, margins, embedded fonts)
- Smart enharmonic spelling (# vs b based on target key, plus manual override)
- Nashville Number System (級數譜 1 - 6m - 4 - 5) & Dual Chord+Number mode
- Automatic metadata header key update (e.g., "Upper Room/活水/D" -> "Upper Room/活水/C")
- Capo recommendations for guitarists
"""

import argparse
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional, Tuple, List, Dict

NOTES_SHARP = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
NOTES_FLAT  = ['C', 'Db', 'D', 'Eb', 'E', 'F', 'Gb', 'G', 'Ab', 'A', 'Bb', 'B']

NOTE_TO_SEMITONE = {
    'C': 0, 'B#': 0,
    'C#': 1, 'Db': 1, 'C♯': 1, 'D♭': 1,
    'D': 2,
    'D#': 3, 'Eb': 3, 'D♯': 3, 'E♭': 3,
    'E': 4, 'Fb': 4,
    'F': 5, 'E#': 5,
    'F#': 6, 'Gb': 6, 'F♯': 6, 'G♭': 6,
    'G': 7,
    'G#': 8, 'Ab': 8, 'G♯': 8, 'A♭': 8,
    'A': 9,
    'A#': 10, 'Bb': 10, 'A♯': 10, 'B♭': 10,
    'B': 11, 'Cb': 11,
}

# Keys that naturally prefer flats (Circle of Fifths: F, Bb, Eb, Ab, Db, Gb and their relative minors)
FLAT_KEYS = {
    'F', 'Bb', 'Eb', 'Ab', 'Db', 'Gb',
    'Dm', 'Gm', 'Cm', 'Fm', 'Bbm', 'Ebm',
    'B♭', 'E♭', 'A♭', 'D♭', 'G♭'
}

# Major scale intervals in semitones from root: 1, 2, 3, 4, 5, 6, 7
SCALE_DEGREES_SHARP = ['1', '#1', '2', '#2', '3', '4', '#4', '5', '#5', '6', 'b7', '7']
SCALE_DEGREES_FLAT  = ['1', 'b2', '2', 'b3', '3', '4', 'b5', '5', 'b6', '6', 'b7', '7']

# Regex for matching a chord token:
# Root: [A-G][#b♯♭]?
# Quality/Extension: (?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\-\^\(\)])*
# Optional Slash Bass: (?:/[A-G][#b♯♭]?)?
CHORD_TOKEN_RE = re.compile(
    r'(?<![A-Za-z0-9])'
    r'([A-G][#b♯♭]?)'
    r'((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)'
    r'(?:/([A-G][#b♯♭]?))?'
    r'(?![A-Za-z0-9])'
)

SECTION_HEADERS = {
    'intro', 'verse', 'verse1', 'verse2', 'verse3', 'verse4',
    'pre-chorus', 'prechorus', 'chorus', 'chorus1', 'chorus2',
    'bridge', 'bridge1', 'bridge2', 'tag', 'outro', 'interlude',
    'ending', 'solo', 'instrumental', 'vamp', 'turnaround'
}


def should_use_flats(target_key: str, accidental_pref: str = 'auto') -> bool:
    if accidental_pref == 'flat':
        return True
    if accidental_pref == 'sharp':
        return False
    return target_key in FLAT_KEYS


def semitone_to_note(semitone: int, use_flats: bool) -> str:
    idx = semitone % 12
    return NOTES_FLAT[idx] if use_flats else NOTES_SHARP[idx]


def semitone_to_number(semitone: int, key_semitone: int, use_flats: bool = True) -> str:
    diff = (semitone - key_semitone) % 12
    return SCALE_DEGREES_FLAT[diff] if use_flats else SCALE_DEGREES_SHARP[diff]


def is_chord_line(line: str) -> bool:
    """
    Determine whether a line is a chord progression line (like 'G - Bm - ｜D - - -')
    rather than normal English/Chinese lyrics or section headers.
    """
    stripped = line.strip()
    if not stripped:
        return False
    low = stripped.lower().rstrip(':')
    if low in SECTION_HEADERS or re.match(r'^(intro|verse|chorus|bridge|tag|outro|interlude|pre-?chorus)\s*\d*$', low):
        return False
    # If line contains Chinese characters and has no barlines or beat dashes, check carefully
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', stripped))
    # Tokenize by whitespace and barlines/separators
    tokens = [t for t in re.split(r'[\s｜|\-–—/、,，\(\)\[\]]+', stripped) if t]
    if not tokens:
        return False
    chord_matches = 0
    for tok in tokens:
        if CHORD_TOKEN_RE.fullmatch(tok):
            chord_matches += 1
    # Check if line has rhythmic markers like '- - -' or '｜' or '|'
    has_grid_markers = ('｜' in stripped or '|' in stripped or '- -' in stripped)
    if has_grid_markers and chord_matches >= 1:
        return True
    if chinese_chars > 0 and not has_grid_markers:
        return False
    # If majority of tokens are valid chords (and not a long English sentence)
    if len(tokens) <= 16 and chord_matches / len(tokens) >= 0.6:
        # Make sure it's not an English title like "To The One Who Is Seated On The Throne"
        english_words = [t for t in tokens if len(t) > 1 and t.isalpha() and not re.match(r'^[A-G](m|maj|min|dim|aug|sus|add)$', t)]
        if len(english_words) >= 2:
            return False
        return True
    return False


def detect_key_from_lines(lines: List[str]) -> str:
    """
    Auto-detect original key from metadata line (e.g. 'Upper Room/活水/D' or 'Key: D')
    or fallback to analyzing first/last chords.
    """
    for line in lines[:8]:
        s = line.strip()
        # Match patterns like 'Upper Room/活水/D' or 'Key: D' or '調性：D'
        m = re.search(r'(?:/|Key\s*[:：]\s*|調性\s*[:：]\s*)([A-G][#b♯♭]?m?)\s*$', s, re.IGNORECASE)
        if m:
            k = m.group(1)
            return k[0].upper() + k[1:]
    # Fallback: collect all chords
    chords = []
    for line in lines:
        if is_chord_line(line):
            for m in CHORD_TOKEN_RE.finditer(line):
                chords.append(m.group(1) + ('m' if m.group(2).startswith('m') and not m.group(2).startswith('maj') else ''))
    if chords:
        return chords[-1].rstrip('m') if not chords[-1].endswith('m') else chords[0]
    return 'D'


def transpose_chord_token(
    root: str,
    quality: str,
    bass: Optional[str],
    src_semitone: int,
    dst_semitone: int,
    use_flats: bool,
    mode: str = 'chord'  # 'chord', 'number', 'dual'
) -> str:
    root_semi = NOTE_TO_SEMITONE.get(root)
    if root_semi is None:
        return f"{root}{quality}" + (f"/{bass}" if bass else "")

    new_root_semi = (root_semi - src_semitone + dst_semitone) % 12
    new_root_note = semitone_to_note(new_root_semi, use_flats)
    new_root_num = semitone_to_number(new_root_semi, dst_semitone, use_flats=True)

    new_bass_note = None
    new_bass_num = None
    if bass:
        bass_semi = NOTE_TO_SEMITONE.get(bass)
        if bass_semi is not None:
            new_bass_semi = (bass_semi - src_semitone + dst_semitone) % 12
            new_bass_note = semitone_to_note(new_bass_semi, use_flats)
            new_bass_num = semitone_to_number(new_bass_semi, dst_semitone, use_flats=True)
        else:
            new_bass_note = bass
            new_bass_num = bass

    chord_str = f"{new_root_note}{quality}" + (f"/{new_bass_note}" if new_bass_note else "")
    num_str = f"{new_root_num}{quality}" + (f"/{new_bass_num}" if new_bass_num else "")

    if mode == 'number':
        return num_str
    elif mode == 'dual':
        return f"{chord_str}({num_str})"
    return chord_str


INLINE_BRACKET_CHORD_RE = re.compile(
    r'\[([A-G][#b♯♭]?)((?:maj|min|m|M|dim|aug|sus|add|no|[0-9]|[\+\^\(\)])*)(?:/([A-G][#b♯♭]?))?\]'
)


def transpose_line(
    line: str,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> str:
    src_semi = NOTE_TO_SEMITONE.get(src_key.rstrip('m'), 2)
    dst_semi = NOTE_TO_SEMITONE.get(dst_key.rstrip('m'), src_semi)
    use_flats = should_use_flats(dst_key, accidental_pref)

    # Check if this is a metadata line ending with '/<src_key>' like 'Upper Room/活水/D'
    meta_match = re.search(r'(/|Key\s*[:：]\s*|調性\s*[:：]\s*)([A-G][#b♯♭]?m?)(\s*)$', line)
    if meta_match and not is_chord_line(line):
        prefix_end = meta_match.start(2)
        suffix_start = meta_match.end(2)
        label = f"{dst_key} (級數)" if mode == 'number' else dst_key
        return line[:prefix_end] + label + line[suffix_start:]

    # Check if line has inline bracketed chords like [G]我想認識[Bm]祢
    if INLINE_BRACKET_CHORD_RE.search(line):
        def bracket_repl(m: re.Match) -> str:
            root, quality, bass = m.group(1), m.group(2) or '', m.group(3)
            transposed = transpose_chord_token(root, quality, bass, src_semi, dst_semi, use_flats, mode)
            return f"[{transposed}]"
        return INLINE_BRACKET_CHORD_RE.sub(bracket_repl, line)

    if not is_chord_line(line):
        return line

    def repl(m: re.Match) -> str:
        root, quality, bass = m.group(1), m.group(2) or '', m.group(3)
        return transpose_chord_token(root, quality, bass, src_semi, dst_semi, use_flats, mode)

    matches = list(CHORD_TOKEN_RE.finditer(line))
    if not matches:
        return line

    result = []
    cursor = 0
    debt = 0  # positive means previous replacement was longer by `debt` chars

    for m in matches:
        gap = line[cursor:m.start()]
        if debt > 0:
            m_spaces = re.search(r' {2,}', gap)
            if m_spaces:
                s_start, s_end = m_spaces.span()
                consume = min(debt, s_end - s_start - 1)
                gap = gap[:s_start] + (' ' * (s_end - s_start - consume)) + gap[s_end:]
                debt -= consume
            if debt > 0:
                trail = len(gap) - len(gap.rstrip(' '))
                consume = min(debt, max(0, trail - 1))
                if consume > 0:
                    gap = gap[:-consume]
                    debt -= consume
        elif debt < 0:
            pad = -debt
            pipe_pos = gap.find('｜') if '｜' in gap else gap.find('|')
            if pipe_pos >= 0:
                gap = gap[:pipe_pos] + (' ' * pad) + gap[pipe_pos:]
            else:
                gap = gap + (' ' * pad)
            debt = 0

        result.append(gap)
        old_tok = m.group(0)
        new_tok = repl(m)
        result.append(new_tok)
        debt += len(new_tok) - len(old_tok)
        cursor = m.end()

    tail = line[cursor:]
    if debt > 0:
        m_spaces = re.search(r' {2,}', tail)
        if m_spaces:
            s_start, s_end = m_spaces.span()
            consume = min(debt, s_end - s_start - 1)
            tail = tail[:s_start] + (' ' * (s_end - s_start - consume)) + tail[s_end:]
    elif debt < 0:
        tail = tail + (' ' * (-debt))
    result.append(tail)
    return ''.join(result)


def get_capo_suggestions(target_key: str) -> List[Dict[str, str]]:
    """Return practical acoustic guitar Capo positions for the target sounding key."""
    target_semi = NOTE_TO_SEMITONE.get(target_key.rstrip('m'), 0)
    easy_guitar_keys = [
        ('C', 0), ('G', 7), ('D', 2), ('A', 9), ('E', 4), ('F', 5)
    ]
    suggestions = []
    for play_key, play_semi in easy_guitar_keys:
        capo = (target_semi - play_semi) % 12
        if 0 <= capo <= 9:
            suggestions.append({
                'capo': capo,
                'play_key': play_key,
                'label': f"不夾 Capo (直接彈 {play_key} 調)" if capo == 0 else f"Capo {capo} (彈 {play_key} 調指法)"
            })
    suggestions.sort(key=lambda x: (0 if x['play_key'] in ('C', 'G', 'D') and x['capo'] <= 5 else 1, x['capo']))
    return suggestions


def transpose_docx(
    input_path: str,
    output_path: str,
    target_key: str,
    source_key: Optional[str] = None,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> Dict[str, str]:
    """
    Losslessly transpose a .docx chord sheet while keeping all XML styling,
    embedded fonts, shading, alignments, and runs intact.
    """
    with zipfile.ZipFile(input_path, 'r') as zin:
        xml_bytes = zin.read('word/document.xml')
        namelist = zin.namelist()
        other_files = {name: zin.read(name) for name in namelist if name != 'word/document.xml'}

    # Parse XML while preserving original namespace prefixes
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tree = ET.fromstring(xml_bytes)

    # First pass: extract paragraph text to detect source key if not provided
    paragraphs = tree.findall('.//w:p', ns)
    para_texts = []
    for p in paragraphs:
        t_list = [t.text for t in p.findall('.//w:t', ns) if t.text]
        para_texts.append(''.join(t_list))

    detected_src = source_key or detect_key_from_lines(para_texts)

    # Second pass: update paragraph text runs in-place
    for p, full_text in zip(paragraphs, para_texts):
        if not full_text.strip():
            continue
        new_full_text = transpose_line(full_text, detected_src, target_key, accidental_pref, mode)
        if new_full_text == full_text:
            continue

        # Check if metadata line with multiple runs (e.g. Run 0: "Upper Room", Run 1: "/活水/D")
        t_elems = p.findall('.//w:t', ns)
        if not t_elems:
            continue

        if is_chord_line(full_text):
            # If chord line has a single <w:t> or multiple <w:t>, put transposed text in first <w:t>
            # or transpose per <w:t> if each <w:t> is self-contained
            all_self_contained = True
            for t_el in t_elems:
                if t_el.text and ('/' in t_el.text and not CHORD_TOKEN_RE.search(t_el.text)):
                    all_self_contained = False
            if len(t_elems) == 1 or not all_self_contained:
                t_elems[0].text = new_full_text
                for extra_t in t_elems[1:]:
                    extra_t.text = ''
            else:
                for t_el in t_elems:
                    if t_el.text:
                        t_el.text = transpose_line(t_el.text, detected_src, target_key, accidental_pref, mode)
        else:
            # Metadata line (e.g. "/活水/D"): transpose the last non-empty run so run-level fonts stay intact
            replaced = False
            for t_el in reversed(t_elems):
                if t_el.text and t_el.text.strip():
                    updated_run = transpose_line(t_el.text, detected_src, target_key, accidental_pref, mode)
                    if updated_run != t_el.text:
                        t_el.text = updated_run
                        replaced = True
                        break
            if not replaced:
                t_elems[0].text = new_full_text
                for extra_t in t_elems[1:]:
                    extra_t.text = ''

    # Wait: ElementTree rewrites namespace prefixes unless we preserve them or operate on the XML string/DOM carefully.
    # Let's make sure we preserve the exact XML declaration and namespaces of word/document.xml!
    new_xml_bytes = transpose_docx_xml_bytes(xml_bytes, detected_src, target_key, accidental_pref, mode)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in other_files.items():
            zout.writestr(name, data)
        zout.writestr('word/document.xml', new_xml_bytes)

    return {
        'source_key': detected_src,
        'target_key': target_key,
        'mode': mode,
        'output_path': output_path
    }


def transpose_docx_xml_bytes(
    xml_bytes: bytes,
    src_key: str,
    dst_key: str,
    accidental_pref: str = 'auto',
    mode: str = 'chord'
) -> bytes:
    """
    Operate directly on raw word/document.xml string so 100% of XML namespaces,
    attributes, declarations, and formatting tags remain byte-for-byte untouched!
    """
    xml_str = xml_bytes.decode('utf-8')

    p_pattern = re.compile(r'(<w:p\b[^>]*>)(.*?)(</w:p>)', re.DOTALL)
    t_pattern = re.compile(r'(<w:t(?:\s+[^>]*)?>)(.*?)(</w:t>)', re.DOTALL)

    def unescape_xml(s: str) -> str:
        return (s.replace('&amp;', '&')
                 .replace('&lt;', '<')
                 .replace('&gt;', '>')
                 .replace('&quot;', '"')
                 .replace('&apos;', "'"))

    def escape_xml(s: str) -> str:
        return (s.replace('&', '&amp;')
                 .replace('<', '&lt;')
                 .replace('>', '&gt;'))

    def ensure_xml_space_preserve(tag: str) -> str:
        if 'xml:space="preserve"' in tag or "xml:space='preserve'" in tag:
            return tag
        return re.sub(r'<w:t\b([^>]*)>', r'<w:t xml:space="preserve"\1>', tag)

    def process_paragraph(p_match: re.Match) -> str:
        p_open, p_body, p_close = p_match.group(1), p_match.group(2), p_match.group(3)
        t_matches = list(t_pattern.finditer(p_body))
        if not t_matches:
            return p_match.group(0)

        full_text = ''.join(unescape_xml(m.group(2)) for m in t_matches)
        if not full_text.strip():
            return p_match.group(0)

        new_full_text = transpose_line(full_text, src_key, dst_key, accidental_pref, mode)
        if new_full_text == full_text:
            return p_match.group(0)

        # If single <w:t> in paragraph, replace directly
        if len(t_matches) == 1:
            m0 = t_matches[0]
            new_t = ensure_xml_space_preserve(m0.group(1)) + escape_xml(new_full_text) + m0.group(3)
            return p_open + p_body[:m0.start()] + new_t + p_body[m0.end():] + p_close

        # If metadata line (not chord line), try updating the specific <w:t> run that has the key suffix
        if not is_chord_line(full_text):
            for idx in range(len(t_matches) - 1, -1, -1):
                m_run = t_matches[idx]
                run_txt = unescape_xml(m_run.group(2))
                new_run_txt = transpose_line(run_txt, src_key, dst_key, accidental_pref, mode)
                if new_run_txt != run_txt:
                    new_t = ensure_xml_space_preserve(m_run.group(1)) + escape_xml(new_run_txt) + m_run.group(3)
                    return p_open + p_body[:m_run.start()] + new_t + p_body[m_run.end():] + p_close

        # Otherwise place full transposed text in the first <w:t> and empty subsequent <w:t> in this paragraph
        pieces = []
        last_idx = 0
        for idx, m_run in enumerate(t_matches):
            pieces.append(p_body[last_idx:m_run.start()])
            if idx == 0:
                pieces.append(ensure_xml_space_preserve(m_run.group(1)) + escape_xml(new_full_text) + m_run.group(3))
            else:
                pieces.append(m_run.group(1) + m_run.group(3))
            last_idx = m_run.end()
        pieces.append(p_body[last_idx:])
        return p_open + ''.join(pieces) + p_close

    updated_xml = p_pattern.sub(process_paragraph, xml_str)
    return updated_xml.encode('utf-8')


def main():
    parser = argparse.ArgumentParser(description='Lossless .docx & Text Worship Chord Sheet Transposer')
    parser.add_argument('input', help='Input .docx or .txt file path')
    parser.add_argument('--to', '-t', required=True, help='Target key (e.g. C, Db, D, Eb, E, F, F#, G, Ab, A, Bb, B)')
    parser.add_argument('--from-key', '-f', default=None, help='Source key (auto-detected if omitted)')
    parser.add_argument('--accidental', '-a', choices=['auto', 'sharp', 'flat'], default='auto',
                        help='Accidental preference: auto (by key), sharp (#), or flat (b)')
    parser.add_argument('--mode', '-m', choices=['chord', 'number', 'dual'], default='chord',
                        help='Output mode: chord (standard), number (Nashville 1-7), or dual (Chord + Number)')
    parser.add_argument('--output', '-o', default=None, help='Output file path')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(args.input)[1].lower()
    if ext == '.docx':
        out_path = args.output
        if not out_path:
            base = os.path.splitext(os.path.basename(args.input))[0]
            # Replace trailing key in filename if present (e.g. "坐在寶座上的那位 D" -> "坐在寶座上的那位 C")
            suffix = f"{args.to}_級數" if args.mode == 'number' else args.to
            new_base = re.sub(r'\s+[A-G][#b♯♭]?m?$', f' {suffix}', base)
            if new_base == base:
                new_base = f"{base} ({suffix})"
            out_path = os.path.join(os.path.dirname(args.input) or '.', f"{new_base}.docx")

        res = transpose_docx(args.input, out_path, args.to, args.from_key, args.accidental, args.mode)
        print(f"✅ Transposed .docx: {res['source_key']} -> {res['target_key']} (mode={res['mode']})")
        print(f"📄 Saved to: {res['output_path']}")
        capos = get_capo_suggestions(args.to)
        if capos:
            print("🎸 Guitar Capo Suggestions:", " | ".join(c['label'] for c in capos[:3]))
    else:
        with open(args.input, 'r', encoding='utf-8') as f:
            lines = f.read().splitlines()
        src_key = args.from_key or detect_key_from_lines(lines)
        out_lines = [transpose_line(line, src_key, args.to, args.accidental, args.mode) for line in lines]
        out_text = '\n'.join(out_lines)
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(out_text + '\n')
            print(f"✅ Saved transposed text ({src_key} -> {args.to}) to {args.output}")
        else:
            print(out_text)


if __name__ == '__main__':
    main()
