#!/usr/bin/env python3
"""설교문 성경 인용 추출기 (CLI 버전)

.docx 설교문에서 성경 인용/언급 구절을 등장 순서대로 찾아 정리합니다.

사용법:
    pip install python-docx
    python bible_quotes.py 설교문.docx
    python bible_quotes.py 설교문.docx -o 결과.txt
"""

import argparse
import os
import re
import sys

try:
    import docx
except ImportError:
    sys.exit("python-docx가 설치되어 있지 않습니다. 먼저 'pip install python-docx'를 실행하세요.")

U1 = "\x01"  # 밑줄(직접 인용) 시작 표시
U2 = "\x02"  # 밑줄(직접 인용) 끝 표시

# ---------------- 성경책 사전 ----------------
BOOKS = [
    ("창세기", ["창세기", "창"]), ("출애굽기", ["출애굽기", "출"]), ("레위기", ["레위기", "레"]),
    ("민수기", ["민수기", "민"]), ("신명기", ["신명기", "신"]), ("여호수아", ["여호수아", "수"]),
    ("사사기", ["사사기", "삿"]), ("룻기", ["룻기", "룻"]), ("사무엘상", ["사무엘상", "삼상"]),
    ("사무엘하", ["사무엘하", "삼하"]), ("열왕기상", ["열왕기상", "왕상"]), ("열왕기하", ["열왕기하", "왕하"]),
    ("역대상", ["역대상", "대상"]), ("역대하", ["역대하", "대하"]), ("에스라", ["에스라", "스"]),
    ("느헤미야", ["느헤미야", "느"]), ("에스더", ["에스더", "에"]), ("욥기", ["욥기", "욥"]),
    ("시편", ["시편", "시"]), ("잠언", ["잠언", "잠"]), ("전도서", ["전도서", "전"]),
    ("아가", ["아가"]), ("이사야", ["이사야", "사"]), ("예레미야", ["예레미야", "렘"]),
    ("예레미야애가", ["예레미야애가", "애가"]), ("에스겔", ["에스겔", "겔"]), ("다니엘", ["다니엘", "단"]),
    ("호세아", ["호세아", "호"]), ("요엘", ["요엘", "욜"]), ("아모스", ["아모스", "암"]),
    ("오바댜", ["오바댜", "옵"]), ("요나", ["요나", "욘"]), ("미가", ["미가"]),
    ("나훔", ["나훔"]), ("하박국", ["하박국", "합"]), ("스바냐", ["스바냐", "습"]),
    ("학개", ["학개", "학"]), ("스가랴", ["스가랴", "슥"]), ("말라기", ["말라기", "말"]),
    ("마태복음", ["마태복음", "마태", "마"]), ("마가복음", ["마가복음", "마가", "막"]),
    ("누가복음", ["누가복음", "누가", "눅"]), ("요한복음", ["요한복음", "요"]),
    ("사도행전", ["사도행전", "행"]), ("로마서", ["로마서", "롬"]),
    ("고린도전서", ["고린도전서", "고전"]), ("고린도후서", ["고린도후서", "고후"]),
    ("갈라디아서", ["갈라디아서", "갈"]), ("에베소서", ["에베소서", "엡"]),
    ("빌립보서", ["빌립보서", "빌"]), ("골로새서", ["골로새서", "골"]),
    ("데살로니가전서", ["데살로니가전서", "살전"]), ("데살로니가후서", ["데살로니가후서", "살후"]),
    ("디모데전서", ["디모데전서", "딤전"]), ("디모데후서", ["디모데후서", "딤후"]),
    ("디도서", ["디도서", "딛"]), ("빌레몬서", ["빌레몬서", "몬"]), ("히브리서", ["히브리서", "히"]),
    ("야고보서", ["야고보서", "약"]), ("베드로전서", ["베드로전서", "벧전"]),
    ("베드로후서", ["베드로후서", "벧후"]), ("요한일서", ["요한일서", "요일"]),
    ("요한이서", ["요한이서", "요이"]), ("요한삼서", ["요한삼서", "요삼"]),
    ("유다서", ["유다서", "유"]), ("요한계시록", ["요한계시록", "계시록", "계"]),
]
ALIAS_MAP = {}
ALL_ALIASES = []
for _name, _aliases in BOOKS:
    for _a in _aliases:
        ALIAS_MAP[_a] = _name
        ALL_ALIASES.append(_a)
ALL_ALIASES.sort(key=len, reverse=True)
BOOK_ALT = "|".join(re.escape(a) for a in ALL_ALIASES)

RE_FULL = re.compile(
    rf"({BOOK_ALT})\s*(\d{{1,3}})\s*장(?:\s*(\d{{1,3}})\s*(?:[~\-–,]\s*(\d{{1,3}})\s*)?절)?"
)
RE_COLON = re.compile(
    rf"({BOOK_ALT})\s*(\d{{1,3}})\s*[:：]\s*(\d{{1,3}})(?:\s*[~\-–]\s*(\d{{1,3}}))?"
)
RE_HEADER = re.compile(
    rf"본문\s*[:：]?\s*({BOOK_ALT})\s*(\d{{1,3}})\s*장\s*(\d{{1,3}})(?:\s*[~\-–,]\s*(\d{{1,3}}))?\s*절"
)
RE_VERSE_ONLY = re.compile(r"(\d{1,3})\s*(?:[~\-–,]\s*(\d{1,3}))?\s*절")


def alias_to_canonical(alias):
    return ALIAS_MAP.get(alias.strip(), alias.strip())


# ---------------- 텍스트 유틸 ----------------
def strip_markers(s):
    return s.replace(U1, "").replace(U2, "")


def clean_text(s):
    s = strip_markers(s)
    s = re.sub(r"[【】「」『』]", "", s)
    s = re.sub(r'^[\s"\'“”‘’]+|[\s"\'“”‘’]+$', "", s)
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s*\n\s*", " ", s)
    return s.strip()


def extract_underline_segments(para):
    segments = []
    for m in re.finditer(re.escape(U1) + r"([^" + re.escape(U2) + r"]*)" + re.escape(U2), para):
        text = m.group(1).strip()
        if text:
            segments.append({"index": m.start(), "text": text})
    return segments


def find_quote_mark_span(para):
    plain = strip_markers(para)
    patterns = [
        r"「([^」]{4,400})」", r"『([^』]{4,400})』", r"【([^】]{4,400})】",
        r'"([^"]{4,400})"', r"“([^”]{4,400})”", r"‘([^’]{4,400})’",
    ]
    for pat in patterns:
        m = re.search(pat, plain)
        if m:
            return m.group(1)
    return None


def split_verses(raw_content, v_start, v_end):
    content = strip_markers(raw_content)
    nums = list(range(v_start, v_end + 1))
    if len(nums) < 2:
        return None
    positions = []
    search_from = 0
    for n in nums:
        pat = re.compile(rf"(?:^|[^0-9])({n})(?=[^0-9]|$)")
        m = pat.search(content, search_from)
        if not m:
            return None
        num_idx = m.start(1)
        positions.append(num_idx)
        search_from = num_idx + len(str(n))
    parts = []
    for i, pos in enumerate(positions):
        start = pos + len(str(nums[i]))
        end = positions[i + 1] if i + 1 < len(positions) else len(content)
        text = content[start:end].strip()
        if not text:
            return None
        parts.append({"num": nums[i], "text": text})
    return parts


def extract_sentence_around(text, idx):
    enders = ".?!\n"
    start = 0
    for i in range(idx, -1, -1):
        if text[i] in enders:
            start = i + 1
            break
    end = len(text)
    for i in range(idx, len(text)):
        if text[i] in enders:
            end = i + 1
            break
    return text[start:end].strip()


def looks_like_new_citation(para):
    plain = strip_markers(para)
    return bool(RE_FULL.search(plain)) or bool(RE_COLON.search(plain))


# ---------------- docx -> 문단 (밑줄은 마커로 표시) ----------------
def docx_to_paragraphs(path):
    document = docx.Document(path)
    paragraphs = []
    for p in document.paragraphs:
        text = ""
        for run in p.runs:
            if run.underline:
                text += U1 + run.text + U2
            else:
                text += run.text
        text = text.strip()
        if strip_markers(text).strip():
            paragraphs.append(text)
    return paragraphs


# ---------------- 본문 기준(헤더) 감지 ----------------
def detect_header_context(paragraphs):
    for para in paragraphs[:20]:
        plain = strip_markers(para)
        m = RE_HEADER.search(plain)
        if m:
            return {"book": alias_to_canonical(m.group(1)), "chapter": m.group(2)}
    return {"book": "", "chapter": ""}


# ---------------- 인용 내용 추출 ----------------
def extract_content(paragraphs, p_idx, m_end, m_index, v_start, v_end):
    para = paragraphs[p_idx]
    expected_count = v_end - v_start + 1
    segments = extract_underline_segments(para)

    if segments:
        if expected_count > 1 and len(segments) == expected_count:
            lines = [f"{v_start + i} {clean_text(s['text'])}" for i, s in enumerate(segments)]
            return "\n".join(lines), False
        joined = " ".join(s["text"] for s in segments)
        if expected_count > 1:
            split = split_verses(joined, v_start, v_end)
            if split:
                return "\n".join(f"{p['num']} {clean_text(p['text'])}" for p in split), False
        return clean_text(joined), False

    q_span = find_quote_mark_span(para)
    if q_span:
        if expected_count > 1:
            split = split_verses(q_span, v_start, v_end)
            if split:
                return "\n".join(f"{p['num']} {clean_text(p['text'])}" for p in split), False
        return clean_text(q_span), False

    if expected_count > 1:
        after = para[m_end:]
        split = split_verses(after, v_start, v_end)
        if split:
            return "\n".join(f"{p['num']} {clean_text(p['text'])}" for p in split), False

    plain_para = strip_markers(para).strip()
    if len(plain_para) <= 40:
        for j in range(p_idx + 1, min(p_idx + 3, len(paragraphs))):
            nxt = paragraphs[j]
            next_plain = strip_markers(nxt).strip()
            if not next_plain:
                continue
            if looks_like_new_citation(nxt):
                break
            segs2 = extract_underline_segments(nxt)
            if segs2:
                if expected_count > 1 and len(segs2) == expected_count:
                    lines = [f"{v_start + i} {clean_text(s['text'])}" for i, s in enumerate(segs2)]
                    return "\n".join(lines), False
                joined2 = " ".join(s["text"] for s in segs2)
                if expected_count > 1:
                    split2 = split_verses(joined2, v_start, v_end)
                    if split2:
                        return "\n".join(f"{p['num']} {clean_text(p['text'])}" for p in split2), False
                return clean_text(joined2), False
            if expected_count > 1:
                split3 = split_verses(next_plain, v_start, v_end)
                if split3:
                    return "\n".join(f"{p['num']} {clean_text(p['text'])}" for p in split3), False
            return clean_text(next_plain), False

    plain = strip_markers(para)
    sentence = extract_sentence_around(plain, m_index)
    return clean_text(sentence), True


# ---------------- 인용 목록 추출 ----------------
def extract_citations(paragraphs, ctx):
    items = []
    last_book = ctx.get("book", "")
    last_chapter = ctx.get("chapter", "")

    for p_idx, para in enumerate(paragraphs):
        matches = []
        for m in RE_FULL.finditer(para):
            matches.append({
                "index": m.start(), "end": m.end(), "has_verse": bool(m.group(3)),
                "book": m.group(1), "chapter": m.group(2),
                "v_start": m.group(3), "v_end": m.group(4) or m.group(3),
            })
        for m in RE_COLON.finditer(para):
            matches.append({
                "index": m.start(), "end": m.end(), "has_verse": True,
                "book": m.group(1), "chapter": m.group(2),
                "v_start": m.group(3), "v_end": m.group(4) or m.group(3),
            })
        matches.sort(key=lambda x: x["index"])

        cleaned = []
        last_end = -1
        for mm in matches:
            if mm["index"] >= last_end:
                cleaned.append(mm)
                last_end = mm["end"]

        for m in RE_VERSE_ONLY.finditer(para):
            overlaps = any(m.start() < mm["end"] and m.end() > mm["index"] for mm in cleaned)
            if not overlaps:
                cleaned.append({
                    "index": m.start(), "end": m.end(), "has_verse": True,
                    "book": None, "chapter": None,
                    "v_start": m.group(1), "v_end": m.group(2) or m.group(1),
                })
        cleaned.sort(key=lambda x: x["index"])

        for mm in cleaned:
            book = alias_to_canonical(mm["book"]) if mm["book"] else (last_book or "")
            chapter = mm["chapter"] or last_chapter or ""
            last_book = book
            last_chapter = chapter
            if not mm["has_verse"]:
                continue
            v_start = int(mm["v_start"])
            v_end = int(mm["v_end"]) if mm["v_end"] else v_start
            content, is_paraphrase = extract_content(paragraphs, p_idx, mm["end"], mm["index"], v_start, v_end)
            items.append({
                "book": book, "chapter": chapter, "v_start": v_start, "v_end": v_end,
                "content": content, "is_paraphrase": is_paraphrase,
            })
    return items


# ---------------- 출력 포맷 ----------------
def build_output_text(items):
    if not items:
        return ""
    blocks = []
    for i, it in enumerate(items, start=1):
        if it["v_end"] and it["v_end"] != it["v_start"]:
            verse_label = f"{it['v_start']} - {it['v_end']}절"
        else:
            verse_label = f"{it['v_start']}절"
        title = f"**{i}. {it['book']} {it['chapter']}장 {verse_label}**"
        if it["is_paraphrase"]:
            title += " (직접 인용은 아니고 본문 언급)"
        blocks.append(title + "\n" + it["content"].strip())
    return "\n\n".join(blocks)


def _clean_dropped_path(raw):
    """터미널에 파일을 끌어다 놓았을 때 붙는 장식을 제거합니다.

    PowerShell은 '& "경로"' 형태로 붙이고, cmd/일반 터미널은 따옴표만 붙입니다.
    """
    s = raw.strip()
    if s.startswith("&"):
        s = s[1:].strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
        s = s[1:-1]
    return s.strip()


def _finish(message, interactive, code=1):
    print(message, file=sys.stderr)
    if interactive:
        input("\nEnter를 누르면 창이 닫힙니다...")
    sys.exit(code)


def main():
    parser = argparse.ArgumentParser(description="설교문 .docx에서 성경 인용을 추출합니다.")
    parser.add_argument("docx_path", nargs="?", help="설교문 .docx 파일 경로 (생략하면 실행 중 입력받습니다)")
    parser.add_argument("-o", "--output", help="결과를 저장할 .txt 파일 경로 (생략 시 자동 저장 위치를 사용)")
    args = parser.parse_args()

    interactive = args.docx_path is None
    docx_path = args.docx_path
    if interactive:
        raw = input("설교문 .docx 파일 경로를 입력하거나, 파일을 이 창에 끌어다 놓은 뒤 Enter를 누르세요: ")
        docx_path = _clean_dropped_path(raw)

    if not docx_path or not os.path.isfile(docx_path):
        _finish(f"파일을 찾을 수 없습니다: {docx_path}", interactive)

    paragraphs = docx_to_paragraphs(docx_path)
    if not paragraphs:
        _finish("문서에서 텍스트를 찾을 수 없습니다.", interactive)

    ctx = detect_header_context(paragraphs)
    if ctx["book"]:
        print(f"[본문 기준 자동 인식] {ctx['book']} {ctx['chapter']}장\n", file=sys.stderr)
    else:
        print("[알림] 본문 기준(예: '본문: 사도행전 18장 9~10절')을 찾지 못했습니다. "
              "책 이름이 생략된 인용은 정확히 추론되지 않을 수 있습니다.\n", file=sys.stderr)

    items = extract_citations(paragraphs, ctx)
    if not items:
        _finish("인용을 찾지 못했습니다.", interactive)

    output = build_output_text(items)
    print(output)

    out_path = args.output
    if not out_path and interactive:
        base, _ = os.path.splitext(docx_path)
        out_path = base + "_인용정리.txt"
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"\n[저장됨] {out_path}", file=sys.stderr)

    if interactive:
        input("\n완료했습니다. Enter를 누르면 창이 닫힙니다...")


if __name__ == "__main__":
    main()
