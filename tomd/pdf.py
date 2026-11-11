"""PDF 변환 (텍스트가 들어 있는 PDF 전용, OCR 없음)."""
from . import table


def convert_pdf(path):
    import pdfplumber

    out = []
    has_text = False
    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for number, page in enumerate(pdf.pages, 1):
            try:
                tables = page.find_tables()
            except Exception:
                tables = []
            boxes = [t.bbox for t in tables]

            def outside(obj, boxes=boxes):
                if "x0" not in obj or "top" not in obj:
                    return True
                cx = (obj["x0"] + obj["x1"]) / 2
                cy = (obj["top"] + obj["bottom"]) / 2
                return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in boxes)

            text_page = page.filter(outside) if boxes else page
            items = []
            for line in text_page.extract_text_lines():
                items.append((line["top"], "text", line))
            for t in tables:
                items.append((t.bbox[1], "table", t))
            items.sort(key=lambda x: x[0])

            blocks, para, last_bottom = [], [], None
            for _, kind, obj in items:
                if kind == "table":
                    if para:
                        blocks.append("\n".join(para))
                        para = []
                    grid = [[(c or "") for c in row] for row in obj.extract()]
                    rendered = table.render(table.from_grid(grid))
                    if rendered:
                        blocks.append(rendered)
                    last_bottom = None
                    continue
                text = obj["text"].strip()
                if not text:
                    continue
                height = max(obj["bottom"] - obj["top"], 1)
                if para and last_bottom is not None and obj["top"] - last_bottom > height * 0.8:
                    blocks.append("\n".join(para))
                    para = []
                para.append(text)
                last_bottom = obj["bottom"]
            if para:
                blocks.append("\n".join(para))

            if blocks:
                has_text = True
                if total > 1:
                    out.append(f"<!-- {number}쪽 -->")
                out.extend(blocks)
    if not has_text:
        return (
            "> 이 PDF에서는 글자를 찾지 못했습니다. 스캔한 이미지로 된 PDF로 보이며, "
            "글자 인식(OCR)은 지원하지 않습니다."
        )
    return "\n\n".join(out)
