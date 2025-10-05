# extract_structure.py
import json
import re
from docx import Document
import argparse

def _heading_level(p):
    name = (p.style.name or "").lower()
    m = re.search(r'(heading|überschrift|عنوان)\s*(\d+)$', name)
    if m: return int(m.group(2))
    return None

def paragraph_to_json(p):
    return {
        "text": p.text.strip(),
        "style": p.style.name if p.style else "None"
    }

def build_tree(doc):
    roots, stack = [], []
    for p in doc.paragraphs:
        lvl = _heading_level(p)
        if lvl:
            node = {
                "title": p.text.strip(),
                "level": lvl,
                "style": p.style.name if p.style else "None",
                "children": []
            }
            while stack and stack[-1]["level"] >= lvl:
                stack.pop()
            if stack:
                stack[-1]["children"].append(node)
            else:
                roots.append(node)
            stack.append(node)
    return roots

def main():
    ap = argparse.ArgumentParser(description="Extract the structure and styles of a Word document into a JSON file.")
    ap.add_argument("input_docx", help="The Word document to analyze.")
    ap.add_argument("output_json", help="The path to the output JSON file.")
    args = ap.parse_args()

    doc = Document(args.input_docx)
    structure = {
        "styles": [s.name for s in doc.styles],
        "chapters": build_tree(doc)
    }

    with open(args.output_json, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    print(f"Successfully extracted structure to {args.output_json}")

if __name__ == "__main__":
    main()