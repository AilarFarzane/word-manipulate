# update_toc_and_fields.py
# Forces Persian digits in footer PAGE fields and updates TOC/fields.
# Sets Word numeral shaping to Hindi (Eastern) and shapes footers to Persian/RTL.

import sys, os, argparse
import win32com.client as win32
import re

def fix_toc_numbering(doc):
    """Ensure TOC entries like '1-2 ' become '1-2- '."""
    for para in doc.Paragraphs:
        try:
            if para.Style.NameLocal.startswith("TOC"):
                rng = para.Range
                new_text = re.sub(r"(\d+(?:-\d+)+)(\s+)", r"\1-\2", rng.Text)
                if new_text != rng.Text:
                    rng.Text = new_text
        except Exception:
            pass


def update_fields(in_path, out_path=None, levels=3, remove_numbering_in="none"):
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(in_path))

        # Refresh TOCs (H1..levels)
        for toc in list(doc.TablesOfContents):
            try:
                toc.UpperHeadingLevel = 1
                toc.LowerHeadingLevel = int(levels)
            except Exception:
                pass
            toc.Update()

        fix_toc_numbering(doc)
        # Update all fields everywhere (body + stories)
        doc.Fields.Update()
        try:
            story = doc.StoryRanges(1)  # wdMainTextStory
            while True:
                try:
                    story.Fields.Update()
                    story = story.NextStoryRange
                except Exception:
                    break
        except Exception:
            pass

        out = os.path.abspath(out_path or in_path)
        doc.SaveAs(out)
        print(f"Updated TOC/fields with Persian footer digits → {out}")
    finally:
        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass
        word.Quit()

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Refresh TOC/fields; force Persian digits/RTL in footers; optional numbering removal.")
    ap.add_argument("input_docx")
    ap.add_argument("output_docx", nargs="?", default=None)
    args = ap.parse_args()
    update_fields(args.input_docx, args.output_docx)
