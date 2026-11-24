
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import io, re

st.set_page_config(page_title="Smart Teacher PPT Maker — Free", page_icon="📚", layout="centered")

NAVY=RGBColor(22,45,82)
BLUE=RGBColor(40,100,170)
TEAL=RGBColor(0,135,145)
BG=RGBColor(247,249,252)
DARK=RGBColor(38,48,62)
MUTED=RGBColor(100,112,128)
WHITE=RGBColor(255,255,255)

st.title("📚 Smart Teacher PPT Maker — Free")
st.caption("Class 6–10 | Mathematics | No API key required")
st.info("Yeh version bina OpenAI API key ke PPT file banata hai. Chapter ka lesson content ChatGPT se banwa kar neeche paste karein; yeh app usse PowerPoint mein badal dega.")

with st.form("ppt_form"):
    c1,c2=st.columns(2)
    with c1:
        grade=st.selectbox("Class", ["6","7","8","9","10"], index=4)
        board=st.selectbox("Board / Curriculum", ["CBSE / NCERT", "State Board", "Other"])
        language=st.selectbox("Language", ["Hindi + English (Hinglish)", "Hindi", "English"])
    with c2:
        slide_count=st.select_slider("Target number of slides", options=[10,15,20,25,30], value=25)
        theme=st.selectbox("PPT design", ["Clean classroom", "Modern colourful", "Minimal / exam-focused"])
        notes_on=st.checkbox("Add teacher notes where supplied", value=True)
    chapter=st.text_input("Chapter / Topic", placeholder="e.g. Circles / त्रिभुज / वास्तविक संख्याएँ")
    instructions=st.text_area("Special instructions (optional)", placeholder="e.g. Include theorem statements, full proofs, diagrams to draw, solved examples and board questions.")
    prompt = f"""Create exactly {slide_count} slides for Class {grade} Mathematics, {board}, chapter: {chapter or '[chapter]'}.
Language: {language}. Special instructions: {instructions or 'Include concepts, definitions, worked examples, practice and recap.'}
Return slide content in this exact plain-text format (no markdown table):
[SLIDE]
TITLE: Slide title
BULLET: First short point
BULLET: Second short point
EXAMPLE: Optional example or leave blank
TEACHER NOTE: Optional teacher explanation
[SLIDE]
TITLE: Next slide title
BULLET: ...
Use accurate mathematics. Include theorem statement and proof steps when relevant. Do not claim verbatim NCERT wording unless I supplied the source. Keep slides concise."""
    st.text_area("Copy this prompt into ChatGPT to prepare the lesson content", value=prompt, height=190)
    lesson_text=st.text_area("Paste the slide content from ChatGPT here", height=360,
        placeholder="[SLIDE]\nTITLE: Introduction to Circles\nBULLET: A circle is...\nBULLET: Centre is...\nEXAMPLE: ...\nTEACHER NOTE: Explain with a compass.\n[SLIDE]\nTITLE: ...")
    submitted=st.form_submit_button("📥 Create PowerPoint", use_container_width=True)

def parse_slides(raw):
    chunks=re.split(r'\[SLIDE\]', raw, flags=re.I)
    slides=[]
    for chunk in chunks:
        chunk=chunk.strip()
        if not chunk: continue
        item={"title":"","bullets":[],"example":"","note":""}
        for line in chunk.splitlines():
            line=line.strip()
            if not line: continue
            if line.upper().startswith("TITLE:"):
                item["title"]=line.split(":",1)[1].strip()
            elif line.upper().startswith("BULLET:"):
                item["bullets"].append(line.split(":",1)[1].strip())
            elif line.upper().startswith("EXAMPLE:"):
                item["example"]=line.split(":",1)[1].strip()
            elif line.upper().startswith("TEACHER NOTE:"):
                item["note"]=line.split(":",1)[1].strip()
            elif not item["title"]:
                item["title"]=line
            else:
                item["bullets"].append(line)
        if item["title"] or item["bullets"]:
            slides.append(item)
    return slides

def make_ppt(slides, grade, chapter, board, notes_on):
    prs=Presentation()
    prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
    for i,item in enumerate(slides,1):
        s=prs.slides.add_slide(prs.slide_layouts[6])
        bg=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb=BG; bg.line.fill.background()
        top=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,Inches(.16))
        top.fill.solid(); top.fill.fore_color.rgb=TEAL; top.line.fill.background()
        title=s.shapes.add_textbox(Inches(.65),Inches(.42),Inches(12),Inches(.75))
        p=title.text_frame.paragraphs[0]; p.text=item["title"] or f"Slide {i}"
        p.font.name="Aptos Display"; p.font.size=Pt(29); p.font.bold=True; p.font.color.rgb=NAVY
        body=s.shapes.add_textbox(Inches(.85),Inches(1.5),Inches(11.7),Inches(4.2))
        tf=body.text_frame; tf.word_wrap=True
        bullets=item["bullets"] or ["(Add lesson point here)"]
        for j,b in enumerate(bullets):
            p=tf.paragraphs[0] if j==0 else tf.add_paragraph()
            p.text="•  "+b
            p.font.name="Aptos"; p.font.size=Pt(21 if len(bullets)<=4 else 18); p.font.color.rgb=DARK
            p.space_after=Pt(13)
        ex=item["example"]
        if ex:
            card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(.85), Inches(5.85), Inches(11.65), Inches(.72))
            card.fill.solid(); card.fill.fore_color.rgb=RGBColor(230,242,246)
            card.line.color.rgb=RGBColor(215,228,235)
            tf=card.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE
            p=tf.paragraphs[0]; p.text="Example: "+ex; p.font.name="Aptos"; p.font.size=Pt(15); p.font.color.rgb=TEAL
        foot=s.shapes.add_textbox(Inches(.65), Inches(7.08), Inches(12), Inches(.22))
        p=foot.text_frame.paragraphs[0]
        p.text=f"CLASS {grade} • MATHEMATICS • {chapter.upper()}                                      {i:02d}/{len(slides):02d}"
        p.font.size=Pt(9); p.font.color.rgb=MUTED
        if notes_on and item["note"]:
            try: s.notes_slide.notes_text_frame.text=item["note"]
            except Exception: pass
    out=io.BytesIO(); prs.save(out); out.seek(0); return out

if submitted:
    if not chapter.strip():
        st.error("Chapter / Topic ka naam likhiye.")
    elif not lesson_text.strip():
        st.error("Pehle ChatGPT se slide content banwa kar neeche paste karein.")
    else:
        slides=parse_slides(lesson_text)
        if not slides:
            st.error("Format samajh nahi aaya. Har slide ke pehle [SLIDE] aur phir TITLE:/BULLET: lines honi chahiye.")
        else:
            ppt=make_ppt(slides,grade,chapter,board,notes_on)
            safe=re.sub(r"[^A-Za-z0-9_-]+","_",chapter).strip("_") or "Maths"
            st.success(f"{len(slides)} slides ki PowerPoint ready hai!")
            st.download_button("⬇️ Download PowerPoint", data=ppt,
                file_name=f"Class_{grade}_{safe}_{len(slides)}_Slides.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
            st.warning("Teacher review zaroor karein: formulas, diagrams, proofs aur calculations verify kar lein.")
st.divider()
st.caption("Free no-API builder: lesson content is pasted by the teacher; this app formats it into an editable PPT.")
