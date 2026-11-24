
import streamlit as st
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import json, re, io

st.set_page_config(page_title="Smart Teacher PPT Maker", page_icon="📚", layout="centered")

NAVY=RGBColor(22, 45, 82)
BLUE=RGBColor(40, 100, 170)
TEAL=RGBColor(0, 135, 145)
BG=RGBColor(247,249,252)
DARK=RGBColor(38,48,62)
MUTED=RGBColor(100,112,128)
WHITE=RGBColor(255,255,255)

st.title("📚 Smart Teacher PPT Maker")
st.caption("Class 6–10 | Mathematics | AI-powered PowerPoint generator")
st.info("Mobile aur computer dono ke browser mein chalega. API key sirf is session mein use hoti hai; app use save nahi karti.")

with st.form("ppt_form"):
    c1,c2=st.columns(2)
    with c1:
        grade=st.selectbox("Class", ["6","7","8","9","10"], index=4)
        subject=st.selectbox("Board / Curriculum", ["CBSE / NCERT", "State Board", "Other"])
        language=st.selectbox("Presentation language", ["Hindi + English (Hinglish)", "Hindi", "English"])
    with c2:
        slide_count=st.select_slider("Number of slides", options=[10,15,20,25,30], value=25)
        style=st.selectbox("PPT style", ["Clean classroom", "Modern colourful", "Minimal / exam-focused"])
        include_notes=st.checkbox("Teacher explanation notes include karein", value=True)
    chapter=st.text_input("Chapter name", placeholder="e.g. Circles / त्रिभुज / वास्तविक संख्याएँ")
    focus=st.text_area("Special instructions (optional)", placeholder="e.g. Theorems ke complete proof, 5 solved examples, board exam focus...")
    reference=st.text_area("NCERT / textbook reference (optional)", height=100,
        placeholder="Agar exact textbook alignment chahiye, relevant definitions/questions ya source text yahan paste karein.")
    api_key=st.text_input("OpenAI API key", type="password", help="API key platform.openai.com se milti hai. Ise kisi ke saath share na karein.")
    submitted=st.form_submit_button("✨ Generate PowerPoint", use_container_width=True)

def make_prompt():
    return f"""
Create a classroom-ready mathematics lesson presentation as valid JSON only.
Class: {grade}; curriculum: {subject}; chapter: {chapter}; language: {language};
exact slide count: {slide_count}; style: {style}; teacher notes: {include_notes}.
Special instructions: {focus or 'None'}.
Reference supplied by teacher (treat as source material, do not invent quotations): {reference or 'No source text supplied'}.

Return exactly this JSON structure:
{{
 "title": "presentation title",
 "slides": [
  {{
   "title": "short slide title",
   "bullets": ["short, teachable point", "..."],
   "example": "optional worked example or leave empty",
   "teacher_note": "optional explanation for teacher"
  }}
 ]
}}
Rules:
- Exactly {slide_count} slides.
- Accurate school-level mathematics. Check calculations and algebra before responding.
- Use standard mathematical notation in plain text (e.g. x², √, ∠, π, ⟂, =).
- Include learning goals, concepts, definitions, worked examples, common mistakes, practice questions and recap as appropriate.
- If the topic contains theorems, include statement, labelled proof steps and conclusion.
- Do not claim a question is verbatim NCERT unless it appears in the supplied reference text.
- Keep each slide concise: 3–5 bullets, each preferably under 18 words.
- If information is uncertain, label it for teacher verification rather than inventing.
- Output valid JSON only, with no markdown fences.
"""

def clean_json(raw):
    raw=raw.strip()
    raw=re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I)
    start=raw.find("{"); end=raw.rfind("}")
    if start<0 or end<start: raise ValueError("AI response mein valid JSON nahi mila.")
    return json.loads(raw[start:end+1])

def build_ppt(data, grade, chapter, style, include_notes):
    prs=Presentation()
    prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
    for idx,item in enumerate(data["slides"],1):
        slide=prs.slides.add_slide(prs.slide_layouts[6])
        bg=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb=BG; bg.line.fill.background()
        band=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width, Inches(.15))
        band.fill.solid(); band.fill.fore_color.rgb=TEAL; band.line.fill.background()
        tb=slide.shapes.add_textbox(Inches(.65), Inches(.42), Inches(12), Inches(.7))
        p=tb.text_frame.paragraphs[0]; p.text=item.get("title",""); p.font.name="Aptos Display"; p.font.size=Pt(28); p.font.bold=True; p.font.color.rgb=NAVY
        bullets=item.get("bullets",[]) or []
        body=slide.shapes.add_textbox(Inches(.9), Inches(1.55), Inches(11.5), Inches(4.3))
        tf=body.text_frame; tf.word_wrap=True
        for j,b in enumerate(bullets):
            p=tf.paragraphs[0] if j==0 else tf.add_paragraph()
            p.text=b; p.font.name="Aptos"; p.font.size=Pt(21 if len(bullets)<=4 else 18); p.font.color.rgb=DARK
            p.space_after=Pt(14); p.level=0
            p.text="•  "+p.text
        ex=item.get("example","")
        if ex:
            card=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(.9), Inches(5.85), Inches(11.5), Inches(.72))
            card.fill.solid(); card.fill.fore_color.rgb=RGBColor(230,242,246); card.line.color.rgb=RGBColor(215,228,235)
            tf=card.text_frame; tf.clear(); tf.vertical_anchor=MSO_ANCHOR.MIDDLE
            p=tf.paragraphs[0]; p.text="Example: "+ex; p.font.name="Aptos"; p.font.size=Pt(15); p.font.color.rgb=TEAL
        foot=slide.shapes.add_textbox(Inches(.65), Inches(7.08), Inches(12), Inches(.22))
        p=foot.text_frame.paragraphs[0]; p.text=f"CLASS {grade} • MATHEMATICS • {chapter.upper()}                                      {idx:02d}/{len(data['slides']):02d}"
        p.font.size=Pt(9); p.font.color.rgb=MUTED
        note=item.get("teacher_note","")
        if include_notes and note:
            try:
                slide.notes_slide.notes_text_frame.text = note
            except Exception:
                pass
    out=io.BytesIO()
    prs.save(out); out.seek(0)
    return out

if submitted:
    if not chapter.strip():
        st.error("Pehle chapter ka naam likhiye.")
    elif not api_key.strip():
        st.error("OpenAI API key enter karein. Key ke bina AI content generate nahi hoga.")
    else:
        try:
            with st.spinner("AI lesson content bana raha hai aur slides assemble ho rahi hain..."):
                client=OpenAI(api_key=api_key.strip())
                response=client.responses.create(
                    model="gpt-5.5",
                    instructions="You are an experienced CBSE mathematics teacher and meticulous mathematical reviewer.",
                    input=make_prompt(),
                    max_output_tokens=10000
                )
                data=clean_json(response.output_text)
                if len(data.get("slides",[])) != slide_count:
                    raise ValueError(f"AI ne {len(data.get('slides',[]))} slides di; expected {slide_count}. Dobara generate karein.")
                ppt=build_ppt(data,grade,chapter,style,include_notes)
            safe=re.sub(r"[^A-Za-z0-9_-]+","_",chapter).strip("_") or "Maths"
            st.success(f"{len(data['slides'])} slides ki PPT ready hai!")
            st.download_button("⬇️ Download PowerPoint", data=ppt,
                file_name=f"Class_{grade}_{safe}_{slide_count}_Slides.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
            st.warning("Class mein dikhane se pehle formulas, proofs aur examples ek baar verify kar lein. Exact NCERT wording ke liye textbook reference paste karein.")
        except Exception as e:
            st.error(f"Generation mein dikkat aayi: {e}")
            st.caption("API key, internet connection aur API billing/credits check karein.")
st.divider()
st.caption("Smart Teacher PPT Maker • Mathematics only • Drafts are AI-generated and should be teacher-checked.")
