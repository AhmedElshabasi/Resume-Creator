import os
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
import re
from datetime import datetime
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
import json
import json
import re
from docx.shared import Inches

def first_existing_style(doc: Document, candidates: list[str]):
    """Return the first style object that exists in doc.styles from candidates, else None."""
    for name in candidates:
        try:
            return doc.styles[name]
        except KeyError:
            continue
    return None

def replace_placeholder_with_bullets(doc: Document, placeholder: str, items: list[str]) -> bool:
    """
    Replace a placeholder paragraph (e.g. "[skills123]") with a bullet list.
    Uses an existing bullet/list style if found; otherwise falls back to a
    manual bullet glyph and hanging indent.
    """
    bullet_style_obj = first_existing_style(
        doc,
        # Try common English names first; add your own if needed
        ["List Bullet", "List Paragraph", "Bulleted List", "Bullet", "Normal"]
    )

    def iter_paragraphs(d: Document):
        for p in d.paragraphs:
            yield p
        for tbl in d.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

    for p in iter_paragraphs(doc):
        if placeholder in p.text:
            if not items:
                # nothing to insert, just clear placeholder
                set_paragraph_text(p, "")
                return True

            # Decide rendering path: real style vs manual bullets
            use_manual_bullets = (bullet_style_obj is None or bullet_style_obj.name in ("Normal", "List Paragraph"))

            def style_as_bullet(par):
                if bullet_style_obj is not None:
                    par.style = bullet_style_obj
                if use_manual_bullets:
                    # manual bullet look
                    par.paragraph_format.left_indent = Inches(0.25)
                    par.paragraph_format.first_line_indent = Inches(-0.15)

            # First item replaces the placeholder line
            style_as_bullet(p)
            set_paragraph_text(p, (("• " if use_manual_bullets else "") + items[0]))

            # Remaining items
            anchor = p
            for it in items[1:]:
                anchor = insert_paragraph_after(anchor)
                style_as_bullet(anchor)
                set_paragraph_text(anchor, (("• " if use_manual_bullets else "") + it))
            return True

    return False

def extract_json(text: str) -> str:
    """
    Try a few strategies to pull a valid JSON object from a model response.
    1) If it's inside a code fence ```json ... ```, grab that.
    2) If it looks like `resume_data = { ... }`, grab the {...}.
    3) Fallback: take the first balanced {...} block.
    """
    if not text or not text.strip():
        raise ValueError("Empty API response")

    # 1) ```json ... ``` or ``` ... ```
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, flags=re.DOTALL)
    if m:
        return m.group(1)

    # 2) resume_data = { ... }
    m = re.search(r"resume_data\s*=\s*(\{.*\})", text, flags=re.DOTALL)
    if m:
        return m.group(1)

    # 3) first balanced { ... } block (greedy from first { to last })
    first = text.find("{")
    last = text.rfind("}")
    if first != -1 and last != -1 and last > first:
        return text[first:last+1]

    # If all else fails
    raise ValueError("Could not find JSON object in response")


# Load .env file
load_dotenv()

client = OpenAI(
    api_key = os.getenv("OPENAI_API_KEY"),
)

with open("input.txt", "r", encoding="utf-8") as file:
    job_descrip = file.read()

print("Job description loaded. and about to call the api")

completion = client.chat.completions.create(
    model="gpt-4o",
    messages=[
        {
            "role": "user",
            "content": f"""Ahmed Elshabasi
(587)-435-9647 – Calgary, AB – ahmed.elshabasi@ucalgary.ca – LinkedIn: linkedin.com/in/ahmed-elshabasi/
Skills
•	Programming languages: Java, Python, JavaScript, C, C++, HTML, CSS, React, Node.JS, Express.JS, SQL
•	Software tools: VS Code, Git, Github, Gitlab, Unity, Unreal Engine
•	Algorithm and Data Structures: Studied different algorithms and structures in university
•	Professional Skills: Adaptability, Communication, Detail-oriented, Leadership, and Time Management
Experience
Undergraduate Research Assistant (Node, React, JS)					                           May 2024 – Sep 2024
University of Calgary,     Calgary, AB
•	Developed an automated workflow using Node and React for extracting in-depth information about the data in the corpus completing tasks within set deadlines.
•	Collected first-person videos, spoken recordings, and biometric data for a corpus of information needs for outdoor running.
•	Quickly learned new tools and technologies (Node, React, and smart capture devices) to automate data analysis workflows.
Projects
Self-Checkout Machine Software (Java)								          Sep 2023 – Dec 2023
•	Collaborated with a team of 20 to designed and develop the software for a self-checkout machine
•	Focused on user-friendly interface design and efficient transaction handling to ensure smooth customer experience.
•	Integrated functionalities that simulate real-world use cases.
Educational Assessment Web App (JS, CSS, HTML)						            Jan 2024 – Apr 2024
•	Collaborated with a team of 5 to design a web application for dynamic assessments.
•	Implemented functionality to generate random questions per session, offering immediate grading and feedback.
•	Emphasized designing a user-friendly interface to provide smooth navigation and create an engaging test experience.
Full-stack Financial Assistant| Hackathon Project (Node, React, JS)	   		     	                              Feb 2024
•	Led a team of 4 to build a full-stack prompt-based financial assistant.
•	 Used ChatGPT’s API for real-time financial insights and assistance.
•	Ensured seamless deployment within the time constraints of the hackathon (24 hours).
Education
Bachelor of Computer Science									   2022 – 2026 [Expected]
University of Calgary,  Calgary, AB									    GPA: 3.68
•	Awards:
o	PURE award
o	President’s Admission Scholarship
o	University of Calgary International Undergraduate Award
•	Certificates:
o	Google Cybersecurity Professional Certificate		
o	Ready for Research Micro credential
Volunteering
Setup Crew											          Jan 2022 – May 2022
G.N.P. Hospital,    Jeddah, Saudi Arabia
•	Assisted medical students with a staff team of 5 and documented the students’ progress.
•	Collaboratively smoothed the experience for the students by getting their feedback and answering inquiries.
•	Recorded students’ progress and managed their attendance for their hospital’s academy training.
Executive Team Member									          Dec 2021 - Apr 2022
Model United Nations (MUN) at Dar Jana International School
•	Managed and participated with the MUN executive team in the organization of documents and preparation of participants
•	Prepared the hall in which the event takes place and ensured that the event proceeded as expected.
•	As one of the spokespersons (chairmen) of the event, fulfilled my duties of planning and managing the procedures of the delegates.
Extracurricular Clubs
Event Coordinator										         Sep 2023 – May 2024
Infosec Club											
•	Coordinated hands-on security labs and ethical hacking workshops, providing practical learning experiences for members.
Member											         Sep 2023 – May 2024
Competitive Programming Club									 
•	Engaged in practical programming challenges, honing my problem-solving skills through regular workshop attendance.




this is my resume and bs coverletter keep that in mind cause i am about to ask you to do stuff

Here is the job description: {job_descrip}

i want you to infiltrate my experiences with details that aid my core skills and support my ability to do the major aspects of what's required from the job description (it doesn't have to be legitimate but it has to be powerful), then create a resume.

i want you to output your answer in a very strict format of a json array that looks like:

resume_data = {{
    "skills123": [
        "result1",
        "result2",
        "result3",
        "result4.... and so on as necessary"
    ],
     "education123": [
         "result1",
        "result2",
        "result3",
        "result4.... and so on as necessary"
    ],
    "experience123": [
         "result1",
        "result2",
        "result3",
        "result4.... and so on as necessary"
    ],
    "projects123": [
         "result1",
        "result2",
        "result3",
        "result4.... and so on as necessary"
    ],
}}
"""
        }
    ]
)

api_response = completion.choices[0].message.content

# ---------- helpers ----------
def insert_paragraph_after(paragraph: Paragraph) -> Paragraph:
    """Create a new empty paragraph directly after `paragraph`."""
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)                  # type: ignore[attr-defined]
    return Paragraph(new_p, paragraph._parent)

def set_paragraph_text(p: Paragraph, text: str):
    """Replace all runs in a paragraph with a single run containing `text`."""
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.add_run(text)


# ---------- usage ----------


api_response = completion.choices[0].message.content
json_str = extract_json(api_response)
resume_data = json.loads(json_str)

doc = Document("resume_template.docx")           # your .docx version of the 2nd screenshot
replace_placeholder_with_bullets(doc, "[skills123]", resume_data["skills123"])
doc.save("resume_filled_skills.docx")
