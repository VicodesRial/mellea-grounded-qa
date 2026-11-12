from pathlib import Path

import mellea
from docx import Document

SYLLABI_DIR = Path("CSCourseSyllabi")


def read_file(course_code: str) -> str:
    path = SYLLABI_DIR / f"{course_code}.docx"
    if not path.is_file():
        raise FileNotFoundError(f"Syllabus not found: {path}")

    document = Document(path)
    sections = []

    sections.extend(
        paragraph.text.strip()
        for paragraph in document.paragraphs
        if paragraph.text.strip()
    )

    for table in document.tables:
        for row in table.rows:
            cells = []

            for cell in row.cells:
                text = cell.text.strip()

                if text and text not in cells:
                    cells.append(text)

            if cells:
                sections.append(" | ".join(cells))

    return "\n".join(sections)


def ask_question(course_code: str, question: str) -> str:
    if not isinstance(question, str) or not question.strip():
        raise ValueError("question cannot be empty")

    syllabus_text = read_file(course_code)
    session = mellea.start_session()
    prompt = (
        "Answer the question using only the provided syllabus context. "
        "If the answer is not in the syllabus, say 'I don't know.'\n\n"
        f"Question:\n{question.strip()}"
    )
    response = session.instruct(
        prompt,
        grounding_context={course_code: syllabus_text},
        strategy=None,
    )

    return str(response)
