
# backend/fae_core/requirements.py
import io, os, tempfile
import pdfplumber
from docx import Document
from .llm import call_llm, extract_first_json_block

def summarize_requirement(req_text: str) -> str:
    """
    Convierte el requisito literal del baseline en el título técnico corto
    como el que aparece en el Excel de compliance.
    """
    from .llm import call_llm

    prompt = f"""
Rewrite this requirement as a SHORT technical compliance title (max 12 words).
Do NOT invent anything. Just summarize the core intent.

Requirement:
{req_text}
"""

    try:
        result = call_llm([
            {"role": "system", "content": "You write compliance requirement titles."},
            {"role": "user", "content": prompt}
        ], max_tokens=60)

        return result.strip().replace("\n", " ")
    except Exception:
        return req_text


import re

def extract_basic_requirements(text: str, document: str, include_will: bool = False) -> list[dict]:
    rows = []
    if not text:
        return rows

    lines = text.replace("\r", "\n").split("\n")

    current_id = None
    current_text = []

    for line in lines:
        line = line.strip()

        # Detectar inicio por ID
        match = re.match(r'^(\d{6})\s+(.*)', line)
        if match:
            # guardar anterior
            if current_id and current_text:
                full_text = " ".join(current_text)
                short_text = summarize_requirement(full_text)

                rows.append({
                    "Documento": document,
                    "Id": current_id,
                    "Requisito": short_text,
                    "Tipo": "REQ",
                    "Ambiguo": False
                })


            current_id = match.group(1)
            current_text = [match.group(2)]

        else:
            if current_id:
                if line:
                    current_text.append(line)

    # último
    if current_id and current_text:
        rows.append({
            "Documento": document,
            "Id": current_id,
            "Requisito": " ".join(current_text),
            "Tipo": "REQ",
            "Ambiguo": False
        })

    return rows





def read_doc_to_text(file_name: str, file_bytes: bytes) -> str:
    """
    Convierte un DOCX o PDF a texto usando librerías robustas.
    - DOCX: python-docx
    - PDF: pdfplumber (evita .decode('latin1') y problemas de encoding)
    """
    name = file_name.lower()

    if name.endswith(".docx"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(file_bytes)
            tmp.flush()
            path = tmp.name
        try:
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs)
        finally:
            try:
                os.remove(path)
            except:
                pass

    if name.endswith(".pdf"):
        all_text = []

        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                txt = page.extract_text(layout=True)
                if txt:
                    all_text.append(txt)

        return "\n".join(all_text)


    # Otros tipos no soportados en este paso
    return ""
def extract_by_id_blocks(text: str, document: str) -> list[dict]:
    import re

    print("TEXT LENGTH:", len(text))
    print("FIRST 2000 CHARS:\n", text[:2000])
    print("LAST 2000 CHARS:\n", text[-2000:])

    pattern = r'(?m)^\s*(17\d{4}|18\d{4}|20\d{4})\b'
    matches = list(re.finditer(pattern, text))

    print("TOTAL MATCHES:", len(matches))
    print("IDS FOUND:", [m.group(1) for m in matches])

    rows = []

    for i in range(len(matches)):
        start = matches[i].start()
        end = matches[i+1].start() if i+1 < len(matches) else len(text)

        block = text[start:end].strip()
        req_id = matches[i].group(1)

        rows.append({
            "Documento": document,
            "Id": req_id,
            "Requisito": block,
            "Tipo": "BASELINE_ID",
            "Ambiguo": False
        })

    return rows

def extract_requirements_semantic(text: str, document: str) -> list[dict]:
    # 1️⃣ Intentar extracción estructural por ID
    id_rows = extract_by_id_blocks(text, document)

    # Si parece un baseline estructurado (>=30 IDs), usar esta extracción
    if len(id_rows) >= 30:
        return id_rows
    

    """
    Extrae requisitos por significado, no por 'shall'.
    Genera títulos técnicos tipo compliance matrix.
    """
    from .llm import call_llm, extract_first_json_block

    prompt = f"""
You are creating a Cybersecurity Compliance Matrix.

From the document below, generate the checklist items that a supplier must demonstrate.

IMPORTANT:
- These are NOT sentences from the document.
- These are COMPLIANCE CHECK ITEMS like in a matrix.
- You may combine multiple related requirements into one checklist item.
- Include hardening, configuration, and best practice obligations even if not written as 'shall'.

Return ONLY a JSON array like:
[
  {{
    "requirement": "short compliance checklist item"
  }}
]

DOCUMENT:
{text[:12000]}
"""


    try:
        raw = call_llm([
            {"role": "system", "content": "You extract technical compliance requirements."},
            {"role": "user", "content": prompt}
        ], max_tokens=2000)

        data = extract_first_json_block(raw)

        rows = []
        for r in data:
            rows.append({
                "Documento": document,
                "Requisito": r.get("requirement", ""),
                "Tipo": "SEMANTIC",
                "Ambiguo": False
            })

        return rows

    except Exception:
        return []

    
# --- Añadir al final de backend/fae_core/requirements.py ---

def split_text(text: str, chunk_size: int = 9000, overlap: int = 500) -> list[str]:
    """
    Divide un texto largo en bloques solapados para procesarlos por IA.
    chunk_size: tamaño objetivo de cada bloque (~9k recomendado)
    overlap:   solape entre bloques para no cortar párrafos (p.ej. 500)
    """
    chunks = []
    i = 0
    n = len(text)
    while i < n:
        j = min(i + chunk_size, n)
        chunks.append(text[i:j])
        if j == n:
            break
        i = max(0, j - overlap)
    return chunks


def extract_if244_with_ai_multichunk(full_text: str, document: str) -> list[dict]:
    """
    Recorre el documento en bloques y fusiona los requisitos encontrados por IA.
    """
    all_rows: list[dict] = []

    for idx, chunk in enumerate(split_text(full_text)):
        prompt = f"""
You are a senior automotive validation engineer.
Extract ALL IF-244 relevant requirements from the document below.
RULES:
- Extract explicit AND implicit requirements.
- Rewrite them as clear, testable requirements.
- If no IF-244 requirements exist, return an empty list.
- Do NOT invent requirements.
Return ONLY a valid JSON array.
NO explanations. NO markdown.
JSON FORMAT:
[
  {{
    "description": "...",
    "priority": "High | Medium | Low",
    "category": "ENVIRONMENTAL | ELECTRICAL | MECHANICAL | QUALITY | OTHER",
    "if244_section": "...",
    "objective": "...",
    "validation_criteria": "..."
  }}
]
DOCUMENT NAME: {document}
DOCUMENT CONTENT:
{chunk}
""".strip()

        try:
            raw = call_llm([
                {"role": "system", "content": "Expert IF-244 requirements engineer"},
                {"role": "user",   "content": prompt}
            ], max_tokens=1500)

            # (Opcional de depuración)
            # print("[IF244-IA] RAW:", (raw or "")[:400])

            data = extract_first_json_block(raw)
            for r in data:
                all_rows.append({
                    "Description": r.get("description", ""),
                    "Priority": r.get("priority", ""),
                    "Category": r.get("category", ""),
                    "IF-244 Section": r.get("if244_section", ""),
                    "Objective": r.get("objective", ""),
                    "Validation criteria": r.get("validation_criteria", ""),
                    "Source document": f"{document} (chunk {idx+1})"
                })
        except Exception:
            # Si un bloque falla, seguimos con los demás
            continue

    # Deduplicación simple por Description
    seen = set()
    unique_rows: list[dict] = []
    for row in all_rows:
        key = row["Description"].strip().lower()
        if key in seen:
            continue
        seen.add(key)
        unique_rows.append(row)

    return unique_rows

def extract_ens_requirements_ai(full_text: str, document: str) -> list[dict]:
    """
    Extrae todos los Test Requirements, Acceptance Criteria, Severity Levels,
    Conditions, Functional Status, Test Duration, Temperature, Voltage,
    Operating Conditions, y Test Procedures del documento ENS.
    """
    from .llm import call_llm, extract_first_json_block

    prompt = f"""
You are a senior validation engineer with expertise in ENS, ISO 16750, ISO 20653 and EMC testing.

Your task is to extract ALL test-related requirements from the document below:
- Acceptance Criteria
- Severity levels
- Test conditions
- Environmental levels (temperature, humidity, voltage, vibration, dust, water ingress)
- Durations
- Functional status (Class A/B/C)
- DUT operating modes (B01, C11, FT0, FT1, FT2…)
- Procedures
- Notes and constraints
- Any detail defining PASS/FAIL
- Any limits or mandatory behavior of the DUT during the test.

VERY IMPORTANT:
• Extract even if “shall”, “must” or “will” do NOT appear.
• ENS310 uses descriptive sections like “Purpose”, “Effects”, “Procedure”, “Acceptance criteria” → THESE ARE THE REQUIREMENTS.
• Make each requirement concise and testable.
• Group them into units (one row per requirement).

RETURN ONLY a valid JSON array like this:
[
  {{
    "test_name": "Temperature Cycle Test",
    "section": "10.1.3",
    "requirement": "The DUT shall operate through the entire temperature sweep without impaired function.",
    "conditions": "10 cycles, -40°C to +125°C depending on location",
    "acceptance_criteria": "Class A functions at end of test",
    "location": "1A, 2C, 3B...",
    "functional_status": "FT2",
    "notes": "Full I/O monitoring required"
  }}
]

DOCUMENT NAME: {document}
DOCUMENT CONTENT (TRUNCATED TO FIRST 20k CHARS):
{full_text[:20000]}
"""

    try:
        raw = call_llm([
            {"role": "system", "content": "Expert ENS requirements extraction"},
            {"role": "user", "content": prompt}
        ], max_tokens=1800)
        


        data = extract_first_json_block(raw)
        rows = []

        for r in data:
            rows.append({
                "Test Name": r.get("test_name", ""),
                "Section": r.get("section", ""),
                "Requirement": r.get("requirement", ""),
                "Conditions": r.get("conditions", ""),
                "Acceptance Criteria": r.get("acceptance_criteria", ""),
                "Location": r.get("location", ""),
                "Functional Status": r.get("functional_status", ""),
                "Notes": r.get("notes", ""),
                "Source document": document
            })

        return rows

    except Exception as e:
        print("ERROR ENS extraction:", e)
        return []



