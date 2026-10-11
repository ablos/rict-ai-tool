"""Ask the LLM which methods a paper describes, and what it says about each one.

The prompt wording lives in prompts/*.md and the ontology in config/ontology.yaml;
this file only fills them in, calls the model, and flattens the answer into rows.
"""

from pathlib import Path
from mistralai.client import Mistral
from mistralai.extra import response_format_from_pydantic_model
from pydantic import BaseModel
from ontology import build_output_model, describe_ontology, relations

EXTRACTION_MODEL = "mistral-large-4-0"
TEMPERATURE = 0.0
TOP_P = 1.0 # The API requires 1 when temperature is 0
RANDOM_SEED = 42
REASONING_EFFORT = "high"
PROMPT_DIR = Path("prompts")

class MethodList(BaseModel):
    methods: list[str]

def paper_text(ocr: dict) -> str:
    """All OCR pages as one text, with a marker before each page so the model can cite pages."""
    return "\n\n".join(f"=== PAGE {page['page']} ===\n{page['markdown']}" for page in ocr["pages"])

def load_prompt(name: str) -> str:
    return (PROMPT_DIR / f"{name}.md").read_text(encoding="utf-8")

def _ask(client: Mistral, system: str, ocr: dict, output_model: type[BaseModel]) -> BaseModel:
    """Send the prompt plus the paper, and parse the answer into output_model."""
    response = client.chat.complete(
        model=EXTRACTION_MODEL,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": paper_text(ocr)},
        ],
        response_format=response_format_from_pydantic_model(output_model),
        temperature=TEMPERATURE,
        top_p=TOP_P,
        random_seed=RANDOM_SEED,
        reasoning_effort=REASONING_EFFORT,
    )
    choice = response.choices[0]

    # "length" means the answer was cut off, so the JSON is incomplete.
    if choice.finish_reason != "stop":
        raise RuntimeError(f"Model stopped early ({choice.finish_reason}) on {ocr['pdf']}")

    # A reasoning model answers with "thinking" chunks first; the JSON is in the "text" chunks.
    content = choice.message.content

    if isinstance(content, list):
        content = "".join(chunk.text for chunk in content if chunk.type == "text")

    return output_model.model_validate_json(content)

def identify_methods(client: Mistral, ocr: dict, ontology: dict) -> list[str]:
    """Names of the methods the paper describes, for when none are given."""
    subject = ontology["subject"]
    system = load_prompt("identify").format(
        method_definition=subject["definition"],
        criteria="\n".join(f"- {criterion}" for criterion in subject["criteria"]),    
    )
    return _ask(client, system, ocr, MethodList).methods

def extract_claims(client: Mistral, ocr: dict, method: str, ontology: dict) -> list[dict]:
    """One row per claim, plus one row for each relation without claims."""
    system = load_prompt("extract").format(method=method, ontology=describe_ontology(ontology))
    answer = _ask(client, system, ocr, build_output_model(ontology))

    rows = []
    for concept, relation in relations(ontology):
        result = getattr(answer, relation["name"])
        base = {
            "publication": ocr["pdf"],
            "pdf_sha256": ocr["sha256"],
            "method": method,
            "relation": relation["name"],
            "concept": concept["label"],
        }

        # Trust the claims over the status: a "reported" relation without claims
        # counts as not reported, and claims are kept even if marked "not_reported".
        if not result.claims:
            rows.append({**base, "status": "not_reported"})

        for claim in result.claims:
            rows.append({**base, "status": "reported", **claim.model_dump()})

    return rows

if __name__ == "__main__":
    import os
    import sys

    from dotenv import load_dotenv

    from ocr import ocr_pdf
    from ontology import load_ontology

    load_dotenv()
    client = Mistral(api_key=os.environ["MISTRAL_API_KEY"])
    ontology = load_ontology("config/ontology.yaml")
    ocr = ocr_pdf(client, sys.argv[1])

    # Method names after the PDF path are optional; without them the model picks.
    methods = sys.argv[2:] or identify_methods(client, ocr, ontology)
    print(f"Methods: {methods}\n")

    for method in methods:
        print(f"===== {method} =====")
        for row in extract_claims(client, ocr, method, ontology):
            if row["status"] == "not_reported":
                print(f"{row['relation']:<22} (not reported)")
            else:
                print(f"{row['relation']:<22} {row['value']}  [p{row['page']}, {row['section']}]")
                print(f"{'':<22} \"{row['quote']}\"")
        print()
