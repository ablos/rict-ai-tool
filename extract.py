"""Ask the LLM which methods a paper describes, and what it says about each one.

The prompt wording lives in prompts/*.md and the ontology in config/ontology.yaml;
this file only fills them in, calls the model, and flattens the answer into rows.
"""

from functools import cache
from pathlib import Path
from mistralai.client import Mistral
from mistralai.extra import response_format_from_pydantic_model
from pydantic import BaseModel
from ontology import build_output_model, build_screening_model, describe_ontology, relations

EXTRACTION_MODEL = "mistral-large-4-0"
TEMPERATURE = 0.0
TOP_P = 1.0 # The API requires 1 when temperature is 0
RANDOM_SEED = 42
REASONING_EFFORT = "high"
TIMEOUT_MS = 20 * 60 * 1000  # the SDK's default of 5 minutes is too short for long papers with reasoning
PROMPT_DIR = Path("prompts")

def paper_text(ocr: dict) -> str:
    """All OCR pages as one text, with a marker before each page so the model can cite pages."""
    return "\n\n".join(f"=== PAGE {page['page']} ===\n{page['markdown']}" for page in ocr["pages"])

@cache  # read once per run: editing a prompt during a run does not change that run
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
        timeout_ms=TIMEOUT_MS,
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

def screen_methods(client: Mistral, ocr: dict, ontology: dict) -> list[dict]:
    """The candidate methods in the paper, each checked against every inclusion criterion.

    A candidate is included only if it meets all criteria. That rule is applied here,
    not left to the model, so it is always applied the same way.
    """
    subject = ontology["subject"]
    system = load_prompt("identify").format(
        method_definition=subject["definition"],
        criteria="\n".join(f"- {criterion['id']}: {criterion['text']}" for criterion in subject["criteria"]),
    )
    answer = _ask(client, system, ocr, build_screening_model(ontology))
    candidates = []

    for candidate in answer.candidates:
        checks = {criterion["id"]: getattr(candidate, criterion["id"]).model_dump() for criterion in subject["criteria"]}
        candidates.append({
            "name": candidate.name,
            "criteria": checks,
            "included": all(check["met"] for check in checks.values()),
        })

    return candidates

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

    # Method names after the PDF path are optional; without them the model screens the paper.
    methods = sys.argv[2:]

    if not methods:
        for candidate in screen_methods(client, ocr, ontology):
            print(f"{candidate['name']}: {'included' if candidate['included'] else 'rejected'}")

            for criterion_id, check in candidate["criteria"].items():
                print(f"  {criterion_id} {'met' if check['met'] else 'NOT met'}: {check['reason']}")

            if candidate["included"]:
                methods.append(candidate["name"])

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
