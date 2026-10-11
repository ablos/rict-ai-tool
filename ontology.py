"""Load the ontology config and derive the prompt text and output schema from it.

Nothing in this file names a concept or relation: everything comes from the YAML,
so changing the ontology means editing config/ontology.yaml
"""

import hashlib
from pathlib import Path
from typing import Literal
import yaml
from pydantic import BaseModel, Field, create_model

def load_ontology(path: str | Path) -> dict:
    """Read the config, drop disabled concepts, and check that it is usable."""
    raw = Path(path).read_bytes()
    ontology = yaml.safe_load(raw)

    ontology["concepts"] = [c for c in ontology["concepts"] if c.get("enabled", True)]
    ontology["hash"] = hashlib.sha256(raw).hexdigest()[:12]

    _check(ontology)
    return ontology

def _check(ontology: dict) -> None:
    """Fail on config mistakes now, instead of halfway through a paid API call."""
    seen = set()

    for concept in ontology["concepts"]:
        for key in ("id", "label", "definition", "relations"):
            if key not in concept:
                raise ValueError(f"Concept '{concept.get('id', '?')}' is missing '{key}'")

        for relation in concept["relations"]:
            name = relation["name"]

            # Relation names become keys in the JSON output, so keep them simple
            if not name.isidentifier():
                raise ValueError(f"Relation name '{name}' must be one word, like 'usedIn'")

            if name in seen:
                raise ValueError(f"Relation '{name}' appears more than once")

            seen.add(name)

def relations(ontology: dict) -> list[tuple[dict, dict]]:
    """All (concept, relation) pairs, in config order."""
    return [(concept, relation) for concept in ontology["concepts"] for relation in concept["relations"]]

def describe_ontology(ontology: dict) -> str:
    """The ontology as readable text, to paste into the extraction prompt."""
    subject = ontology["subject"]
    lines = [f"{subject['label']}: {subject['definition']}", ""]

    for concept in ontology["concepts"]:
        lines.append(f"## {concept['label']}")
        lines.append(concept['definition'])

        if concept.get("examples"):
            lines.append(f"Examples: {', '.join(concept['examples'])}")
        
        lines.append("Relations:")

        for relation in concept['relations']:
            line = f"- {relation['name']}: {relation['definition']}"

            if relation.get("examples"):
                line += f" (examples: {', '.join(relation['examples'])})"
            
            lines.append(line)

        lines.append("")

    return "\n".join(lines)

class Claim(BaseModel):
    # The quote comes first so the model commits to evidence before writing a value
    quote: str = Field(description="Verbatim text copied from the paper that supports the value.")
    value: str = Field(description="The extracted value, in the paper's own wording.")
    scope: str = Field(
        description='"general" if the claim is about the method itself, '
        "otherwise the name of the application case it comes from."
    )
    page: int = Field(description="Number of the page the quote is on, from the page markers.")
    section: str | None = Field(description="Heading of the section the quote is in, or null.")

class RelationResult(BaseModel):
    status: Literal["reported", "not_reported"]
    claims: list[Claim]

def build_output_model(ontology: dict) -> type[BaseModel]:
    """A Pydantic model with one required field per relation.

    Because every relation is a required key, the model cannot silently skip one:
    it has to say either "reported" (with claims) or "not_reported".
    """

    fields = {
        relation["name"]: (RelationResult, Field(description=relation["definition"]))
        for _, relation in relations(ontology)
    }

    return create_model("MethodExtraction", **fields)

if __name__ == "__main__":
    import json

    ontology = load_ontology("config/ontology.yaml")
    print(f"Ontology version {ontology['version']}, hash {ontology['hash']}\n")
    print(describe_ontology(ontology))
    print(json.dumps(build_output_model(ontology).model_json_schema(), indent=2))
