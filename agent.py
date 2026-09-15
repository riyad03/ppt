"""CLI: generates a regulatory-note .pptx via the LangGraph + Qwen graph.

    python agent.py --request "..." [--context-file path.txt] [--tenant default] [--out dir]

With no arguments, uses a sample request to verify the full chain
(plan → render → qa) end-to-end against a real model.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from atlas_deck.graph import build_graph
from atlas_deck.llm import get_llm

# Kept in French: this is the sample business content fed to the LLM (a
# French regulatory-note request), not code.
DEFAULT_REQUEST = (
    "Note d'impact de la circulaire Bank Al-Maghrib n° 5/W/2022 sur le dispositif "
    "de gestion du risque opérationnel de la banque."
)

DEFAULT_CONTEXT = """\
Circulaire BAM n° 5/W/2022, articles 8 à 14 : les métiers deviennent responsables
de la collecte exhaustive des incidents opérationnels, sans filtrage préalable par
la fonction Risques. Le seuil de déclaration obligatoire est abaissé. La fonction
Risques conserve la validation méthodologique et la consolidation vers le régulateur.

Article 12 : est réputé significatif tout incident opérationnel dont la perte
brute excède le seuil fixé par l'établissement, ou dont les effets affectent la
continuité d'un service essentiel pendant plus de deux heures.

Situation de l'entité : seuil interne non encore validé par le comité des risques,
42 correspondants incidents à désigner dans les métiers, supervision technique non
raccordée à la base incidents. Échéance réglementaire : 18 mois.
"""


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--request", default=DEFAULT_REQUEST)
    parser.add_argument("--context-file", type=Path, default=None)
    parser.add_argument("--tenant", default="default")
    parser.add_argument("--out", dest="output_dir", default="./sorties")
    args = parser.parse_args()

    context = (
        args.context_file.read_text(encoding="utf-8")
        if args.context_file
        else DEFAULT_CONTEXT
    )

    llm = get_llm()
    graph = build_graph(llm)
    state = graph.invoke(
        {
            "request": args.request,
            "context": context,
            "tenant": args.tenant,
            "output_dir": args.output_dir,
        }
    )

    print(f"File: {state['pptx_path']}")
    if state.get("warnings"):
        print("Layout warnings:")
        for w in state["warnings"]:
            print(f"  - {w}")
    else:
        print("No layout warnings.")


if __name__ == "__main__":
    main()
