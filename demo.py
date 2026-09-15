"""Renderer test without any LLM call — a hand-built deck.

    python demo.py [output_path.pptx]
"""

import sys

from atlas_deck.renderer import render_deck
from atlas_deck.schema import (
    BulletsSlide,
    Deck,
    ImpactRow,
    ImpactsSlide,
    Kpi,
    KpiSlide,
    Milestone,
    Recommendation,
    ReferenceSlide,
    SectionSlide,
    SummarySlide,
    TimelineSlide,
    TitleSlide,
)
from atlas_deck.theme import Theme

deck = Deck(
    title="Renforcement du dispositif de gestion du risque opérationnel",
    entity="Maghreb Universal Bank",
    slides=[
        TitleSlide(
            title="Renforcement du dispositif de gestion du risque opérationnel",
            subtitle="Analyse d'impact et trajectoire de mise en conformité",
            reference="Circulaire Bank Al-Maghrib n° 5/W/2022",
        ),
        SectionSlide(number=1, title="Ce que le texte change"),
        BulletsSlide(
            action_title="Le texte déplace la charge de la preuve vers la première ligne de défense",
            points=[
                "Les métiers deviennent responsables de la collecte exhaustive des incidents, "
                "sans filtrage préalable par la fonction Risques.",
                "Le seuil de déclaration obligatoire est abaissé, ce qui élargit mécaniquement "
                "l'assiette des incidents à documenter.",
                "La fonction Risques conserve la validation méthodologique et la consolidation "
                "vers le régulateur.",
            ],
            source="Circulaire BAM n° 5/W/2022, articles 8 à 14",
        ),
        ReferenceSlide(
            action_title="La notion d'incident significatif est désormais définie par des critères cumulatifs",
            reference="Circulaire BAM n° 5/W/2022, article 12",
            excerpt=(
                "Est réputé significatif tout incident opérationnel dont la perte brute "
                "excède le seuil fixé par l'établissement, ou dont les effets affectent "
                "la continuité d'un service essentiel pendant plus de deux heures."
            ),
            interpretation=[
                "Le seuil interne devient un paramètre réglementaire opposable, à faire valider "
                "par le comité des risques.",
                "La mesure de l'indisponibilité impose un horodatage fiable des incidents SI.",
            ],
        ),
        SectionSlide(number=2, title="Impacts sur le dispositif actuel"),
        ImpactsSlide(
            action_title="Quatre chantiers conditionnent la conformité à l'échéance",
            rows=[
                ImpactRow(
                    requirement="Collecte décentralisée",
                    impact="Désigner et former des correspondants incidents dans chaque métier",
                    entity="Risques / RH",
                    criticality="Élevée",
                ),
                ImpactRow(
                    requirement="Seuil de significativité",
                    impact="Faire valider le seuil interne par le comité des risques",
                    entity="Comité des risques",
                    criticality="Moyenne",
                ),
                ImpactRow(
                    requirement="Horodatage des incidents SI",
                    impact="Raccorder la supervision technique à la base incidents",
                    entity="DSI",
                    criticality="Élevée",
                ),
                ImpactRow(
                    requirement="Reporting régulateur",
                    impact="Adapter le format de remontée trimestrielle",
                    entity="Conformité",
                    criticality="Faible",
                ),
            ],
            source="Analyse BFS Consulting",
        ),
        KpiSlide(
            action_title="L'effort porte d'abord sur la couverture organisationnelle",
            kpis=[
                Kpi(value="18", label="mois pour atteindre la conformité pleine"),
                Kpi(value="42", label="correspondants incidents à désigner"),
                Kpi(value="4", label="chantiers structurants"),
            ],
            source="Analyse BFS Consulting",
        ),
        TimelineSlide(
            action_title="La trajectoire s'organise autour de quatre jalons de gouvernance",
            milestones=[
                Milestone(date="T4 2026", label="Validation du seuil par le comité des risques"),
                Milestone(date="T1 2027", label="Désignation et formation des correspondants"),
                Milestone(date="T3 2027", label="Raccordement de la supervision SI"),
                Milestone(date="T1 2028", label="Premier reporting au format cible"),
            ],
        ),
        SummarySlide(
            action_title="Trois décisions sont attendues du comité des risques",
            recommendations=[
                Recommendation(
                    title="Arrêter le seuil de significativité",
                    detail="Retenir un seuil aligné sur l'appétence au risque déjà approuvée, "
                    "et le documenter comme paramètre réglementaire opposable.",
                ),
                Recommendation(
                    title="Mandater la DSI sur l'horodatage",
                    detail="Le raccordement de la supervision technique conditionne la mesure "
                    "de l'indisponibilité, donc la qualification des incidents.",
                ),
                Recommendation(
                    title="Valider le calendrier de déploiement",
                    detail="La trajectoire proposée laisse six mois de marge avant l'échéance "
                    "réglementaire.",
                ),
            ],
        ),
    ],
)

if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else "demo-note-reglementaire.pptx"
    result = render_deck(deck, Theme.load("default"), output)
    print(f"File: {result.path}")
    if result.warnings:
        print("Layout warnings:")
        for w in result.warnings:
            print(f"  - {w}")
    else:
        print("No layout warnings.")
