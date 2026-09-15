# atlas_deck — génération de notes réglementaires (.pptx) pour agent LangGraph

Module à brancher dans Atlas Guardian. Le LLM ne produit **jamais** de PPTX ni de XML :
il produit un objet `Deck` validé, que le renderer transforme en fichier de façon
déterministe.

```
atlas_deck/
├── schema.py            # contrat de deck (Pydantic) — 8 types de slides
├── renderer.py          # moteur de rendu python-pptx, un helper par type
├── theme.py             # résolution du thème par tenant
├── themes/default.yaml  # charte : un fichier par client
├── llm.py               # fabrique du modèle (Ollama local ↔ endpoint compatible OpenAI)
└── graph.py             # sous-graphe LangGraph : plan → render → qa
demo.py                  # deck d'exemple, sans appel LLM
agent.py                 # CLI : génère une note via le graphe + Qwen
```

## Installation

Environnement virtuel (`venv` suffit ici — module unique, pas de publication
sur un registre, donc pas besoin de Poetry) :

```bash
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
python demo.py sortie.pptx      # vérifie la chaîne de rendu, sans LLM
```

### Modèle Qwen

Par défaut (`LLM_PROVIDER=ollama`), le graphe appelle un Qwen local via
[Ollama](https://ollama.com) — gratuit, hors-ligne :

```bash
ollama pull qwen2.5:7b-instruct
cp .env.example .env             # ajuster si besoin
python agent.py                  # demande d'exemple intégrée
python agent.py --request "..." --context-file extraits.txt --tenant default
```

Pour basculer sur un endpoint compatible OpenAI (palier gratuit OpenRouter
aujourd'hui, vLLM sur Runpod plus tard), il suffit de changer le `.env` — pas
de code à toucher :

```bash
LLM_PROVIDER=openai_compatible
QWEN_MODEL=qwen/qwen-2.5-72b-instruct   # ou le nom exposé par le serveur vLLM
OPENAI_BASE_URL=https://openrouter.ai/api/v1
OPENAI_API_KEY=...
```

## Les 8 types de slides

| Type | Usage dans une note réglementaire |
|---|---|
| `title` | Couverture — titre, sous-titre, texte réglementaire de rattachement |
| `section` | Intercalaire numéroté |
| `bullets` | Titre-message + 2 à 5 points |
| `reference` | Extrait verbatim du texte à gauche, lecture / portée à droite |
| `impacts` | Tableau exigence / impact / entité responsable / criticité |
| `kpi` | 2 à 4 chiffres clés (délais, volumétrie, seuils) |
| `timeline` | Échéancier de mise en conformité, 2 à 5 jalons |
| `summary` | 2 à 4 recommandations numérotées — ferme le deck |

L'inventaire est **fermé**. Ajouter un type = ajouter un modèle dans `schema.py`,
un helper dans `renderer.py`, une entrée dans `_HELPERS`. Rien d'autre.

## Brancher dans le graphe

```python
from atlas_deck.graph import build_graph, build_note_tool
from atlas_deck.llm import get_llm

graph = build_graph(get_llm())   # Qwen (Ollama ou endpoint compatible OpenAI)
state = graph.invoke({
    "request": "Note d'impact de la circulaire BAM 5/W/2022 sur le dispositif RO",
    "context": extraits_reglementaires + situation_entite,
    "tenant": "maghreb-universal-bank",
})
print(state["pptx_path"])
```

`build_graph` accepte n'importe quel modèle LangChain supportant
`with_structured_output` (Qwen via `get_llm()`, mais aussi Claude, GPT, etc.
si besoin un jour).

Ou, pour l'exposer comme outil à l'agent principal :

```python
agent = create_react_agent(llm, tools=[build_note_tool(llm), ...])
```

## Ajouter un client

Copier `themes/default.yaml` en `themes/<tenant>.yaml`, ajuster les couleurs et le
logo. Aucun code à modifier : le renderer ne connaît que des noms logiques
(`primary`, `accent`, `muted`, `block_background`). En production, remplacer la
lecture disque par une lecture en base, la signature de `Theme.load` ne change pas.

## Boucle qualité

`render_deck` retourne des avertissements de mise en page (titre trop long, slide
trop dense, cellule qui déborde) calculés à partir d'une estimation pessimiste de
la largeur des caractères en Arial. Le nœud `qa` les renvoie au planificateur, qui
reprend la structure — scinder une slide, raccourcir un libellé. Deux itérations
maximum, sinon le pipeline boucle sur des cas insolubles.

Pour ajouter la relecture visuelle dans le nœud `qa` :

```bash
soffice --headless --convert-to pdf deck.pptx
pdftoppm -jpeg -r 150 deck.pdf slide
```

puis passer les images à un modèle vision et concaténer ses remarques aux
avertissements existants. En SaaS, faire tourner LibreOffice dans un worker
séparé, jamais dans le process de l'API.

## Choix de conception à connaître

- **Pas de `.potx`.** Le style appartenant au client, un template par tenant serait
  ingérable ; la géométrie vient de la grille du renderer, les couleurs du YAML.
- **Arial partout**, disponible sur les postes des institutions financières. Le
  rendu LibreOffice utilise une substitution : contrôler sous PowerPoint avant
  livraison.
- **Contraintes de longueur dans le schéma.** C'est ce qui empêche les
  débordements, bien plus efficacement qu'une consigne dans le prompt.
- **Aucun filet décoratif, aucune bande de couleur.** Un seul accent, réservé aux
  chiffres clés, aux numéros de recommandation et à la criticité élevée.
