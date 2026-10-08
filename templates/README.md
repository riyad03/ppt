# Templates clients

Déposer ici le `.pptx` d'un client — un vrai template, un ancien deck, un
export d'un autre outil : aucune structure particulière n'est requise.

Test rapide, sans appel au modèle (~1 seconde) :

```bash
python demo.py sortie.pptx --template templates/client.pptx
```

Génération complète avec le modèle :

```bash
python agent.py --template templates/client.pptx
```

Les `.pptx` de ce dossier sont ignorés par git (données client).
