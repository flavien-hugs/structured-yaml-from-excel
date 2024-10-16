# Génération de Structure de Données à partir de Fichiers Excel

## Contexte

Ce projet permet de générer une structure de données YAML à partir de feuilles de calcul Excel contenant
des informations sur des divisions géographiques. Les données sont extraites de plusieurs onglets dans le fichier
Excel (comme les délégations régionales, régions, départements, sous-préfectures, localités, et quartiers) et
organisées de manière hiérarchique dans un fichier YAML. Cette solution est particulièrement utile pour
les applications nécessitant des données géographiques structurées sous format lisible.

## Outils Utilisés

Le projet utilise les outils et bibliothèques suivants :

- **Python** : Langage principal du projet.
- **Pandas** : Bibliothèque pour la manipulation et l’analyse de données, utilisée pour lire et transformer les 
  données Excel.
- **PyYAML** : Bibliothèque pour la gestion du format YAML, utilisée pour générer le fichier de sortie en YAML.
- **Typer** : Bibliothèque Python permettant de créer des interfaces en ligne de commande (CLI) avec une syntaxe
  simple et intuitive.
- **Pathlib** : Fournit des classes pour manipuler des chemins de fichiers de manière flexible.

### Prérequis

1. Assurez-vous d'avoir Python installé (version 3.7 ou plus récente).
2. Installez les dépendances du projet en utilisant `pip` ou `poetry`.

```bash
# Si vous utilisez pip
pip install pandas pyyaml typer openpyxl

# Si vous utilisez poetry
poetry install
```

## Utilisation
```bash
poetry run generate-data --filepath <path/to/your/data.xlsx> --output-file <path/to/output/data.yml>
```
