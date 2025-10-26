# CRA — Extracteur de Compte Rendu d'Activité

## Présentation

Ce petit outil Python extrait les données de plusieurs fichiers Excel (.xlsm) formatés selon le gabarit CRA (Compte Rendu d'Activité) et génère un fichier CSV contenant, par collaborateur et par activité, le temps passé par catégorie. Il lit la feuille nommée dans le script, parcourt les colonnes correspondant aux jours/semaines et produit un CSV par fichier source.

## Pré-requis

- macOS (ou tout système avec Python 3)
- Python 3.8+
- module Python : openpyxl
- Facultatif : utiliser un environnement virtuel (venv)

## Installation

1. Créer et activer un environnement virtuel (recommandé) :
   - python3 -m venv .venv
   - source .venv/bin/activate
2. Installer la dépendance :
   - pip install openpyxl

## Configuration

Les variables globales se trouvent au début de `main.py`. Les plus importantes :

- FOLDER_PREFIX (ex. `'CRA'`)  
  Ne parcours que les dossier dont le nom commence par ce préfixe — les sous-dossiers correspondants sont parcourus pour trouver les .xlsm.

- FILE_EXTENSION (ex. `'.xlsm'`)  
  Extension des fichiers à rechercher. (exclus les autres fichiers)

- FILE_NAME  
  Nom d'exemple (utilisé pour tests/manuels) — généralement non nécessaire si vous utilisez `BASE_DIR`.

- WORKSHEET_NAME (ex. `'MàJ CRA'`)  
  Nom de la feuille Excel à lire dans chaque fichier.

- BASE_DIR  
  Chemin racine où commencer la recherche des dossiers `CRA`. Mettre `None` pour utiliser le dossier contenant le script :
  - Exemple : `BASE_DIR = '/Users/jleroy/Documents/dev/CRA/TEST/CRA TS'`
  - Ou : `BASE_DIR = None` (utiliser le répertoire du script)

Modification : si vous changez ces variables, enregistrez `main.py` puis lancez le script.

## Utilisation

1. Placez vos fichiers `.xlsm` dans des dossiers dont le nom commence par le préfixe (par défaut `CRA`) sous `BASE_DIR`.
2. Depuis le terminal (dans le dossier du projet) :
   - source .venv/bin/activate  # si vous utilisez un venv
   - python3 main.py
3. Pour chaque fichier .xlsm détecté, un fichier CSV nommé `<original>.xlsm.csv` sera créé à côté du fichier source.
4. Le CSV utilise `;` comme séparateur et contient l'en-tête :
   - annee;collaborateur;matricule;date_saisie;id_categorie;libelle_categorie;detail_activite;nb_jour

## Remarques techniques

- openpyxl renvoie souvent des objets `datetime.datetime` pour les cellules date. Le script doit convertir en ISO date sans l'heure, par exemple :
  - `val.date().isoformat()` ou `val.date()` puis format `YYYY-MM-DD`.
- Si une cellule date est un nombre Excel, utilisez `openpyxl.utils.datetime.from_excel()` pour la convertir.
- Le script actuel parcourt les colonnes jour/semaines à partir de la colonne D (index 4). Adaptez si votre gabarit diffère.
- Pour inclure la dernière valeur `offset = 5`, utilisez `range(1, 6)` (borne supérieure exclusive).

## Dépannage rapide

- Erreur `UnboundLocalError: local variable 'BASE_DIR' referenced before assignment` : ne réaffectez pas la variable globale `BASE_DIR` directement dans la fonction ; utilisez un paramètre local (`base_dir`) ou référencez la variable globale sans réaffectation.
- Feuille introuvable : vérifiez que `WORKSHEET_NAME` correspond exactement (accents et casse).
- Dates affichées avec l'heure : convertissez avec `.date().isoformat()` avant d'écrire dans le CSV.
- Permissions / fichiers non trouvés : vérifiez `BASE_DIR` et droits d'accès.
