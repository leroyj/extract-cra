# CRA — Extracteur de Compte Rendu d'Activité

## Présentation

Ce petit outil Python extrait les données de plusieurs fichiers Excel (.xlsm) formatés selon le gabarit CRA (Compte Rendu d'Activité) et génère un fichier CSV contenant, par collaborateur et par activité, le temps passé par catégorie. Il lit la feuille nommée dans le script, parcourt les colonnes correspondant aux jours/semaines et produit un CSV par fichier source.

## Utilisation

Voir la section installation avant la première installation.

* Copier le répertoire CRA situé sur I: dans le répertoire du code
* Pensez à configurer la variable `BASE_DIR` en conséquence (seuls les répretoires commençant par `FOLDER_PREFIX` sont parcourus).

* Lancer le script principal :

```bash
uv run main.py
```

* Pour chaque fichier .xlsm détecté, un fichier CSV nommé `<original>.xlsm.csv` sera créé à côté du fichier source.
* Le CSV utilise `;` comme séparateur et contient l'en-tête :
  * annee;collaborateur;matricule;date_saisie;id_categorie;libelle_categorie;detail_activite;nb_jour

Pour concaténer tous les CSV générés en un seul fichier `all_cra.csv` :

```bash
# remplacez le chemin par celui de votre répertoire CRA
cat TEST/CRA\ TS/CRA\ */*.csv > CRA.csv
```

## Pré-requis

* macOS (ou tout système avec Python 3)
* Python 3.8+
* module Python : openpyxl
* Facultatif : utiliser un environnement virtuel (venv)

## Installation

### avec uv tools (recommandé)

* Cloner le dépôt : `git clone https://github.com/leroyj/extract-cra.git`
* Installer uv tools si ce n'est pas déjà fait : `pip install uv-tools`

c'est tout !

### Manuellement avec pip

```bash
#Créer et activer un environnement virtuel (recommandé) :
python3 -m venv .venv`
source .venv/bin/activate
#Installer la dépendance :
pip install -r requirements.txt
```

## Configuration

Les variables globales se trouvent au début de `main.py`. Les plus importantes :

* `FOLDER_PREFIX` (ex. `'CRA'`)
  Ne parcours que les dossier dont le nom commence par ce préfixe — les sous-dossiers correspondants sont parcourus pour trouver les .xlsm.

* `FILE_EXTENSION` (ex. `'.xlsm'`)
  Extension des fichiers à rechercher. (exclus les autres fichiers)

* `FILE_NAME`
  Nom d'exemple (utilisé pour tests/manuels) — généralement non nécessaire si vous utilisez `BASE_DIR`.

* `WORKSHEET_NAME` (ex. `'MàJ CRA'`)
  Nom de la feuille Excel à lire dans chaque fichier.

* `BASE_DIR`
  Chemin racine où commencer la recherche des dossiers `CRA`. Mettre `None` pour utiliser le dossier contenant le script :
  * Exemple : `BASE_DIR = '/Users/jleroy/Documents/dev/CRA/TEST/CRA TS'`
  * Ou : `BASE_DIR = None` (utiliser le répertoire du script)

Modification : si vous changez ces variables, enregistrez `main.py` puis lancez le script.

## Limitations

* Gérer les années bissextiles (366 jours) n'est pas encore implémenté.
* Calculer le nombre de jours de congés n'est pas encore implémenté.

## Dépannage rapide

* Erreur `UnboundLocalError: local variable 'BASE_DIR' referenced before assignment` : ne réaffectez pas la variable globale `BASE_DIR` directement dans la fonction ; utilisez un paramètre local (`base_dir`) ou référencez la variable globale sans réaffectation.
* Feuille introuvable : vérifiez que `WORKSHEET_NAME` correspond exactement (accents et casse).
* Dates affichées avec l'heure : convertissez avec `.date().isoformat()` avant d'écrire dans le CSV.
* Permissions / fichiers non trouvés : vérifiez `BASE_DIR` et droits d'accès.
