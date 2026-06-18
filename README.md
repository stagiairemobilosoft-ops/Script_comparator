# Script_Synchronisation magasins Nickel
Script Python pour comparer les fichiers mensuels Nickel (~13 000 magasins) et produire le fichier Excel attendu par l'equipe IT (depot FTP / Filezilla).

## Installation

```powershell
cd nickel
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

# requirement
pandas>=2.2.0
openpyxl>=3.1.0


## Workflow mensuel

1. Recevoir le nouveau fichier Nickel (ex. `Nickel_Gmaps_20260615.xlsx`)
2. Comparer avec le fichier du mois precedent (deja synchronise)
3. Deposer `output/resultat/fichier_synchronisation.xlsx` sur Filezilla
4. Conserver le nouveau fichier : il servira de base au mois suivant

# CMD
```powershell
python main.py compare FICHIER_MOIS_PRECEDENT.xlsx FICHIER_NOUVEAU.xlsx
```

Exemple :

```powershell
python main.py validate Nickel_Gmaps_20260211_.xlsx  
```


- 1er argument : fichier du mois precedent
- 2e argument : nouveau fichier recu

Option : `-o chemin/fichier.xlsx` pour changer le fichier de sortie.

Pendant l'execution, un indicateur affiche l'avancement (ex. `Comparaison : 5000 / 12861`). Option `-q` pour le masquer.

Important : compare toujours mois N-1 vs mois N. Ne pas comparer un fichier très ancien avec le nouveau, sinon des changements deja synchronises ressortiraient.

## Controle qualite (CS)

Verifie un fichier sans lancer de comparaison.

```powershell
python main.py validate Nickel_Gmaps_20260211_.xlsx  
```

## Fichiers produits

```
output/
├── resultat/
│   └── fichier_synchronisation.xlsx   # a deposer sur Filezilla
└── log/
    ├── analyse_YYYYMMDD_HHMMSS.xlsx     # detail de la comparaison
    ├── compare_YYYYMMDD_HHMMSS.log      # resume texte
    └── problemes_qualite.xlsx           # rapport validate
```

## Classification (flow IT)


| Situation                        | ACTION             | STATUS |
| -------------------------------- | ------------------ | ------ |
| Magasin inchange                 | (exclu du fichier) | -      |
| Magasin modifie                  | UPDATE             | OPEN   |
| Nouveau magasin                  | CREATE             | OPEN   |
| Magasin absent du fichier client | UPDATE             | CLOSED |


Les doublons sur `Code_de_magasin` sont fusionnes (derniere ligne conservee).

## Format de sortie

Le fichier resultat suit le modele `CLIENT_EXPORT_DATA` :

`SHOP_CODE`, `ACTION`, `STATUS`, `LOCATION_NAME`, `LOCATION_LANGUAGE`, `PHONE_NUMBER`, `EMAIL`, `WEBSITE_URL`, `ADR_STREET_NAME`, `ADR_STREET_NUMBER`, `ADR_ZIPCODE`, `ADR_CITY`, `ADR_COUNTRY`, `ADR_LATITUDE`, `ADR_LONGITUDE`, `REGULAR_HOURS`, `SPECIAL_HOURS (OPTIONAL)`

## Fichiers sources acceptes

- Export Nickel Gmaps (feuille `all_in` ou `Feuille 1`)

## Notes

- Fermer `fichier_synchronisation.xlsx` dans Excel avant de relancer une comparaison.
- Certains magasins n'ont pas d'horaires dans le fichier source : `REGULAR_HOURS` reste vide.
- Certains champs ne sont pas fournis par Nickel.
                  
# Refactorisation

|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|       main.py                |                             |- main()                     | 

## constants/
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|       constant.py            |variable :STORE_ID_COLUMN / FILE_COLUMNS / SYNC_FIELDS /   |  
|                              |SCHEDULE_COLUMNS /CLIENT_COLUMN_RENAME / OPTIONAL_COLUMNS/ | 
|                              |RESULT_COLUMNS / STATUS_OPEN / STATUS_CLOSED / OUTPUT_DIR/ |
|                              |EXPORT_ACTION_MAP /OUTPUT_RESULTAT / OUTPUT_LOG /          |
|                              |DEFAULT_COMPARE_OUTPUT / DEFAULT_VALIDATE_OUTPUT /         |
|                              |ACTION_NO_CHANGE / ACTION_UPDATE / ACTION_CREATE /         |
|                              |ACTION_CLOSE / ACTION_COLUMN                               | 
## normalize/ 
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|      file_loader.py          |-LoadError                   | - _is_client_format()       |
|                              |                             | - _find_data_sheet()        | 
|                              |                             | - _normalize_raw_dataframe()|
|                              |                             | - load_store_file()         |

## comparison/
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|   compare.py                 |-ComparisonResult            |- _index_by_store()          | 
|                              |                             |- _rows_to_dataframe()       | 
|                              |                             |- compare_store_files()      |          

 
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|          progress.py         |-ProgressReporter            |- update                     | 
|                              |                             |- finish                     |


|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|        utils.py              |                             |- print-stage                | 
|                              |                             |- print_setp_done            |
|                              |                             |- normalize_cell             | 

## export/
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|     excel_export.py          |                             |- _cell_display_length()     | 
|                              |                             |- autofit_columns()          |
|                              |                             |- write_excel()              | 
|                              |                             |- write_workbook()           |
|                              |                             |- split_address()            | 
|                              |                             |- _join_regular_hours()      |
|                              |                             |- to_result_layout()         |
|                              |                             |- build_sync_file()          |
|                              |                             |- _sync_columns              |
|                              |                             |- with_action                |


|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|         log.py               |                             |- build_analysis_summary()   | 
|                              |                             |- write_analysis_workbook()  | 
|                              |                             |- timestamp()                | 
|                              |                             |- timestamped_log_path()     |
|                              |                             |- timestamped_analysis_path()|
|                              |                             |- write_log()                |


|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|      cmd.py                  |                             |- cmd_compare()              | 
|                              |                             |- cmd_validate()             | 
|                              |                             |- _ensure_output_dirs ()     | 



## validation/
|       Fichier                |           Classes           |          Fonctions          |
|------------------------------|-----------------------------|-----------------------------|
|        quality_check.py      |                             |- _issue()                   | 
|                              |                             |- _to_float                  | 
|                              |                             |- deduplicate_stores         | 
|                              |                             |- find_data_quality_issues() | 

