# Openpyxl - Boardgames

## Dictionnaire des données

Nous travaillons sur 3 jeux de données ayant pour thème les jeux de société. 
Toutes les données font référence à la classification des jeux de société disponible sur le site [Board Game Geek (BGG)](https://boardgamegeek.com/).

Chaque enregistrement correspond à un jeu de société. 
Voici les champs utilisés, leur type et leur description :

- ID, int64, Identification interne de BGG
- Name, object, Nom du jeu
- Year Published, int64, Année de publication du jeu
- Min Players, int64, Nombre de joueurs minimal
- Max Players, int64, Nombre de joueurs maximal
- Play Time, int64, Temps de jeu d'une partie
- Min Age, int64, Age minimum conseillé pour jouer
- Users Rated, int64, Nombre d'utilisateurs ayant noté le jeu
- Rating Average, float64, Note moyenne des avis (calcul BGG)
- BGG Rank, int64, Rang global BGG
- Complexity Average, float64, Note de complexité
- Owned Users, int64, Nombre d'utilisateurs possédant le jeu
- Mechanics, object, Mécaniques du jeu
- Domains, object, Famille du jeu
- artist, object, Artiste
- designer, object, Designer
- publisher, object, Maison d'édition
- is_expansion, int64, Est une extension d'un autre jeu
- abstracts_rank, int64, Rang dans la famille des jeux abstraits
- cgs_rank, int64, Rang dans la famille des jeux de cartes
- childrensgames_rank, int64, Rang dans la famille des jeux pour enfants
- familygames_rank, int64, Rang dans la famille des jeux familiaux
- partygames_rank, int64, Rang dans la famille des jeux d'ambiance
- strategygames_rank, int64, Rang dans la famille des jeux de stratégie
- thematic_rank, int64, Rang dans la famille des jeux thématiques
- wargames_rank, int64, Rang dans la famille des wargames

## Source

Le jeu de données principal est disponible sur le site [IEE DataPort](https://ieee-dataport.org/open-access/boardgamegeek-dataset-board-games).

Le second jeu de données est disponible sur [Kaggle](https://www.kaggle.com/datasets/sujaykapadnis/board-games).

Le dernier jeu de données est accessible sur [BGG](https://boardgamegeek.com/wiki/page/BGG_XML_API2).

## Objectifs du projet

L'objectif du projet est de découvrir ce qui caractérise les différentes famille de jeu de société à travers des graphiques et des indicateurs.

## Schéma du projet

<img src='https://github.com/xifar08/Openpyxl-Boardgames/blob/main/assets/sch%C3%A9ma%20initial.png'>

## Prérequis

- uv
- Python >= 3.13

## Installation

```
git clone https://github.com/xifar08/Openpyxl-Boardgames.git
cd Openpyxl-Boardgames
uv sync
```

## Utilisation

```
uv run src/openpyxl_boardgames/main.py
```

Le reporting est produit dans results/reporting.xlsx .

## Structure du projet

```
Openpyxl-Boardgames/
├── pyproject.toml
├── uv.lock
├── .python-version
│
├── assets/                      # Images / schémas
│
├── notebooks/
│   ├── explore.ipynb            # Notebook d'exploration
│   ├── feat-1.ipynb
│   ├── feat-2.ipynb
│   ├── feat-3.ipynb
│   └── feat-4.ipynb
│  
│
└── src/openpyxl_boardgames/
    ├── __init__.py
    ├── config.py                # Constantes : chemins, références Excel, formules, config des données
    ├── data.py                  # Chargement et nettoyage des données source
    ├── dashboard.py              # Orchestration : assemblage du classeur Excel complet
    ├── main.py                  # Point d'entrée CLI : génération du reporting Excel
    │
    ├── components/
    │   ├── __init__.py
    │   ├── charts.py             # Création des différents graphiques
    │   ├── data_validation.py    # Filtres sous forme de validation de données
    │   └── utils.py              # Nettoyage et normalisation des données
    │
    └── sheets/
        ├── __init__.py
        ├── tableau.py             # Construction de la feuille TABLEAU (titres + formules)
        └── tdb.py                 # Construction de la feuille TDB (titres + formules)
```