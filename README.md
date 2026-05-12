# Openpyxl - Boardgames

## Dictionnaire des données
Nous travaillons sur 3 jeux de données ayant pour thèmes les jeux de société. 
Toutes les données font références à la classification des jeux de société disponible sur le site [Board Game Geek (BGG)](https://boardgamegeek.com/).

Chaque enregistrement correspond à un jeu de société. 
Voici les champs utilisés, leur type et leur description :

- ID, int64, Identification interne à la classification faite par BGG
- Name, object, Nom du jeu
- Year Published, int64, Année de publication du jeu
- Min Players, int64, Nombre de joueurs minimal
- Max Players, int64, Nombre de joueurs maximal
- Play Time, int64, Temps de jeu d'une partie
- Min Age, int64, Age minimum conseillé pour jouer
- Users Rated, int64, Note des utilisateurs sur BGG
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
- cgs_rank, int64, 
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

## Schéma du projet