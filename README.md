# le beau rôle. — Pilotage du podcast

Dashboard autonome de suivi des données du podcast : **un seul fichier, `index.html`, à ouvrir dans un navigateur** (double-clic suffit). Aucune installation, aucune connexion internet requise : tous les listés graphiques sont dessinés à la main en SVG.

## Cinq onglets

| Onglet | Contenu |
|---|---|
| **Aperçu** | 4 grands indicateurs (écoutes relevées 30 s, estimation ajustée, lectures vérifiées, temps d'écoute), courbes cumulées SFC/ajusté, donut de répartition par plateforme, gains mensuels, funnel de découverte Spotify, production en cours |
| **Écoutes des épisodes** | Graphique par épisode (3 modes : relevé 30 s / ajusté 30 s / lectures play), écoutes par auditeur, taux de conversion, comparaison 60 s vs 30 s, durée moyenne d'écoute, registre complet triable et filtrable |
| **Rétention & qualité** | Rétention YouTube 30 s / 60 s, taux de complétion J0–J7, durée moyenne d'écoute Spotify et YouTube |
| **Découverte — Spotify** | Impressions mensuelles, conversion, origine de la découverte (Rechercher / Accueil / Bibliothèque), tops épisodes par impressions (Août, Septembre, Octobre 2026) |
| **Production** | Épisodes publiés / en cours / planifiés, timeline à venir avec avancement des checklists, suivi des déclarations SCAM |

## Enregistrer les données

1. Dans **Écoutes des épisodes**, cliquer sur une ligne du registre (ou « + Saisir les relevés »).
2. Le bouton **Modifier les relevés** ouvre le formulaire : Spotify 30 s, Autre 30 s, auditeurs uniques, Audioblog, YouTube, rétentions, complétion, temps de lecture, statut…
3. À l'enregistrement, **tous les totaux et estimations sont recalculés automatiquement** (même logique que le classeur : ajusté = total 30 s + plays × rétention YouTube moyenne ; YouTube estimé = lectures × rétention de l'épisode).
4. Les modifications sont conservées dans le navigateur (localStorage) et pilotables via la barre latérale :
   - **Exporter** → télécharge un instantané JSON complet (pour_exchange entre machines ou sauvegarde) ;
   - **Importer** → recharge un instantané
   - **Réinitialiser** → retour aux données d'origine du classeur Excel.

> Le fichier embarque une copie des données au 15 novembre 2026 (É30). Pour repartir du classeur Excel mis à jour, ré-extrayez `data.json` (script `extract.py`) et réinjectez-le dans premier `<script id="seed-data">` de `index.html`.

## Structure des données

- `Suivi épisodes` : 1 ligne par épisode, 32 colonnes de relevés
- `TDB` : séries mensuelles (écoutes cumulées, abonnés, conversion…)
- `Plateformes` : relevé brut vs estimation ajustée
- `Découverte` : impressions, conversion et origines mensuelles Spotify + tops épisodes
- `Exx` : checklists de production par épisode
