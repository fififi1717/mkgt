# CHANGELOG v4.19 → v4.20 (30/08/2026, crash-test double client — correctifs)

⚠ Fichiers PAS ENCORE POUSSÉS sur mkgt (droits d'écriture manquants). À faire par Marc :
1. Remplacer les 4 fichiers ci-joints sur le dépôt (finalize_deck.py, check_template_residue.py,
   remove_solution_slot.py, personalize_canvas.py) + template_positions.json.
2. Mettre à jour CHECKSUMS.sha256 avec les nouveaux hash listés ci-dessous.

## Nouveaux hash (sha256)
- finalize_deck.py            : 0c6f46175af542803e1296c2da69f5b24f5fe1e553725a59491888120b970037
- check_template_residue.py   : ca4c452151a778e244ec17d287c562b09faf02ab19f4ad26757fe16da3afc8d3
- remove_solution_slot.py     : cbcc1d0123126d20140fcdf5e602f8f974fad8f19e6db6e141421bfc1c4f4d33
- personalize_canvas.py       : 3ba44d97ae1e1b3c4b8dfcc436022744ee5ac8c9c85ff3eadcd74d7fde46492e
- template_positions.json     : b5de38d71d75f2eccdc327f1e5adc15071d5ec433af8afead296ed2181173fe3
- check_charts.py             : INCHANGÉ (4f7ebd13...) — pas de modification nécessaire, cf. correctif 1 ci-dessous.

## Correctifs (5), tous re-testés en réel sur les 2 dossiers fictifs du crash-test du 30/08/2026

**1. Déduplication allocation_pct/liquidite_pct (finalize_deck.py)**
Avant : ces deux clés devaient être saisies à l'identique dans plan.json (pour personalize_canvas.py,
labels texte slide 6) ET dans dossier_client.json (pour check_charts.py, graphique natif slide 6).
Un oubli dans l'un des deux reproduit le bug historique R14 (légende correcte, donut resté en démo)
— confirmé en crash-test sur les 2 dossiers avant correctif.
Après : `finalize_deck.py` accepte un argument optionnel `--plan plan.json`. Si allocation_pct/
liquidite_pct sont absents de --dossier, ils sont repris depuis --plan (dossier reste prioritaire
si les deux sont fournis ; si les deux sont fournis et DIFFÈRENT, la clé est ignorée des deux côtés
avec un avertissement, plutôt que de choisir arbitrairement une source).
Testé : les 2 dossiers ont eu leurs 2 donuts corrigés automatiquement en fournissant uniquement
--plan, sans dupliquer allocation_pct/liquidite_pct dans dossier_client.json.

**2. Slide 17 "Accompagnement" absente de la personnalisation (personalize_canvas.py,
template_positions.json)**
Avant : la clé "17" (bloc "Engagements spécifiques à ce dossier", 6 placeholders) n'était documentée
nulle part et absente de template_positions.json["slides_personnalisees"] — elle sortait donc
TOUJOURS avec "[Engagement 1]" etc. visibles, sur tous les dossiers générés depuis v4.19, sans
qu'aucun contrôle ne le signale comme anormal.
Après : "17" ajoutée à slides_personnalisees + documentée dans le docstring de personalize_canvas.py
(6 valeurs, ordre : 3x [intitulé engagement, détail]).
Testé : les 2 dossiers sortent avec un contenu réel sur cette slide, sans intervention manuelle.

**3. remove_solution_slot.py silencieux quand il ne trouve rien à faire**
Avant : sur le Canvas Master actuel, la slide 15 n'a plus qu'UN SEUL emplacement solution par axe
(le "double placement" documenté dans template_positions.json["slots_variables"]["15"]
["solution_slot_2_y"] ne correspond à aucune forme réelle) — le script tournait en no-op silencieux
(0 forme supprimée, aucune erreur), donnant une fausse impression de traitement effectué.
Après : avertissement explicite ⚠ dès que 0 forme est supprimée alors qu'une suppression était
demandée, renvoyant vers la note ajoutée dans template_positions.json.
Testé : avertissement bien affiché sur les 2 dossiers (0 forme supprimée dans les deux cas).

**4. template_positions.json["slots_variables"]["8"] obsolète**
Avant : décrivait encore 3 lignes "Enfant 1/2/3" individuelles, alors que le Cans_Mstr.pptx en
production utilise déjà le format résolu "Par enfant (×N)" à une seule ligne (confirmé par
inspection XML directe le 30/08/2026 — les 3 offsets Y documentés correspondent en réalité aux
lignes "Par enfant (×N)" / "Conjoint" / "Capitaux décès", pas à 3 enfants). Utiliser cette entrée
via nb_elements_reels aurait supprimé la ligne "Capitaux décès" au lieu d'un enfant excédentaire.
Après : entrée retirée. Aucune action requise pour la slide 8 quel que soit le nombre d'enfants
(déjà résolu v4.14/15, reconfirmé).

**5. Faux positif R16 sur la slide Synthèse (check_template_residue.py)**
Avant : le script flaguait systématiquement 6 runs en couleur C4B5A5 sur la slide "Synthèse de la
stratégie" comme résidus non recolorés — alors que R16 (corrigée v4.19, confirmée par Marc) précise
explicitement que C4B5A5 coexiste comme couleur de hiérarchie volontaire sur cette slide. QA visuelle
réelle (rendu image) confirme un texte parfaitement lisible, pas un résidu "beige sur beige". Le
script n'avait jamais été mis à jour après la correction R16 de v4.19.
Après : détection par marqueur de contenu ("STRATÉGIE · SYNTHÈSE", eyebrow fixe) plutôt que par
numéro de slide — plus robuste, les numéros de fichier internes ne correspondant pas de façon
stable à la position narrative après assemblage (constaté en crash-test).
Testé : 11 résidus → 5 résidus sur les 2 dossiers (les 6 supprimés étaient bien tous des faux
positifs Synthèse ; les 5 restants — notes de conception R15 — sont de vrais résidus, inchangés).

## Non traité dans ce lot (reporté, cf. tableau de synthèse du 30/08/2026)
- check_montants.py : faux positifs bénins sur les seuils de barème IR fixes (11 497/29 315/
  83 823/180 294 €) — mineur, non corrigé ici.
- Slide 14 "Impact" : mise en page fixe à 3 catégories, sans lien avec nb_elements_reels["12"] —
  mineur, non corrigé ici (chantier plus large : ajouter 14 à slots_variables).
- assemble.py : clé "patrimoine" du spec.json est un vestige mort (imprime un message, n'injecte
  rien) — cosmétique, non corrigé ici.
- template_positions.json["canvas_sha256"] (champ interne informatif) toujours désynchronisé du
  hash réel de Cans_Mstr.pptx — cosmétique, non corrigé ici.
- 5 résidus R15 restants (notes de conception interne, ex. "Adapter le nombre de colonnes") —
  toujours en lecture seule (check_template_residue.py ne corrige jamais), à nettoyer manuellement
  avant chaque envoi client comme aujourd'hui.
