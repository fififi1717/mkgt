#!/usr/bin/env python3
"""
Affilior — Traduit "nombre d'éléments réels pour ce dossier" en positions XML
à supprimer, à partir du référentiel générique template_positions.json
(commun à tous les clients, lié à la version de Cans_Mstr.pptx).

C'est la seule chose qui varie d'un client à l'autre pour cette étape : le
NOMBRE de projets / moyens / axes réellement retenus. Les coordonnées elles-
mêmes ne bougent jamais (cf. template_positions.json).

Usage :
    from select_slots import compute_slots_a_supprimer
    slots = compute_slots_a_supprimer(template_positions, {"12": 2, "13": 2, "15": 2})
    # -> {"12": {"x": [5897880]}, "13": {"y": [2715768, 3026664, 3337560]}, "15": {"x": [5897880]}}
"""
import json


def compute_slots_a_supprimer(template_positions: dict, nb_elements_reels: dict) -> dict:
    slots_variables = template_positions["slots_variables"]
    result = {}
    for slide_n, nb_reel in nb_elements_reels.items():
        if slide_n not in slots_variables:
            raise ValueError(f"slide {slide_n} absente de template_positions['slots_variables']")
        conf = slots_variables[slide_n]
        positions = conf["positions"]
        max_slots = conf["max_slots"]
        if nb_reel > max_slots:
            raise ValueError(
                f"slide {slide_n} ({conf['description']}) : {nb_reel} éléments réels "
                f"mais le template n'a que {max_slots} slots — cas non géré automatiquement, "
                f"cf. limite dure documentée (comme slide 8 / 3 enfants)."
            )
        # on garde les nb_reel premiers slots, on supprime le reste
        a_supprimer = positions[nb_reel:]
        if a_supprimer:
            result[slide_n] = {conf["axe"]: a_supprimer}
    return result


if __name__ == "__main__":
    import sys
    tp = json.load(open(sys.argv[1], encoding="utf-8"))
    # exemple : dossier avec 2 projets (slide12), 2 moyens (slide13), 2 axes (slide15)
    nb = {"12": 2, "13": 2, "15": 2}
    slots = compute_slots_a_supprimer(tp, nb)
    print(json.dumps(slots, ensure_ascii=False, indent=2))
