#!/usr/bin/env python3
"""
Affilior — Suppression ciblée d'un emplacement "solution" sur la slide 15
(règle 5.2sexies, ajouté le 28/08/2026).

POURQUOI CE SCRIPT EXISTE (à ne pas fusionner naïvement avec remove_unused_slots.py)
------------------------------------------------------------------------------------
La slide 15 (Synthèse) refondue en cartes pleines porte, dans le pied de chaque
carte, DEUX emplacements solution empilés, pour le cas où un même objectif est
servi par deux placements (ex. assurance-vie française + assurance-vie
luxembourgeoise). Quand une colonne n'a qu'un seul placement, il faut retirer
l'emplacement n°2 de CETTE colonne seulement.

remove_unused_slots.py ne peut pas faire ça, par construction :
  - il matche par offset X  -> supprimerait la colonne entière ;
  - il matche par offset Y  -> supprimerait l'emplacement n°2 des TROIS colonnes
    d'un coup, puisque les trois cartes partagent les mêmes Y.
Le matching X OU Y est un "ou" logique : il ne sait pas cibler une cellule.

Ce script fait donc un matching X **ET** Y (intersection) = une cellule précise.
Il est volontairement séparé pour ne pas modifier la sémantique de
remove_unused_slots.py, utilisé par ailleurs sur les slides 8/12/13/15.

CONTRAINTE DE GABARIT À NE PAS CASSER
-------------------------------------
Toutes les formes d'une colonne de la slide 15 doivent rester calées exactement
sur l'offset X de la colonne (marge visuelle obtenue via lIns/rIns du texte,
jamais via un décalage de x). Sinon elles survivent à la suppression de colonne
par remove_unused_slots.py — bug réel constaté le 28/08/2026 : seul le fond de
carte disparaissait et tout le texte restait affiché.

ORDRE D'EXÉCUTION
-----------------
À lancer AVANT remove_unused_slots.py (donc avant finalize_deck.py, ou intégré
en tête de chaîne) : on retire d'abord les emplacements solution inutiles, puis
on retire les colonnes d'axes inutiles. L'inverse fonctionne aussi mais fait
travailler ce script sur des colonnes déjà supprimées, pour rien.

Usage :
    python3 remove_solution_slot.py deck.pptx --out deck2.pptx \
        --template-positions template_positions.json \
        --colonnes-un-seul-placement 2 3

    (numéros de colonne = 1, 2, 3 de gauche à droite)
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

TOLERANCE_EMU = 20000  # même tolérance que remove_unused_slots.py, par cohérence


def remove_shapes_at_cell(xml: str, x_values: set, y_values: set) -> tuple:
    """Supprime les formes dont l'offset X ET l'offset Y correspondent (à la
    tolérance près) — c'est-à-dire une cellule précise, pas une ligne ni une
    colonne entière. Retourne (xml_modifie, nb_supprimees, details)."""
    def near(val, values):
        return any(abs(val - v) <= TOLERANCE_EMU for v in values)

    shapes = re.findall(r'<p:sp>.*?</p:sp>', xml, re.DOTALL)
    removed = 0
    details = []
    for sp in shapes:
        m = re.search(r'<a:off x="(\d+)" y="(\d+)"/>', sp)
        if not m:
            continue
        x, y = int(m.group(1)), int(m.group(2))
        if near(x, x_values) and near(y, y_values):   # ET logique, pas OU
            nm = re.search(r'<p:cNvPr id="\d+" name="([^"]*)"', sp)
            details.append((nm.group(1) if nm else '?', x, y))
            xml = xml.replace(sp, '', 1)
            removed += 1
    return xml, removed, details


def apply(src_pptx, dst_pptx, slide_n, x_values, y_values):
    zin = zipfile.ZipFile(src_pptx, 'r')
    items = zin.infolist()
    data = {it.filename: zin.read(it.filename) for it in items}
    zin.close()

    fname = f'ppt/slides/slide{slide_n}.xml'
    if fname not in data:
        sys.exit(f"ERREUR : {fname} absent de {src_pptx}")

    xml = data[fname].decode('utf-8')
    xml, removed, details = remove_shapes_at_cell(xml, set(x_values), set(y_values))
    data[fname] = xml.encode('utf-8')

    print(f'slide{slide_n}.xml : {removed} forme(s) supprimée(s) '
          f'(cellules x={sorted(x_values)} ET y={sorted(y_values)})')
    for name, x, y in details:
        print(f'    - {name:14s} x={x} y={y}')

    # CORRECTIF v4.20 (30/08/2026, crash-test double client) : un résultat à
    # 0 forme supprimée s'affichait auparavant de façon aussi discrète qu'un
    # résultat normal ("0 forme(s) supprimée(s)" ressemble à une ligne de log
    # anodine) — repéré en crash-test : sur les 2 dossiers testés, les
    # offsets 'solution_slot_2_y' de template_positions.json ne correspondent
    # à AUCUNE forme réelle du Cans_Mstr.pptx en production (la slide 15 n'a
    # plus qu'un seul emplacement solution par axe depuis une refonte non
    # répercutée dans template_positions.json — cf. avertissement dans ce
    # fichier). Sans ce correctif, le script « réussit » silencieusement sans
    # rien faire, ce qui peut laisser croire à tort que le double placement a
    # été traité. Avertissement explicite désormais si x_values/y_values sont
    # non vides mais qu'aucune forme n'a matché.
    if removed == 0 and (x_values or y_values):
        print("⚠ ATTENTION : 0 forme supprimée alors qu'une suppression était demandée. "
              "Cela signifie très probablement que les offsets fournis (cf. "
              "template_positions.json['slots_variables']['15']) ne correspondent à AUCUNE "
              "forme réelle de ce fichier — vérifier si le Canvas Master a changé de "
              "structure (cf. avertissement v4.20 dans template_positions.json) avant de "
              "considérer cette étape comme traitée.")

    zout = zipfile.ZipFile(dst_pptx, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        zout.writestr(it, data[it.filename])
    zout.close()
    print(f'OK -> {dst_pptx}')
    return removed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template-positions", required=True)
    ap.add_argument("--colonnes-un-seul-placement", nargs="*", type=int, default=[],
                    help="Numéros de colonnes (1-3) n'ayant QU'UN placement : "
                         "leur emplacement solution n°2 sera retiré.")
    ap.add_argument("--slide", type=int, default=15)
    args = ap.parse_args()

    tp = json.load(open(args.template_positions, encoding="utf-8"))

    key = str(args.slide)
    conf = tp.get("slots_variables", {}).get(key)
    if conf is None:
        sys.exit(f"ERREUR : slide {key} absente de template_positions['slots_variables']")

    positions = conf["positions"]
    sol2 = conf.get("solution_slot_2_y")
    if sol2 is None:
        sys.exit(
            f"ERREUR : template_positions['slots_variables']['{key}'] ne contient pas "
            f"'solution_slot_2_y'. Ce gabarit de slide {key} ne gère pas le double "
            f"placement — régénérer template_positions.json contre le Cans_Mstr.pptx courant."
        )

    cols = args.colonnes_un_seul_placement
    if not cols:
        print("Aucune colonne déclarée à un seul placement — rien à supprimer.")
        import shutil
        shutil.copy(args.src, args.out)
        print(f'OK -> {args.out}')
        return

    for c in cols:
        if not 1 <= c <= len(positions):
            sys.exit(f"ERREUR : colonne {c} hors bornes (1..{len(positions)})")

    x_values = [positions[c - 1] for c in cols]
    apply(args.src, args.out, args.slide, x_values, sol2)


if __name__ == "__main__":
    main()
