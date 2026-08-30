#!/usr/bin/env python3
"""
fix_footer_date.py — Corrige le résidu de template confirmé le 30/08/2026 :
le footer '[MOIS ANNÉE]' n'était remplacé QUE sur les slides Canvas explicitement
traitées par le remplissage texte (slides personnalisées de template_positions.json).
Les slides bibliothèque insérées R8 "telles quelles" portent leur PROPRE footer
'[MOIS ANNÉE]' jamais substitué, de même que des slides Canvas non couvertes
par la liste slides_personnalisees (ex. slide 2, structurelle mais avec footer).

Ce script s'applique après assemblage (assemble.py), sur TOUTES les slideN.xml
du paquet final, sans distinction Canvas/bibliothèque — le footer n'est jamais
une donnée métier variable par slide, une seule substitution globale suffit.

Usage : python3 fix_footer_date.py IN.pptx OUT.pptx --mois-annee "AOÛT 2026"
"""
import argparse
import re
import shutil
import zipfile
import os


def fix_footer_date(pptx_in, pptx_out, mois_annee):
    tmp_dir = pptx_out + ".__tmp_footer"
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir)
    with zipfile.ZipFile(pptx_in) as z:
        z.extractall(tmp_dir)

    slides_dir = os.path.join(tmp_dir, "ppt", "slides")
    n_fixed = 0
    for fname in sorted(os.listdir(slides_dir)):
        if not (fname.startswith("slide") and fname.endswith(".xml")):
            continue
        path = os.path.join(slides_dir, fname)
        xml = open(path, encoding="utf-8").read()
        if "[MOIS ANNÉE]" in xml:
            xml = xml.replace("[MOIS ANNÉE]", mois_annee)
            open(path, "w", encoding="utf-8").write(xml)
            n_fixed += 1

    if os.path.exists(pptx_out):
        os.remove(pptx_out)
    with zipfile.ZipFile(pptx_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(tmp_dir):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, tmp_dir)
                zf.write(full, arc)
    shutil.rmtree(tmp_dir)
    print(f"OK -> {pptx_out} ({n_fixed} slide(s) corrigée(s))")
    return n_fixed


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx_in")
    ap.add_argument("pptx_out")
    ap.add_argument("--mois-annee", default="AOÛT 2026")
    args = ap.parse_args()
    fix_footer_date(args.pptx_in, args.pptx_out, args.mois_annee)
