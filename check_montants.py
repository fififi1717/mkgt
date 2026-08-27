#!/usr/bin/env python3
"""
Affilior — Contrôle de cohérence des montants (non bloquant).
Compare tous les montants en euros trouvés dans le PPTX final à une liste de
valeurs de référence tracées (données du CR + calculs explicitement validés
dans la spec). Signale (⚠) tout montant "orphelin" — présent dans le deck
mais introuvable dans les sources — sans jamais bloquer la génération.

Usage :
    python3 check_montants.py deck.pptx --spec spec.json

spec.json doit contenir une clé "chiffres_source" (liste de nombres, en euros,
sans espace/symbole) rassemblant :
  - tous les montants extraits du compte rendu (patrimoine, fiscalité, succession)
  - tous les montants validés explicitement par le consultant en Étape 2
Toute somme/pourcentage dérivé (ex. un cumul) doit être ajouté explicitement
à cette liste au moment où il est validé — ce script ne fait aucun calcul,
il vérifie seulement la traçabilité.
"""
import argparse
import json
import re
import sys
import zipfile

AMOUNT_RE = re.compile(r'(\d[\d\s\u202f\u00a0]*)\s*€')

def normalize(s: str) -> int:
    digits = re.sub(r'[^\d]', '', s)
    return int(digits) if digits else None

def extract_amounts_from_pptx(path):
    z = zipfile.ZipFile(path)
    found = []  # (slide_file, raw_text, value)
    for name in sorted(z.namelist()):
        if not (name.startswith('ppt/slides/slide') and name.endswith('.xml')):
            continue
        xml = z.read(name).decode('utf-8', 'ignore')
        texts = re.findall(r'<a:t>([^<]*)</a:t>', xml)
        full = ' '.join(texts)
        for m in AMOUNT_RE.finditer(full):
            val = normalize(m.group(1))
            if val is not None and val > 0:
                found.append((name, m.group(0).strip(), val))
    return found

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx')
    ap.add_argument('--spec', required=True)
    ap.add_argument('--tolerance', type=int, default=0,
                     help="écart absolu en euros toléré pour un arrondi (défaut 0)")
    ap.add_argument('--slides', default='1,4,5,7,8,9,10,12,13,14,15,16,18',
                     help="positions Canvas Master personnalisées à contrôler (par défaut : "
                          "toutes sauf 2/3/6/11/17/19 qui sont fixes) — les slides bibliothèque "
                          "insérées (R8, copiées telles quelles) ne sont JAMAIS contrôlées, "
                          "leurs montants ne viennent pas du dossier client.")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding='utf-8'))
    reference = set(spec.get('chiffres_source', []))
    if not reference:
        print("⚠ spec sans 'chiffres_source' — contrôle non significatif, tout sera signalé.",
              file=sys.stderr)

    allowed_slides = {f"ppt/slides/slide{n}.xml" for n in
                       (int(x) for x in args.slides.split(','))}

    found = [(s, r, v) for (s, r, v) in extract_amounts_from_pptx(args.pptx)
             if s in allowed_slides]

    orphans = []
    seen = set()
    for slide, raw, val in found:
        key = (slide, val)
        if key in seen:
            continue
        seen.add(key)
        if not any(abs(val - r) <= args.tolerance for r in reference):
            orphans.append((slide, raw, val))

    print(f"Montants détectés dans le deck : {len(seen)} occurrences distinctes (slide, valeur)")
    print(f"Valeurs de référence tracées   : {len(reference)}")
    print()
    if not orphans:
        print("✓ Aucun montant orphelin détecté — tous les montants du deck sont traçables à une source déclarée.")
        return 0

    print(f"⚠ {len(orphans)} montant(s) présents dans le deck mais introuvables dans les sources déclarées "
          f"(non bloquant — à vérifier avant diffusion) :")
    for slide, raw, val in orphans:
        slide_short = slide.replace('ppt/slides/', '')
        print(f"  ⚠ {slide_short:20s} {raw:>15s}  (valeur normalisée : {val} €)")
    print()
    print("Ce contrôle ne bloque jamais la génération : il signale, il ne juge pas la validité "
          "métier du chiffre. Un montant orphelin peut être légitime (ex. un cumul calculé et "
          "validé mais pas encore ajouté à 'chiffres_source') — mais il peut aussi être un "
          "montant halluciné (cf. crash-test 27/08/2026, rachat AV). À trancher par le consultant.")
    return 0  # non bloquant par construction — code retour toujours 0

if __name__ == '__main__':
    sys.exit(main())
