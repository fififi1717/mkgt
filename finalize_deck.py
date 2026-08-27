#!/usr/bin/env python3
"""
Affilior — Finalisation post-remplissage d'un dossier (Étape 5, après le
remplissage de contenu qui suit assemble.py).

Chaîne 3 contrôles/corrections issus du crash-test du 27/08/2026, dans cet
ordre (chacun opère sur le fichier produit par le précédent) :

  1. remove_unused_slots  — supprime réellement les formes des slots vides
                             (règle 5.2quater) plutôt que de vider leur texte.
  2. shrink_oversized_text — réduit automatiquement (paliers ~30%, 2 max) le
                             texte qui déborde des zones "gros chiffre" ;
                             signale ⚠ ce qui reste illisible plutôt que de
                             deviner davantage.
  3. check_montants        — contrôle NON BLOQUANT de traçabilité : tout
                             montant du deck introuvable dans les sources
                             déclarées (chiffres_source) est signalé, jamais
                             bloqué (la validation métier reste à l'associé
                             IP — cf. décision du 27/08/2026).

Rien ici ne remplace la QA visuelle (planche contact, §5.6) ni la validation
associé IP (R4). Ce sont des filets de sécurité mécaniques pour les erreurs
que les runs précédents ont montré comme non détectées par validate.py seul.

Usage :
    python3 finalize_deck.py deck.pptx --out deck_final.pptx \
        --template-positions template_positions.json \
        --dossier dossier_client.json

template_positions.json est GÉNÉRIQUE — commun à tous les clients, ne change
que si Cans_Mstr.pptx change de version. Contient les coordonnées XML des
slots variables et la liste fixe des slides personnalisées.

dossier_client.json est PROPRE À CE CLIENT — à reconstruire à chaque dossier :
{
  "nb_elements_reels": {"12": 2, "13": 2, "15": 2},
  "chiffres_source": [5880000, 225783, ...]
}
"""
import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import remove_unused_slots
import shrink_oversized_text
import check_montants
import select_slots


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template-positions", required=True,
                     help="Référentiel générique (identique pour tous les clients, "
                          "cf. template_positions.json)")
    ap.add_argument("--dossier", required=True,
                     help="Fichier propre à CE client (nb_elements_reels + chiffres_source)")
    ap.add_argument("--tolerance", type=int, default=0)
    args = ap.parse_args()

    template_positions = json.load(open(args.template_positions, encoding="utf-8"))
    dossier = json.load(open(args.dossier, encoding="utf-8"))

    slots = select_slots.compute_slots_a_supprimer(
        template_positions, dossier.get("nb_elements_reels", {})
    )
    slides_check = template_positions["slides_personnalisees"]
    chiffres_source = set(dossier.get("chiffres_source", []))

    with tempfile.TemporaryDirectory() as tmp:
        step1 = str(Path(tmp) / "step1_slots.pptx")
        step2 = str(Path(tmp) / "step2_shrink.pptx")

        print("=== [1/3] Suppression des slots vides (5.2quater) ===")
        if slots:
            remove_unused_slots.apply(args.src, step1, slots)
        else:
            shutil.copy(args.src, step1)
            print("Aucun slot déclaré à supprimer — étape ignorée.")

        print("\n=== [2/3] Réduction des zones en débordement ===")
        report_shrink = shrink_oversized_text.apply(step1, step2, slides_check)

        shutil.copy(step2, args.out)
        print(f"\n-> Fichier finalisé : {args.out}")

        print("\n=== [3/3] Contrôle de traçabilité des montants (non bloquant) ===")
        allowed_slides = {f"ppt/slides/slide{n}.xml" for n in slides_check}
        found = [
            (s, r, v)
            for (s, r, v) in check_montants.extract_amounts_from_pptx(args.out)
            if s in allowed_slides
        ]
        seen = set()
        orphans = []
        for slide, raw, val in found:
            key = (slide, val)
            if key in seen:
                continue
            seen.add(key)
            if not any(abs(val - r) <= args.tolerance for r in chiffres_source):
                orphans.append((slide, raw, val))

        print(f"Montants distincts contrôlés : {len(seen)} | Sources tracées : {len(chiffres_source)}")
        if orphans:
            print(f"⚠ {len(orphans)} montant(s) non tracé(s) — à vérifier avant diffusion :")
            for slide, raw, val in orphans:
                print(f"  ⚠ {slide.replace('ppt/slides/', ''):15s} {raw:>15s}")
        else:
            print("✓ Tous les montants contrôlés sont tracés à une source déclarée.")

        print("\n=== Résumé pour le message de livraison (§5.7) ===")
        unresolved = [r for r in report_shrink if not r[4]]
        if unresolved:
            print(f"⚠ {len(unresolved)} zone(s) toujours en débordement après réduction automatique "
                  f"— vérification visuelle obligatoire avant envoi.")
        if orphans:
            print(f"⚠ {len(orphans)} montant(s) non tracé(s) à une source — à confirmer.")
        if not unresolved and not orphans:
            print("✓ Aucun point mécanique résiduel. QA visuelle (§5.6) et validation associé IP (R4) restent requises.")


if __name__ == "__main__":
    main()
