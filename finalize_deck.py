#!/usr/bin/env python3
"""
Affilior — Finalisation post-remplissage d'un dossier (Étape 5, après le
remplissage de contenu qui suit assemble.py).

Chaîne 4 contrôles/corrections, dans cet ordre (chacun opère sur le fichier
produit par le précédent) :

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
  4. check_charts          — contrôle NON BLOQUANT (+ correction optionnelle
                             via --fix-charts) des graphiques natifs (donuts) :
                             assemble.py ne touche jamais aux séries de
                             graphique, seuls les placeholders texte/tableaux
                             sont remplis — un deck peut donc être livré avec
                             les donuts encore aux valeurs de démo du template
                             pendant que le texte autour affiche les vrais
                             chiffres (bug confirmé crash-test 28/08/2026).

CORRECTIF ORDRE (constat critique 29/08/2026) : check_montants passe désormais
AVANT check_charts (inversion de l'ordre v4.14). Raison : --fix-charts recharge
le deck avec python-pptx (Presentation().save()), ce qui RENUMÉROTE les
fichiers ppt/slides/slideN.xml selon l'ordre visuel de la présentation —
alors qu'assemble.py conservait jusque-là les noms de fichiers d'origine du
Canvas Master indépendamment de l'ordre narratif (§5.2ter). Conséquence
trouvée en crash-test réel (2 clients fictifs, 29/08/2026) : après un
--fix-charts, "slide18.xml" pouvait correspondre à une slide bibliothèque
pédagogique insérée (jamais censée être contrôlée, R8) au lieu de la vraie
slide Canvas "Prochaines étapes" — faux positifs sur des seuils légaux ET
disparition silencieuse du contrôle sur les vraies slides personnalisées.
Le fix-charts ne touchant jamais le texte des slides, contrôler les montants
AVANT cette étape (sur des noms de fichiers encore fiables) est rigoureusement
équivalent et élimine le problème sans toucher à check_montants lui-même.

Rien ici ne remplace la QA visuelle (planche contact, §5.6) ni la validation
associé IP (R4). Ce sont des filets de sécurité mécaniques pour les erreurs
que les runs précédents ont montré comme non détectées par validate.py seul.
Voir aussi check_template_residue.py (script séparé, non chaîné ici) pour
2 défauts supplémentaires du même ordre (couleur placeholder non corrigée,
notes de conception internes visibles) — à exécuter en plus, pas à la place.

Usage :
    python3 finalize_deck.py deck.pptx --out deck_final.pptx \
        --template-positions template_positions.json \
        --dossier dossier_client.json [--fix-charts]

template_positions.json est GÉNÉRIQUE — commun à tous les clients, ne change
que si Cans_Mstr.pptx change de version. Contient les coordonnées XML des
slots variables et la liste fixe des slides personnalisées.

dossier_client.json est PROPRE À CE CLIENT — à reconstruire à chaque dossier :
{
  "nb_elements_reels": {"12": 2, "13": 2, "15": 2},
  "chiffres_source": [5880000, 225783, ...],
  "allocation_pct": {"Immobilier": 45, "Financier": 38, "Retraite": 12, "Liquidités": 5},
  "liquidite_pct": {"Liquide": 20, "Illiquide": 80}
}
allocation_pct et liquidite_pct sont OPTIONNELS (ajoutés v4.14, cf. check_charts.py) :
sans eux, l'étape 3/4 se contente de signaler ⚠ si un graphique est resté en
valeurs de démo, sans pouvoir le corriger automatiquement même avec --fix-charts.

CORRECTIF v4.20 (30/08/2026, crash-test double client) : allocation_pct et
liquidite_pct devaient jusqu'ici être saisis DEUX FOIS à l'identique — une
fois dans plan.json (pour personalize_canvas.py, qui corrige les 6 LABELS
texte de la slide 6) et une seconde fois dans dossier_client.json (pour
check_charts.py, qui corrige le GRAPHIQUE natif). Un oubli dans l'un des deux
reproduit exactement le bug historique R14 : légende correcte, donut resté
aux valeurs de démo — repéré en crash-test le 30/08/2026 sur les 2 dossiers
testés. `--plan` (optionnel, nouveau) permet de fournir une seule fois le
plan.json déjà utilisé par personalize_canvas.py : si allocation_pct/
liquidite_pct sont absents de --dossier, ils sont repris depuis --plan.
S'ils sont présents aux DEUX endroits et diffèrent, un avertissement bloque
la correction automatique plutôt que de choisir silencieusement une source
(cf. merge_allocation_sources ci-dessous).
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
import check_charts
# select_slots.py fusionné dans remove_unused_slots.py le 28/08/2026 (nettoyage dépôt) —
# compute_slots_a_supprimer y est désormais définie directement.


def merge_allocation_sources(dossier, plan, warnings):
    """CORRECTIF v4.20 — fusionne allocation_pct/liquidite_pct de --plan dans
    dossier UNIQUEMENT si absents de dossier (dossier reste prioritaire, pour
    compatibilité avec les dossiers existants qui les fournissent déjà en
    double). Si les deux sources sont présentes et diffèrent, retire la clé
    de dossier plutôt que de choisir arbitrairement : mieux vaut un ⚠ "aucune
    donnée fournie" (déjà géré par check_charts.py) qu'une correction basée
    sur une source silencieusement incohérente."""
    if not plan:
        return dossier
    for key in ("allocation_pct", "liquidite_pct"):
        plan_val = plan.get(key)
        dossier_val = dossier.get(key)
        if dossier_val and plan_val and dossier_val != plan_val:
            warnings.append(
                f"⚠ {key} présent à la fois dans --dossier et --plan avec des valeurs "
                f"différentes — ignoré des deux côtés pour cette correction automatique, "
                f"à corriger en amont (source unique attendue)."
            )
            dossier.pop(key, None)
        elif not dossier_val and plan_val:
            dossier[key] = plan_val
    return dossier


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template-positions", required=True,
                     help="Référentiel générique (identique pour tous les clients, "
                          "cf. template_positions.json)")
    ap.add_argument("--dossier", required=True,
                     help="Fichier propre à CE client (nb_elements_reels + chiffres_source)")
    ap.add_argument("--plan", default=None,
                     help="Optionnel (v4.20) : plan.json déjà utilisé pour personalize_canvas.py. "
                          "Si allocation_pct/liquidite_pct sont absents de --dossier, ils sont repris "
                          "depuis --plan pour éviter la double saisie (cf. correctif v4.20 ci-dessus).")
    ap.add_argument("--tolerance", type=int, default=0)
    ap.add_argument("--fix-charts", action="store_true",
                     help="Corrige automatiquement les graphiques natifs (donuts) restés aux "
                          "valeurs de démo, si allocation_pct/liquidite_pct sont fournis dans "
                          "--dossier et sans ambiguïté (cf. check_charts.py).")
    args = ap.parse_args()

    template_positions = json.load(open(args.template_positions, encoding="utf-8"))
    dossier = json.load(open(args.dossier, encoding="utf-8"))
    plan = json.load(open(args.plan, encoding="utf-8")) if args.plan else None
    allocation_warnings = []
    dossier = merge_allocation_sources(dossier, plan, allocation_warnings)
    for w in allocation_warnings:
        print(w)

    # CORRECTIF (bug C, crash-test #3 27/08/2026 — toujours présent avant ce
    # correctif, confirmé en direct le 29/08/2026 lors du crash-test à deux
    # clients) : compute_slots_a_supprimer lève un ValueError brut (dépassement
    # de plafond 12/13/15, ou nb_elements_reels négatif) qui remontait jusqu'ici
    # en traceback Python complet sur stderr — pas de message clair pour le
    # consultant, contrairement au style die() d'assemble.py. Encadré désormais.
    try:
        slots = remove_unused_slots.compute_slots_a_supprimer(
            template_positions, dossier.get("nb_elements_reels", {})
        )
    except ValueError as e:
        print(f"Erreur : {e}", file=sys.stderr)
        print("Erreur : dossier_client.json invalide (nb_elements_reels) — "
              "corriger la valeur en amont (Étape 2) avant de relancer la génération. "
              "Aucun fichier n'a été produit.", file=sys.stderr)
        sys.exit(1)

    slides_check = template_positions["slides_personnalisees"]
    chiffres_source = set(dossier.get("chiffres_source", []))

    with tempfile.TemporaryDirectory() as tmp:
        step1 = str(Path(tmp) / "step1_slots.pptx")
        step2 = str(Path(tmp) / "step2_shrink.pptx")

        print("=== [1/4] Suppression des slots vides (5.2quater) ===")
        if slots:
            remove_unused_slots.apply(args.src, step1, slots)
        else:
            shutil.copy(args.src, step1)
            print("Aucun slot déclaré à supprimer — étape ignorée.")

        print("\n=== [2/4] Réduction des zones en débordement ===")
        report_shrink = shrink_oversized_text.apply(step1, step2, slides_check)

        # CORRECTIF (constat critique 29/08/2026) : le contrôle de traçabilité
        # des montants (étape [4/4] ci-dessous) tourne maintenant sur `step2`
        # — c'est-à-dire AVANT toute correction de graphique — et non plus sur
        # le fichier final. Raison : la correction --fix-charts recharge le
        # deck avec python-pptx (Presentation(...).save(...)), qui RENUMÉROTE
        # les fichiers ppt/slides/slideN.xml selon l'ORDRE VISUEL de la
        # présentation, alors qu'assemble.py conservait jusque-là les noms de
        # fichiers d'origine du Canvas Master indépendamment de l'ordre
        # narratif. Résultat observé en crash-test réel (2 clients, 29/08/2026) :
        # après un --fix-charts, "ppt/slides/slide18.xml" pouvait correspondre
        # à une slide bibliothèque pédagogique insérée (jamais censée être
        # contrôlée, R8) au lieu de la vraie slide Canvas "Prochaines étapes"
        # — génère de faux positifs sur des seuils légaux ET fait disparaître
        # silencieusement le contrôle sur les vraies slides personnalisées.
        # Le fix-charts ne touche JAMAIS le texte des slides (uniquement les
        # séries de graphique), donc les montants à contrôler sont rigoureu-
        # sement identiques avant/après — faire le contrôle sur step2 (noms de
        # fichiers encore fiables) élimine le problème sans toucher à la
        # logique de check_montants elle-même.
        print("\n=== [3/4] Contrôle de traçabilité des montants (non bloquant) ===")
        allowed_slides = {f"ppt/slides/slide{n}.xml" for n in slides_check}
        found = [
            (s, r, v)
            for (s, r, v) in check_montants.extract_amounts_from_pptx(step2)
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

        # --- [4/4] Graphiques natifs (donuts) — après les montants (cf. correctif
        # ci-dessus) puisque cette étape est la seule à faire un aller-retour
        # python-pptx susceptible de renuméroter les fichiers slideN.xml.
        step3 = str(Path(tmp) / "step3_charts.pptx")
        print("\n=== [4/4] Contrôle des graphiques natifs (donuts, non bloquant) ===")
        chart_alerts, chart_fixes = check_charts.check(step2, dossier)
        chart_fixed_slides = set()
        if chart_alerts:
            for slide_n, name, cats, vals, msg in chart_alerts:
                print(f"  slide {slide_n:<3} série « {name} » {dict(zip(cats, vals))}")
                print(f"    {msg}")
        else:
            print("✓ Aucun graphique laissé aux valeurs de démonstration du template.")

        if args.fix_charts and chart_fixes:
            # CORRECTIF (déduplication, constat 29/08/2026) : appelle désormais
            # check_charts.apply_fixes(), implémentation unique partagée avec
            # la CLI check_charts.py --fix (auparavant recopiée ici à l'identique,
            # avec le risque de divergence silencieuse entre les deux versions).
            applied, per_series = check_charts.apply_fixes(step2, dossier, step3)
            fully_fixed_slides = set()
            partially_left_slides = set()
            for slide_n, name, ok, reason in per_series:
                if ok:
                    fully_fixed_slides.add(slide_n)
                else:
                    partially_left_slides.add(slide_n)
                    print(f"  ⚠ slide {slide_n} « {name} » NON corrigé — {reason}")
            chart_fixed_slides = fully_fixed_slides
            # CORRECTIF (constat mineur 29/08/2026) : message désormais précis
            # PAR SÉRIE — auparavant "slide(s) corrigée(s)" pouvait laisser croire
            # que TOUS les graphiques d'une slide étaient corrigés alors qu'un
            # seul des deux donuts pouvait l'être (cas d'une répartition fournie
            # mais ambiguë à côté d'une répartition valide sur la même slide).
            if applied:
                print(f"✓ {applied} graphique(s) corrigé(s) automatiquement "
                      f"(slide(s) {sorted(fully_fixed_slides)}).")
            ambiguous_only_slides = partially_left_slides - fully_fixed_slides
            if ambiguous_only_slides:
                print(f"⚠ slide(s) {sorted(ambiguous_only_slides)} : au moins un graphique "
                      f"reste en valeurs de démo, non corrigé — voir détail ci-dessus.")
        else:
            shutil.copy(step2, step3)
            if chart_fixes and not args.fix_charts:
                print("  (correction disponible mais --fix-charts non passé — donuts laissés tels quels)")

        shutil.copy(step3, args.out)
        print(f"\n-> Fichier finalisé : {args.out}")

        print("\n=== Résumé pour le message de livraison (§5.7) ===")
        unresolved = [r for r in report_shrink if not r[4]]
        if unresolved:
            print(f"⚠ {len(unresolved)} zone(s) toujours en débordement après réduction automatique "
                  f"— vérification visuelle obligatoire avant envoi.")
        if orphans:
            print(f"⚠ {len(orphans)} montant(s) non tracé(s) à une source — à confirmer.")
        if chart_alerts and not chart_fixed_slides:
            print(f"⚠ {len(chart_alerts)} graphique(s) natif(s) resté(s) en valeurs de démo — non corrigé(s).")
        if not unresolved and not orphans and not (chart_alerts and not chart_fixed_slides):
            print("✓ Aucun point mécanique résiduel. QA visuelle (§5.6) et validation associé IP (R4) restent requises.")

        # --- Périmètre restreint pour la QA visuelle (§5.6) ---
        # Seules les slides effectivement touchées par ce run ont besoin d'être
        # rendues en image pour vérification humaine — pas les 28 slides du
        # deck. Évite un rendu soffice/pdftoppm complet à chaque run.
        import re as _re
        touched_slides = set(slides_check) | chart_fixed_slides | {
            int(_re.search(r"slide(\d+)\.xml$", s).group(1)) for s, _, _ in found
        }
        if touched_slides:
            page_list = ",".join(str(n) for n in sorted(touched_slides))
            print(f"\n=== Périmètre QA visuelle recommandé (§5.6) ===")
            print(f"Slides à rendre en image (pas le deck entier) : {page_list}")
            print(f"  soffice --headless --convert-to pdf {args.out}")
            print(f"  pdftoppm -jpeg -r 150 -f <min> -l <max> {Path(args.out).stem}.pdf apercu")
            print("  (ou extraire page par page avec -f N -l N pour chaque slide de la liste ci-dessus)")


if __name__ == "__main__":
    main()
