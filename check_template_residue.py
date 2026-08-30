#!/usr/bin/env python3
"""
Affilior — Contrôle des résidus de template (non bloquant). Ajouté v4.12
suite au crash-test double consultant du 28/08/2026, qui a révélé trois
défauts "silencieux" (invisibles pour validate.py et check_montants.py,
et facilement manqués en relecture rapide) :
  1. Couleur de police placeholder (C4B5A5) jamais recolorée après saisie
     de contenu réel (R16 du skill) -> rendu "beige sur beige" illisible.
  2. Notes de conception internes laissées visibles dans des shapes du
     Canvas Master (ex. "Adapter le nombre de colonnes...") -> confusion
     et manque de professionnalisme côté client (R15 du skill).
  3. Graphiques natifs (camembert/donut, barres) dont les séries de
     données n'ont jamais été mises à jour -> incohérence visuelle avec
     les pourcentages affichés en légende (R14 du skill).
     ⚠ CE POINT N'EST PLUS COUVERT ICI depuis v4.14 — voir note plus bas,
     couvert par check_charts.py (appelé dans finalize_deck.py, étape 3/4).

Usage :
    python3 check_template_residue.py deck.pptx

Ce script est un CONTRÔLE DE LECTURE SEULE. Il n'écrit jamais dans le
fichier .pptx analysé, ne supprime ni ne modifie aucun shape, aucun texte,
aucune donnée. Il se contente d'imprimer un rapport avec des symboles ⚠ /
✓ / ℹ, exactement comme check_montants.py.

⚠️ IMPORTANT — ne pas confondre "résidu de template" et "champ volontairement
laissé à compléter" :
- Un texte contenant "[À compléter]" ou "[MANQUANT]" est un état normal du
  workflow (§2.0 du skill : le consultant a répondu "je ne sais pas" ou une
  section du compte rendu était absente) — CE N'EST PAS UN BUG. Ce script le
  liste à part, en information (ℹ), jamais dans la même catégorie que les
  vrais résidus (⚠). Il ne le signale même que si --show-pending est passé,
  pour ne pas noyer le rapport avec des états déjà connus et déjà tracés
  ailleurs dans le récapitulatif Étape 4.
- Un texte entre crochets qui NE correspond PAS à un marqueur de workflow
  connu (ex. "[X €]", "[PROJET 1 — intitulé]") est un placeholder de
  template non rempli — cela reste un signal utile (génération incomplète),
  mais SANS PRÉJUGER qu'il faille le supprimer : c'est au consultant de
  décider si ce champ doit être rempli ou volontairement laissé pour un
  usage ultérieur. Ce script ne recommande jamais de suppression.
Correctif v4.14 (28/08/2026) : la vérification des graphiques natifs a été RETIRÉE de ce
script. Un script dédié plus robuste (`check_charts.py`, catégories nommées et non un
tuple positionnel) a été produit en parallèle dans une autre session et fait
maintenant partie du pipeline `finalize_deck.py` (étape [3/4]). Le maintenir ici
en double aurait dupliqué la liste de référence des valeurs de démo avec un
risque de divergence — matérialisé concrètement : la signature Liquidité de
CE script était (65.0, 35.0) alors que la vraie valeur de démo du Canvas
Master est (35.0, 65.0) (Liquide, Illiquide) — bug qui faisait manquer la
moitié des graphiques non mis à jour. cf. check_charts.py pour ce contrôle.
"""
import argparse
import re
import sys
import zipfile

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

PLACEHOLDER_COLOR = "C4B5A5"  # couleur de saisie du template Canvas Master, jamais destinée au rendu final

# CORRECTIF v4.20 (30/08/2026, crash-test double client) : R16 (corrigée v4.19,
# confirmée par Marc) précise explicitement que C4B5A5 N'EST PAS entièrement
# une couleur de placeholder — elle coexiste comme couleur de hiérarchie
# volontaire sur la slide "Synthèse de la stratégie" (cf. SKILL.md, note R16 :
# "coexiste avec le bronze fort... cf. slide 15 où il reste présent
# volontairement"). Ce script n'avait jamais été mis à jour après cette
# correction et flaguait donc un FAUX POSITIF systématique sur cette slide à
# chaque dossier généré — confirmé en crash-test le 30/08/2026 : QA visuelle
# réelle (rendu image) montre un texte parfaitement lisible (contraste correct
# sur les cartes pleines et sur fond crème), pas un résidu "beige sur beige".
# Détection PAR CONTENU plutôt que par numéro de slide : les numéros de
# fichier internes ne correspondent PAS de façon stable à la position
# narrative après assemblage (constaté en crash-test — le contenu "Synthèse"
# peut se retrouver dans un fichier slideN.xml dont N ne correspond ni à sa
# position Canvas Master d'origine (15) ni à son folio affiché). Le repère
# fixe "STRATÉGIE · SYNTHÈSE" (eyebrow de section, identique pour tous les
# clients, jamais traduit ni personnalisé) est présent sur ce fichier quel
# que soit son nom — c'est un ancrage robuste.
SLIDES_INTENTIONAL_C4B5A5_MARKERS = [
    "STRATÉGIE · SYNTHÈSE",
]

# Marqueurs de workflow légitimes (§2.0, §GESTION DES CAS LIMITES du skill) —
# jamais traités comme un résidu, seulement listés à part en --show-pending.
WORKFLOW_MARKERS = ("À compléter", "MANQUANT")

# Phrases de conception interne connues, laissées par erreur dans des shapes
# visibles du Canvas Master (liste fermée, pas un pattern générique — pour
# éviter tout faux positif sur du texte client légitime contenant des mots
# communs comme "adapter" ou "personnaliser").
KNOWN_INTERNAL_PHRASES = [
    "Adapter le nombre de colonnes",
    "Section gauche : contenu fixe Affilior",
    "supprimer les lignes non utilisées",
    "supprimer les colonnes non utilisées",
    "à personnaliser selon le dossier",
    "Illustration — valeurs à personnaliser",
    "5 remarques maximum recommandées",
]

def iter_slide_xml(pptx_path):
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted(
            [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda n: int(re.search(r"slide(\d+)\.xml", n).group(1)),
        )
        for n in names:
            slide_num = int(re.search(r"slide(\d+)\.xml", n).group(1))
            yield slide_num, z.read(n).decode("utf-8", errors="ignore")


def check_placeholder_color(slide_num, xml, report):
    """R16 — cherche des runs de texte non vide encore en couleur C4B5A5.

    CORRECTIF v4.20 : ignore les slides portant un marqueur de contenu connu
    où C4B5A5 est une couleur de hiérarchie volontaire (cf. constante
    SLIDES_INTENTIONAL_C4B5A5_MARKERS ci-dessus et R16 du skill)."""
    if any(marker in xml for marker in SLIDES_INTENTIONAL_C4B5A5_MARKERS):
        return
    for m in re.finditer(
        r'<a:rPr[^>]*>.*?<a:solidFill><a:srgbClr val="' + PLACEHOLDER_COLOR + r'"/>.*?</a:rPr>\s*<a:t>([^<]*)</a:t>',
        xml, re.DOTALL,
    ):
        text = m.group(1).strip()
        if not text:
            continue
        if any(marker in text for marker in WORKFLOW_MARKERS):
            report["pending"].append((slide_num, text))
        else:
            report["color_residue"].append((slide_num, text))


def check_internal_phrases(slide_num, xml, report):
    """R15 — cherche les phrases de conception interne connues."""
    texts = re.findall(r"<a:t>([^<]*)</a:t>", xml)
    joined = " ".join(texts)
    for phrase in KNOWN_INTERNAL_PHRASES:
        if phrase in joined:
            report["internal_phrase"].append((slide_num, phrase))


def check_unfilled_placeholders(slide_num, xml, report):
    """Recense les placeholders entre crochets non reconnus comme marqueurs
    de workflow légitimes — signal de génération incomplète, PAS une
    recommandation de suppression."""
    texts = re.findall(r"<a:t>([^<]*)</a:t>", xml)
    for t in texts:
        if "[" in t and "]" in t and not any(marker in t for marker in WORKFLOW_MARKERS):
            report["unfilled"].append((slide_num, t.strip()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", help="fichier .pptx généré à contrôler")
    ap.add_argument("--show-pending", action="store_true",
                     help="afficher aussi les marqueurs de workflow légitimes ([À compléter], [MANQUANT]) — "
                          "masqués par défaut car déjà tracés en Étape 4, pas un défaut à corriger")
    args = ap.parse_args()

    report = {"color_residue": [], "internal_phrase": [], "unfilled": [], "pending": []}

    for slide_num, xml in iter_slide_xml(args.pptx):
        check_placeholder_color(slide_num, xml, report)
        check_internal_phrases(slide_num, xml, report)
        check_unfilled_placeholders(slide_num, xml, report)

    total_issues = len(report["color_residue"]) + len(report["internal_phrase"])

    print("=== Contrôle des résidus de template (non bloquant, lecture seule) ===")
    print("ℹ Le contrôle des graphiques natifs (R14) est délégué à check_charts.py "
          "(appelé dans finalize_deck.py, étape 3/4) — pas dupliqué ici depuis v4.14.\n")

    if report["color_residue"]:
        print(f"⚠ {len(report['color_residue'])} run(s) en couleur placeholder ({PLACEHOLDER_COLOR}) "
              f"jamais recolorés (R16) :")
        for slide_num, text in report["color_residue"]:
            print(f"   slide{slide_num}: \"{text[:70]}\"")
        print()
    else:
        print(f"✓ Aucun texte rempli encore en couleur placeholder ({PLACEHOLDER_COLOR}).\n")

    if report["internal_phrase"]:
        print(f"⚠ {len(report['internal_phrase'])} note(s) de conception interne visible(s) (R15) :")
        for slide_num, phrase in report["internal_phrase"]:
            print(f"   slide{slide_num}: \"{phrase}\"")
        print()
    else:
        print("✓ Aucune note de conception interne connue détectée.\n")

    if report["unfilled"]:
        print(f"ℹ {len(report['unfilled'])} placeholder(s) entre crochets non reconnu(s) comme marqueur de "
              f"workflow — génération probablement incomplète, à vérifier (PAS une recommandation de suppression) :")
        for slide_num, text in report["unfilled"]:
            print(f"   slide{slide_num}: \"{text[:70]}\"")
        print()

    if args.show_pending and report["pending"]:
        print(f"ℹ {len(report['pending'])} marqueur(s) de workflow légitime ([À compléter]/[MANQUANT]) — "
              f"état normal, déjà tracé en Étape 4, affiché ici pour information uniquement :")
        for slide_num, text in report["pending"]:
            print(f"   slide{slide_num}: \"{text[:70]}\"")
        print()

    print("=== Résumé ===")
    if total_issues == 0:
        print("✓ Aucun résidu de template détecté.")
    else:
        print(f"⚠ {total_issues} résidu(s) de template à vérifier avant diffusion. "
              f"Rappel : ce script ne modifie jamais le fichier — toute correction reste manuelle.")

    return 0  # non bloquant, toujours


if __name__ == "__main__":
    sys.exit(main())
