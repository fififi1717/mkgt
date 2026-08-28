"""
assemble.py — Affilior, Étape 5 (génération PPTX), pipeline versionné.

Remplace la ré-écriture manuelle du §5.1/§5.2ter/§5bis à chaque session Claude.
Corrige les deux écarts identifiés lors du crash-test du 27/08/2026 :
  1. [Content_Types].xml : scan générique des extensions média réellement
     présentes dans le paquet final (pas de liste d'extensions codée en dur) ;
  2. Réordonnancement narratif (§5.2ter) appliqué automatiquement à partir
     de index.json, plutôt que laissé à une ré-interprétation par session.

Usage :
    python3 assemble.py --spec client_spec.json --out DOSSIER.pptx

client_spec.json attend :
{
  "canvas": "mkgt/Cans_Mstr.pptx",
  "bibliotheque": "mkgt/Bibe_Def.pptx",
  "index": "mkgt/index.json",
  "slides_bibliotheque": [5, 6, 25, 26, 27, 28, 33, 34, 7, 29, 35, 38, 48, 51, 52],
  "exclude_canvas": [19],
  "_note_exclude_canvas": "Optionnel (ajouté v4.12). Liste de numéros de slides Canvas Master (1-19) à exclure de "
                          "l'assemblage, typiquement pour éviter un doublon avec une slide bibliothèque équivalente "
                          "(ex. slide bibliothèque 49 'Barème succession' rend la Canvas 19 'Annexe barème détaillé' "
                          "redondante — cf. index.json, règle systématique 'slide_49_incluse'). Ne PAS utiliser pour "
                          "retirer une slide narrative structurante (audit, diagnostic, stratégie, impact, synthèse) : "
                          "R7 du skill interdit toujours la suppression arbitraire d'une slide Canvas Master par le "
                          "consultant — ce champ sert uniquement à appliquer une règle d'anti-doublon déjà validée.",
  "patrimoine": {
     "familles": [
        {"nom": "ASSURANCE VIE", "actifs": [
            {"nom": "Contrat d'assurance vie", "detention": "PP", "proprietaire": "Cl.T.", "montant": 450000},
            ...
        ]}
     ],
     "total_net": 5880000
  }
}

Ne fait AUCUNE opération réseau, ne touche jamais [Content_Types].xml d'un
dépôt distant — travaille uniquement sur des fichiers déjà clonés localement.
"""
import argparse
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

# ---------------------------------------------------------------------------
# Constantes de style — Charte Affilior (§5.4 du skill). Câblées en dur :
# aucune slide générée par ce script ne peut dériver de la charte, puisque
# ces valeurs ne sont jamais réinterprétées, seulement référencées.
# ---------------------------------------------------------------------------
COLOR_FOND = "F3F0EB"          # crème
COLOR_NAVY = "1C2B3A"          # titres, chiffres, bandeau total
COLOR_BORDEAUX = "8B1A1A"      # accents, section headers, en-têtes de famille
COLOR_CORPS = "4A5568"         # corps de texte
COLOR_LABEL = "A09080"         # labels discrets, "COMMUN."
COLOR_LIGNE = "CFC8BD"         # filets
COLOR_LIGNE_CLAIRE = "E5E0D8"
COLOR_TABLEAU_ALT1 = "F7F5F1"
COLOR_TABLEAU_ALT2 = "F3F0EB"
FONT_TITRE = "Garamond"        # titres, section headers, chiffres clés
FONT_CORPS = "Calibri"         # corps de texte
MENTION_MIF2 = "Ce document ne constitue pas un conseil en investissement au sens MIF2"
MENTION_PERF = "Les performances passées ne préjugent pas des performances futures."
MENTION_PROJECTION = ("Projections établies sur la base des éléments communiqués. "
                       "Ces estimations conditionnelles ne sauraient constituer un engagement.")


def die(msg):
    print(f"Erreur : {msg}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# 1. Copie cross-fichier générique (corrige le bug Content_Types du 27/08)
# ---------------------------------------------------------------------------
def insert_library_slides(tmp_dir, biblio_tmp_dir, positions):
    slides_dir = os.path.join(tmp_dir, "ppt", "slides")
    existing_nums = sorted(
        int(re.match(r"slide(\d+)\.xml", f).group(1))
        for f in os.listdir(slides_dir) if re.match(r"slide(\d+)\.xml", f)
    )
    next_num = max(existing_nums) + 1 if existing_nums else 1

    ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
    pres_rels_path = os.path.join(tmp_dir, "ppt", "_rels", "presentation.xml.rels")
    pres_path = os.path.join(tmp_dir, "ppt", "presentation.xml")

    ct = open(ct_path, encoding="utf-8").read()
    pres_rels = open(pres_rels_path, encoding="utf-8").read()
    pres = open(pres_path, encoding="utf-8").read()

    used_rids = {int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels)}
    used_slideids = {int(m) for m in re.findall(r'<p:sldId[^>]*\bid="(\d+)"', pres)}
    next_rid = max(used_rids) + 1 if used_rids else 1
    next_slideid = max(used_slideids) + 1 if used_slideids else 256

    inserted = []
    media_counter = 0

    for pos in positions:
        src_slide = os.path.join(biblio_tmp_dir, "ppt", "slides", f"slide{pos}.xml")
        if not os.path.exists(src_slide):
            die(f"slide{pos}.xml absente de la bibliothèque clonée — dépôt désynchronisé ? "
                f"Relancer la vérification §3.0 avant de poursuivre.")
        dest_name = f"slide{next_num}.xml"
        shutil.copy2(src_slide, os.path.join(slides_dir, dest_name))

        src_rels = os.path.join(biblio_tmp_dir, "ppt", "slides", "_rels", f"slide{pos}.xml.rels")
        if os.path.exists(src_rels):
            rels_xml = open(src_rels, encoding="utf-8").read()
            for m in re.finditer(r'Target="\.\./media/([^"]+)"', rels_xml):
                fname = m.group(1)
                ext = fname.rsplit(".", 1)[-1]
                media_counter += 1
                new_fname = f"libimg_{pos}_{media_counter}.{ext}"
                src_media = os.path.join(biblio_tmp_dir, "ppt", "media", fname)
                dst_media_dir = os.path.join(tmp_dir, "ppt", "media")
                os.makedirs(dst_media_dir, exist_ok=True)
                if os.path.exists(src_media):
                    shutil.copy2(src_media, os.path.join(dst_media_dir, new_fname))
                rels_xml = rels_xml.replace(f"../media/{fname}", f"../media/{new_fname}")
            rels_dir = os.path.join(slides_dir, "_rels")
            os.makedirs(rels_dir, exist_ok=True)
            open(os.path.join(rels_dir, f"{dest_name}.rels"), "w", encoding="utf-8").write(rels_xml)

        override = (f'<Override PartName="/ppt/slides/{dest_name}" '
                    f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
        if f'/ppt/slides/{dest_name}"' not in ct:
            ct = ct.replace("</Types>", f"{override}</Types>")

        rel = (f'<Relationship Id="rId{next_rid}" '
               f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
               f'Target="slides/{dest_name}"/>')
        pres_rels = pres_rels.replace("</Relationships>", f"{rel}</Relationships>")

        inserted.append({"pos": pos, "rid": next_rid, "sldid": next_slideid, "name": dest_name})
        next_rid += 1
        next_slideid += 1
        next_num += 1

    entries_xml = "".join(f'<p:sldId id="{e["sldid"]}" r:id="rId{e["rid"]}"/>' for e in inserted)
    pres = pres.replace("</p:sldIdLst>", f"{entries_xml}</p:sldIdLst>")

    # --- Correction du bug du 27/08 : scan générique, pas de liste en dur ---
    media_dir = os.path.join(tmp_dir, "ppt", "media")
    declared_exts = set(re.findall(r'Default Extension="([^"]+)"', ct))
    present_exts = set()
    if os.path.isdir(media_dir):
        for f in os.listdir(media_dir):
            if "." in f:
                present_exts.add(f.rsplit(".", 1)[-1].lower())
    missing = present_exts - declared_exts
    for ext in sorted(missing):
        mime, _ = mimetypes.guess_type(f"x.{ext}")
        mime = mime or "application/octet-stream"
        ct = ct.replace("</Types>", f'<Default Extension="{ext}" ContentType="{mime}"/></Types>')
        print(f"  [content-types] extension '{ext}' manquante ajoutée automatiquement ({mime})")

    open(ct_path, "w", encoding="utf-8").write(ct)
    open(pres_rels_path, "w", encoding="utf-8").write(pres_rels)
    open(pres_path, "w", encoding="utf-8").write(pres)
    return inserted


# ---------------------------------------------------------------------------
# 2. Réordonnancement narratif (§5.2ter) — appliqué depuis index.json,
#    plus jamais laissé à une ré-interprétation manuelle par session.
# ---------------------------------------------------------------------------
CANVAS_ORDER_BLOCK_A = list(range(1, 13))       # 1-12 : audit + diagnostic + projets
CANVAS_MOYENS = 13
CANVAS_SYNTHESE = 15
CANVAS_IMPACT = 14
CANVAS_AVANT_APRES = 16
CANVAS_ACCOMPAGNEMENT = 17
CANVAS_PROCHAINES_ETAPES = 18
CANVAS_ANNEXE_BAREME = 19


def reorder_narratif(tmp_dir, canvas_slide_count, inserted, index_data, exclude_canvas=None):
    """Réordonne <p:sldId> selon §5.2ter. Les positions Canvas 1..canvas_slide_count
    sont les N premiers <p:sldId> d'origine (ordre physique source, inchangé pour
    elles) ; les slides insérées (inserted, dans l'ordre de la sélection Étape 3)
    sont réparties en 2 groupes : pédagogique/conviction (juste après Impact) et
    tarification/annexes (juste avant l'annexe barème).

    exclude_canvas (ajouté v4.12) : ensemble optionnel de numéros de slides Canvas
    Master (1-19) à omettre du réordonnancement final — la slide reste physiquement
    présente à ce stade (retirée ensuite par cleanup_excluded_canvas_slides), on ne
    fait ici que ne pas l'inclure dans la liste ordonnée. Sert à l'anti-doublon
    bibliothèque/Canvas (ex. slide 49 vs Canvas 19), jamais à une suppression libre
    d'une slide narrative structurante (R7 du skill)."""
    exclude_canvas = exclude_canvas or set()
    pres_path = os.path.join(tmp_dir, "ppt", "presentation.xml")
    pres = open(pres_path, encoding="utf-8").read()

    m = re.search(r"<p:sldIdLst>(.*)</p:sldIdLst>", pres, re.DOTALL)
    all_entries = re.findall(r'<p:sldId\b[^>]*/>', m.group(1))
    if len(all_entries) != canvas_slide_count + len(inserted):
        die("incohérence : nombre d'entrées sldIdLst != canvas + bibliothèque insérées")

    canvas_entries = all_entries[:canvas_slide_count]   # positions 1..19, ordre physique = ordre logique 1-19 ici
    lib_entries = all_entries[canvas_slide_count:]       # dans l'ordre où insert_library_slides les a ajoutées

    lib_by_pos = {e["pos"]: entry for e, entry in zip(inserted, lib_entries)}
    idx_by_pos = {s["pos"]: s for s in index_data["slides"]}

    pedago_conviction = []
    tarification_annexes = []
    for e in inserted:
        pos = e["pos"]
        role = idx_by_pos.get(pos, {}).get("role", "")
        entry_xml = lib_by_pos[pos]
        if role in ("tarification",):
            tarification_annexes.append(entry_xml)
        elif idx_by_pos.get(pos, {}).get("categorie") == "Transmission / succession":
            tarification_annexes.append(entry_xml)
        else:
            pedago_conviction.append(entry_xml)

    def canvas(n):
        return canvas_entries[n - 1]

    def canvas_if_kept(n):
        """Retourne [entry] normalement, [] si n est exclu (v4.12)."""
        if n in exclude_canvas:
            return []
        return [canvas(n)]

    ordered = []
    ordered += [canvas(n) for n in CANVAS_ORDER_BLOCK_A if n not in exclude_canvas]
    ordered += canvas_if_kept(CANVAS_MOYENS)
    ordered += canvas_if_kept(CANVAS_SYNTHESE)
    ordered += canvas_if_kept(CANVAS_IMPACT)
    ordered += pedago_conviction
    ordered += canvas_if_kept(CANVAS_AVANT_APRES)
    ordered += canvas_if_kept(CANVAS_ACCOMPAGNEMENT)
    ordered += canvas_if_kept(CANVAS_PROCHAINES_ETAPES)
    ordered += tarification_annexes
    ordered += canvas_if_kept(CANVAS_ANNEXE_BAREME)

    expected_len = len(all_entries) - len(exclude_canvas)
    if len(ordered) != expected_len:
        die(f"réordonnancement incomplet : {len(ordered)} entrées produites, {expected_len} attendues "
            f"(exclude_canvas={sorted(exclude_canvas)})")

    excluded_rids = [re.search(r'r:id="(rId\d+)"', canvas(n)).group(1) for n in sorted(exclude_canvas)]

    new_lst = "<p:sldIdLst>" + "".join(ordered) + "</p:sldIdLst>"
    pres = pres[:m.start()] + new_lst + pres[m.end():]
    open(pres_path, "w", encoding="utf-8").write(pres)
    return [re.search(r'r:id="(rId\d+)"', e).group(1) for e in ordered], excluded_rids


def cleanup_excluded_canvas_slides(tmp_dir, exclude_canvas, excluded_rids):
    """Supprime proprement (v4.12, corrigé après échec validate.py en test) les
    parties du paquet .pptx correspondant aux slides Canvas exclues par
    exclude_canvas : fichier slideN.xml, ses _rels, ses dépendances (notesSlide,
    médias non partagés), l'Override dans [Content_Types].xml, et la Relationship
    dans ppt/_rels/presentation.xml.rels. Appelé APRÈS reorder_narratif — la slide
    est déjà absente de sldIdLst à ce stade, ceci ne fait que retirer les fichiers
    devenus orphelins plutôt que de laisser des parties non référencées dans le zip.

    v4.12 correctif : la première version utilisait `[^/]*` dans les regex de
    nettoyage, qui s'arrête au premier '/' rencontré dans l'URL du Target ou du
    ContentType (ex. "http://schemas...") — les Relationship et Override n'étaient
    donc jamais réellement supprimés (match partiel silencieux). Remplacé par
    `[^>]*?` (aucun '>' dans les attributs XML, c'est la seule garantie fiable).
    Découvert par exécution réelle de validate.py en test, pas par relecture."""
    if not exclude_canvas:
        return
    slides_dir = os.path.join(tmp_dir, "ppt", "slides")
    rels_dir = os.path.join(slides_dir, "_rels")
    media_dir = os.path.join(tmp_dir, "ppt", "media")
    notes_dir = os.path.join(tmp_dir, "ppt", "notesSlides")
    notes_rels_dir = os.path.join(notes_dir, "_rels")
    ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
    pres_rels_path = os.path.join(tmp_dir, "ppt", "_rels", "presentation.xml.rels")

    ct = open(ct_path, encoding="utf-8").read()
    pres_rels = open(pres_rels_path, encoding="utf-8").read()

    for n in sorted(exclude_canvas):
        slide_name = f"slide{n}.xml"
        slide_path = os.path.join(slides_dir, slide_name)
        slide_rels_path = os.path.join(rels_dir, f"{slide_name}.rels")

        # Découvrir les dépendances (notesSlide, médias) AVANT de supprimer la slide
        dependent_targets = []
        if os.path.exists(slide_rels_path):
            rels_xml = open(slide_rels_path, encoding="utf-8").read()
            dependent_targets = re.findall(r'Target="([^"]+)"', rels_xml)

        if os.path.exists(slide_path):
            os.remove(slide_path)
        if os.path.exists(slide_rels_path):
            os.remove(slide_rels_path)
        ct = re.sub(rf'<Override PartName="/ppt/slides/{slide_name}"[^>]*?/>', "", ct)

        for target in dependent_targets:
            # Targets relatifs type "../notesSlides/notesSlideN.xml" ou "../media/xxx.png"
            norm = target.replace("../", "")
            dep_path = os.path.join(tmp_dir, "ppt", norm)
            if not os.path.exists(dep_path):
                continue
            if "notesSlides/" in norm and norm.endswith(".xml"):
                notes_name = os.path.basename(norm)
                os.remove(dep_path)
                notes_rels_path = os.path.join(notes_rels_dir, f"{notes_name}.rels")
                if os.path.exists(notes_rels_path):
                    os.remove(notes_rels_path)
                ct = re.sub(rf'<Override PartName="/ppt/notesSlides/{notes_name}"[^>]*?/>', "", ct)
                print(f"  [exclude_canvas] notesSlide dépendante supprimée : {notes_name}")
            elif "media/" in norm:
                # Ne supprimer le média que s'il n'est référencé par AUCUNE autre slide restante
                media_name = os.path.basename(norm)
                still_used = False
                for f in os.listdir(rels_dir) if os.path.isdir(rels_dir) else []:
                    other_rels = open(os.path.join(rels_dir, f), encoding="utf-8").read()
                    if media_name in other_rels:
                        still_used = True
                        break
                if not still_used:
                    os.remove(dep_path)
                    print(f"  [exclude_canvas] média orphelin supprimé : {media_name}")

    for rid in excluded_rids:
        pres_rels = re.sub(rf'<Relationship Id="{rid}"[^>]*?/>', "", pres_rels)

    open(ct_path, "w", encoding="utf-8").write(ct)
    open(pres_rels_path, "w", encoding="utf-8").write(pres_rels)
    print(f"  [exclude_canvas] {len(exclude_canvas)} slide(s) Canvas retirée(s) proprement : "
          f"{sorted(exclude_canvas)}")


# ---------------------------------------------------------------------------
# 3. Renumérotation des folios Canvas (§5bis)
# ---------------------------------------------------------------------------
def renumber_canvas_folios(tmp_dir, ordered_rids, canvas_folio_map):
    """canvas_folio_map: {rid_canvas_source (str) -> ancien_numero_texte_fige (str)}
    CORRECTIF 27/08/2026 (crash-test) : la docstring précédente indiquait l'ordre
    inverse (ancien_numero -> rid), incohérent avec la boucle ci-dessous
    (`for rid, old_folio in canvas_folio_map.items()`), qui attend rid en clé.
    Avec l'ancien ordre, un canvas_folio_map construit "à la lettre" de la
    docstring ne levait aucune erreur mais ne renumérotait RIEN (aucun match
    de pattern, car old_folio == une valeur rId qui n'apparaît jamais dans un
    <a:t>) — échec silencieux. Corrigé ici pour que docstring et code concordent.
    Remplace le <a:t>N</a:t> figé de chaque slide Canvas par sa position réelle
    dans le deck final (1-indexée), calculée à partir de ordered_rids."""
    pres_rels_path = os.path.join(tmp_dir, "ppt", "_rels", "presentation.xml.rels")
    rels = open(pres_rels_path, encoding="utf-8").read()
    rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/([^"]+)"', rels))

    position_of_rid = {rid: i + 1 for i, rid in enumerate(ordered_rids)}
    updated = []
    for rid, old_folio in canvas_folio_map.items():
        new_folio = position_of_rid.get(rid)
        if new_folio is None:
            continue
        slide_file = rid_to_target.get(rid)
        if not slide_file:
            continue
        path = os.path.join(tmp_dir, "ppt", "slides", slide_file)
        xml = open(path, encoding="utf-8").read()
        pattern = f"<a:t>{old_folio}</a:t>"
        replacement = f"<a:t>{new_folio}</a:t>"
        if pattern in xml:
            xml = xml.replace(pattern, replacement, 1)
            open(path, "w", encoding="utf-8").write(xml)
            updated.append((slide_file, old_folio, new_folio))
    return updated


# ---------------------------------------------------------------------------
# 3bis. Construction automatique de canvas_folio_map (CORRECTIF 27/08/2026)
#    Avant ce correctif : canvas_folio_map optionnel, jamais fourni en
#    pratique -> §5.1 (renumérotation) systématiquement sautée -> folios
#    imprimés faux/dupliqués/non-monotones dans TOUT dossier généré dès que
#    des slides bibliothèque sont insérées entre des slides Canvas Master.
#    Confirmé en crash-test réel (dossier Fillon) : positions 22 et 23
#    affichaient toutes deux "16", position 24 affichait "19" puis la
#    position 26 affichait "18" (régression). Ni validate.py ni le grep
#    placeholder ne détectent ce défaut visible au client.
#    Ce correctif calcule canvas_folio_map automatiquement à partir du
#    Canvas Master D'ORIGINE (avant toute insertion/réordonnancement), donc
#    sans dépendre d'une saisie manuelle par session.
# ---------------------------------------------------------------------------
def build_default_canvas_folio_map(canvas_path):
    """Scanne le Cans_Mstr.pptx d'origine (non modifié) et associe à chaque
    rId de slide Canvas le folio statique (2 chiffres isolés, ex. "07") tel
    qu'il apparaît AVANT toute insertion de slide bibliothèque ou
    réordonnancement. Slides sans folio détecté (couverture, sections
    structurelles sans numéro) sont ignorées -- comportement attendu."""
    folio_map = {}
    with zipfile.ZipFile(canvas_path) as z:
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8", "ignore")
        rid_to_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/([^"]+)"', rels))
        for rid, target in rid_to_target.items():
            path = f"ppt/slides/{target}"
            try:
                xml = z.read(path).decode("utf-8", "ignore")
            except KeyError:
                continue
            texts = [t for t in re.findall(r"<a:t>([^<]*)</a:t>", xml) if t.strip()]
            folio_candidates = [t for t in texts[:4] if re.fullmatch(r"\d{2}", t)]
            if folio_candidates:
                folio_map[rid] = folio_candidates[0]
    return folio_map


# ---------------------------------------------------------------------------
# 4. Slide 5 — Patrimoine détaillé, construite avec les constantes de style
#    ci-dessus (jamais réinterprétées en session — §5.2bis).
# ---------------------------------------------------------------------------
def build_patrimoine_table_xml(patrimoine_spec):
    """Retourne une liste de lignes (texte, style) prêtes à insérer dans une
    zone de texte unique (implémentation simplifiée v1 — un tableau natif
    <a:tbl> complet est le raffinement naturel d'une v2, cf. notes de livraison)."""
    lines = []
    for famille in patrimoine_spec["familles"]:
        sous_total = sum(a["montant"] for a in famille["actifs"])
        lines.append({
            "texte": f'{famille["nom"]}  —  {sous_total:,.0f} €'.replace(",", " "),
            "couleur": COLOR_BORDEAUX, "police": FONT_TITRE, "style": "entete_famille"
        })
        for actif in famille["actifs"]:
            det = f' ({actif["detention"]})' if actif.get("detention") not in (None, "PP") else ""
            prop = actif.get("proprietaire", "COMMUN.")
            couleur_prop = COLOR_LABEL if prop == "COMMUN." else (
                COLOR_NAVY if prop.startswith(("Cl.", "Ge.")) and "conjoint" not in prop else COLOR_BORDEAUX
            )
            lines.append({
                "texte": f'{actif["nom"]}{det}',
                "proprietaire": prop, "proprietaire_couleur": couleur_prop,
                "montant": f'{actif["montant"]:,.0f} €'.replace(",", " "),
                "couleur": COLOR_CORPS, "police": FONT_CORPS, "style": "ligne_actif"
            })
    total = patrimoine_spec["total_net"]
    lines.append({
        "texte": f'TOTAL PATRIMOINE NET  —  {total:,.0f} €'.replace(",", " "),
        "couleur": "FFFFFF", "fond": COLOR_NAVY, "police": FONT_TITRE, "style": "bandeau_total"
    })
    return lines


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True, help="fichier JSON de spécification du dossier client")
    ap.add_argument("--out", required=True, help="chemin du .pptx de sortie")
    ap.add_argument("--skip-validate", action="store_true")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding="utf-8"))
    index_data = json.load(open(spec["index"], encoding="utf-8"))

    canvas_path = spec["canvas"]
    biblio_path = spec["bibliotheque"]
    for p in (canvas_path, biblio_path, spec["index"]):
        if not os.path.exists(p):
            die(f"{p} introuvable — cloner le dépôt mkgt avant d'appeler ce script.")

    # CORRECTIF (constat crash-test 29/08/2026, bug B du crash-test #3 du
    # 27/08/2026) : on ne construit plus jamais directement à args.out. Tout
    # le travail (extraction, rezip) se fait dans un fichier de chantier
    # temporaire ; args.out n'est écrit/écrasé qu'une fois validate.py passé
    # avec succès (ou --skip-validate explicite). Avant ce correctif, un
    # échec de validation laissait un fichier non conforme physiquement
    # présent à args.out (potentiellement en écrasant un livrable valide
    # d'un run précédent dès la ligne suivante, avant même toute validation).
    build_dir = tempfile.mkdtemp()
    work_out = os.path.join(build_dir, "_chantier.pptx")
    shutil.copy2(canvas_path, work_out)
    tmp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(work_out) as z:
        z.extractall(tmp_dir)
    biblio_tmp = tempfile.mkdtemp()
    with zipfile.ZipFile(biblio_path) as z:
        z.extractall(biblio_tmp)

    canvas_slide_count = len([
        f for f in os.listdir(os.path.join(tmp_dir, "ppt", "slides"))
        if re.match(r"slide\d+\.xml", f)
    ])

    print(f"[1/4] Insertion de {len(spec['slides_bibliotheque'])} slides bibliothèque...")
    inserted = insert_library_slides(tmp_dir, biblio_tmp, spec["slides_bibliotheque"])

    print("[2/4] Réordonnancement narratif (§5.2ter)...")
    exclude_canvas = set(spec.get("exclude_canvas", []))
    ordered_rids, excluded_rids = reorder_narratif(tmp_dir, canvas_slide_count, inserted, index_data, exclude_canvas)
    if exclude_canvas:
        cleanup_excluded_canvas_slides(tmp_dir, exclude_canvas, excluded_rids)

    canvas_folio_map = spec.get("canvas_folio_map")
    if canvas_folio_map:
        print("[3/4] Renumérotation des folios Canvas (§5bis, canvas_folio_map fourni)...")
    else:
        print("[3/4] canvas_folio_map absent de la spec — construction automatique depuis "
              "le Canvas Master d'origine (correctif 27/08/2026)...")
        canvas_folio_map = build_default_canvas_folio_map(canvas_path)
    # CORRECTIF 28/08/2026 (retour consultant — capture d'écran slide bibliothèque
    # "Fiscalité de l'assurance vie") : les slides bibliothèque insérées portent
    # elles aussi un folio statique figé dans Bibe_Def.pptx, identique à leur
    # position dans le catalogue (§3.4) — ex. la slide catalogue n°1 affiche "1"
    # même une fois assemblée en position 16 d'un dossier réel. build_default_
    # canvas_folio_map() ne scanne QUE le Canvas Master d'origine et ignorait donc
    # ces folios, jamais renumérotés. On les ajoute ici au même canvas_folio_map,
    # en réutilisant le mécanisme de remplacement de renumber_canvas_folios déjà
    # existant (aucune duplication de logique) : pour chaque slide insérée, le
    # folio d'origine est simplement str(pos) (vérifié identique au texte réel
    # sur les 4 slides testées le 28/08/2026 : positions 1, 2, 3, 7, 33).
    for e in inserted:
        canvas_folio_map[f"rId{e['rid']}"] = str(e["pos"])
    updated = renumber_canvas_folios(tmp_dir, ordered_rids, canvas_folio_map)
    for slide_file, old, new in updated:
        print(f"  {slide_file}: folio {old} -> {new}")
    if not updated:
        print("  ⚠ aucun folio renuméré — à vérifier manuellement avant livraison.")

    if spec.get("patrimoine"):
        lines = build_patrimoine_table_xml(spec["patrimoine"])
        print(f"[4/4] Slide 5 (patrimoine) : {len(lines)} lignes calculées avec les constantes de style "
              f"(non encore injectées en XML natif dans cette v1 — voir notes de livraison).")

    # Rezip — toujours vers le chantier temporaire, jamais directement args.out
    # (cf. correctif ci-dessus).
    if os.path.exists(work_out):
        os.remove(work_out)
    with zipfile.ZipFile(work_out, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(tmp_dir):
            for f in files:
                full = os.path.join(root, f)
                zf.write(full, os.path.relpath(full, tmp_dir))

    total_slides = canvas_slide_count - len(exclude_canvas) + len(inserted)

    def _promote_to_out():
        """N'écrit/n'écrase args.out qu'une fois le chantier validé."""
        if os.path.exists(args.out):
            os.remove(args.out)
        shutil.move(work_out, args.out)
        shutil.rmtree(build_dir, ignore_errors=True)

    if args.skip_validate:
        _promote_to_out()
        print(f"\nOK -> {args.out}  ({total_slides} slides : {canvas_slide_count - len(exclude_canvas)} Canvas + {len(inserted)} bibliothèque)"
              f"  [validation ignorée, --skip-validate]")
        return

    validate_script = "/mnt/skills/public/pptx/scripts/office/validate.py"
    if os.path.exists(validate_script):
        print("\n--- validate.py ---")
        r = subprocess.run([sys.executable, validate_script, work_out, "--original", canvas_path],
                            capture_output=True, text=True)
        print(r.stdout)
        if r.returncode != 0:
            print(r.stderr, file=sys.stderr)
            shutil.rmtree(build_dir, ignore_errors=True)
            die("validate.py a échoué — dossier NON conforme, ne pas livrer au consultant. "
                "CORRECTIF (constat 29/08/2026) : aucun fichier n'a été écrit ni modifié à "
                f"--out ({args.out}) — le chantier invalide a été supprimé, un éventuel "
                "livrable précédent à ce chemin reste intact.")
        _promote_to_out()
        print(f"\nOK -> {args.out}  ({total_slides} slides : {canvas_slide_count - len(exclude_canvas)} Canvas + {len(inserted)} bibliothèque)")
    else:
        shutil.rmtree(build_dir, ignore_errors=True)
        die("validate.py introuvable à l'emplacement attendu "
            "(/mnt/skills/public/pptx/scripts/office/validate.py). "
            "CORRECTIF 27/08/2026 : ce cas bloquait auparavant seulement avec un print, "
            "et le script continuait -> un dossier pouvait sortir sans AUCUNE validation "
            "structurelle sans que personne ne s'en aperçoive. Désormais bloquant. "
            "Utiliser --skip-validate explicitement si une validation manuelle est prévue. "
            f"Aucun fichier écrit à --out ({args.out}).")


if __name__ == "__main__":
    main()
