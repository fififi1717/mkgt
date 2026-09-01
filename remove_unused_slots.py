#!/usr/bin/env python3
"""
Affilior — Suppression des slots non utilisés (règle 5.2quater).
Retire entièrement les formes XML (<p:sp>) d'une colonne/ligne identifiée par
son offset X (ou X+Y pour une ligne), plutôt que de vider leur texte —
corrige le bug "cadres vides visibles" trouvé en crash-test (27/08/2026).

⚠ Ne fait PAS le reflow (élargir les colonnes restantes) : supprime proprement
les shapes en trop, mais laisse l'espace vide. Le reflow réel nécessite des
variantes de mise en page pré-construites (1/2/3 colonnes) — chantier séparé.

FUSION 28/08/2026 (nettoyage dépôt, demande consultant) : select_slots.py
intégré ici — c'était un fichier de 20 lignes, une seule fonction, jamais
appelée seule en production (toujours importée par finalize_deck.py). Fusion
sans changement de comportement, sauf un correctif : nb_elements_reels
négatif est désormais rejeté explicitement (bug identifié au crash-test #3
du 27/08/2026, jamais corrigé jusqu'ici — un slicing Python silencieux
laissait passer un résultat incorrect sans aucune alerte).
"""
import re, sys, zipfile, shutil

TOLERANCE_EMU = 20000  # ~1.6mm — absorbe l'écart texte/rectangle de fond (cf. crash-test : ~9144 EMU observé)


def compute_slots_a_supprimer(template_positions: dict, nb_elements_reels: dict) -> dict:
    """Traduit "nombre d'éléments réels pour ce dossier" en positions XML à
    supprimer, à partir du référentiel générique template_positions.json.

    Usage :
        slots = compute_slots_a_supprimer(template_positions, {"12": 2, "13": 2, "15": 2})
        # -> {"12": {"x": [5897880]}, "13": {"y": [...]}, "15": {"x": [5897880]}}
    """
    slots_variables = template_positions["slots_variables"]
    result = {}
    for slide_n, nb_reel in nb_elements_reels.items():
        if slide_n not in slots_variables:
            raise ValueError(f"slide {slide_n} absente de template_positions['slots_variables']")
        if nb_reel < 0:
            # CORRECTIF 28/08/2026 : un nombre négatif ne devrait jamais arriver en
            # amont, mais rester silencieux ici (slicing positions[nb_reel:] avec un
            # indice négatif) produisait un résultat incorrect sans aucune alerte —
            # trouvé au crash-test #3 (27/08/2026), non corrigé jusqu'ici.
            raise ValueError(
                f"slide {slide_n} : nb_elements_reels={nb_reel} négatif — valeur "
                f"invalide, à corriger en amont (Étape 2) avant de poursuivre."
            )
        conf = slots_variables[slide_n]
        positions = conf["positions"]
        max_slots = conf["max_slots"]
        if nb_reel > max_slots:
            raise ValueError(
                f"slide {slide_n} ({conf['description']}) : {nb_reel} éléments réels "
                f"mais le template n'a que {max_slots} slots — cas non géré automatiquement, "
                f"cf. limite dure documentée (comme slide 8 / 3 enfants)."
            )
        a_supprimer = positions[nb_reel:]
        # CORRECTIF 01/09/2026 (support "slot groupé", cf. slide 4 et slide 12
        # depuis leur refonte en lignes horizontales / puces) : un slot peut être
        # soit un offset unique (int, ancien format colonnes/lignes simples,
        # ex. slide 13/15), soit une LISTE de plusieurs offsets appartenant au
        # même slot logique (ex. slide 12 : mot-clé + phrase + filet, à des Y
        # différents ; slide 4 : barre+texte + filet précédent). On aplatit ici
        # avant de renvoyer, pour que remove_shapes_at_offset() n'ait jamais à
        # savoir si le slide utilise un format simple ou groupé.
        flat = []
        for item in a_supprimer:
            if isinstance(item, list):
                flat.extend(item)
            else:
                flat.append(item)
        if flat:
            result[slide_n] = {conf["axe"]: flat}
    return result

def remove_shapes_at_offset(xml: str, x_values: set = None, y_values: set = None) -> tuple[str, int]:
    """Supprime les formes dont l'offset X (colonnes, ex. slide12/15) OU l'offset
    Y (lignes de table, ex. slide13) correspond à une valeur donnée, à une
    tolérance près. Nécessaire car le rectangle de fond d'une ligne n'a PAS le
    même offset Y exact que les textes qu'il porte (~9144 EMU d'écart observé
    sur ce Canvas Master) — un matching strict laisse les bandes de fond vides
    visibles même quand le texte a bien été supprimé (trouvé en crash-test,
    27/08/2026). Les deux layouts (colonnes / lignes) existent dans le Canvas
    Master — toujours vérifier lequel s'applique avant de construire la config."""
    x_values = x_values or set()
    y_values = y_values or set()
    def near(val, values):
        return any(abs(val - v) <= TOLERANCE_EMU for v in values)
    shapes = re.findall(r'<p:sp>.*?</p:sp>', xml, re.DOTALL)
    removed = 0
    for sp in shapes:
        m = re.search(r'<a:off x="(\d+)" y="(\d+)"/>', sp)
        if not m:
            continue
        x, y = int(m.group(1)), int(m.group(2))
        if near(x, x_values) or near(y, y_values):
            xml = xml.replace(sp, '', 1)
            removed += 1
    return xml, removed

def apply(src_pptx, dst_pptx, slide_slots):
    """slide_slots: {slide_n: {"x": [...]}} ou {slide_n: {"y": [...]}} ou
    {slide_n: [...]} (raccourci = x, pour compatibilité colonnes)"""
    shutil.copy(src_pptx, dst_pptx)
    zin = zipfile.ZipFile(src_pptx, 'r')
    items = zin.infolist()
    data = {it.filename: zin.read(it.filename) for it in items}
    zin.close()

    for slide_n, slots in slide_slots.items():
        fname = f'ppt/slides/slide{slide_n}.xml'
        xml = data[fname].decode('utf-8')
        if isinstance(slots, list):
            x_values, y_values = slots, []
        else:
            x_values, y_values = slots.get('x', []), slots.get('y', [])
        xml, removed = remove_shapes_at_offset(xml, set(x_values), set(y_values))
        axis_desc = f"x={x_values}" if x_values else f"y={y_values}"
        print(f'slide{slide_n}.xml : {removed} formes supprimées ({axis_desc})')
        data[fname] = xml.encode('utf-8')

    zout = zipfile.ZipFile(dst_pptx, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        zout.writestr(it, data[it.filename])
    zout.close()
    print(f'OK -> {dst_pptx}')


# ---------------------------------------------------------------------------
# Centrage dynamique — slide 12 "Vos projets" (AJOUTÉ 01/09/2026, retour Marc
# en crash-test, proposition A retenue parmi 3 maquettes). FUSIONNÉ ici
# plutôt que dans un script séparé (center_projets12.py, existé brièvement
# le temps de la mise au point) — même logique de nettoyage de dépôt que la
# fusion select_slots.py du 28/08/2026 : une seule responsabilité ("gérer les
# lignes de la slide 12 selon nb_elements_reels"), un seul fichier.
#
# Comportement : à 3 objectifs, AUCUN changement (déjà calé pleine hauteur,
# ne déborde pas). À 1 ou 2 objectifs, le(s) objectif(s) restant(s) sont
# centrés verticalement comme un bloc (espace haut/bas égal, RIEN entre eux)
# — et le filet séparateur entre rangées est SUPPRIMÉ (pas déplacé) pour ces
# cas, conformément à la consigne "pas de trait" : un trait entre 2 objectifs
# centrés avec du vide de chaque côté n'aurait plus de rôle de séparation de
# tableau, seulement décoratif et non désiré.
# ---------------------------------------------------------------------------
TOP_PROJETS = 1371600
BOTTOM_PROJETS = 4600000  # marge de sécurité avant le filet de pied de page (4773168)
ENVELOPE_PROJETS = BOTTOM_PROJETS - TOP_PROJETS
ROW_HEIGHT_PROJETS = 570000       # bloc mot-clé + phrase d'une rangée (budget template)
DELTA_PHRASE_PROJETS = 220000     # offset phrase par rapport au mot-clé, fixe

# Positions Y d'origine dans le Canvas Master, par rangée affichée (1-indexé) —
# après compute_slots_a_supprimer, les N premières rangées survivent toujours
# à CES positions d'origine (jamais renumérotées par ce script).
ORIGIN_ROWS_PROJETS = {
    1: {"filet": None, "motcle": 1371600, "phrase": 1591600},
    2: {"filet": 2171600, "motcle": 2271600, "phrase": 2491600},
    3: {"filet": 3071600, "motcle": 3171600, "phrase": 3391600},
}


def _move_shape_at_offset(xml, old_y, new_y, tolerance=TOLERANCE_EMU):
    """Déplace (contrairement à remove_shapes_at_offset qui supprime) la forme
    dont le <a:off> a un y à ±tolerance de old_y. Ne touche jamais x. Lève si
    0 ou >1 correspondance (sécurité — même esprit que le reste du fichier)."""
    pattern = re.compile(r'(<a:off x="\d+" y=")(\d+)("/>)')
    matches = [m for m in pattern.finditer(xml) if abs(int(m.group(2)) - old_y) <= tolerance]
    if len(matches) != 1:
        raise ValueError(f"attendu exactement 1 forme à y≈{old_y}, trouvé {len(matches)}")
    m = matches[0]
    return xml[:m.start()] + m.group(1) + str(new_y) + m.group(3) + xml[m.end():]


def _target_positions_projets(n):
    """n rangées à afficher -> liste de {motcle, phrase} centrées dans
    l'enveloppe, une bande égale par rangée."""
    if n <= 0:
        return []
    band_h = ENVELOPE_PROJETS / n
    targets = []
    for i in range(n):
        band_top = TOP_PROJETS + i * band_h
        motcle_y = round(band_top + (band_h - ROW_HEIGHT_PROJETS) / 2)
        targets.append({"motcle": motcle_y, "phrase": motcle_y + DELTA_PHRASE_PROJETS})
    return targets


def center_slide12_projets(xml, nb_objectifs):
    if nb_objectifs >= 3 or nb_objectifs <= 0:
        return xml, "aucun changement (3 objectifs ou 0 — comportement historique conservé)"
    targets = _target_positions_projets(nb_objectifs)
    moved, deleted = 0, 0
    for i in range(1, nb_objectifs + 1):
        origin, target = ORIGIN_ROWS_PROJETS[i], targets[i - 1]
        if origin["filet"] is not None:
            xml, n = remove_shapes_at_offset(xml, y_values={origin["filet"]})
            deleted += n
        for key in ("motcle", "phrase"):
            if origin[key] != target[key]:
                xml = _move_shape_at_offset(xml, origin[key], target[key])
                moved += 1
    return xml, f"{moved} forme(s) repositionnée(s), {deleted} filet(s) supprimé(s), pour {nb_objectifs} objectif(s)"


def apply_center_slide12(src_pptx, dst_pptx, nb_objectifs):
    """À appeler APRÈS apply() (suppression des rangées en trop) — édition
    XML brute, jamais via python-pptx : aucun risque de renumérotation des
    fichiers slideN.xml (cf. §5bis)."""
    with zipfile.ZipFile(src_pptx) as z:
        xml = z.read("ppt/slides/slide12.xml").decode("utf-8")
    new_xml, status = center_slide12_projets(xml, nb_objectifs)
    print(status)
    tmp_out = dst_pptx + ".__tmp"
    with zipfile.ZipFile(src_pptx) as zin, zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/slides/slide12.xml":
                data = new_xml.encode("utf-8")
            zout.writestr(item, data)
    import os
    os.replace(tmp_out, dst_pptx)
    print(f"OK -> {dst_pptx}")


if __name__ == '__main__':
    if len(sys.argv) == 3:
        # Test : slide 13 (Moyens), retirer la dernière ligne (y=3337560,
        # cf. template_positions.json). Ancien exemple (slide 12, x=5897880)
        # obsolète depuis la refonte en lignes horizontales du 01/09/2026 —
        # slide 12 n'utilise plus l'axe X (colonnes) mais l'axe Y (groupé).
        apply(sys.argv[1], sys.argv[2], {13: {"y": [3337560]}})
    elif len(sys.argv) == 2:
        # Mode select_slots.py d'origine : affiche les slots calculés sans les appliquer
        import json
        tp = json.load(open(sys.argv[1], encoding="utf-8"))
        nb = {"12": 2, "13": 2, "15": 2}
        slots = compute_slots_a_supprimer(tp, nb)
        print(json.dumps(slots, ensure_ascii=False, indent=2))
    else:
        print("Usage : remove_unused_slots.py <src.pptx> <dst.pptx>   (applique la suppression)")
        print("    ou : remove_unused_slots.py <template_positions.json>   (calcule les slots, exemple)")
        sys.exit(1)
