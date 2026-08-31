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
