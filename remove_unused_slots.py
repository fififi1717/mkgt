#!/usr/bin/env python3
"""
Affilior — Suppression des slots non utilisés (règle 5.2quater).
Retire entièrement les formes XML (<p:sp>) d'une colonne/ligne identifiée par
son offset X (ou X+Y pour une ligne), plutôt que de vider leur texte —
corrige le bug "cadres vides visibles" trouvé en crash-test (27/08/2026).

⚠ Ne fait PAS le reflow (élargir les colonnes restantes) : supprime proprement
les shapes en trop, mais laisse l'espace vide. Le reflow réel nécessite des
variantes de mise en page pré-construites (1/2/3 colonnes) — chantier séparé.
"""
import re, sys, zipfile, shutil

TOLERANCE_EMU = 20000  # ~1.6mm — absorbe l'écart texte/rectangle de fond (cf. crash-test : ~9144 EMU observé)

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
    # Test : slide 12, retirer la colonne 3 (x=5897880, cf. inspection XML)
    apply(sys.argv[1], sys.argv[2], {12: {"x": [5897880]}})
