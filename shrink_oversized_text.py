#!/usr/bin/env python3
"""
Affilior — Réduction automatique des zones "gros chiffre" en débordement.
Corrige le bug de débordement trouvé en crash-test (27/08/2026) : un montant
long ou un [MANQUANT] injecté dans une zone calibrée pour une valeur courte
(TMI, pourcentages, montants clés — sz ≥ 32pt) chevauche la légende sous elle.

Méthode : pour chaque run de texte avec sz ≥ SZ_THRESHOLD, estime le nombre
de caractères qui tiennent dans la largeur réelle de la forme (<a:ext cx>),
avec un ratio caractère/police calibré empiriquement sur ce Canvas Master
(K≈0.55, Garamond). Si le texte dépasse, réduit sz par paliers de 30 %
(cf. skill §5.6 : "paliers ~30 % ... ne pas descendre en dessous"),
jusqu'à 2 paliers maximum. Au-delà, NE DEVINE PAS davantage : signale ⚠
pour vérification visuelle plutôt que de produire du texte illisible.

⚠ Heuristique, pas une mesure de rendu réelle (pas de moteur de layout ici).
Reste un filet de sécurité — la QA visuelle (planche contact, §5.6) demeure
nécessaire sur les slides à risque, comme déjà prévu par le skill.
"""
import re, shutil, sys, zipfile

SZ_THRESHOLD = 3200        # 32pt — au-delà, on considère "zone gros chiffre"
CHAR_WIDTH_RATIO = 0.55    # ratio empirique largeur-caractère / taille police (Garamond)
MAX_STEPS = 2              # paliers de réduction ~30% max (cf. skill §5.6)
STEP_FACTOR = 0.70
EMU_PER_PT = 12700

def max_chars(width_emu: int, sz_hundredths: int) -> float:
    width_pt = width_emu / EMU_PER_PT
    sz_pt = sz_hundredths / 100
    return width_pt / (CHAR_WIDTH_RATIO * sz_pt)

def process_slide_xml(xml: str, slide_label: str):
    shapes = re.findall(r'<p:sp>.*?</p:sp>', xml, re.DOTALL)
    report = []
    for sp in shapes:
        ext_m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', sp)
        if not ext_m:
            continue
        width_emu = int(ext_m.group(1))

        # Traite chaque run individuellement (rPr sz + a:t associé)
        def run_repl(m):
            sz = int(m.group('sz'))
            text = m.group('text')
            if sz < SZ_THRESHOLD or not text.strip():
                return m.group(0)
            cur_sz = sz
            fits = len(text) <= max_chars(width_emu, cur_sz)
            steps = 0
            while not fits and steps < MAX_STEPS:
                cur_sz = int(cur_sz * STEP_FACTOR)
                fits = len(text) <= max_chars(width_emu, cur_sz)
                steps += 1
            if cur_sz != sz:
                report.append((slide_label, text[:40], sz, cur_sz, fits))
                new_run = m.group(0).replace(f'sz="{sz}"', f'sz="{cur_sz}"')
                if not fits:
                    report[-1] = (slide_label, text[:40], sz, cur_sz, False)
                return new_run
            return m.group(0)

        new_sp = re.sub(
            r'(?P<full><a:rPr[^>]*sz="(?P<sz>\d+)"[^>]*/?>.*?<a:t>(?P<text>[^<]*)</a:t>)',
            run_repl, sp, flags=re.DOTALL
        )
        if new_sp != sp:
            xml = xml.replace(sp, new_sp, 1)
    return xml, report

def apply(src_pptx, dst_pptx, slide_numbers):
    shutil.copy(src_pptx, dst_pptx)
    zin = zipfile.ZipFile(src_pptx, 'r')
    items = zin.infolist()
    data = {it.filename: zin.read(it.filename) for it in items}
    zin.close()

    full_report = []
    for n in slide_numbers:
        fname = f'ppt/slides/slide{n}.xml'
        if fname not in data:
            continue
        xml = data[fname].decode('utf-8')
        new_xml, report = process_slide_xml(xml, f'slide{n}')
        data[fname] = new_xml.encode('utf-8')
        full_report.extend(report)

    zout = zipfile.ZipFile(dst_pptx, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        zout.writestr(it, data[it.filename])
    zout.close()

    print(f'OK -> {dst_pptx}')
    if full_report:
        print(f'\n{len(full_report)} run(s) réduit(s) automatiquement :')
        for slide, txt, sz0, sz1, fits in full_report:
            status = '✓ tient sur une ligne' if fits else '⚠ toujours trop long — vérifier visuellement'
            print(f'  {slide:10s} {sz0/100:.0f}pt → {sz1/100:.0f}pt  "{txt}"  [{status}]')
    else:
        print('Aucune réduction nécessaire.')
    return full_report

if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('src'); ap.add_argument('dst')
    ap.add_argument('--slides', default='1,4,5,7,8,9,10,12,13,14,15,16,18')
    args = ap.parse_args()
    apply(args.src, args.dst, [int(x) for x in args.slides.split(',')])
