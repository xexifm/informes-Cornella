"""Converteix el Word de Llicencia a ESTRUCTURALS/LLIC.json.

LLIC no copia cap text de REQ1: per cada punt hi desa nomes el que es propi de
Llicencia (el "No es disposa", el "Es disposa" i el "Quan:") i una CLAU que
apunta a l'item de REQ1 ("Seccio::Titol"). El text surt de REQ1 en viu.

Els punts que NO existeixen a REQ1 van a la seccio PROPIS amb el text sencer.
Al final s'imprimeix el REPARTIMENT perque es pugui repassar.
"""
import json, re, sys
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
DOC = '/tmp/claude-0/-home-user-informes-Cornella/60d2770b-bc99-5da3-a36e-111509e6fdfd/scratchpad/llic/req/word/document.xml'
REQ1 = '/home/user/informes-Cornella/ESTRUCTURALS/REQ1.json'
OUT = '/home/user/informes-Cornella/ESTRUCTURALS/LLIC.json'


def llegeix_paragrafs(path):
    t = ET.parse(path)
    out = []
    for p in t.getroot().iter(W + 'p'):
        txt = ''.join(x.text or '' for x in p.iter(W + 't'))
        out.append(txt)
    return out


P = llegeix_paragrafs(DOC)


def para(txt, url=False):
    """Un paragraf del format estandard: llista de runs."""
    d = {'runs': [{'t': txt}]}
    if url:
        d['url'] = True
    return d


def cos(idxs):
    """Cos a partir d'indexs de paragraf del Word. Els que son un enllac es
    marquen com a url (el format estandard ja ho preveu)."""
    out = []
    for i in idxs:
        t = P[i].strip()
        if not t:
            continue
        out.append(para(t, url=t.startswith('http')))
    return out


# --- ABANS: punt del Word -> clau de REQ1 -------------------------------------
# (idx_titol, [idx del "No es disposa"], [idx del "Es disposa"], clau REQ1)
S = 'Autoritzacions / Informes preceptius::'
ABANS = [
    (26, [27, 28], [29], S + 'Sanitat'),
    (31, [32, 33], [34], S + 'Educació'),
    (36, [37, 38], [39], S + 'Comerç'),
    (41, [42, 43], [44], S + 'Esport'),
    (46, [47, 48], [49], S + 'Incendis'),
    (53, [54], [55], S + 'Impacte ambiental'),
    (57, [58, 59], [60], S + 'Vector Aigua (ACA) - Zona inundable fluvial T500'),
    (64, [65], [66], S + 'Vector Aigua (ACA) - abocaments'),
    (68, [69, 70], [71], S + 'Vector Aigua (AMB) - abocaments'),
    (73, [74], [75], S + 'Vector Residus (ARC)'),
    (77, [78, 79], [80], S + 'Vector Residus - SDR'),
    (82, [85], [86], S + 'Vector Atmosfera - Grup A/B'),
    (89, [90], [91], S + 'Vector Sòl'),
    (93, [94], [95], S + 'Mobilitat'),
    (97, [98, 99], [100], 'Registres::RASIC'),
]

# --- DESPRES: punt del Word -> clau de REQ1 + el seu "Quan:" ------------------
# (idx_titol, idx_quan, clau REQ1 o None si es PROPI, [idx dels sub-punts])
# Els sub-punts son les llistes que pengen d'un punt (per exemple, quines
# instal·lacions s'han de legalitzar): si no es recollien, es perdien.
DESPRES = [
    (114, 115, None, []),                                              # Certificat Final d'Activitat
    (118, 120, None, [119]),                                              # Acte comprovació incendis EC-PCAA
    (122, 125, 'Incendis::SP 136 Model A - limitació propagació', []),
    (127, 130, 'Incendis::SP 136 Model C - protecció estructures', []),
    (132, 135, 'Incendis::ITC SP 108 - lluernes en coberta', []),
    (137, None, 'Instal·lacions::Insp. periòdica - PCI', []),
    (140, 142, None, [141]),                                              # Sanitat - autorització funcionament
    (144, 146, None, [145]),                                              # Educació - obertura
    (148, 150, None, [149]),                                              # Comerç - inici activitat
    (152, 153, None, []),                                              # Vector Aigua (AMB) - anàlisi
    (156, 157, "Pla d'Autoprotecció::PAU", []),
    (160, 162, 'Controls inicials::Annex II Llei 20/2009 - control inicial', [161]),
    (164, 170, 'Controls inicials::Vector Atmosfera - control inicial', [165, 166]),
    (173, 184, None, list(range(174, 184))),                                              # RITSIC - legalitzacions (llista)
    (186, 191, None, list(range(187, 191))),                                              # Inspecció inicial (llista)
]

# --- Els dos punts CONDICIONALS ----------------------------------------------
COND = [
    ('annexii',     17, [18], [19]),   # nomes si NO es llicencia provisional
    ('provisional', 22, [23], [24]),   # nomes si SI ho es
]

# --- Comprovacio: totes les claus han d'existir a REQ1 ------------------------
req = json.load(open(REQ1, encoding='utf-8'))
claus_req1 = set()
for sec in req['nodes']:
    st = sec.get('titol', '')
    for el in sec.get('fills', []):
        if el.get('tipus') == 'subseccio':
            for x in el.get('fills', []):
                if x.get('titol'):
                    claus_req1.add(f"{st}::{x['titol']}")
        elif el.get('titol'):
            claus_req1.add(f"{st}::{el['titol']}")

orfes = []
for _, _, _, k in ABANS:
    if k not in claus_req1:
        orfes.append(k)
for _, _, k, _s in DESPRES:
    if k and k not in claus_req1:
        orfes.append(k)
if orfes:
    print('ERROR: claus que NO existeixen a REQ1:')
    for k in orfes:
        print('   ', k)
    sys.exit(1)

# --- Muntatge del JSON --------------------------------------------------------
def item_abans(idx_t, no_idx, si_idx, clau):
    return {
        'tipus': 'item',
        'titol': clau.split('::')[-1],
        'clau': clau,
        'cos': [],
        'fills': [
            {'tipus': 'nodisposa', 'titol': '', 'cos': cos(no_idx), 'fills': []},
            {'tipus': 'sidisposa', 'titol': '', 'cos': cos(si_idx), 'fills': []},
        ],
    }


nodes = []

nodes.append({
    'tipus': 'seccio',
    'titol': 'ABANS',
    'cos': [],
    'fills': [item_abans(*a) for a in ABANS],
})

fills_despres = []
for idx_t, idx_q, clau, subs in DESPRES:
    n = {
        'tipus': 'item',
        'titol': clau.split('::')[-1] if clau else P[idx_t].strip()[:60],
        'cos': [] if clau else cos([idx_t]),
        'fills': [],
    }
    if clau:
        n['clau'] = clau
    for i in subs:
        t = P[i].strip()
        if t:
            n['fills'].append({'tipus': 'subitem', 'titol': '', 'cos': cos([i]), 'fills': []})
    if idx_q is not None:
        quan = P[idx_q].strip()
        quan = re.sub(r'^Quan:\s*', '', quan)
        n['fills'].append({'tipus': 'quan', 'titol': '', 'cos': [para(quan)], 'fills': []})
    fills_despres.append(n)

nodes.append({'tipus': 'seccio', 'titol': 'DESPRES', 'cos': [], 'fills': fills_despres})

fills_propis = []
for nom, idx_t, no_idx, si_idx in COND:
    fills_propis.append({
        'tipus': 'item',
        'titol': P[idx_t].strip()[:60],
        'condicio': nom,
        'cos': cos([idx_t]),
        'fills': [
            {'tipus': 'nodisposa', 'titol': '', 'cos': cos(no_idx), 'fills': []},
            {'tipus': 'sidisposa', 'titol': '', 'cos': cos(si_idx), 'fills': []},
        ],
    })
nodes.append({'tipus': 'seccio', 'titol': 'PROPIS', 'cos': [], 'fills': fills_propis})

doc = {'tipus': 'LLIC', 'familia': 'llicencia', 'intro': [], 'nodes': nodes}
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(doc, f, ensure_ascii=False, indent=2)
    f.write('\n')

# --- Repartiment, per repassar-lo --------------------------------------------
print(f"LLIC.json escrit: {OUT}\n")
print('=== ABANS (autoritzacions) — lligats a REQ1 ===')
for idx_t, _, _, k in ABANS:
    print(f"  {P[idx_t].strip()[:46]:<48s} -> {k}")
print('\n=== DESPRES ===')
for idx_t, idx_q, k, subs in DESPRES:
    dest = k if k else '(PROPI: no hi ha equivalent a REQ1)'
    q = P[idx_q].strip()[:40] if idx_q is not None else '(sense Quan)'
    print(f"  {P[idx_t].strip()[:42]:<44s} -> {dest}")
    extra = f"   ({len(subs)} sub-punts)" if subs else ''
    print(f"     {'':44s}    {q}{extra}")
print('\n=== PROPIS (condicionals) ===')
for nom, idx_t, _, _ in COND:
    print(f"  [{nom}] {P[idx_t].strip()[:70]}")
print(f"\nTOTAL: {len(ABANS)} abans, {len(DESPRES)} despres "
      f"({sum(1 for _,_,k,_s in DESPRES if k)} lligats a REQ1, "
      f"{sum(1 for _,_,k,_s in DESPRES if not k)} propis), {len(COND)} condicionals")
