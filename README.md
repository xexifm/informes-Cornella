# Informes Cornellà — Generador

Script PowerShell que munta informes de deficiències per a llicències
d'activitat de l'Ajuntament de Cornellà.

## Estructura

```
informes-Cornella/
├── GenerarInforme.ps1            ← script principal
├── ESTRUCTURALS/
│   ├── 0 CAPCALERA.docx          ← capçalera fixa (placeholders <<NOM>>)
│   ├── 0 CONCLUSIONS.docx        ← un paràgraf per conclusió alternativa
│   └── REQ1.docx                 ← catàleg de deficiències (mostra 3 seccions)
└── Informes generats/            ← s'autogenera; els .docx finals hi van
```

Els .docx d'`ESTRUCTURALS/` que comencen per `0 ` són fixos. Tots els
altres es consideren **catàlegs** i apareixen al Pas 1 si n'hi ha més
d'un (si només hi ha REQ1, el script salta el Pas 1 i el fa servir
directament).

## Convencions del catàleg (REQ1.docx)

Estils Word que el parser reconeix:

| Estil      | Significat                                             |
|------------|--------------------------------------------------------|
| Heading 1  | Títol de secció (Autoritzacions, Controls inicials...) |
| Heading 2  | Nom curt de l'ítem (apareix al TreeView)               |
| Normal     | Cos: 1a línia = text principal, següents = URLs/extra  |

**Sub-bullets:** un Heading 2 que comenci per `::CHILD:: ` es tracta
com a fill de l'ítem Heading 2 anterior. Al TreeView apareix indentat
i seleccionable per separat.

**Placeholders al cos:**

- `[CAMP: nom_camp]` — el script demana el valor un sol cop per nom.
- `[CAMP: nom_camp (text d'ajuda)]` — el text entre parèntesis surt
  com a hint a sota del camp del formulari.

Exemples reals al fitxer:
- `Grup CAPCA: [CAMP: Grup CAPCA].` (apareix en 3 ítems → es demana 1 cop)
- `[CAMP: Òrgan homologació PAU (Protecció civil de Catalunya / Protecció civil local)]`

## Capçalera (CAPCALERA.docx)

Placeholders amb format `<<NOM>>` substituïts pels valors del Pas 2.
Els camps actualment demanats:

`ID_GIA`, `EXP_NUM`, `ADRECA`, `ACTIVITAT`, `PETICIONARI`, `DATA`,
`DECISIO`, `TECNIC`.

## Flux d'execució

1. **Pas 1** — Selecció de catàleg (si n'hi ha més d'un).
2. **Pas 2** — Formulari de dades de la capçalera.
3. **Pas 3** — TreeView amb checkboxes per cada secció / ítem / fill.
4. **Pas 4** — Formulari amb un camp per cada `[CAMP: …]` únic detectat.
5. **Pas 5** — Checkboxes per a les conclusions a incloure.
6. Es genera `Informes generats/Informe - <Activitat> - <data>.docx` i
   s'obre amb Word.

## Numeració

Tots els ítems seleccionats (pares i fills) es numeren **globalment**
1, 2, 3, … al document final, agrupats per secció. Els títols de
secció apareixen abans del seu grup d'ítems com a Heading 1.

## Estat actual

REQ1.docx conté **3 seccions** com a mostra inicial:

1. Autoritzacions / Informes preceptius (19 ítems, alguns amb fills)
2. Pla d'Autoprotecció (1 ítem)
3. Controls inicials (3 ítems)

Falten per afegir les seccions: Controls periòdics, Instal·lacions,
Registres, Incendis (constructiu), Projecte, Activitat, Restauració,
Cuina, Denúncies (soroll/calor/accessibilitat/emmagatzematge),
Documentació/Rètols i Accessibilitat.
