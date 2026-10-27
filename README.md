# ContadorVisitas_LPT

**Contador Digital / Digital Counter** — a VB6 visitor-counting device built by
Mediterrania Multimedia (2007/2008) to control the flow of people entering
exhibitions. Three independent counters (A/B/C) register pulses coming from
**optical sensors** wired to a custom **parallel-port (LPT, 0x378) hardware
board** (the classic "door counter station" for museum/exhibition openers,
exiting its head with a Flash interface). It was used in the exhibition
**"Sorolla"** at Bancaja's headquarters (Valencia) and afterwards in
**"Guerreros de Xian"** at the Museo de la Ciudad de las Ciencias
(Valencia).

Each counter can be configured through `config.txt`
(`config_explain.txt` explains every field: gate negation n0–n2, direction
and subtract-mode per counter, time-limit and refresh clock times), works
add/sub direction, keeps its accumulated total in `acumulado.txt` across
sessions and writes every event to timestamped logs in `regs/`. The
`interfaz.fla/swf` visual front-end talks to the VB core the same way as did
the other Mediterrania Multimedia kiosk projects of the era.

> **This project is shared ONLY for educational purposes. Its use, direct or derived, for professional or economic purposes is not permitted.**

For security/integrity reasons, no part of the project needed to be edited or removed: everything is kept as it was left in 2007/2008, including the third-party tools with their attributions.

## Requirements / How to run

- **Windows (XP era)** with VB6 runtime:
  run `contador/Contador1.exe` (core) and open `contador/interfaz.swf` with a
  Flash Player 8-era runtime (the "Ver Interfaz" button in the VB form opens
  the Flash front-end).
- **Hardware**: parallel-port I/O needs the user-mode driver
  `inpout32.dll` (freeware parallel-port access driver of the era; a copy of
  its sources and documentation survived in the original project research
  folders) — kept unmodified in the folders. The optical-sensor hardware
  board itself is not included.
- `contador/BAK/` keeps an older iteration; `ContadorSwap.exe` is a
  companion tool of the counter core.
- **Release**: `release/` contains the packaged **"Contador Digital
  _ version 3.exe"**; the full `ContadorDigital_v3.zip` shipment is attached
  as a **GitHub Release asset** (tag `v3`).

## Third-party tools

- `inpout32.dll` — parallel-port access driver, freeware of the era. Kept
  as-is.
- `scrrun.dll` — Microsoft Scripting Runtime redistributable of the era.
- `interfaz.fla` — Flash 8 source; requires a FLA-era authoring tool.

## DISCLAIMER

This software is provided "as is", without warranty of any kind, and no
liability is accepted for errors or damages resulting from its use.

> **AI training**: the use of the content of this repository (code, assets,
> documentation) to train or fine-tune machine learning models or AI systems
> is expressly **not permitted**. All of the repository is licensed under
> CC-BY-NC-4.0 (see the `LICENSE` file).

---

# ContadorVisitas_LPT

**Contador Digital** — dispositivo de conteo de visitas construido en VB6 por
Mediterrania Multimedia (2007/2008) para controlar el paso de personas que
entran a exposiciones. Tres contadores independientes (A/B/C) registran los
pulsos procedentes de **sensores ópticos** cableados a una **placa de
hardware del puerto paralelo (LPT, 0x378)** — la clásica "estación de
conteo en la puerta" de museos/exposiciones, con su interfaz gráfica en
Flash. Se usó en la exposición **"Sorolla"** en la sede central de Bancaja
(Valencia) y posteriormente en **"Guerreros de Xian"** en el Museo de la
Ciudad de las Ciencias (Valencia).

Cada contador se configura mediante `config.txt` (`config_explain.txt`
explica todos los campos: negación de puertas n0–n2, dirección y
modo-rester por contador, tiempo límite y tiempos de reloj de refresco IN/OUT),
funciona en dirección suma/resta, guarda el total acumulado entre sesiones en
`acumulado.txt` y registra cada evento con marca de tiempo en los logs de
`regs/`. La interfaz visual `interfaz.fla/swf` dialoga con el motor VB igual
que en los kioscos de Mediterrania Multimedia de la época.

> **Este proyecto se comparte SOLO con valor educativo. No está permitido su uso (directo o derivado) para fines profesionales o económicos.**

Por motivos de seguridad no fue necesario editar ni eliminar ninguna parte del proyecto: se conserva tal y como quedó en 2007/2008, incluyendo las herramientas de terceros con sus atribuciones.

## Requisitos / Cómo ejecutar

- **Windows (época XP)** con runtime VB6:
  ejecutar `contador/Contador1.exe` (motor) y abrir `contador/interfaz.swf`
  con un runtime Flash Player 8 de época (el botón "Ver Interfaz" del form
  VB abre la interfaz Flash).
- **Hardware**: la E/S por puerto paralelo necesita el driver en modo usuario
  `inpout32.dll` (driver freeware de acceso al puerto paralelo de la época; su
  fuente y documentación se conservaron en la investigación del proyecto) —
  conservado sin modificar en las carpetas. La placa de sensores ópticos en
  sí no está incluida.
- `contador/BAK/` guarda una iteración antigua; `ContadorSwap.exe` es una
  utilidad acompañante del motor del contador.
- **Entrega**: `release/` contiene el instalado **"Contador Digital _
  version 3.exe"**; el envío completo `ContadorDigital_v3.zip` se adjunta
  como **asset de Release en GitHub** (tag `v3`).

## Herramientas de terceros

- `inpout32.dll` — driver de acceso al puerto paralelo, freeware de la época.
  Conservada tal cual.
- `scrrun.dll` — redistribuible de Microsoft Scripting Runtime de la época.
- `interfaz.fla` — fuente de Flash 8; requiere una herramienta de autoría
  de la era FLA.

## AVISO LEGAL

Este software se proporciona "tal cual", sin garantía de ningún tipo, y no se
asume ninguna responsabilidad por errores o daños derivados de su uso.

> **Entrenamiento de IA**: no se permite expresamente el uso del contenido de
> este repositorio (código, recursos, documentación) para entrenar o ajustar
> (*fine-tune*) modelos de aprendizaje automático o sistemas de IA. Todo el
> repositorio se publica con licencia CC-BY-NC-4.0 (ver el archivo `LICENSE`).
