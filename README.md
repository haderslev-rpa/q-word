# q-word

`q-word` udfylder `[[PLADSHOLDERE]]` i Word DOCX-skabeloner og returnerer
en ny DOCX som bytes. Biblioteket bruger kun Pythons standardbibliotek.

## Fast ansvar

```text
DOCX-skabelon som bytes + værdier -> q-word -> udfyldt DOCX som bytes
```

`q-word` konverterer ikke til PDF og har ingen dependency til andre interne
repositories.

## Formatering

Den indsatte værdi arver formateringen fra pladsholderens første tegn.
Hvis hele `[[SAGSNUMMER]]` er fed i skabelonen, bliver værdien også fed.
Det samme princip gælder kursiv, understregning, skrifttype, størrelse,
farve og fremhævning.

Markér hele pladsholderen inklusive `[[` og `]]`, og anvend samme formatering
på hele pladsholderen. Hvis Word har opdelt teksten internt, bruges
formateringen fra det tekstområde, der indeholder det første `[`.

## Pladsholdere

```text
[[BORGER_NAVN]]
[[SAGSNUMMER]]
[[DATO]]
[[ADRESSE]]
```

Feltnavne bruger store bogstaver, tal, underscore samt Æ, Ø og Å. Første
tegn skal være et bogstav.

Funktionen behandler brødtekst, tabeller, sidehoveder, sidefødder, fodnoter
og slutnoter. `
` i en værdi indsættes som et Word-linjeskift.

## Brug

```python
from q_word.udfyld_skabelon import udfyld_word_skabelon

udfyldt_docx_bytes = udfyld_word_skabelon(
    skabelon_bytes=skabelon_bytes,
    vaerdier={
        "BORGER_NAVN": "Eksempel Navn",
        "SAGSNUMMER": "123456",
        "DATO": "6. oktober 2026",
        "ADRESSE": "Eksempelvej 1
6100 Eksempelby",
    },
)
```

Outputtet er altid en DOCX-fil som bytes. Originalskabelonen ændres ikke.
Manglende værdier giver fejl. `None` betyder bevidst tom tekst.

## Hvis processen skal bruge PDF

Funktionen laver bevidst ikke PDF. Processen skal selv foretage næste trin.
I vores processer kan den efterfølgende konvertering eksempelvis udføres
ved et separat kald til `q-gotenberg`:

```python
from q_gotenberg.fil_til_pdf import konverter_fil_til_pdf
from q_word.udfyld_skabelon import udfyld_word_skabelon

udfyldt_docx_bytes = udfyld_word_skabelon(
    skabelon_bytes=skabelon_bytes,
    vaerdier={"SAGSNUMMER": "123456"},
)

pdf_bytes = konverter_fil_til_pdf(
    filnavn="brev.docx",
    filindhold=udfyldt_docx_bytes,
)
```

Koblingen findes kun i processen. `q-word` importerer ikke `q-gotenberg`,
og `q-gotenberg` står ikke i `q-word`-projektets dependencies.

## Installation i en proces

```toml
[project]
dependencies = ["q-word"]

[tool.uv.sources]
q-word = { git = "https://github.com/haderslev-rpa/q-word.git", branch = "main" }
```

Kør `uv lock` og `uv sync`. Hvis processen også skal lave PDF, tilføjer
processen selv sin PDF-løsning som en separat dependency.




