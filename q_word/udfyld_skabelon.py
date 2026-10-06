"""Udfyld [[PLADSHOLDERE]] i DOCX-skabeloner.

Modulet arbejder direkte med DOCX-filer som bytes.

Det bruger kun Pythons standardbibliotek og har ingen relation
til PDF-konvertering eller q-gotenberg.
"""

from __future__ import annotations

from io import BytesIO
import re
from typing import Any
from xml.dom import Node
from xml.dom import minidom
from zipfile import BadZipFile
from zipfile import ZIP_DEFLATED
from zipfile import ZipFile


WORD_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

XML_NAMESPACE = (
    "http://www.w3.org/XML/1998/namespace"
)

PLACEHOLDER_PATTERN = re.compile(
    r"\[\[([A-ZÆØÅ][A-Z0-9_ÆØÅ]*)\]\]"
)

FIELD_NAME_PATTERN = re.compile(
    r"^[A-ZÆØÅ][A-Z0-9_ÆØÅ]*$"
)

WORD_TEXT_PART_PATTERN = re.compile(
    r"^word/(?:"
    r"document|"
    r"header\d+|"
    r"footer\d+|"
    r"footnotes|"
    r"endnotes"
    r")\.xml$"
)


def udfyld_word_skabelon(
    skabelon_bytes: bytes,
    vaerdier: dict[str, Any],
) -> bytes:
    """Udfyld pladsholdere i en DOCX-skabelon.

    Pladsholdere skrives som:

        [[BORGER_NAVN]]
        [[SAGSNUMMER]]
        [[DATO]]

    Den indsatte værdi anvender det Word-run, som indeholder
    pladsholderens første tegn. Dermed bevares formateringen fra
    pladsholderens første tegn.

    Funktionen behandler blandt andet:

    - almindelig Word-tekst
    - tabeller
    - sidehoveder
    - sidefødder
    - fodnoter
    - slutnoter

    None bliver erstattet med tom tekst.

    Hvis en værdi indeholder linjeskift, bliver linjeskiftene
    indsat som Word-linjeskift.

    Funktionen returnerer altid en DOCX som bytes.

    Hvis dokumentet skal bruges som PDF, skal den kaldende proces
    selv udføre PDF-konverteringen bagefter, eksempelvis ved et
    separat kald til q-gotenberg.

    q-word importerer eller kalder ikke q-gotenberg.
    """

    _validate_inputs(
        skabelon_bytes=skabelon_bytes,
        vaerdier=vaerdier,
    )

    normalized_values = {
        field_name: (
            ""
            if value is None
            else str(value)
        )
        for field_name, value in vaerdier.items()
    }

    try:
        with ZipFile(
            BytesIO(skabelon_bytes),
            "r",
        ) as source_zip:
            entries = source_zip.infolist()

            file_contents = {
                entry.filename: source_zip.read(
                    entry.filename
                )
                for entry in entries
                if not entry.is_dir()
            }

    except (BadZipFile, OSError) as error:
        raise ValueError(
            "skabelon_bytes er ikke en gyldig DOCX-fil."
        ) from error

    if "word/document.xml" not in file_contents:
        raise ValueError(
            "DOCX-filen mangler word/document.xml."
        )

    parsed_parts: dict[
        str,
        minidom.Document,
    ] = {}

    found_placeholders: set[str] = set()

    for part_name, xml_bytes in file_contents.items():
        if not WORD_TEXT_PART_PATTERN.fullmatch(
            part_name
        ):
            continue

        try:
            document = minidom.parseString(
                xml_bytes
            )

        except Exception as error:
            raise ValueError(
                f"DOCX-delen {part_name!r} "
                "indeholder ugyldig XML."
            ) from error

        parsed_parts[part_name] = document

        found_placeholders.update(
            _find_placeholders(document)
        )

    missing_fields = sorted(
        found_placeholders.difference(
            normalized_values
        )
    )

    if missing_fields:
        raise ValueError(
            "Skabelonen indeholder pladsholdere "
            "uden værdier: "
            + ", ".join(missing_fields)
            + "."
        )

    for part_name, document in parsed_parts.items():
        _replace_placeholders_in_document(
            document=document,
            values=normalized_values,
        )

        file_contents[part_name] = (
            document.toxml(
                encoding="utf-8"
            )
        )

    output = BytesIO()

    with ZipFile(
        output,
        mode="w",
        compression=ZIP_DEFLATED,
    ) as target_zip:
        for entry in entries:
            if entry.is_dir():
                target_zip.writestr(
                    entry,
                    b"",
                )
                continue

            target_zip.writestr(
                entry,
                file_contents[entry.filename],
            )

    result = output.getvalue()

    _validate_generated_docx(result)

    return result


def _validate_inputs(
    *,
    skabelon_bytes: bytes,
    vaerdier: dict[str, Any],
) -> None:
    """Kontrollér den offentlige funktions input."""

    if not isinstance(
        skabelon_bytes,
        bytes,
    ):
        raise TypeError(
            "skabelon_bytes skal være bytes."
        )

    if not skabelon_bytes:
        raise ValueError(
            "skabelon_bytes må ikke være tom."
        )

    if not isinstance(
        vaerdier,
        dict,
    ):
        raise TypeError(
            "vaerdier skal være en dictionary."
        )

    for field_name in vaerdier:
        if not isinstance(
            field_name,
            str,
        ):
            raise TypeError(
                "Alle nøgler i vaerdier "
                "skal være tekst."
            )

        if not FIELD_NAME_PATTERN.fullmatch(
            field_name
        ):
            raise ValueError(
                f"Ugyldigt feltnavn: "
                f"{field_name!r}. "
                "Brug store bogstaver, tal, "
                "underscore samt Æ, Ø og Å. "
                "Første tegn skal være et bogstav."
            )


def _find_placeholders(
    document: minidom.Document,
) -> set[str]:
    """Find pladsholdere i alle Word-afsnit."""

    placeholders: set[str] = set()

    paragraphs = document.getElementsByTagNameNS(
        WORD_NAMESPACE,
        "p",
    )

    for paragraph in paragraphs:
        paragraph_text = _get_paragraph_text(
            paragraph
        )

        placeholders.update(
            match.group(1)
            for match in PLACEHOLDER_PATTERN.finditer(
                paragraph_text
            )
        )

    return placeholders


def _replace_placeholders_in_document(
    *,
    document: minidom.Document,
    values: dict[str, str],
) -> None:
    """Erstat pladsholdere i én Word XML-del."""

    paragraphs = list(
        document.getElementsByTagNameNS(
            WORD_NAMESPACE,
            "p",
        )
    )

    for paragraph in paragraphs:
        _replace_placeholders_in_paragraph(
            paragraph=paragraph,
            values=values,
        )


def _replace_placeholders_in_paragraph(
    *,
    paragraph: minidom.Element,
    values: dict[str, str],
) -> None:
    """Erstat pladsholdere i ét Word-afsnit."""

    while True:
        text_elements = _get_text_elements(
            paragraph
        )

        paragraph_text = "".join(
            _get_element_text(text_element)
            for text_element in text_elements
        )

        matches = list(
            PLACEHOLDER_PATTERN.finditer(
                paragraph_text
            )
        )

        if not matches:
            return

        # Behandl sidste pladsholder først.
        # Så ændres positionen ikke for tidligere
        # pladsholdere i samme afsnit.
        match = matches[-1]

        field_name = match.group(1)

        _replace_text_range(
            paragraph=paragraph,
            text_elements=text_elements,
            start=match.start(),
            end=match.end(),
            replacement=values[field_name],
        )


def _get_paragraph_text(
    paragraph: minidom.Element,
) -> str:
    """Returnér den samlede synlige tekst i et afsnit."""

    return "".join(
        _get_element_text(text_element)
        for text_element in _get_text_elements(
            paragraph
        )
    )


def _get_text_elements(
    paragraph: minidom.Element,
) -> list[minidom.Element]:
    """Returnér alle w:t-elementer i et afsnit."""

    return list(
        paragraph.getElementsByTagNameNS(
            WORD_NAMESPACE,
            "t",
        )
    )


def _get_element_text(
    element: minidom.Element,
) -> str:
    """Returnér teksten fra et w:t-element."""

    text_parts: list[str] = []

    for child in element.childNodes:
        if child.nodeType in (
            Node.TEXT_NODE,
            Node.CDATA_SECTION_NODE,
        ):
            text_parts.append(
                child.data
            )

    return "".join(text_parts)


def _set_element_text(
    element: minidom.Element,
    value: str,
) -> None:
    """Erstat teksten i et w:t-element."""

    children = list(
        element.childNodes
    )

    for child in children:
        element.removeChild(child)

    text_node = (
        element.ownerDocument.createTextNode(
            value
        )
    )

    element.appendChild(text_node)

    _update_xml_space(
        element=element,
        value=value,
    )


def _update_xml_space(
    *,
    element: minidom.Element,
    value: str,
) -> None:
    """Bevar foranstillede og efterstillede mellemrum."""

    if (
        value.startswith(" ")
        or value.endswith(" ")
    ):
        element.setAttributeNS(
            XML_NAMESPACE,
            "xml:space",
            "preserve",
        )
        return

    if element.hasAttributeNS(
        XML_NAMESPACE,
        "space",
    ):
        element.removeAttributeNS(
            XML_NAMESPACE,
            "space",
        )


def _replace_text_range(
    *,
    paragraph: minidom.Element,
    text_elements: list[minidom.Element],
    start: int,
    end: int,
    replacement: str,
) -> None:
    """Erstat et tegninterval i et Word-afsnit."""

    element_ranges = _build_element_ranges(
        text_elements
    )

    (
        start_element_index,
        start_offset,
    ) = _locate_character(
        element_ranges=element_ranges,
        position=start,
    )

    (
        end_element_index,
        end_offset,
    ) = _locate_character(
        element_ranges=element_ranges,
        position=end - 1,
    )

    end_offset += 1

    start_element = text_elements[
        start_element_index
    ]

    end_element = text_elements[
        end_element_index
    ]

    start_text = _get_element_text(
        start_element
    )

    end_text = _get_element_text(
        end_element
    )

    prefix = start_text[:start_offset]
    suffix = end_text[end_offset:]

    if start_element_index == end_element_index:
        _set_replacement_in_start_run(
            paragraph=paragraph,
            start_element=start_element,
            prefix=prefix,
            replacement=replacement,
            suffix=suffix,
        )
        return

    _set_replacement_in_start_run(
        paragraph=paragraph,
        start_element=start_element,
        prefix=prefix,
        replacement=replacement,
        suffix="",
    )

    for element_index in range(
        start_element_index + 1,
        end_element_index,
    ):
        _set_element_text(
            text_elements[element_index],
            "",
        )

    _set_element_text(
        end_element,
        suffix,
    )


def _set_replacement_in_start_run(
    *,
    paragraph: minidom.Element,
    start_element: minidom.Element,
    prefix: str,
    replacement: str,
    suffix: str,
) -> None:
    """Indsæt værdien i pladsholderens første Word-run."""

    run = _find_parent_run(
        element=start_element,
        paragraph=paragraph,
    )

    normalized_replacement = (
        replacement
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )

    lines = normalized_replacement.split(
        "\n"
    )

    if len(lines) == 1:
        _set_element_text(
            start_element,
            prefix + lines[0] + suffix,
        )
        return

    _set_element_text(
        start_element,
        prefix + lines[0],
    )

    reference_node: Node = start_element

    for line_index, line in enumerate(
        lines[1:],
        start=1,
    ):
        break_element = (
            start_element.ownerDocument
            .createElementNS(
                WORD_NAMESPACE,
                "w:br",
            )
        )

        _insert_after(
            parent=run,
            new_node=break_element,
            reference_node=reference_node,
        )

        reference_node = break_element

        text_element = (
            start_element.ownerDocument
            .createElementNS(
                WORD_NAMESPACE,
                "w:t",
            )
        )

        text_value = line

        if line_index == len(lines) - 1:
            text_value += suffix

        _set_element_text(
            text_element,
            text_value,
        )

        _insert_after(
            parent=run,
            new_node=text_element,
            reference_node=reference_node,
        )

        reference_node = text_element


def _find_parent_run(
    *,
    element: minidom.Element,
    paragraph: minidom.Element,
) -> minidom.Element:
    """Find det w:r-element, der ejer et w:t-element."""

    current_node = element.parentNode

    while current_node is not None:
        if current_node is paragraph:
            break

        if (
            current_node.nodeType
            == Node.ELEMENT_NODE
            and current_node.namespaceURI
            == WORD_NAMESPACE
            and current_node.localName
            == "r"
        ):
            return current_node

        current_node = current_node.parentNode

    raise ValueError(
        "Pladsholderens tekst ligger ikke i "
        "et normalt Word-run."
    )


def _insert_after(
    *,
    parent: Node,
    new_node: Node,
    reference_node: Node,
) -> None:
    """Indsæt en XML-node efter en eksisterende node."""

    next_node = reference_node.nextSibling

    if next_node is None:
        parent.appendChild(new_node)
    else:
        parent.insertBefore(
            new_node,
            next_node,
        )


def _build_element_ranges(
    text_elements: list[minidom.Element],
) -> list[tuple[int, int]]:
    """Beregn tegnintervallet for hvert w:t-element."""

    ranges: list[tuple[int, int]] = []

    position = 0

    for text_element in text_elements:
        text_length = len(
            _get_element_text(text_element)
        )

        ranges.append(
            (
                position,
                position + text_length,
            )
        )

        position += text_length

    return ranges


def _locate_character(
    *,
    element_ranges: list[tuple[int, int]],
    position: int,
) -> tuple[int, int]:
    """Find XML-element og lokalt offset for et tegn."""

    for element_index, (
        element_start,
        element_end,
    ) in enumerate(element_ranges):
        if (
            element_start
            <= position
            < element_end
        ):
            return (
                element_index,
                position - element_start,
            )

    raise ValueError(
        "Kunne ikke finde pladsholderens "
        "position i Word-dokumentet."
    )


def _validate_generated_docx(
    docx_bytes: bytes,
) -> None:
    """Udfør en grundlæggende teknisk kontrol af outputfilen."""

    try:
        with ZipFile(
            BytesIO(docx_bytes),
            "r",
        ) as archive:
            bad_file = archive.testzip()

            if bad_file is not None:
                raise ValueError(
                    "Den genererede DOCX indeholder "
                    f"en beskadiget fil: {bad_file}."
                )

            names = set(
                archive.namelist()
            )

            if "word/document.xml" not in names:
                raise ValueError(
                    "Den genererede DOCX mangler "
                    "word/document.xml."
                )

            for name in names:
                if not WORD_TEXT_PART_PATTERN.fullmatch(
                    name
                ):
                    continue

                minidom.parseString(
                    archive.read(name)
                )

    except BadZipFile as error:
        raise ValueError(
            "Den genererede DOCX er ikke "
            "en gyldig ZIP-fil."
        ) from error