
from __future__ import annotations
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional

# Namespace used by Microsoft Word bibliography files
NS = {"b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"}

@dataclass
class Person:
    first: str = ""
    middle: str = ""
    last: str = ""
    suffix: str = ""

    def format_bibtex(self) -> str:
        parts = []
        if self.last:
            parts.append(self.last)
        given = " ".join([p for p in [self.first, self.middle, self.suffix] if p])
        if given:
            parts.append(given)
        return ", ".join(parts).strip()

    def format_ris(self) -> str:
        given = " ".join([p for p in [self.first, self.middle] if p]).strip()
        if self.suffix:
            given = (given + (" " if given else "") + self.suffix).strip()
        return f"{self.last}, {given}".strip(", ")

@dataclass
class Source:
    tag: str = ""
    type: str = ""
    title: str = ""
    year: str = ""
    journal: str = ""
    book_title: str = ""
    publisher: str = ""
    city: str = ""
    volume: str = ""
    issue: str = ""
    pages: str = ""
    doi: str = ""
    url: str = ""
    authors: List[Person] = field(default_factory=list)
    editors: List[Person] = field(default_factory=list)

    def bibtex_key(self) -> str:
        # lastnameYYYYFirstWord
        last = self.authors[0].last if self.authors else "anon"
        year = re.findall(r"\d{4}", self.year or "")
        y = year[0] if year else "n.d."
        fw = re.sub(r"\W+", "", (self.title or "untitled").split()[0])[:12]
        return f"{last}{y}{fw}".lower()

    def entry_type_bibtex(self) -> str:
        t = (self.type or "").lower()
        mapping = {
            "journalarticle": "article",
            "book": "book",
            "booksection": "incollection",
            "conferenceproceedings": "inproceedings",
            "report": "techreport",
            "thesis": "phdthesis",
            "mastersthesis": "mastersthesis",
            "internet": "misc",
            "webpage": "misc",
            "film": "misc",
            "art": "misc",
            "patent": "misc",
        }
        return mapping.get(t, "misc")

    def entry_type_ris(self) -> str:
        t = (self.type or "").lower()
        mapping = {
            "journalarticle": "JOUR",
            "book": "BOOK",
            "booksection": "CHAP",
            "conferenceproceedings": "CPAPER",
            "report": "RPRT",
            "thesis": "THES",
            "mastersthesis": "THES",
            "internet": "ELEC",
            "webpage": "ELEC",
            "film": "MPCT",
            "art": "GEN",
            "patent": "PAT",
        }
        return mapping.get(t, "GEN")

def _text(node: Optional[ET.Element]) -> str:
    return (node.text or "").strip() if node is not None else ""

def _find(node: ET.Element, path: str) -> Optional[ET.Element]:
    return node.find(path, NS)

def parse_sources_xml(xml_bytes: bytes) -> List[Source]:
    """Parse Microsoft Word Sources.xml to a list of Source objects."""
    tree = ET.fromstring(xml_bytes)
    if tree.tag.endswith("Sources"):
        root = tree
    else:
        # maybe the file had a BOM or wrapper; try to find matching node
        root = tree.find(".//b:Sources", NS) or tree

    sources = []
    for s in root.findall("b:Source", NS):
        src = Source()
        src.tag = _text(_find(s, "b:Tag"))
        src.type = _text(_find(s, "b:SourceType"))
        src.title = _text(_find(s, "b:Title"))
        src.year = _text(_find(s, "b:Year"))
        src.journal = _text(_find(s, "b:JournalName"))
        src.book_title = _text(_find(s, "b:BookTitle"))
        src.publisher = _text(_find(s, "b:Publisher"))
        src.city = _text(_find(s, "b:City"))
        src.volume = _text(_find(s, "b:Volume"))
        src.issue = _text(_find(s, "b:Number"))
        src.pages = _text(_find(s, "b:Pages"))
        src.doi = _text(_find(s, "b:DOI"))
        src.url = _text(_find(s, "b:URL")) or _text(_find(s, "b:LCID"))  # LCID misused sometimes

        # Authors
        for person in s.findall(".//b:Author/b:Author/b:NameList/b:Person", NS):
            p = Person(
                first=_text(_find(person, "b:First")),
                middle=_text(_find(person, "b:Middle")),
                last=_text(_find(person, "b:Last")),
                suffix=_text(_find(person, "b:Suffix")),
            )
            src.authors.append(p)
        # Corporate author
        for corp in s.findall(".//b:Author/b:Author/b:Corporate", NS):
            src.authors.append(Person(last=_text(corp)))

        # Editors (rare but present)
        for person in s.findall(".//b:Editor/b:Editor/b:NameList/b:Person", NS):
            p = Person(
                first=_text(_find(person, "b:First")),
                middle=_text(_find(person, "b:Middle")),
                last=_text(_find(person, "b:Last")),
                suffix=_text(_find(person, "b:Suffix")),
            )
            src.editors.append(p)

        sources.append(src)
    return sources

def _escape_bibtex(value: str) -> str:
    # Minimal escaping for BibTeX
    value = value.replace("\\", "\\\\")
    for ch in ["{", "}", "\""]:
        value = value.replace(ch, f"\\{ch}")
    return value

def to_bibtex(sources: List[Source]) -> str:
    entries = []
    for s in sources:
        entry_type = s.entry_type_bibtex()
        key = s.bibtex_key()
        fields: Dict[str, str] = {}

        if s.title: fields["title"] = s.title
        if s.year: fields["year"] = s.year
        if s.journal: fields["journal"] = s.journal
        if s.book_title: fields["booktitle"] = s.book_title
        if s.publisher: fields["publisher"] = s.publisher
        if s.city: fields["address"] = s.city
        if s.volume: fields["volume"] = s.volume
        if s.issue: fields["number"] = s.issue
        if s.pages: fields["pages"] = s.pages
        if s.doi: fields["doi"] = s.doi
        if s.url: fields["url"] = s.url

        if s.authors:
            fields["author"] = " and ".join([a.format_bibtex() for a in s.authors])
        if s.editors:
            fields["editor"] = " and ".join([e.format_bibtex() for e in s.editors])

        # Build entry
        body = ",\n".join([f"  {k} = {{{_escape_bibtex(v)}}}" for k, v in fields.items() if v])
        entries.append(f"@{entry_type}{{{key},\n{body}\n}}")
    return "\n\n".join(entries).strip() + ("\n" if entries else "")

def to_ris(sources: List[Source]) -> str:
    lines: List[str] = []
    def add(tag: str, val: Optional[str]):
        if val:
            lines.append(f"{tag}  - {val}")
    for s in sources:
        add("TY", s.entry_type_ris())
        if s.authors:
            for a in s.authors:
                add("AU", a.format_ris())
        if s.editors:
            for e in s.editors:
                add("ED", e.format_ris())
        add("TI", s.title)
        add("PY", s.year)
        add("JO", s.journal)
        add("T2", s.book_title)
        add("VL", s.volume)
        add("IS", s.issue)
        add("SP", s.pages.split("-")[0] if "-" in (s.pages or "") else s.pages)
        add("EP", s.pages.split("-")[1] if "-" in (s.pages or "") else None)
        add("DO", s.doi)
        add("UR", s.url)
        add("PB", s.publisher)
        add("CY", s.city)
        lines.append("ER  - ")
        lines.append("")
    return "\n".join(lines).strip() + ("\n" if lines else "")

def to_rows(sources: List[Source]):
    rows = []
    for s in sources:
        rows.append({
            "key": s.bibtex_key(),
            "type": s.type,
            "title": s.title,
            "year": s.year,
            "journal": s.journal,
            "book_title": s.book_title,
            "publisher": s.publisher,
            "city": s.city,
            "volume": s.volume,
            "issue": s.issue,
            "pages": s.pages,
            "doi": s.doi,
            "url": s.url,
            "authors": "; ".join([a.format_ris() for a in s.authors]),
        })
    return rows
