import os
import shutil
import zipfile
import tempfile
from xml.etree import ElementTree as ET
from typing import Dict, Tuple, Optional, Set

# WordprocessingML namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}
ET.register_namespace('w', NS['w'])
ET.register_namespace('r', NS['r'])

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('', CT_NS)

HEADER_TYPES = ['default', 'first', 'even']
FOOTER_TYPES = ['default', 'first', 'even']


def _read_zip_xml(z: zipfile.ZipFile, path: str) -> Optional[ET.ElementTree]:
    try:
        with z.open(path) as f:
            return ET.parse(f)
    except KeyError:
        return None


def _write_zip_xml(z: zipfile.ZipFile, path: str, tree: ET.ElementTree):
    xml_bytes = ET.tostring(tree.getroot(), encoding='utf-8', xml_declaration=True)
    # zipfile has no direct overwrite; we rebuild via temp dir write later
    z.writestr(path, xml_bytes)


def _collect_headers_footers_from_template(tpl_zip: zipfile.ZipFile) -> Tuple[Dict[str, bytes], Dict[str, bytes]]:
    headers: Dict[str, bytes] = {}
    footers: Dict[str, bytes] = {}

    for item in tpl_zip.namelist():
        if item.startswith('word/header') and item.endswith('.xml'):
            # header1.xml, header2.xml, ... we can't know types by index, so we map later by rels
            headers[item] = tpl_zip.read(item)
        elif item.startswith('word/footer') and item.endswith('.xml'):
            footers[item] = tpl_zip.read(item)
    return headers, footers


def _pick_first_part(parts: Dict[str, bytes]) -> Tuple[Optional[str], Optional[bytes]]:
    if not parts:
        return None, None
    # pick deterministic by sorted name
    name = sorted(parts.keys())[0]
    return name, parts[name]


def _map_rel_types(doc_zip: zipfile.ZipFile, rels_path: str) -> Dict[str, str]:
    # Map part path -> type (default/first/even) using sectPr rels in document.xml.rels
    mapping: Dict[str, str] = {}
    try:
        with doc_zip.open(rels_path) as f:
            rels_tree = ET.parse(f)
    except KeyError:
        return mapping

    for rel in rels_tree.getroot().iterfind('Relationship'):
        r_type = rel.get('Type', '')
        target = rel.get('Target', '')
        if 'header' in target:
            # We need to inspect document.xml to see which type maps to which Id; but simpler path is to
            # rely on sectPr in document.xml instead of rels types. We'll not fill here.
            pass
    return mapping


def _ensure_media_dependencies(src_zip: zipfile.ZipFile, tpl_zip: zipfile.ZipFile, out_zip: zipfile.ZipFile, parts: Dict[str, bytes]):
    # Copy related media (images) used by header/footer parts from template to output if missing
    # Heuristic: copy entire word/media from template if not present in source
    for item in tpl_zip.namelist():
        if item.startswith('word/media/'):
            if item not in out_zip.namelist():
                out_zip.writestr(item, tpl_zip.read(item))

def _replace_header_footer_parts(src_zip: zipfile.ZipFile, tpl_zip: zipfile.ZipFile, out_zip: zipfile.ZipFile):
    # Improved strategy:
    # - Skip copying header/footer parts and related rels from source
    # - Buffer document.xml, document.xml.rels, and [Content_Types].xml for modification
    # - Copy all other parts as-is
    # - Choose template header/footer (first ones), copy as /word/header1.xml and /word/footer1.xml
    #   and copy their rels to /word/_rels/header1.xml.rels, /word/_rels/footer1.xml.rels
    # - Copy word/media from template
    # - Update document.xml.rels to point to header1.xml/footer1.xml
    # - Update document.xml sectPr to reference the new r:ids
    # - Ensure content types overrides exist

    names = src_zip.namelist()

    doc_xml_bytes = None
    doc_rels_bytes = None
    content_types_bytes = None

    def is_header_footer_part(n: str) -> bool:
        return (n.startswith('word/header') or n.startswith('word/footer')) and n.endswith('.xml')

    def is_header_footer_rels(n: str) -> bool:
        base = os.path.basename(n)
        return n.startswith('word/_rels/') and (base.startswith('header') or base.startswith('footer')) and base.endswith('.rels')

    for n in names:
        if n == 'word/document.xml':
            with src_zip.open(n) as f:
                doc_xml_bytes = f.read()
            continue
        if n == 'word/_rels/document.xml.rels':
            with src_zip.open(n) as f:
                doc_rels_bytes = f.read()
            continue
        if n == '[Content_Types].xml':
            with src_zip.open(n) as f:
                content_types_bytes = f.read()
            continue
        if is_header_footer_part(n) or is_header_footer_rels(n):
            # skip copying these from source
            continue
        # copy others
        out_zip.writestr(n, src_zip.read(n))

    # Template parts
    tpl_headers, tpl_footers = _collect_headers_footers_from_template(tpl_zip)
    tpl_header_name, tpl_header_bytes = _pick_first_part(tpl_headers)
    tpl_footer_name, tpl_footer_bytes = _pick_first_part(tpl_footers)

    # Copy media from template
    _ensure_media_dependencies(src_zip, tpl_zip, out_zip, {})

    # Write selected header/footer parts and their rels
    header_rel_id = None
    footer_rel_id = None

    if tpl_header_bytes is not None:
        out_zip.writestr('word/header1.xml', tpl_header_bytes)
        # copy rels if presents
        if tpl_header_name:
            tpl_header_base = os.path.basename(tpl_header_name)
            tpl_header_rels = f'word/_rels/{tpl_header_base}.rels'
            try:
                data = tpl_zip.read(tpl_header_rels)
                out_zip.writestr('word/_rels/header1.xml.rels', data)
            except KeyError:
                pass

    if tpl_footer_bytes is not None:
        out_zip.writestr('word/footer1.xml', tpl_footer_bytes)
        if tpl_footer_name:
            tpl_footer_base = os.path.basename(tpl_footer_name)
            tpl_footer_rels = f'word/_rels/{tpl_footer_base}.rels'
            try:
                data = tpl_zip.read(tpl_footer_rels)
                out_zip.writestr('word/_rels/footer1.xml.rels', data)
            except KeyError:
                pass

    # Update/compose document.xml.rels
    if doc_rels_bytes is not None:
        rels_root = ET.fromstring(doc_rels_bytes)
    else:
        rels_root = ET.Element(f'{{{REL_NS}}}Relationships')

    # Collect existing ids and remove header/footer relationships
    used_ids: Set[str] = set()
    for rel in list(rels_root):
        if rel.tag.endswith('Relationship'):
            rid = rel.get('Id')
            if rid:
                used_ids.add(rid)
            rtype = rel.get('Type', '')
            if rtype.endswith('/header') or rtype.endswith('/footer'):
                rels_root.remove(rel)

    def new_rid(prefix: str) -> str:
        i = 9001
        while True:
            cand = f'rId{i}'
            if cand not in used_ids:
                used_ids.add(cand)
                return cand
            i += 1

    if tpl_header_bytes is not None:
        header_rel_id = new_rid('H')
        rel = ET.SubElement(rels_root, f'{{{REL_NS}}}Relationship')
        rel.set('Id', header_rel_id)
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
        rel.set('Target', 'header1.xml')

    if tpl_footer_bytes is not None:
        footer_rel_id = new_rid('F')
        rel = ET.SubElement(rels_root, f'{{{REL_NS}}}Relationship')
        rel.set('Id', footer_rel_id)
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
        rel.set('Target', 'footer1.xml')

    # Update/compose document.xml (sectPr refs)
    if doc_xml_bytes is not None:
        doc_root = ET.fromstring(doc_xml_bytes)
    else:
        # Create a minimal document if missing (shouldn't happen)
        doc_root = ET.Element(f'{{{NS["w"]}}}document')
        body = ET.SubElement(doc_root, f'{{{NS["w"]}}}body')
        ET.SubElement(body, f'{{{NS["w"]}}}sectPr')

    body = doc_root.find('w:body', NS)
    if body is None:
        body = ET.SubElement(doc_root, f'{{{NS["w"]}}}body')
    sect_prs = body.findall('.//w:sectPr', NS)
    if not sect_prs:
        sect_prs = [ET.SubElement(body, f'{{{NS["w"]}}}sectPr')]

    for sect in sect_prs:
        # remove existing header/footer references
        for child in list(sect):
            if child.tag == f'{{{NS["w"]}}}headerReference' or child.tag == f'{{{NS["w"]}}}footerReference':
                sect.remove(child)
        if header_rel_id is not None:
            h = ET.SubElement(sect, f'{{{NS["w"]}}}headerReference')
            h.set(f'{{{NS["w"]}}}type', 'default')  # note: in OOXML this is w:type attribute
            h.set(f'{{{NS["r"]}}}id', header_rel_id)
        if footer_rel_id is not None:
            f = ET.SubElement(sect, f'{{{NS["w"]}}}footerReference')
            f.set(f'{{{NS["w"]}}}type', 'default')
            f.set(f'{{{NS["r"]}}}id', footer_rel_id)

    # Update/compose [Content_Types].xml
    if content_types_bytes is not None:
        ct_root = ET.fromstring(content_types_bytes)
    else:
        ct_root = ET.Element(f'{{{CT_NS}}}Types')

    # Ensure overrides present
    def ensure_override(part_name: str, content_type: str):
        for ov in ct_root.findall(f'{{{CT_NS}}}Override'):
            if ov.get('PartName') == part_name:
                ov.set('ContentType', content_type)
                return
        ov = ET.SubElement(ct_root, f'{{{CT_NS}}}Override')
        ov.set('PartName', part_name)
        ov.set('ContentType', content_type)

    if tpl_header_bytes is not None:
        ensure_override('/word/header1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
    if tpl_footer_bytes is not None:
        ensure_override('/word/footer1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')

    # Finally, write modified core parts
    out_zip.writestr('[Content_Types].xml', ET.tostring(ct_root, encoding='utf-8', xml_declaration=True))
    out_zip.writestr('word/_rels/document.xml.rels', ET.tostring(rels_root, encoding='utf-8', xml_declaration=True))
    out_zip.writestr('word/document.xml', ET.tostring(doc_root, encoding='utf-8', xml_declaration=True))


def apply_template_to_docx(src_path: str, tpl_path: str, dst_path: str):
    """
    Replace all header/footer parts in src with those from template, preserving other content. Outputs to dst_path.
    Only supports .docx (ZIP-based Office Open XML).
    """
    if not src_path.lower().endswith('.docx') or not tpl_path.lower().endswith('.docx'):
        raise ValueError('仅支持 .docx 文件')

    tmp_fd, tmp_out = tempfile.mkstemp(suffix='.docx')
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(src_path, 'r') as src_zip, \
             zipfile.ZipFile(tpl_path, 'r') as tpl_zip, \
             zipfile.ZipFile(tmp_out, 'w', compression=zipfile.ZIP_DEFLATED) as out_zip:

            _replace_header_footer_parts(src_zip, tpl_zip, out_zip)

        # move to destination
        os.makedirs(os.path.dirname(os.path.abspath(dst_path)), exist_ok=True)
        shutil.move(tmp_out, dst_path)
    except Exception:
        if os.path.exists(tmp_out):
            os.remove(tmp_out)
        raise
