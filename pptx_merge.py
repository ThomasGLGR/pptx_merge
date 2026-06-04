from __future__ import annotations
import argparse
import hashlib
import posixpath
import re
import sys
import zipfile
from collections import deque
from typing import Dict, List, Optional, Set, Tuple
from lxml import etree

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PRES_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_TYPE_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
RT_SLIDE = REL_TYPE_BASE + "slide"
RT_SLIDE_MASTER = REL_TYPE_BASE + "slideMaster"
RT_NOTES_MASTER = REL_TYPE_BASE + "notesMaster"
CONTENT_TYPES_PART = "[Content_Types].xml"
PRESENTATION_PART = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
PRESENTATION_CHILD_ORDER = [
    "sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst",
    "sldSz", "notesSz", "smartTags", "embeddedFontLst", "custShowLst",
    "photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle",
    "modifyVerifier", "extLst",
]
XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def rels_path_for(part: str) -> str:
    d = posixpath.dirname(part)
    b = posixpath.basename(part)
    return f"{d}/_rels/{b}.rels" if d else f"_rels/{b}.rels"


def resolve_target(target: str, base_dir: str) -> str:
    return posixpath.normpath(posixpath.join(base_dir, target))


def relative_target(target_abs: str, base_dir: str) -> str:
    return posixpath.relpath(target_abs, base_dir)


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _mul_attr(el, name: str, factor: float) -> None:
    v = el.get(name)
    if v is None:
        return
    try:
        el.set(name, str(int(round(float(v) * factor))))
    except (TypeError, ValueError):
        pass


def scale_drawing_xml(data: bytes, fx: float, fy: float, fs: float) -> bytes:
    root = etree.fromstring(data)
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        local = etree.QName(el).localname
        if local in ("off", "chOff"):
            _mul_attr(el, "x", fx); _mul_attr(el, "y", fy)
        elif local in ("ext", "chExt"):
            _mul_attr(el, "cx", fx); _mul_attr(el, "cy", fy)
        elif local in ("rPr", "defRPr", "endParaRPr"):
            _mul_attr(el, "sz", fs)
        elif local == "ln":
            _mul_attr(el, "w", fs)
        elif local == "spcPts":
            _mul_attr(el, "val", fs)
        elif local == "bodyPr":
            _mul_attr(el, "lIns", fx); _mul_attr(el, "rIns", fx)
            _mul_attr(el, "tIns", fy); _mul_attr(el, "bIns", fy)
        elif local == "tab":
            _mul_attr(el, "pos", fx)
        elif local == "gridCol":
            _mul_attr(el, "w", fx)
        elif local == "tr":
            _mul_attr(el, "h", fy)
        for name in ("marL", "marR", "indent", "defTabSz"):
            if el.get(name) is not None:
                _mul_attr(el, name, fx)
        for name in ("dist", "blurRad"):
            if el.get(name) is not None:
                _mul_attr(el, name, fs)
    return XML_DECL + etree.tostring(root)


def scale_factors(src_cx, src_cy, tgt_cx, tgt_cy):
    fx = tgt_cx / src_cx
    fy = tgt_cy / src_cy
    fs = fx if abs(fx - fy) < 1e-6 else min(fx, fy)
    return fx, fy, fs


def _is_identity(fx, fy):
    return abs(fx - 1.0) < 1e-9 and abs(fy - 1.0) < 1e-9


class Package:
    def __init__(self, path: str):
        self.path = path
        self.parts: Dict[str, bytes] = {}
        with zipfile.ZipFile(path, "r") as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                name = info.filename.replace("\\", "/")
                self.parts[name] = z.read(info)

    def rels_of(self, part):
        rels_name = rels_path_for(part)
        if rels_name not in self.parts:
            return []
        root = etree.fromstring(self.parts[rels_name])
        out = []
        for rel in root.findall(qn(RELS_NS, "Relationship")):
            out.append((rel.get("Id"), rel.get("Type"), rel.get("Target"), rel.get("TargetMode")))
        return out

    def content_types(self):
        defaults, overrides = {}, {}
        root = etree.fromstring(self.parts[CONTENT_TYPES_PART])
        for d in root.findall(qn(CT_NS, "Default")):
            defaults[d.get("Extension").lower()] = d.get("ContentType")
        for o in root.findall(qn(CT_NS, "Override")):
            overrides[o.get("PartName").lstrip("/")] = o.get("ContentType")
        return defaults, overrides

    def ordered_slide_parts(self):
        return self._ordered_parts("sldIdLst", "sldId", RT_SLIDE)

    def ordered_master_parts(self):
        return self._ordered_parts("sldMasterIdLst", "sldMasterId", RT_SLIDE_MASTER)

    def _ordered_parts(self, lst_tag, item_tag, rel_type):
        root = etree.fromstring(self.parts[PRESENTATION_PART])
        lst = root.find(qn(PRES_NS, lst_tag))
        if lst is None:
            return []
        rid_to_target = {}
        for rid, rtype, target, mode in self.rels_of(PRESENTATION_PART):
            if mode != "External":
                rid_to_target[rid] = resolve_target(target, "ppt")
        result = []
        for item in lst.findall(qn(PRES_NS, item_tag)):
            rid = item.get(qn(R_NS, "id"))
            if rid in rid_to_target:
                result.append(rid_to_target[rid])
        return result

    def notes_master_part(self):
        for name in self.parts:
            if re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", name):
                return name
        return None

    def slide_size(self):
        root = etree.fromstring(self.parts[PRESENTATION_PART])
        sz = root.find(qn(PRES_NS, "sldSz"))
        return int(sz.get("cx")), int(sz.get("cy"))

    def media_hashes(self):
        out = {}
        for name, data in self.parts.items():
            if name.startswith("ppt/media/"):
                out[hashlib.sha256(data).hexdigest()] = name
        return out


class Numberer:
    CATEGORIES = [
        ("slideLayout", "ppt/slideLayouts", "slideLayout", ".xml"),
        ("slideMaster", "ppt/slideMasters", "slideMaster", ".xml"),
        ("notesSlide", "ppt/notesSlides", "notesSlide", ".xml"),
        ("notesMaster", "ppt/notesMasters", "notesMaster", ".xml"),
        ("slide", "ppt/slides", "slide", ".xml"),
        ("theme", "ppt/theme", "theme", ".xml"),
    ]

    def __init__(self, existing):
        self.existing = set(existing)
        self.next_index = {}
        self.misc_counter = 0
        for cat, folder, prefix, ext in self.CATEGORIES:
            pat = re.compile(re.escape(folder) + "/" + re.escape(prefix) + r"(\d+)" + re.escape(ext) + "$")
            mx = 0
            for name in self.existing:
                m = pat.match(name)
                if m:
                    mx = max(mx, int(m.group(1)))
            self.next_index[cat] = mx + 1

    @staticmethod
    def classify(part):
        base = posixpath.basename(part)
        folder = posixpath.dirname(part)
        if folder == "ppt/media":
            return ("media", folder, "image", posixpath.splitext(base)[1])
        for cat, _f, prefix, ext in sorted(Numberer.CATEGORIES, key=lambda c: -len(c[2])):
            if re.match(re.escape(prefix) + r"\d+\.xml$", base):
                return (cat, folder, prefix, ext)
        return None

    def assign(self, part):
        info = self.classify(part)
        if info is None:
            folder = posixpath.dirname(part)
            base = posixpath.basename(part)
            while True:
                self.misc_counter += 1
                candidate = f"{folder}/imp{self.misc_counter}_{base}"
                if candidate not in self.existing:
                    break
        elif info[0] == "media":
            _cat, folder, prefix, ext = info
            while True:
                idx = self.next_index.get("media", 1)
                self.next_index["media"] = idx + 1
                candidate = f"{folder}/{prefix}{idx}{ext}"
                if candidate not in self.existing:
                    break
        else:
            cat, folder, prefix, ext = info
            while True:
                idx = self.next_index[cat]
                self.next_index[cat] = idx + 1
                candidate = f"{folder}/{prefix}{idx}{ext}"
                if candidate not in self.existing:
                    break
        self.existing.add(candidate)
        return candidate


class Merger:
    def __init__(self, pkg_a, pkg_b, slides_a=None, slides_b=None,
                 insert_index=-1, target_size="first"):
        self.a = pkg_a
        self.b = pkg_b
        self.slides_a = slides_a
        self.slides_b = slides_b
        self.insert_index = insert_index
        self.a_size = pkg_a.slide_size()
        self.b_size = pkg_b.slide_size()
        self.target_size = self._choose_target(target_size)
        self.fa = scale_factors(*self.a_size, *self.target_size)
        self.fb = scale_factors(*self.b_size, *self.target_size)
        self.merged = dict(pkg_a.parts)
        self.numberer = Numberer(set(self.merged.keys()))
        self.rename_map = {}
        self.to_copy = set()
        self.reused = set()
        self.a_media_hashes = pkg_a.media_hashes()
        self.a_notes_master = pkg_a.notes_master_part()
        self.imported_notes_master_b = None
        self.defaults, self.overrides = pkg_a.content_types()
        self.b_defaults, self.b_overrides = pkg_b.content_types()

    def _choose_target(self, policy):
        a, b = self.a_size, self.b_size
        if policy == "first":
            return a
        if policy == "second":
            return b
        if policy == "largest":
            return a if a[0] * a[1] >= b[0] * b[1] else b
        if policy == "smallest":
            return a if a[0] * a[1] <= b[0] * b[1] else b
        raise ValueError(f"target_size inconnu : {policy!r}")

    @staticmethod
    def _is_scalable_category(part):
        info = Numberer.classify(part)
        return info is not None and info[0] in ("slide", "slideLayout", "slideMaster")

    def _is_notes_master(self, part):
        return re.match(r"ppt/notesMasters/notesMaster\d+\.xml$", part) is not None

    def _content_type_of_b_part(self, part):
        if part in self.b_overrides:
            return self.b_overrides[part]
        return None

    def _assign_or_reuse(self, part):
        if part in self.rename_map:
            return
        if part.startswith("ppt/media/"):
            digest = hashlib.sha256(self.b.parts[part]).hexdigest()
            if digest in self.a_media_hashes:
                self.rename_map[part] = self.a_media_hashes[digest]
                self.reused.add(part)
                return
        if self._is_notes_master(part) and self.a_notes_master is not None:
            self.rename_map[part] = self.a_notes_master
            self.reused.add(part)
            return
        new_name = self.numberer.assign(part)
        self.rename_map[part] = new_name
        self.to_copy.add(part)
        if self._is_notes_master(part):
            self.imported_notes_master_b = part

    def _collect_closure(self):
        all_b_slides = self.b.ordered_slide_parts()
        if self.slides_b is not None:
            slide_order = [all_b_slides[i - 1] for i in self.slides_b if 1 <= i <= len(all_b_slides)]
        else:
            slide_order = all_b_slides
        queue = deque(slide_order)
        seen = set()
        while queue:
            part = queue.popleft()
            if part in seen:
                continue
            seen.add(part)
            if part not in self.b.parts:
                continue
            self._assign_or_reuse(part)
            if part in self.reused:
                continue
            for _rid, _rtype, target, mode in self.b.rels_of(part):
                if mode == "External":
                    continue
                tgt_abs = resolve_target(target, posixpath.dirname(part))
                if tgt_abs not in seen:
                    queue.append(tgt_abs)
        return slide_order

    def _copy_parts_and_rewrite_rels(self):
        fbx, fby, fbs = self.fb
        scale_b = not _is_identity(fbx, fby)
        for old_part in self.to_copy:
            new_part = self.rename_map[old_part]
            data = self.b.parts[old_part]
            if scale_b and self._is_scalable_category(old_part):
                data = scale_drawing_xml(data, fbx, fby, fbs)
            self.merged[new_part] = data
            ct = self._content_type_of_b_part(old_part)
            if ct is not None:
                self.overrides[new_part] = ct
            old_rels = rels_path_for(old_part)
            if old_rels in self.b.parts:
                new_rels = rels_path_for(new_part)
                self.merged[new_rels] = self._rewrite_rels(
                    self.b.parts[old_rels], owner_old=old_part, owner_new=new_part)

    def _scale_base_if_needed(self):
        fax, fay, fas = self.fa
        if _is_identity(fax, fay):
            return
        for name in list(self.merged.keys()):
            if name in self.a.parts and self._is_scalable_category(name):
                self.merged[name] = scale_drawing_xml(self.merged[name], fax, fay, fas)

    def _rewrite_rels(self, data, owner_old, owner_new):
        root = etree.fromstring(data)
        base_old = posixpath.dirname(owner_old)
        base_new = posixpath.dirname(owner_new)
        for rel in root.findall(qn(RELS_NS, "Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            tgt_abs = resolve_target(target, base_old)
            new_abs = self.rename_map.get(tgt_abs)
            if new_abs is None:
                continue
            rel.set("Target", relative_target(new_abs, base_new))
        return XML_DECL + etree.tostring(root)

    def _merge_content_types(self):
        for ext, ct in self.b_defaults.items():
            self.defaults.setdefault(ext, ct)
        nsmap = {None: CT_NS}
        root = etree.Element(qn(CT_NS, "Types"), nsmap=nsmap)
        for ext in sorted(self.defaults):
            d = etree.SubElement(root, qn(CT_NS, "Default"))
            d.set("Extension", ext)
            d.set("ContentType", self.defaults[ext])
        for pn in sorted(self.overrides):
            o = etree.SubElement(root, qn(CT_NS, "Override"))
            o.set("PartName", "/" + pn)
            o.set("ContentType", self.overrides[pn])
        self.merged[CONTENT_TYPES_PART] = XML_DECL + etree.tostring(root)

    def _next_rid(self, used):
        mx = 0
        for rid in used:
            m = re.match(r"rId(\d+)$", rid)
            if m:
                mx = max(mx, int(m.group(1)))
        n = mx + 1
        while f"rId{n}" in used:
            n += 1
        used.add(f"rId{n}")
        return f"rId{n}"

    def _ensure_child_in_order(self, root, tag):
        el = root.find(qn(PRES_NS, tag))
        if el is not None:
            return el
        el = etree.Element(qn(PRES_NS, tag))
        order = PRESENTATION_CHILD_ORDER
        idx_self = order.index(tag)
        insert_pos = len(root)
        for i, child in enumerate(root):
            local = etree.QName(child).localname
            if local in order and order.index(local) > idx_self:
                insert_pos = i
                break
        root.insert(insert_pos, el)
        return el

    def _merge_presentation(self, slide_order):
        rels_root = etree.fromstring(self.merged[PRESENTATION_RELS])
        used_rids = {rel.get("Id") for rel in rels_root.findall(qn(RELS_NS, "Relationship"))}

        def add_pres_rel(target_abs, rel_type):
            rid = self._next_rid(used_rids)
            rel = etree.SubElement(rels_root, qn(RELS_NS, "Relationship"))
            rel.set("Id", rid)
            rel.set("Type", rel_type)
            rel.set("Target", relative_target(target_abs, "ppt"))
            return rid

        pres_root = etree.fromstring(self.merged[PRESENTATION_PART])
        sz = pres_root.find(qn(PRES_NS, "sldSz"))
        if sz is not None:
            sz.set("cx", str(self.target_size[0]))
            sz.set("cy", str(self.target_size[1]))

        sld_lst = self._ensure_child_in_order(pres_root, "sldIdLst")
        original_a_slds = list(sld_lst.findall(qn(PRES_NS, "sldId")))
        max_sld_id = 255
        for item in original_a_slds:
            try:
                max_sld_id = max(max_sld_id, int(item.get("id")))
            except (TypeError, ValueError):
                pass

        if self.slides_a is not None:
            filtered_a_slds = []
            seen_ids = set()
            for i in self.slides_a:
                if 1 <= i <= len(original_a_slds):
                    sld = original_a_slds[i - 1]
                    sld_id = sld.get("id")
                    if sld_id not in seen_ids:
                        filtered_a_slds.append(sld)
                        seen_ids.add(sld_id)
        else:
            filtered_a_slds = original_a_slds

        kept_rids = {sld.get(qn(R_NS, "id")) for sld in filtered_a_slds}
        for sld in original_a_slds:
            rid = sld.get(qn(R_NS, "id"))
            if rid not in kept_rids:
                for rel in rels_root.findall(qn(RELS_NS, "Relationship")):
                    if rel.get("Id") == rid:
                        rels_root.remove(rel)

        cust_show = pres_root.find(qn(PRES_NS, "custShowLst"))
        if cust_show is not None:
            pres_root.remove(cust_show)

        for child in original_a_slds:
            sld_lst.remove(child)

        idx = self.insert_index
        if idx is None or idx < 0 or idx > len(filtered_a_slds):
            idx = len(filtered_a_slds)

        for sld in filtered_a_slds[:idx]:
            sld_lst.append(sld)
        for b_slide in slide_order:
            new_part = self.rename_map[b_slide]
            rid = add_pres_rel(new_part, RT_SLIDE)
            max_sld_id += 1
            new_sld = etree.SubElement(sld_lst, qn(PRES_NS, "sldId"))
            new_sld.set("id", str(max_sld_id))
            new_sld.set(qn(R_NS, "id"), rid)
        for sld in filtered_a_slds[idx:]:
            sld_lst.append(sld)

        master_lst = self._ensure_child_in_order(pres_root, "sldMasterIdLst")

        GMIN = 2147483648
        b_imported_masters = [
            self.rename_map[bm] for bm in self.b.ordered_master_parts()
            if bm in self.rename_map and bm in self.to_copy
        ]
        b_imported_set = set(b_imported_masters)

        used_global: Set[int] = set()
        for item in master_lst.findall(qn(PRES_NS, "sldMasterId")):
            try:
                used_global.add(int(item.get("id")))
            except (TypeError, ValueError):
                pass
        for mp in [n for n in self.merged
                   if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]:
            if mp in b_imported_set:
                continue  
            mroot = etree.fromstring(self.merged[mp])
            lay_lst = mroot.find(qn(PRES_NS, "sldLayoutIdLst"))
            if lay_lst is not None:
                for it in lay_lst.findall(qn(PRES_NS, "sldLayoutId")):
                    try:
                        used_global.add(int(it.get("id")))
                    except (TypeError, ValueError):
                        pass

        def fresh_id() -> int:
            nid = (max(used_global) + 1) if used_global else GMIN
            if nid < GMIN:
                nid = GMIN
            while nid in used_global:
                nid += 1
            used_global.add(nid)
            return nid

        for b_master in self.b.ordered_master_parts():
            if b_master not in self.rename_map or b_master not in self.to_copy:
                continue
            new_part = self.rename_map[b_master]
            mroot = etree.fromstring(self.merged[new_part])
            lay_lst = mroot.find(qn(PRES_NS, "sldLayoutIdLst"))
            if lay_lst is not None:
                for it in lay_lst.findall(qn(PRES_NS, "sldLayoutId")):
                    it.set("id", str(fresh_id()))
            self.merged[new_part] = XML_DECL + etree.tostring(mroot)
            rid = add_pres_rel(new_part, RT_SLIDE_MASTER)
            m = etree.SubElement(master_lst, qn(PRES_NS, "sldMasterId"))
            m.set("id", str(fresh_id()))
            m.set(qn(R_NS, "id"), rid)

        if self.imported_notes_master_b is not None:
            new_part = self.rename_map[self.imported_notes_master_b]
            rid = add_pres_rel(new_part, RT_NOTES_MASTER)
            notes_lst = self._ensure_child_in_order(pres_root, "notesMasterIdLst")
            nm = etree.SubElement(notes_lst, qn(PRES_NS, "notesMasterId"))
            nm.set(qn(R_NS, "id"), rid)

        self.merged[PRESENTATION_RELS] = XML_DECL + etree.tostring(rels_root)
        self.merged[PRESENTATION_PART] = XML_DECL + etree.tostring(pres_root)
        app_part = "docProps/app.xml"
        if app_part in self.merged:
            try:
                app_root = etree.fromstring(self.merged[app_part])
                for elem in app_root.iter():
                    if elem.tag.endswith("}Slides"):
                        elem.text = str(len(filtered_a_slds) + len(slide_order))
                        break
                self.merged[app_part] = XML_DECL + etree.tostring(app_root)
            except Exception:
                pass

    def merge(self):
        slide_order = self._collect_closure()
        self._scale_base_if_needed()
        self._copy_parts_and_rewrite_rels()
        self._merge_content_types()
        self._merge_presentation(slide_order)
        return self.merged


def write_pptx(parts, output_path):
    ordered_names = []
    if CONTENT_TYPES_PART in parts:
        ordered_names.append(CONTENT_TYPES_PART)
    if "_rels/.rels" in parts:
        ordered_names.append("_rels/.rels")
    for name in parts:
        if name not in (CONTENT_TYPES_PART, "_rels/.rels"):
            ordered_names.append(name)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in ordered_names:
            z.writestr(name, parts[name])


def merge_pptx(file_a, file_b, output, target_size="first"):
    pkg_a = Package(file_a)
    pkg_b = Package(file_b)
    merged_parts = Merger(pkg_a, pkg_b, target_size=target_size).merge()
    write_pptx(merged_parts, output)
    return output


def merge_specific_slides(file_a, file_b, slides_a, slides_b, insert_index,
                          output="fusion_specifique.pptx", target_size="first"):
    pkg_a = Package(file_a)
    pkg_b = Package(file_b)
    merger = Merger(pkg_a, pkg_b, slides_a=slides_a, slides_b=slides_b,
                    insert_index=insert_index, target_size=target_size)
    merged_parts = merger.merge()
    write_pptx(merged_parts, output)
    return output

def _parse_indices(spec: Optional[str]) -> Optional[List[int]]:
    """Convertit "1,3,5-8" en [1,3,5,6,7,8]. None si non fourni."""
    if not spec:
        return None
    out: List[int] = []
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(chunk))
    return out


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fusionne deux fichiers PPTX en un seul (OOXML pur, 100 % Linux)."
    )
    parser.add_argument("file_a", help="Premier PPTX (ses diapos viennent en premier)")
    parser.add_argument("file_b", help="Second PPTX (ajouté après A)")
    parser.add_argument("-o", "--output", default="fusion.pptx",
                        help="Fichier de sortie (.pptx)")
    parser.add_argument("--target-size",
                        choices=["first", "second", "largest", "smallest"],
                        default="first",
                        help="Taille de diapo finale si A et B diffèrent (défaut : first)")
    parser.add_argument("--slides-a", default=None,
                        help="Diapos de A à garder, ex '1,3,5-8' (défaut : toutes)")
    parser.add_argument("--slides-b", default=None,
                        help="Diapos de B à importer, ex '2,4' (défaut : toutes)")
    parser.add_argument("--insert-index", type=int, default=-1,
                        help="Position d'insertion des diapos de B dans A "
                             "(0 = au début ; défaut : à la fin de A)")
    args = parser.parse_args(argv)

    slides_a = _parse_indices(args.slides_a)
    slides_b = _parse_indices(args.slides_b)

    if slides_a is None and slides_b is None and args.insert_index == -1:
        out = merge_pptx(args.file_a, args.file_b, args.output,
                         target_size=args.target_size)
    else:
        out = merge_specific_slides(
            args.file_a, args.file_b,
            slides_a=slides_a, slides_b=slides_b,
            insert_index=args.insert_index,
            output=args.output, target_size=args.target_size,
        )
    print(f"[OK] Fusion écrite : {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())