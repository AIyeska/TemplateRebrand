import base64
import shutil
import zipfile
import os
from docx import Document


def extract_images_docx(path):
    doc = Document(path)
    images = []
    seen = set()

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            part = rel.target_part
            rId = rel.rId
            if rId in seen:
                continue
            seen.add(rId)
            blob = part.blob
            content_type = part.content_type or "image/png"
            ext = content_type.split("/")[-1].replace("jpeg", "jpg")
            b64 = base64.b64encode(blob).decode()
            images.append({"id": rId, "data": b64, "ext": ext,
                            "content_type": content_type, "_part_name": part.partname})

    return images


def replace_images_docx(src, dst, selected_ids, logo_bytes, logo_ext):
    """Replace selected images using direct zip manipulation — avoids python-docx setter limitations."""
    # Build a map of rId → internal zip path from the live Document
    doc = Document(src)
    rid_to_partname = {}
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            rid_to_partname[rel.rId] = str(rel.target_part.partname).lstrip("/")

    target_parts = {rid_to_partname[rid] for rid in selected_ids if rid in rid_to_partname}

    tmp = dst + ".tmp"
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in target_parts:
                zout.writestr(item, logo_bytes)
            else:
                zout.writestr(item, data)

    os.replace(tmp, dst)
