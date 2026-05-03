import base64
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
            images.append({"id": rId, "data": b64, "ext": ext, "content_type": content_type})

    return images


def replace_images_docx(src, dst, replacements: dict):
    """replacements: {image_id (rId) → new_image_bytes}"""
    doc = Document(src)
    rid_to_partname = {}
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            rid_to_partname[rel.rId] = str(rel.target_part.partname).lstrip("/")

    # Build map: zip-internal-path → new bytes
    path_to_bytes = {}
    for rid, new_bytes in replacements.items():
        if rid in rid_to_partname:
            path_to_bytes[rid_to_partname[rid]] = new_bytes

    tmp = dst + ".tmp"
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in path_to_bytes:
                zout.writestr(item, path_to_bytes[item.filename])
            else:
                zout.writestr(item, data)

    os.replace(tmp, dst)
