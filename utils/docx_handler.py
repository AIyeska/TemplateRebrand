import base64
import shutil
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT


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


def replace_images_docx(src, dst, selected_ids, logo_bytes, logo_ext):
    shutil.copy2(src, dst)
    doc = Document(dst)

    mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp"}
    new_content_type = mime_map.get(logo_ext.lower(), "image/png")

    for rel in doc.part.rels.values():
        if "image" in rel.reltype and rel.rId in selected_ids:
            part = rel.target_part
            part._blob = logo_bytes
            part.content_type = new_content_type

    doc.save(dst)
