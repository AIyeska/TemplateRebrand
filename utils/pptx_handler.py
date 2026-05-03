import base64
import zipfile
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extract_images_pptx(path):
    prs = Presentation(path)
    images = []
    seen = set()

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                img = shape.image
                img_hash = hash(img.blob)
                if img_hash in seen:
                    continue
                seen.add(img_hash)
                ext = img.ext.lower().replace("jpeg", "jpg")
                b64 = base64.b64encode(img.blob).decode()
                images.append({
                    "id": f"slide{slide_idx}_{shape.shape_id}",
                    "data": b64,
                    "ext": ext,
                    "content_type": img.content_type,
                })

    return images


def replace_images_pptx(src, dst, replacements: dict):
    """replacements: {image_id (e.g. 'slide0_5') → new_image_bytes}"""
    prs = Presentation(src)
    path_to_bytes = {}

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                key = f"slide{slide_idx}_{shape.shape_id}"
                if key in replacements:
                    rId = shape._element.blipFill.blip.attrib.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )
                    if rId:
                        part_name = str(slide.part.rels[rId].target_part.partname).lstrip("/")
                        path_to_bytes[part_name] = replacements[key]

    tmp = dst + ".tmp"
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in path_to_bytes:
                zout.writestr(item, path_to_bytes[item.filename])
            else:
                zout.writestr(item, data)

    os.replace(tmp, dst)
