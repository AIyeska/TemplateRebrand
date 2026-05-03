import base64
import shutil
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
                    "slide": slide_idx,
                    "shape_id": shape.shape_id,
                })

    return images


def replace_images_pptx(src, dst, selected_ids, logo_bytes, logo_ext):
    shutil.copy2(src, dst)
    prs = Presentation(dst)

    id_set = set(selected_ids)
    mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp"}
    new_content_type = mime_map.get(logo_ext.lower(), "image/png")

    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                key = f"slide{slide_idx}_{shape.shape_id}"
                if key in id_set:
                    img_part = shape._element.blipFill.blip.embed
                    rel = slide.part.rels[img_part]
                    img_target = rel.target_part
                    img_target._blob = logo_bytes
                    img_target.content_type = new_content_type

    prs.save(dst)
