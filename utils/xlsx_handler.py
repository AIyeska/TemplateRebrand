import base64
import zipfile
import os


def extract_images_xlsx(path):
    images = []
    with zipfile.ZipFile(path, "r") as z:
        media_files = [f for f in z.namelist() if f.startswith("xl/media/")]
        for mf in media_files:
            blob = z.read(mf)
            ext = mf.split(".")[-1].lower().replace("jpeg", "jpg")
            b64 = base64.b64encode(blob).decode()
            images.append({"id": mf, "data": b64, "ext": ext, "content_type": f"image/{ext}"})
    return images


def replace_images_xlsx(src, dst, replacements: dict):
    """replacements: {image_id (zip path like 'xl/media/image1.png') → new_image_bytes}"""
    tmp = dst + ".tmp"
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                zout.writestr(item, replacements[item.filename])
            else:
                zout.writestr(item, data)

    os.replace(tmp, dst)
