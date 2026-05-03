import base64
import shutil
import zipfile
import os
import io


def extract_images_xlsx(path):
    images = []
    with zipfile.ZipFile(path, "r") as z:
        media_files = [f for f in z.namelist() if f.startswith("xl/media/")]
        for i, mf in enumerate(media_files):
            blob = z.read(mf)
            ext = mf.split(".")[-1].lower().replace("jpeg", "jpg")
            b64 = base64.b64encode(blob).decode()
            images.append({"id": mf, "data": b64, "ext": ext, "content_type": f"image/{ext}"})
    return images


def replace_images_xlsx(src, dst, selected_ids, logo_bytes, logo_ext):
    shutil.copy2(src, dst)
    id_set = set(selected_ids)

    tmp = dst + ".tmp.zip"
    with zipfile.ZipFile(dst, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in id_set:
                # Keep original extension in filename, replace content
                name_parts = item.filename.rsplit(".", 1)
                new_name = name_parts[0] + "." + logo_ext if len(name_parts) == 2 else item.filename
                # Use same filename to keep XML references intact
                zout.writestr(item, logo_bytes)
            else:
                zout.writestr(item, data)

    os.replace(tmp, dst)
