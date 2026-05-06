from PIL import Image
import os

FOLDER   = "HMI"
OUTPUT   = "data/HMI_pages.pdf"

images = []
files  = sorted([f for f in os.listdir(FOLDER) if f.endswith('.png')])

print(f"Ditemukan {len(files)} file PNG...")

for i, filename in enumerate(files, 1):
    path = os.path.join(FOLDER, filename)
    img  = Image.open(path).convert('RGB')
    images.append(img)
    print(f"  [{i}/{len(files)}] {filename}")

if images:
    images[0].save(
        OUTPUT,
        save_all=True,
        append_images=images[1:]
    )
    print(f"\n✅ Selesai! PDF disimpan di: {OUTPUT}")
else:
    print("❌ Tidak ada file PNG ditemukan!")