import cv2, glob, os
from pyzbar.pyzbar import decode

files = sorted(glob.glob("slip/*.jpg"))
det = cv2.QRCodeDetector()
for f in files:
    img = cv2.imread(f)
    name = os.path.basename(f)
    results = []
    # try pyzbar (handles many symbologies)
    for d in decode(img):
        results.append((d.type, d.data.decode('utf-8','replace')))
    # try opencv multi
    ok, infos, pts, _ = det.detectAndDecodeMulti(img)
    if ok:
        for s in infos:
            if s:
                results.append(("CV_QR", s))
    print(f"=== {name} ({img.shape[1]}x{img.shape[0]}) ===")
    if not results:
        print("  NO CODE FOUND")
    for t, s in results:
        print(f"  [{t}] {s}")
