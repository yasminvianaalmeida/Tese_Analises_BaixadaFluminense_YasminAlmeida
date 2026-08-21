# OCR resumivel: salva cada pagina num JSON-cache; roda em chamadas curtas ate terminar.
import fitz, pytesseract, io, json, os, sys, time, unicodedata, re
from PIL import Image
import _ocr_env
_ok, _info = _ocr_env.configurar_ocr()
if not _ok:
    raise SystemExit(f'OCR indisponivel: {_info}')

def norm(t):
    t=unicodedata.normalize("NFKD",t); t="".join(c for c in t if not unicodedata.combining(c))
    t=t.lower(); t=re.sub(r"-\n","",t); t=re.sub(r"\s*\n\s*"," ",t); return re.sub(r"[ \t]+"," ",t)

pdf=sys.argv[1]; cache=sys.argv[2]; budget=float(sys.argv[3]) if len(sys.argv)>3 else 40
data=json.load(open(cache, encoding="utf-8")) if os.path.exists(cache) else {"paginas":{}}
d=fitz.open(pdf); n=d.page_count; t0=time.time(); feito=0
for i in range(n):
    if str(i) in data["paginas"]: continue
    if time.time()-t0>budget: break
    pg=d[i]; pix=pg.get_pixmap(matrix=fitz.Matrix(3.0,3.0))
    img=Image.open(io.BytesIO(pix.tobytes("png")))
    txt=pytesseract.image_to_string(img,lang="eng",config="--oem 1 --psm 3")
    data["paginas"][str(i)]=norm(txt); feito+=1
    json.dump(data,open(cache,"w",encoding="utf-8"),ensure_ascii=False)
print(f"{os.path.basename(pdf)}: {len(data['paginas'])}/{n} paginas (novas={feito})")
