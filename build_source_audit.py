#!/usr/bin/env python3
"""Audit: ORIGINAL source (website PDF render, or original Word doc for
word-source policies) on the left vs FINAL accessible docx render on the right."""
import glob,os,re,subprocess,html,json,shutil
from docx import Document
from docx.oxml.ns import qn
HERE=os.path.dirname(os.path.abspath(__file__))
BASE=os.path.expanduser("~/Desktop/Online Ed")
FIN=os.path.join(BASE,"Final Accessible Policy Uploads")
IMG=os.path.join(HERE,"img"); os.makedirs(IMG,exist_ok=True)
SOFFICE="/opt/homebrew/bin/soffice"; RES="100"
OLD_DIRS=[os.path.join(BASE,"Policy_Review_Packet","audit_img"), os.path.join(BASE,"Policy_Review_5000","audit_img")]
WORD_ORIG={ "AP3775": os.path.join(BASE,"AP and BP Additions","AP 3775 Artificial Intelligence (AI) 4-27-26 DRAFT (1).docx"),
            "BP3775": os.path.join(BASE,"AP and BP Additions","BP 3775 Artificial Intelligence (AI) 4-27-26 DRAFT (1).docx") }
results={ (r['sub'],r['name']):r for r in json.load(open("/tmp/sf_results.json")) }
def esc(s): return html.escape(str(s),quote=True)
def code(n):
    m=re.match(r'^([A-Z]+)\s*(\d+)',n); return (m.group(1)+m.group(2)) if m else n[:8]

docs=[]
for sub in ("3000","4000","5000"):
    for f in sorted(glob.glob(os.path.join(FIN,sub,"*.docx"))):
        if os.path.basename(f).startswith('~$'): continue
        docs.append((sub,f))

# ---- LEFT: copy cached original renders (or render the original Word doc once) ----
for sub,f in docs:
    c=code(os.path.basename(f))
    if glob.glob(os.path.join(IMG,c+"_src-*.png")): continue
    hits=[]
    for d in OLD_DIRS:
        hits=sorted(glob.glob(os.path.join(d,c+"_old-*.png")))
        if hits: break
    if hits:
        for h in hits:
            n=re.search(r'-(\d+)\.png$',h).group(1)
            shutil.copy2(h, os.path.join(IMG,f"{c}_src-{int(n)}.png"))
    elif c in WORD_ORIG and os.path.exists(WORD_ORIG[c]):
        tmp=os.path.join(HERE,"_t"); os.makedirs(tmp,exist_ok=True)
        subprocess.run([SOFFICE,"--headless","--convert-to","pdf","--outdir",tmp,WORD_ORIG[c]],capture_output=True,timeout=180)
        pdf=glob.glob(os.path.join(tmp,"*.pdf"))
        if pdf: subprocess.run(["pdftoppm","-png","-r",RES,pdf[0],os.path.join(IMG,c+"_src")],capture_output=True)
        shutil.rmtree(tmp,ignore_errors=True)

# ---- RIGHT: fresh render of the current Final docs ----
if not glob.glob(os.path.join(IMG,"*_new-*.png")):
    tmp=os.path.join(HERE,"_pdf"); os.makedirs(tmp,exist_ok=True)
    files=[f for _,f in docs]
    for k in range(0,len(files),20):
        subprocess.run([SOFFICE,"--headless","--convert-to","pdf","--outdir",tmp]+files[k:k+20],capture_output=True,timeout=600)
    for sub,f in docs:
        c=code(os.path.basename(f))
        pdf=os.path.join(tmp,os.path.splitext(os.path.basename(f))[0]+".pdf")
        if os.path.exists(pdf):
            subprocess.run(["pdftoppm","-png","-r",RES,pdf,os.path.join(IMG,c+"_new")],capture_output=True)
    shutil.rmtree(tmp,ignore_errors=True)

def pages(c,side):
    hits=glob.glob(os.path.join(IMG,f"{c}_{side}-*.png"))
    return [os.path.basename(x) for x in sorted(hits,key=lambda x:int(re.search(r'-(\d+)\.png$',x).group(1)))]
def src(name):
    try: v=int(os.path.getmtime(os.path.join(IMG,name)))
    except OSError: v=0
    return f"img/{name}?v={v}"

def metrics(sub,f):
    d=Document(f)
    hs=[int(p.style.name.split()[1]) for p in d.paragraphs if p.style.name.startswith("Heading") and p.text.strip()]
    acc=(hs.count(1)==1 and not any(hs[i]>hs[i-1]+1 for i in range(1,len(hs))))
    r=results.get((sub,os.path.basename(f)),{})
    return dict(h=(hs.count(1),hs.count(2),hs.count(3)),acc=acc,rlists=r.get('real_lists',0),
                flags=r.get('flags',[]),irr=r.get('irregular_marker_paras',0))

secs={"3000":[],"4000":[],"5000":[]}; toc={"3000":[],"4000":[],"5000":[]}
n_acc=n_flag=tot_lists=0
for sub,f in docs:
    c=code(os.path.basename(f)); m=metrics(sub,f)
    n_acc+=m['acc']; tot_lists+=m['rlists']
    if m['flags'] or m['irr']: n_flag+=1
    toc[sub].append(f'<a href="#{c}" style="margin:2px 6px;color:#003366;text-decoration:none;font-size:12.5px;">{esc(c[:2]+" "+c[2:])}</a>')
    badges=(f'<span style="background:{"#1d6b3f" if m["acc"] else "#b3771a"};color:#fff;padding:2px 9px;border-radius:11px;font-size:11px;font-weight:700;">{"ACCESSIBLE" if m["acc"] else "REVIEW"}</span> '
            f'<span style="background:#eef;color:#334;padding:2px 8px;border-radius:11px;font-size:11px;">H1/H2/H3 = {m["h"][0]}/{m["h"][1]}/{m["h"][2]}</span> '
            f'<span style="background:#eef;color:#334;padding:2px 8px;border-radius:11px;font-size:11px;">{m["rlists"]} real lists</span>'
            + (f' <span style="background:#fff3d6;color:#8a5300;padding:2px 8px;border-radius:11px;font-size:11px;">&#9873; review</span>' if (m['flags'] or m['irr']) else ''))
    sp=pages(c,"src"); np_=pages(c,"new"); npg=max(len(sp),len(np_))
    word_src = c in ("AP3775","BP3775","AP4227","AP4228","AP4229","BP4231")
    left_label = "ORIGINAL WORD DOC" if word_src else "ORIGINAL PDF"
    rows=[]
    for i in range(npg):
        l=(f'<img src="{src(sp[i])}" loading="lazy" style="width:100%;border:2px solid #b3771a;display:block;">' if i<len(sp) else '<div style="padding:26px;color:#9aa;text-align:center;border:2px solid #e3c9a0;">(no page)</div>')
        r=(f'<img src="{src(np_[i])}" loading="lazy" style="width:100%;border:2px solid #1d6b3f;display:block;">' if i<len(np_) else '<div style="padding:26px;color:#9aa;text-align:center;border:2px solid #a9cdb8;">(no page)</div>')
        rows.append(f'<tr><td width="50%" style="vertical-align:top;padding:8px;"><div style="background:#b3771a;color:#fff;font-size:11px;font-weight:700;padding:4px 9px;border-radius:5px 5px 0 0;">{left_label} &middot; p{i+1}</div>{l}</td>'
                    f'<td width="50%" style="vertical-align:top;padding:8px;"><div style="background:#1d6b3f;color:#fff;font-size:11px;font-weight:700;padding:4px 9px;border-radius:5px 5px 0 0;">FINAL ACCESSIBLE DOCX &middot; p{i+1}</div>{r}</td></tr>')
    secs[sub].append(f'<section id="{c}" style="margin:0 0 32px;border:1px solid #d7dee7;border-radius:9px;overflow:hidden;">'
        f'<div style="background:#003366;color:#fff;padding:10px 16px;position:sticky;top:0;z-index:5;"><b style="font-size:15px;">{esc(os.path.splitext(os.path.basename(f))[0])}</b>'
        f'<a href="#top" style="float:right;color:#cdd9e8;font-size:12px;text-decoration:none;">&#8679; top</a><div style="margin-top:6px;">{badges}</div></div>'
        f'<table width="100%" style="border-collapse:collapse;">{"".join(rows)}</table></section>')
N=len(docs)
banner=(f'<div style="background:#003366;color:#fff;border-radius:9px;padding:20px 24px;margin-bottom:14px;">'
 f'<div style="font-size:23px;font-weight:700;">Policy Conversion &mdash; Original vs Final Accessible</div>'
 f'<div style="font-size:14px;color:#cdd9e8;margin-top:4px;">{N} policies &middot; LEFT = original source (website PDF, or original Word file) &middot; RIGHT = final accessible Word document &middot; College of the Canyons</div>'
 f'<div style="margin-top:12px;font-size:14px;">'
 f'<span style="background:#1d6b3f;padding:4px 12px;border-radius:12px;margin-right:6px;">{n_acc}/{N} pass accessibility</span>'
 f'<span style="background:#2c5a8a;padding:4px 12px;border-radius:12px;margin-right:6px;">{tot_lists} real lists</span>'
 f'<span style="background:#8a6a00;padding:4px 12px;border-radius:12px;">{n_flag} flagged for review</span></div></div>')
tocblock="".join(f'<div style="background:#fff;border-radius:9px;padding:12px 18px;margin-bottom:10px;"><b style="color:#003366;">{s} series ({len(toc[s])})</b><div style="margin-top:6px;line-height:1.9;">{"".join(toc[s])}</div></div>' for s in ("3000","4000","5000"))
body="".join(secs["3000"]+secs["4000"]+secs["5000"])
page=f'''<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Original vs Final Accessible - Audit</title></head>
<body style="margin:0;background:#eef1f5;font-family:'Segoe UI',Arial,sans-serif;color:#1a2733;" id="top"><div style="max-width:1200px;margin:0 auto;padding:22px 16px;">{banner}{tocblock}{body}</div></body></html>'''
open(os.path.join(HERE,"index.html"),"w").write(page)
missing=[code(os.path.basename(f)) for _,f in docs if not pages(code(os.path.basename(f)),"src")]
print(f"built: {N} docs; missing original render: {missing or 'none'}")
