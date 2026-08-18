#!/usr/bin/env python3
"""CySA+ (CS0-003) course slide deck — Tertiary Infotech all-white house style.

v7: full theory content migrated from the v5 legacy deck (600 slides), reordered
to the four CS0-003 exam domains, followed by the 30 hands-on labs per domain.
Front admin shows the CS0-003 exam domains; the deck closes with a Practice Exam
section and the CompTIA CySA+ exam registration guide.
Components (tile_grid / flow_h / trainer_slide / cards3 / two_col) follow the
wsq-slides reference deck.
"""
import os, re, sys, json, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import config as C
import lab_parser as LP
import paths as P

LABS = {L['num']: L for L in LP.load_all(P.LABS)}
THEORY = {int(k): v for k, v in json.load(open(os.path.join(HERE, "theory.json"))).items()}
IMG = P.IMAGES
ASSETS = P.ASSETS
os.makedirs(P.OUT, exist_ok=True)

# ---------------- palette ----------------
BLUE=RGBColor(0x1F,0x6F,0xEB); TEAL=RGBColor(0x10,0xB9,0x81); AMBER=RGBColor(0xF5,0x9E,0x0B)
INK=RGBColor(0x16,0x1B,0x26); GREY=RGBColor(0x5B,0x63,0x72); LIGHT=RGBColor(0xF5,0xF8,0xFC)
WHITE=RGBColor(0xFF,0xFF,0xFF); LINE=RGBColor(0xE2,0xE8,0xF0); VIOLET=RGBColor(0x7C,0x3A,0xED)
RED=RGBColor(0xE0,0x2D,0x3C)
PALETTE=[BLUE,TEAL,VIOLET,AMBER]

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
SW,SH=prs.slide_width,prs.slide_height
BLANK=prs.slide_layouts[6]
SHORT_TITLE="CompTIA CySA+ (CS0-003)"

def slide(): return prs.slides.add_slide(BLANK)
def page_no(): return len(prs.slides._sldIdLst)
def rect(s,x,y,w,h,color,line=None):
    sp=s.shapes.add_shape(1,x,y,w,h); sp.fill.solid(); sp.fill.fore_color.rgb=color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(1)
    sp.shadow.inherit=False; return sp
def oval(s,x,y,w,h,color):
    sp=s.shapes.add_shape(9,x,y,w,h); sp.fill.solid(); sp.fill.fore_color.rgb=color
    sp.line.fill.background(); sp.shadow.inherit=False; return sp
def txt(s,x,y,w,h,runs,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,space=4):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for i,line in enumerate(runs):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(space)
        for t,sz,col,bold in line:
            r=p.add_run(); r.text=t; r.font.size=Pt(sz); r.font.bold=bold
            r.font.color.rgb=col; r.font.name="Arial"
    return tb
def bullets(s,x,y,w,h,items,size=18,color=INK,gap=10):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    for i,it in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.space_after=Pt(gap)
        lvl=it[1] if isinstance(it,(tuple,list)) else 0
        text=it[0] if isinstance(it,(tuple,list)) else it
        bold=(isinstance(it,(tuple,list)) and len(it)>2 and bool(it[2]))
        item_color=(it[3] if isinstance(it,(tuple,list)) and len(it)>3 and it[3] is not None else None)
        lvl=min(int(lvl),2)
        r=p.add_run(); r.text=("•  " if lvl==0 else ("–  " if lvl==1 else "·  "))+text
        p.level=lvl
        r.font.size=Pt(size if lvl==0 else max(size-2,11))
        r.font.color.rgb=item_color if item_color is not None else (color if lvl==0 else GREY)
        r.font.name="Arial"; r.font.bold=bold
    return tb
def footer(s):
    txt(s,Inches(0.4),Inches(7.05),Inches(7.5),Inches(0.35),
        [[(f"{SHORT_TITLE}  ·  {C.COURSE_CODE}",9,GREY,False)]])
    txt(s,Inches(5.0),Inches(7.05),Inches(3.3),Inches(0.35),
        [[("© 2026 Tertiary Infotech Academy Pte Ltd",9,GREY,False)]],align=PP_ALIGN.CENTER)
    txt(s,Inches(12.4),Inches(7.05),Inches(0.6),Inches(0.35),
        [[(str(page_no()),9,GREY,False)]],align=PP_ALIGN.RIGHT)
def head(s,title,kicker=None,kcolor=BLUE):
    rect(s,0,0,SW,SH,WHITE); rect(s,0,0,Inches(0.28),Inches(1.55),kcolor)
    if kicker: txt(s,Inches(0.85),Inches(0.5),Inches(11.6),Inches(0.4),[[(kicker,14,kcolor,True)]])
    tsize=29 if len(title)<=58 else 24
    txt(s,Inches(0.85),Inches(0.9),Inches(11.9),Inches(0.9),[[(title,tsize,INK,True)]])
    rect(s,Inches(0.85),Inches(1.7),Inches(11.63),Inches(0.02),LINE)
    return s
def _asset(name):
    p=os.path.join(ASSETS,name)
    return p if os.path.exists(p) else None

# ---------------- slide templates ----------------
def cover():
    s=slide(); rect(s,0,0,SW,SH,WHITE)
    rect(s,0,0,SW,Inches(0.22),BLUE); rect(s,0,Inches(7.28),SW,Inches(0.22),TEAL)
    org=_asset("tertiary_logo.png")
    if org: s.shapes.add_picture(org,Inches(0.85),Inches(0.7),height=Inches(1.05))
    badge=_asset("cysa_logo.png")
    if badge:
        s.shapes.add_picture(badge,Inches(10.9),Inches(0.6),height=Inches(1.35))
    else:
        rect(s,Inches(11.0),Inches(0.72),Inches(1.55),Inches(1.0),RED)
        txt(s,Inches(11.0),Inches(0.82),Inches(1.55),Inches(0.5),[[("CySA+",20,WHITE,True)]],align=PP_ALIGN.CENTER)
        txt(s,Inches(11.0),Inches(1.28),Inches(1.55),Inches(0.4),[[("COMPTIA CS0-003",8,WHITE,True)]],align=PP_ALIGN.CENTER)
    txt(s,Inches(0.9),Inches(2.3),Inches(12),Inches(0.6),[[("COURSE SLIDES  ·  WSQ",16,BLUE,True)]])
    txt(s,Inches(0.9),Inches(2.85),Inches(12.0),Inches(1.9),[[(C.COURSE_TITLE,38,INK,True)]])
    rect(s,Inches(0.92),Inches(4.75),Inches(2.4),Inches(0.06),TEAL)
    txt(s,Inches(0.9),Inches(5.05),Inches(12),Inches(1.4),
        [[(f"WSQ Course Code: {C.COURSE_CODE}",16,GREY,False)],
         [(f"WSQ TSC: {C.TSC_TITLE} ({C.TSC_CODE})",14,GREY,False)],
         [("Conducted by Tertiary Infotech Academy Pte Ltd  ·  UEN 201200696W",14,GREY,False)]],space=6)
    txt(s,Inches(0.9),Inches(6.5),Inches(12),Inches(0.4),[[(f"Version v{C.VERSION_PPT.split('.')[0]}  ·  {C.VERSION_DATE}",12,GREY,False)]])
    txt(s,Inches(0.9),Inches(6.85),Inches(12),Inches(0.34),[[("© 2026 Tertiary Infotech Academy Pte Ltd. All rights reserved.  ·  www.tertiarycourses.com.sg",10,GREY,False)]])

def section(kicker,title,n="",sub=""):
    s=slide(); rect(s,0,0,SW,SH,WHITE); rect(s,0,0,Inches(0.28),SH,BLUE)
    rect(s,Inches(0.85),Inches(2.5),Inches(0.14),Inches(2.0),TEAL)
    txt(s,Inches(1.25),Inches(2.55),Inches(11),Inches(0.6),[[(kicker,18,BLUE,True)]])
    txt(s,Inches(1.25),Inches(3.0),Inches(11.4),Inches(1.6),[[(title,40,INK,True)]])
    if sub: txt(s,Inches(1.27),Inches(4.55),Inches(11),Inches(0.8),[[(sub,16,GREY,False)]])
    if n: txt(s,Inches(10.0),Inches(0.7),Inches(2.8),Inches(1.6),[[(str(n),72,LINE,True)]],align=PP_ALIGN.RIGHT)
    footer(s); return s
def content(title,items,kicker=None,size=20):
    s=head(slide(),title,kicker); bullets(s,Inches(0.85),Inches(1.95),Inches(11.6),Inches(4.9),items,size=size); footer(s); return s
def two_col(title,left,right,kicker=None,lhead="",rhead=""):
    s=head(slide(),title,kicker)
    rect(s,Inches(0.85),Inches(1.95),Inches(5.7),Inches(4.7),LIGHT); rect(s,Inches(6.95),Inches(1.95),Inches(5.55),Inches(4.7),LIGHT)
    if lhead: txt(s,Inches(1.1),Inches(2.15),Inches(5.2),Inches(0.4),[[(lhead,16,BLUE,True)]])
    if rhead: txt(s,Inches(7.2),Inches(2.15),Inches(5.0),Inches(0.4),[[(rhead,16,TEAL,True)]])
    bullets(s,Inches(1.1),Inches(2.7),Inches(5.2),Inches(3.8),left,size=16)
    bullets(s,Inches(7.2),Inches(2.7),Inches(5.05),Inches(3.8),right,size=16); footer(s); return s
def cards3(title,cards,kicker):
    s=head(slide(),title,kicker); xs=[Inches(0.85),Inches(5.0),Inches(9.15)]
    for i,c in enumerate(cards[:3]):
        x=xs[i]; col=c[0]
        rect(s,x,Inches(1.95),Inches(3.65),Inches(4.7),LIGHT); rect(s,x,Inches(1.95),Inches(3.65),Inches(0.12),col)
        txt(s,x+Inches(0.25),Inches(2.2),Inches(3.2),Inches(0.6),[[(c[1],19,col,True)]])
        bullets(s,x+Inches(0.25),Inches(2.95),Inches(3.2),Inches(3.4),c[2],size=13,gap=8)
    footer(s); return s
def big_statement(line1,line2,kicker,color=BLUE):
    s=slide(); rect(s,0,0,SW,SH,WHITE); rect(s,0,0,Inches(0.28),SH,color)
    txt(s,Inches(1.1),Inches(2.2),Inches(11),Inches(0.5),[[(kicker,16,color,True)]])
    txt(s,Inches(1.1),Inches(2.8),Inches(11.3),Inches(2.4),[[(line1,38,INK,True)]])
    if line2: txt(s,Inches(1.12),Inches(4.9),Inches(11),Inches(1.2),[[(line2,20,GREY,False)]])
    footer(s); return s
def tile_grid(title,items,kicker=None,cols=2,size=15,icons=None,accent=BLUE):
    s=head(slide(),title,kicker,kcolor=accent)
    n=len(items); rows=math.ceil(n/cols)
    X0=Inches(0.85); Y0=Inches(1.95); TOTW=Inches(11.63); AREAH=Inches(4.78)
    gx=Inches(0.3); gy=Inches(0.26)
    cw=int((TOTW-gx*(cols-1))/cols); ch=int((AREAH-gy*(rows-1))/rows)
    bd=Inches(0.6)
    for i,it in enumerate(items):
        r=i//cols; c=i%cols
        x=int(X0+(cw+gx)*c); y=int(Y0+(ch+gy)*r); col=PALETTE[i%len(PALETTE)]
        rect(s,x,y,cw,ch,LIGHT); rect(s,x,y,Inches(0.1),ch,col)
        oval(s,x+Inches(0.28),int(y+ch/2-bd/2),bd,bd,col)
        ic=icons[i] if icons else str(i+1)
        txt(s,x+Inches(0.28),int(y+ch/2-bd/2),bd,bd,[[(ic,19,WHITE,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
        tx=x+Inches(1.08); tw=cw-Inches(1.32)
        if isinstance(it,tuple):
            txt(s,tx,int(y+Inches(0.14)),tw,int(ch-Inches(0.2)),
                [[(it[0],size+2,INK,True)],[(it[1],size-2,GREY,False)]],anchor=MSO_ANCHOR.MIDDLE,space=3)
        else:
            txt(s,tx,int(y+Inches(0.1)),tw,int(ch-Inches(0.16)),[[(it,size,INK,False)]],anchor=MSO_ANCHOR.MIDDLE)
    footer(s); return s
def flow_h(title,steps,kicker=None,color=BLUE):
    s=head(slide(),title,kicker,kcolor=color)
    n=len(steps); X0=Inches(0.85); TOTW=Inches(11.63); gap=Inches(0.34)
    cw=int((TOTW-gap*(n-1))/n); y=Inches(2.55); ch=Inches(3.15); bd=Inches(0.82)
    for i,st in enumerate(steps):
        x=int(X0+(cw+gap)*i)
        rect(s,x,y,cw,ch,LIGHT); rect(s,x,y,cw,Inches(0.1),color)
        oval(s,int(x+cw/2-bd/2),int(y+Inches(0.42)),bd,bd,color)
        txt(s,int(x+cw/2-bd/2),int(y+Inches(0.42)),bd,bd,[[(str(i+1),30,WHITE,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
        txt(s,x+Inches(0.16),int(y+Inches(1.55)),cw-Inches(0.32),int(ch-Inches(1.7)),[[(st,14,INK,False)]],align=PP_ALIGN.CENTER)
        if i<n-1:
            txt(s,int(x+cw-Inches(0.04)),int(y+ch/2-Inches(0.3)),int(gap+Inches(0.08)),Inches(0.6),
                [[("▶",15,color,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    footer(s); return s
def trainer_slide(kicker,name,role,rows,initials,accent=BLUE):
    s=head(slide(),"About the Trainer",kicker,kcolor=accent)
    lx=Inches(0.85); lw=Inches(3.65)
    rect(s,lx,Inches(1.95),lw,Inches(4.7),LIGHT); rect(s,lx,Inches(1.95),lw,Inches(0.12),accent)
    bd=Inches(1.7); ax=int(lx+(lw-bd)/2)
    oval(s,ax,Inches(2.5),bd,bd,accent)
    txt(s,ax,Inches(2.5),bd,bd,[[(initials,44,WHITE,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    txt(s,lx+Inches(0.15),Inches(4.55),lw-Inches(0.3),Inches(0.6),[[(name,21,INK,True)]],align=PP_ALIGN.CENTER)
    txt(s,lx+Inches(0.15),Inches(5.2),lw-Inches(0.3),Inches(1.2),[[(role,13,GREY,False)]],align=PP_ALIGN.CENTER)
    rx=Inches(4.9); rw=Inches(7.6); ry=Inches(1.95); rh=Inches(4.7)
    n=len(rows); gy=Inches(0.2); th=int((rh-gy*(n-1))/n)
    for i,(label,val) in enumerate(rows):
        y=int(ry+(th+gy)*i); col=PALETTE[i%len(PALETTE)]
        rect(s,rx,y,rw,th,LIGHT); rect(s,rx,y,Inches(0.1),th,col)
        vruns=[(val,14,INK,False)] if val else [("____________________________________________",13,LINE,False)]
        txt(s,rx+Inches(0.32),y,rw-Inches(0.6),th,
            [[(label.upper(),11,col,True)],vruns],anchor=MSO_ANCHOR.MIDDLE,space=3)
    footer(s); return s
def brk(kind,dur,color=AMBER):
    s=slide(); rect(s,0,0,SW,SH,WHITE)
    rect(s,0,0,SW,Inches(0.22),color); rect(s,0,Inches(7.28),SW,Inches(0.22),color)
    rect(s,Inches(5.4),Inches(2.35),Inches(2.53),Inches(0.1),color)
    txt(s,0,Inches(2.75),SW,Inches(1.2),[[(kind,48,INK,True)]],align=PP_ALIGN.CENTER)
    txt(s,0,Inches(4.05),SW,Inches(0.8),[[(dur,22,color,True)]],align=PP_ALIGN.CENTER)

def _img_size(path):
    try:
        with Image.open(path) as im: return im.size
    except Exception: return (800,600)
def _place_imgs(s,paths,X0,Y0,TOTW,TOTH):
    """Fit 1..3 images side-by-side inside the given box, centred."""
    paths=[p for p in paths if os.path.exists(p)][:3]
    if not paths: return
    n=len(paths); gap=Inches(0.3)
    cellw=int((TOTW-gap*(n-1))/n)
    for i,p in enumerate(paths):
        w0,h0=_img_size(p)
        sc=min(cellw/ (w0*9525), TOTH/(h0*9525))
        w=int(w0*9525*sc); h=int(h0*9525*sc)
        x=int(X0+(cellw+gap)*i+(cellw-w)/2); y=int(Y0+(TOTH-h)/2)
        s.shapes.add_picture(p,x,y,width=w,height=h)
def diagram(title,img_files,kicker=None):
    s=head(slide(),title,kicker)
    _place_imgs(s,[os.path.join(IMG,f) for f in img_files],Inches(0.85),Inches(1.95),Inches(11.63),Inches(4.85))
    footer(s); return s
def content_img(title,items,img_files,kicker=None,size=15):
    s=head(slide(),title,kicker)
    bullets(s,Inches(0.85),Inches(1.95),Inches(6.0),Inches(4.9),items,size=size)
    _place_imgs(s,[os.path.join(IMG,f) for f in img_files],Inches(7.1),Inches(1.95),Inches(5.4),Inches(4.7))
    footer(s); return s
def steps_slide(title,steps,kicker=None):
    """Numbered lab steps, each with optional description + one-line command chip."""
    s=head(slide(),title,kicker,kcolor=TEAL)
    n=len(steps); Y0=Inches(1.95); AREAH=Inches(4.85); gy=Inches(0.14)
    rh=int((AREAH-gy*(n-1))/n)
    for i,st in enumerate(steps):
        y=int(Y0+(rh+gy)*i)
        rect(s,Inches(0.85),y,Inches(11.63),rh,LIGHT); rect(s,Inches(0.85),y,Inches(0.1),rh,TEAL)
        bd=Inches(0.5)
        oval(s,Inches(1.1),int(y+rh/2-bd/2),bd,bd,TEAL)
        txt(s,Inches(1.1),int(y+rh/2-bd/2),bd,bd,[[(str(i+1),15,WHITE,True)]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
        has_code=bool(st.get("code"))
        tw=Inches(6.0) if has_code else Inches(10.4)
        lines=[[(st["label"],13,INK,True)]]
        if st.get("desc"): lines.append([(st["desc"],11,GREY,False)])
        txt(s,Inches(1.85),y,tw,rh,lines,anchor=MSO_ANCHOR.MIDDLE,space=2)
        if has_code:
            cx=Inches(8.0); cwd=Inches(4.35)
            ch2=min(Inches(0.6),rh-Inches(0.16))
            rect(s,cx,int(y+rh/2-ch2/2),cwd,ch2,RGBColor(0x0B,0x12,0x20))
            txt(s,cx+Inches(0.15),int(y+rh/2-ch2/2),cwd-Inches(0.3),ch2,
                [[(st["code"],10,RGBColor(0x9C,0xDC,0xFE),False)]],anchor=MSO_ANCHOR.MIDDLE)
    footer(s); return s

# ---------------- theory rendering ----------------
EXPLAIN = json.load(open(os.path.join(HERE, "concept_explanations.json")))

def clean(t):
    return re.sub(r'\s+',' ',t).strip()

_IMG_URL = re.compile(r'https?://\S+\.(?:png|jpe?g|gif|bmp|svg)\b', re.I)
def drop_bullet(t):
    """True if a theory bullet is just a raw image URL (legacy leak) worth dropping."""
    t = t.strip()
    return bool(_IMG_URL.fullmatch(t)) or (t.lower().startswith('http') and _IMG_URL.search(t) and len(t.split()) == 1)

def explain_slides(base_title, title, kick):
    """In-depth explainer slide(s) for a concept, from the authored explanations."""
    ex = EXPLAIN.get(base_title.lower())
    if not ex: return False
    ebul = [(clean(s), 0) for s in ex.get("explain", []) if clean(s)]
    if not ebul: return False
    exam = (f"On the exam: {clean(ex['exam'])}", 0, True, TEAL) if ex.get("exam") else None
    # Keep 5 rows/slide (≈4 explanation sentences + the exam line) so the exam line
    # never spills alone onto a continuation slide.
    CH = 5
    # reserve room for the exam line on the LAST chunk
    body = ebul[:]
    chunks = [body[i:i+CH] for i in range(0, len(body), CH)] or [[]]
    if exam:
        if len(chunks[-1]) < CH:
            chunks[-1] = chunks[-1] + [exam]
        else:
            chunks.append([exam])
    for ci, chunk in enumerate(chunks):
        t2 = f"{title} — In Depth" if ci == 0 else f"{title} — In Depth (cont'd)"
        content(t2, chunk, kicker=kick, size=17 if len(chunk) <= 4 else 15)
    return True

def render_theory(dom_num, dom_name, items):
    seen_titles={}
    for it in items:
        base=clean(it["title"]) or dom_name
        title=base
        kick=f"TOPIC {dom_num} · {clean(it['kicker']).upper()[:60]}" if it["kicker"] else f"TOPIC {dom_num} · {dom_name.upper()}"
        # de-duplicate repeated titles with a counter suffix
        key=title.lower()
        seen_titles[key]=seen_titles.get(key,0)+1
        first_occurrence = seen_titles[key]==1
        if seen_titles[key]>1: title=f"{title} ({seen_titles[key]})"
        bl=[(clean(b[0]),b[1]) for b in it["bullets"] if clean(b[0]) and not drop_bullet(b[0])]
        imgs=[f for f in it["images"] if os.path.exists(os.path.join(IMG,f))]
        # in-depth explanation (only on the first slide bearing this concept title)
        had_explain = explain_slides(base, title, kick) if first_occurrence else False
        # then the original keyword / diagram slide as the "Key Points" reference
        kp_title = f"{title} — Key Points" if (had_explain and bl) else title
        if imgs and not bl:
            diagram(title,imgs,kicker=kick)
        elif imgs and bl:
            content_img(kp_title,bl,imgs,kicker=kick,size=15 if len(bl)<=8 else 13)
        elif bl:
            CH=11
            for ci in range(0,len(bl),CH):
                chunk=bl[ci:ci+CH]
                t2=kp_title if ci==0 else f"{title} (cont'd)"
                sz=20 if len(chunk)<=5 else (18 if len(chunk)<=8 else 15)
                content(t2,chunk,kicker=kick,size=sz)

# ---------------- lab rendering ----------------
def sent(text,n=2,cap=260):
    text=LP.md_inline(text)
    parts=re.split(r'(?<=[.!?])\s+',text.strip())
    out=' '.join(parts[:n]).strip()
    return (out[:cap].rsplit(' ',1)[0]+'…') if len(out)>cap else out
def cysa_map(intro_text):
    m=re.search(r'\(CySA\+\s*([^)]*)\)',intro_text)
    return m.group(1).strip() if m else ''
def render_lab(dom_num,dom_name,ln,manifest):
    L=LABS[ln]
    kick=f"TOPIC {dom_num} · HANDS-ON LAB"
    tail=L['title'].split('—',1)[-1].strip() if '—' in L['title'] else L['title']
    intro=next((b[1] for b in L['intro'] if b[0]=='text'),'')
    cm=cysa_map(intro)
    objb=[("Objective",0,True),
          (sent(re.sub(r'\s*\(CySA\+[^)]*\)','',intro),3,320),1),
          (f"Environment: {C.ENV[ln]}",0,True)]
    if cm: objb.append((f"CySA+ mapping: {cm}",1))
    content(f"Lab {ln} — {tail}",objb,kicker=kick,size=17)
    manifest["labs"][ln]={"title":L['title'],"slide":page_no()}
    if ln in C.DIAGRAMS:
        diagram(f"Lab {ln} — Reference Diagram",C.DIAGRAMS[ln],kicker=kick)
    steps=LP.get_steps(L); CH=5
    for i in range(0,len(steps),CH):
        grp=steps[i:i+CH]; sd=[]
        for st in grp:
            htail=re.sub(r'^Step\s*\d+\s*[—-]\s*','',st['heading'])
            snum=re.match(r'^Step\s*(\d+)',st['heading'])
            label=f"Step {snum.group(1)}: {htail}" if snum else htail
            desc=''; code=''
            for b in st['blocks']:
                if b[0]=='text' and not desc: desc=sent(b[1],1,150)
                if b[0]=='code' and not code:
                    c0=b[1].strip().split('\n')[0][:72]; clang=b[2] if len(b)>2 else ''
                    is_shell=clang in ('bash','sh','shell','console','') and not c0.lstrip().startswith(('#','*','-','1.','**'))
                    code=("$ " if is_shell else "")+LP.md_inline(c0)
            sd.append({"label":LP.md_inline(label),"desc":desc,"code":code})
        part=" (cont'd)" if i>0 else ""
        steps_slide(f"Lab {ln} — Hands-On Steps{part}",sd,kicker=kick)
    wyl=LP.get_outcomes(L)
    ob=[("By completing this lab you can:",0,True)]
    if wyl:
        for b in wyl['blocks']:
            if b[0] in ('bullets','numbered'):
                for x in b[1]: ob.append((LP.md_inline(x),0))
            elif b[0]=='text': ob.append((sent(b[1],2,200),0))
    content(f"Lab {ln} — What You Learned",ob,kicker=kick,size=17)

# ============================================================ BUILD
manifest={"domains":[],"labs":{},"ppt_file":None}
cover()

# ---------------- ADMIN ----------------
section("COURSE ADMINISTRATION","Welcome & Housekeeping")
content("Digital Attendance (Mandatory)",[
 "It is mandatory to take the AM, PM and Assessment digital attendance for WSQ-funded courses.",
 "The trainer/administrator displays the digital attendance QR code from the SSG portal.",
 "Scan the QR code with your mobile phone camera and submit your attendance.",
 "A minimum of 75% attendance is required to be eligible for assessment and funding."],
 kicker="TRAQOM · SSG DIGITAL ATTENDANCE")
trainer_slide("YOUR TRAINER · GENERAL","Your Trainer","General Trainer template —\nto be completed by the trainer",
 [("Name",""),("Title / Designation",""),("Qualifications",""),
  ("Areas of expertise",""),("Training & industry experience",""),("Contact","")],
 initials="?",accent=GREY)
trainer_slide("YOUR TRAINER","Dr. Alfred Ang","Principal Trainer\nTertiary Infotech Academy Pte. Ltd.",
 [("Role","Principal Trainer, Tertiary Infotech Academy Pte. Ltd."),
  ("Expertise","Cybersecurity, SOC operations, data science, AI and machine learning."),
  ("Delivers","WSQ courses on CompTIA certifications, cybersecurity, AI & data analytics."),
  ("Founder","Founder and lead instructor at Tertiary Infotech / Tertiary Courses.")],
 initials="AA",accent=BLUE)
content("Let's Know Each Other",[
 "Your name and organisation / role.",
 "Your experience with security operations and SOC tools (if any).",
 "What you want to get out of this CySA+ course.",
 "One security incident or tool you are curious about."],kicker="ICE-BREAKER")
tile_grid("Ground Rules",[
 "Set your mobile phone to silent mode.","Participate actively — no question is too small.",
 "Mutual respect: agree to disagree.","One conversation at a time.",
 "Be punctual; return from breaks on time.","75% attendance is required."],
 kicker="HOUSEKEEPING",cols=2,size=15)
# Download course material — LMS visual
s=head(slide(),"Download Course Material",kicker="COURSE PORTAL · LMS / TMS")
shot=_asset("lms_screenshot.png")
if shot: _place_imgs(s,[shot],Inches(6.4),Inches(1.95),Inches(6.1),Inches(4.7))
bullets(s,Inches(0.85),Inches(2.1),Inches(5.3),Inches(4.6),[
 ("Log in to the Tertiary Infotech LMS/TMS portal:",0),
 ("https://lms-tms.tertiaryinfotech.com",1,True),
 ("Download the Course Slides and Learner Guide.",0),
 ("Take your digital attendance and TRAQOM survey here.",0),
 ("Submit your assessment answers on the portal.",0)],size=16)
footer(s)
# Lesson plan (5 days)
two_col("Lesson Plan — 5 Days, 8 hours/day",[
 ("Day 1",0,True),("Topic 1: Security Operations — theory + Labs 1–10",1),
 ("Day 2",0,True),("Topic 2: Vulnerability Management — theory + Labs 11–19",1),
 ("Day 3",0,True),("Topic 3: Incident Response & Management — theory + Labs 20–25",1)],
 [("Day 4",0,True),("Topic 4: Reporting & Communication — theory + Labs 26–30",1),
  ("Day 5",0,True),("Revision + Practice Exam",1),("Course Feedback & TRAQOM Survey",1),
  ("Final Assessment (WA + Case Study)",1),
  ("Daily timing",0,True),("9:30am–6:30pm · 1-hour lunch · tea breaks within",1)],
 kicker="SCHEDULE",lhead="Days 1–3",rhead="Days 4–5 & timing")
# Skills framework
content("WSQ Skills Framework",[
 (f"TSC Title: {C.TSC_TITLE}",0,True),
 (f"TSC Code: {C.TSC_CODE}",0,True),
 ("This course is aligned to the SkillsFuture Singapore (SSG) Skills Framework Technical Skills & Competency (TSC) above.",0)],
 kicker="SKILLS FRAMEWORK")
content("TSC Knowledge (K)",[(k,0) for k in C.TSC_KNOWLEDGE],kicker="SKILLS FRAMEWORK",size=18)
content("TSC Abilities (A)",[(a,0) for a in C.TSC_ABILITIES],kicker="SKILLS FRAMEWORK",size=17)
tile_grid("Learning Outcomes",[(f"LO{i}",lo) for i,lo in enumerate(C.LEARNING_OUTCOMES,1)],
 kicker="WHAT YOU'LL ACHIEVE",cols=1,size=14)
# ---- CS0-003 exam domains (front of deck) ----
tile_grid("CompTIA CySA+ (CS0-003) Exam Domains",
 [(f"{d['name']}  ·  {d['weight']}%", ", ".join(d['concepts'][:4])+" …") for d in C.DOMAINS],
 kicker="EXAM DOMAINS · CS0-003",cols=1,size=15)
content("How This Course Maps to the Exam",[
 ("The four course topics follow the four CS0-003 exam domains in order and weight:",0),
 ("Topic 1 — Security Operations (33%) · Labs 1–10",1),
 ("Topic 2 — Vulnerability Management (30%) · Labs 11–19",1),
 ("Topic 3 — Incident Response and Management (20%) · Labs 20–25",1),
 ("Topic 4 — Reporting and Communication (17%) · Labs 26–30",1),
 ("Every lab is tagged with the CySA+ exam objective it practises.",0)],
 kicker="EXAM DOMAINS · CS0-003",size=17)
for dom in C.DOMAINS:
    content(f"Course Outline — {dom['topic']}: {dom['name']} ({dom['weight']}%)",
            [(c,0) for c in dom['concepts']],kicker="COURSE OUTLINE",size=16)
content("Briefing for Assessment",[
 "Place phones and other materials under the table or on the floor.",
 "No photos or recording of assessment scripts.","No discussion during the assessment.",
 "Use a black/blue pen for hard-copy assessments. No liquid paper or correction tape.",
 "Assessment scripts will be collected when the time is up."],kicker="FINAL ASSESSMENT")
content("Assessment",[
 "Written Assessment (SAQ) — 1 hour.",
 "Case Study (CS) — 1 hour.",
 "Format: Open Book — slides, Learner Guide and approved materials only.",
 "An appeal process is available for learners assessed as Not Yet Competent."],kicker="FINAL ASSESSMENT")
flow_h("Assessment Flow",[
 "TRAQOM survey — scan the QR code on the LMS",
 "Assessment digital attendance — scan the SSG QR",
 "Sit WA (SAQ) then Case Study — open book",
 "Submit your answers on the LMS",
 "Sign the Assessment Summary Record"],kicker="ON ASSESSMENT DAY")
content("Criteria for Funding",[
 "Minimum attendance rate of 75% based on the SSG Digital Attendance record.",
 "Complete the assessment and be assessed as 'Competent'."],kicker="WSQ FUNDING")

# ---------------- TOPICS: theory + labs ----------------
for di,dom in enumerate(C.DOMAINS,1):
    sec=section(f"TOPIC {di} · {dom['weight']}% OF CS0-003",dom['name'],n=di)
    manifest["domains"].append({"topic":dom['topic'],"name":dom['name'],"weight":dom['weight'],
                                "section_slide":page_no(),"labs":dom['labs']})
    tile_grid(f"Topic {di} Overview — {dom['name']}",dom['concepts'],
              kicker=f"EXAM WEIGHTING {dom['weight']}%",cols=2,size=13)
    render_theory(di,dom['name'],THEORY[di])
    big_statement(f"Hands-On: Labs {dom['labs'][0]}–{dom['labs'][-1]}",
                  f"{len(dom['labs'])} step-by-step labs applying {dom['name']} in the Killercoda playground and free tools.",
                  f"TOPIC {di} · PRACTICE",color=TEAL)
    for ln in dom['labs']:
        render_lab(di,dom['name'],ln,manifest)
    content(f"Recap — {dom['name']}",[(c,0) for c in dom['concepts']],kicker=f"TOPIC {di} RECAP",size=16)

# ---------------- PRACTICE EXAM ----------------
section("PRACTICE EXAM","Test Yourself Before the Real Exam")
manifest["practice_section"]=page_no()
content("Practice Exam",[
 "Sharpen your exam readiness with the Tertiary Infotech CySA+ practice exam:",
 ("https://exams.tertiaryinfotech.com",0,True),
 "Attempt it under timed conditions and review every explanation.",
 "Revisit any topic or lab you miss, then re-take the practice exam.",
 "The following slides give you a warm-up practice quiz — answers at the end."],kicker="TEST YOURSELF")
for qi,q in enumerate(C.PRACTICE_QUESTIONS,1):
    content(f"Practice Question {qi}",
            [(q['q'],0,True)]+[(f"{letter}.  {opt}",1) for letter,opt in zip("ABCD",q['opts'])],
            kicker=f"PRACTICE EXAM · {q['domain'].upper()}",size=16)
ANS=[(f"Q{qi}: {q['ans']} — {q['why']}",0) for qi,q in enumerate(C.PRACTICE_QUESTIONS,1)]
half=(len(ANS)+1)//2
content("Practice Quiz — Answer Key (1/2)",ANS[:half],kicker="PRACTICE EXAM",size=14)
content("Practice Quiz — Answer Key (2/2)",ANS[half:],kicker="PRACTICE EXAM",size=14)

# ---------------- EXAM REGISTRATION ----------------
section("CERTIFICATION","CompTIA CySA+ Exam Registration")
manifest["examreg_section"]=page_no()
tile_grid("CompTIA CySA+ (CS0-003) Exam Facts",[
 ("Exam code","CS0-003 · launched June 2023 (V4/CS0-004 announced — check comptia.org for transition dates)."),
 ("Questions & time","Maximum 85 questions (multiple-choice + performance-based) in 165 minutes."),
 ("Passing score","750 on a scale of 100–900."),
 ("Recommended experience","Network+/Security+ plus ~4 years of hands-on security experience."),
 ("Delivery","Pearson VUE — in-person test centre or OnVUE online proctored."),
 ("Renewal","Valid 3 years; renew via CompTIA Continuing Education (60 CEUs).")],
 kicker="EXAM DETAILS",cols=1,size=13)
flow_h("How to Register for the CySA+ Exam",[
 "Create your CompTIA account at login.comptia.org",
 "Buy the CySA+ exam voucher at store.comptia.org",
 "Schedule your exam at Pearson VUE — test centre or OnVUE online",
 "Sit the CS0-003 exam and pass with 750+",
 "Claim your digital badge and start your CE renewal cycle"],kicker="REGISTRATION STEPS")
content("Exam Registration — Useful Links",[
 ("CySA+ certification page: https://www.comptia.org/certifications/cybersecurity-analyst",0,True),
 ("CompTIA store (vouchers): https://store.comptia.org",0,True),
 ("Pearson VUE scheduling: https://home.pearsonvue.com/comptia",0,True),
 "Tip: bundle the voucher with a retake option if you want a safety net.",
 "Use the practice exam results to decide when you are ready to book."],kicker="REGISTRATION",size=16)

# ---------------- CLOSE ----------------
section("WRAP-UP","Course Summary & Next Steps")
manifest["wrap_section"]=page_no()
tile_grid("What You Achieved",[
 ("Security Operations","Analysed logs, network traffic, email and IoCs; used threat intelligence and MITRE ATT&CK."),
 ("Vulnerability Management","Discovered assets, ran vulnerability scans, prioritised with CVSS and hardened systems."),
 ("Incident Response","Applied attack frameworks, forensics, containment and IR playbooks."),
 ("Reporting & Communication","Produced vulnerability, incident, metrics and compliance reports for stakeholders.")],
 kicker="LEARNING OUTCOMES",cols=2,size=14)
content("Cert & TRAQOM Survey (Mandatory)",[
 "Complete the mandatory TRAQOM survey and claim your certificate via the LMS/TMS:",
 ("https://lms-tms.tertiaryinfotech.com",0,True)],kicker="FEEDBACK")
tile_grid("Recommended Courses",[
 "WSQ – CompTIA Security+ Training","WSQ – CompTIA PenTest+ Training",
 "WSQ – CompTIA Linux+ Training","WSQ – CompTIA Server+ Training",
 "WSQ – CompTIA A+ Training (Core 1 & 2)","WSQ – Certified Ethical Hacking (CEH) Training"],
 kicker="WHAT'S NEXT",cols=2,size=14)
content("Support",[
 "If you have any enquiries during and after the class, contact us:",
 (f"Email: {C.CONTACT['email']}",0,True),
 (f"Tel: {C.CONTACT['tel']}",0,True),
 (f"Website: {C.CONTACT['web']}",0,True)],kicker="WE'RE HERE TO HELP")
content("Assessment",[
 "Written Assessment (SAQ) — 1 hour.  Case Study (CS) — 1 hour.",
 "Open book: slides, Learner Guide and approved materials only.",
 "Remember to take the Assessment digital attendance (TRAQOM).",
 "Submit your completed answers on the LMS at https://lms-tms.tertiaryinfotech.com"],kicker="FINAL ASSESSMENT")
flow_h("Assessment Flow",[
 "TRAQOM survey — scan the QR code on the LMS",
 "Assessment digital attendance — scan the SSG QR",
 "Sit WA (SAQ) then Case Study — open book",
 "Submit your answers on the LMS",
 "Sign the Assessment Summary Record"],kicker="ON ASSESSMENT DAY")
content("Digital Attendance (Mandatory)",[
 "It is mandatory to take the AM, PM and Assessment digital attendance for WSQ-funded courses.",
 "The trainer/administrator displays the digital attendance QR code from the SSG portal.",
 "Scan the QR code with your mobile phone camera and submit your attendance.",
 "A minimum of 75% attendance is required to be eligible for assessment and funding."],
 kicker="TRAQOM · SSG DIGITAL ATTENDANCE")
big_statement("Thank You!","You are now ready to work as a cybersecurity analyst — and to sit the CompTIA CySA+ exam.","SEE YOU IN THE SOC",color=TEAL)

out_name=f"{C.COURSE_TITLE}-v{C.VERSION_PPT.split('.')[0]}.pptx"
prs.save(os.path.join(P.OUT,out_name))
manifest["ppt_file"]=out_name
manifest["total_slides"]=page_no()
with open(P.MANIFEST,"w") as f: json.dump(manifest,f,indent=2)
print("Saved:",out_name,"| slides:",page_no())
print("Domains:",[(d['topic'],d['section_slide']) for d in manifest['domains']])
