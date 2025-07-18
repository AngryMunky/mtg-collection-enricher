import os, json, threading, io, requests, pandas as pd
from tkinter import Tk, Frame, Label, Button, Entry, StringVar, \
                     filedialog, scrolledtext, ttk, messagebox
from PIL import Image, ImageTk
from openpyxl import load_workbook

# ─── GUI & card sizing ───────────────────────────────
WIN_W, WIN_H = 700, 600
IMG_W, IMG_H = 400, 560
RIGHT_W      = WIN_W - IMG_W - 20
IMAGE_REFRESH = 10_000

# ─── filenames & column order ────────────────────────
BULK_JSON  = "default-cards.json"
BULK_META  = "bulk_meta.txt"
CARD_BACK  = "card_back.jpg"

COLOR_MAP = {"W":"White","U":"Blue","B":"Black","R":"Red","G":"Green"}

COL_ORDER = [
    "Card Name","Color","Rarity","Mana Value",
    "Card","Type","Set Name",
    "Foil","Quantity","Collector Number","ManaBox ID","Scryfall ID"
]

VERSION_TEXT = "Version 2.75"

# ─── helpers ─────────────────────────────────────────
def id_to_names(ci):
    if pd.isna(ci) or not ci: return "Colorless"
    return ", ".join(COLOR_MAP.get(c,c) for c in str(ci).split(","))

def split_type_line(tline):
    if pd.isna(tline) or not tline: return "",""
    if "—" in tline: left,right=tline.split("—",1)
    elif " - " in tline: left,right=tline.split(" - ",1)
    else: left,right=tline,""
    return left.strip(),right.strip()

def get_bulk_meta():
    meta=requests.get("https://api.scryfall.com/bulk-data",timeout=30).json()
    entry=next(d for d in meta["data"] if d["type"]=="default_cards")
    return entry["updated_at"], entry["download_uri"]

def download_bulk(uri, stamp, log, bar):
    r=requests.get(uri,stream=True,timeout=60); r.raise_for_status()
    total=int(r.headers.get("Content-Length",0)); bar["maximum"]=total//1024
    done=0
    with open(BULK_JSON,"wb") as f:
        for chunk in r.iter_content(1024*64):
            if chunk:
                f.write(chunk); done+=len(chunk)
                bar["value"]=done//1024; bar.update()
    with open(BULK_META,"w") as f: f.write(stamp)
    bar["value"]=0; log("✅ Bulk DB downloaded.")

def ensure_bulk(log, bar):
    try: stamp, uri=get_bulk_meta()
    except Exception as e:
        log(f"❌ Cannot reach Scryfall: {e}"); return False
    need=(not os.path.exists(BULK_JSON) or
          not os.path.exists(BULK_META) or
          open(BULK_META).read().strip()!=stamp)
    if need:
        log("ℹ️ Bulk DB missing/outdated → downloading (~500 MB)…")
        download_bulk(uri,stamp,log,bar)
    else: log("✅ Bulk DB up-to-date.")
    return True

def load_lookup():
    with open(BULK_JSON,"r",encoding="utf-8") as f:
        return {c["id"]:c for c in json.load(f)}

def fetch_random_image():
    try:
        d=requests.get("https://api.scryfall.com/cards/random",timeout=20).json()
        img=requests.get(d["image_uris"]["normal"],timeout=20).content
        return Image.open(io.BytesIO(img)).resize((IMG_W,IMG_H))
    except Exception:
        return Image.open(CARD_BACK).resize((IMG_W,IMG_H))

def autosize(ws):
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = \
            max(len(str(c.value)) if c.value else 0 for c in col)+2

# ─── enrichment ──────────────────────────────────────
def enrich(path_in, path_out, log):
    df=pd.read_excel(path_in) if path_in.endswith(".xlsx") else pd.read_csv(path_in)
    if "Scryfall ID" not in df.columns:
        log("❌ 'Scryfall ID' column missing."); return
    df=df.dropna(subset=["Scryfall ID"]).reset_index(drop=True)
    log(f"📄 {len(df)} rows loaded.")

    prev=pd.read_excel(path_out) if os.path.exists(path_out) else pd.DataFrame()
    prev_ids=set(prev["Scryfall ID"]) if not prev.empty else set()
    df_new=df[~df["Scryfall ID"].isin(prev_ids)]
    log(f"🆕 New cards: {len(df_new)}")

    lookup=load_lookup(); rows=[]
    for i,sid in enumerate(df_new["Scryfall ID"]):
        cd=lookup.get(str(sid),{})
        left,right=split_type_line(cd.get("type_line",""))
        rows.append({
            "Card Name": cd.get("name",""),
            "Mana Value": cd.get("cmc",""),
            "Color": id_to_names(",".join(cd.get("color_identity",[]))),
            "Rarity": cd.get("rarity","").capitalize(),
            "Card": left,
            "Type": right,
            "Set Name": cd.get("set_name","")
        })
        if (i+1)%50==0: log(f"   ⏳ {i+1}/{len(df_new)}")

    df_meta=pd.DataFrame(rows)
    df_new=pd.concat([df_meta, df_new.reset_index(drop=True)], axis=1)
    df_all=pd.concat([prev, df_new], ignore_index=True)
    df_all=df_all.rename(columns={"Collector number":"Collector Number"})
    df_all=df_all.loc[:,~df_all.columns.duplicated()]
    for col in COL_ORDER:
        if col not in df_all.columns: df_all[col]=""
    df_all=df_all[COL_ORDER]

    with pd.ExcelWriter(path_out,engine="openpyxl") as w:
        df_all.to_excel(w,"MtG Collection",index=False)

    wb=load_workbook(path_out)
    ws=wb["MtG Collection"]
    autosize(ws)
    ws.auto_filter.ref = ws.dimensions
    wb.save(path_out)
    log(f"✅ Completed. File saved to:\n{path_out}")

# ─── GUI ─────────────────────────────────────────────
class MTGGUI:
    def __init__(self):
        self.root=Tk(); self.root.title("MTG Collection Enricher")
        self.root.geometry(f"{WIN_W}x{WIN_H}"); self.root.resizable(False,False)

        body=Frame(self.root); body.pack()
        left=Frame(body,width=IMG_W,height=IMG_H); left.pack(side="left",padx=5,pady=20)
        self.right=Frame(body,width=RIGHT_W); self.right.pack(side="right",fill="y",pady=60)

        self.img_label=Label(left); self.img_label.pack()

        Label(self.right,text="Select your ManaBox\ncollection export file").pack(pady=6)
        frm=Frame(self.right); frm.pack()
        self.file_var=StringVar()
        Entry(frm,textvariable=self.file_var,width=26).pack(side="left",padx=4)
        Button(frm,text="Browse",command=self.browse).pack(side="left")

        Button(self.right,text="RUN",bg="green",fg="white",
               width=14,command=self.run).pack(pady=8)

        self.progress=ttk.Progressbar(self.right,length=220,mode="determinate")
        self.progress.pack(pady=3)

        Label(self.right,text="Status / Log:").pack()
        self.log_box=scrolledtext.ScrolledText(self.right,width=34,height=12)
        self.log_box.pack(padx=4,pady=4)

        Label(self.right,text=VERSION_TEXT,fg="gray").pack(pady=(2,0))

        self.cycle_image()
        self.root.mainloop()

    def log(self,msg): self.log_box.insert("end",msg+"\n"); self.log_box.see("end")
    def browse(self):
        p=filedialog.askopenfilename(
            filetypes=[("Excel files","*.xlsx"),("CSV files","*.csv")])
        if p: self.file_var.set(p)

    def run(self):
        if not self.file_var.get():
            messagebox.showerror("Error","Choose an input file first."); return

        # prompt user before save dialog
        messagebox.showinfo("Save Location",
                            "Choose a location for your updated Excel file.")
        save_path=filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx")],
            initialfile="mtg_collection_enriched.xlsx"
        )
        if not save_path: return  # user cancelled

        def task():
            if ensure_bulk(self.log,self.progress):
                enrich(self.file_var.get(),save_path,self.log)
                self.progress["value"]=0
        threading.Thread(target=task,daemon=True).start()

    def cycle_image(self):
        img=fetch_random_image(); self.tk_img=ImageTk.PhotoImage(img)
        self.img_label.configure(image=self.tk_img)
        self.root.after(IMAGE_REFRESH,self.cycle_image)

if __name__=="__main__":
    MTGGUI()
