import os, json, threading, io, time, requests, pandas as pd
from tkinter import Tk, Frame, Label, Button, Entry, StringVar, filedialog, scrolledtext, ttk, messagebox
from PIL import Image, ImageTk
from openpyxl import load_workbook

# GUI & card sizing
IMG_W, IMG_H = 400, 560
WIN_W = 700
WIN_H = IMG_H + 40  # match image height + top/bottom border
RIGHT_W = WIN_W - IMG_W - 20
IMAGE_REFRESH = 10_000

# filenames & column order
BULK_JSON = "default-cards.json"
BULK_META = "bulk_meta.txt"
CARD_BACK = "card_back.jpg"
ICON_FILE = "mtg.ico"  # Window and EXE icon

COLOR_MAP = {"W": "White", "U": "Blue", "B": "Black", "R": "Red", "G": "Green"}

COL_ORDER = [
    "Card Name", "Color", "Rarity", "Mana Value", "Power", "Toughness",
    "Card", "Type", "Set Name",
    "Foil", "Quantity", "Collector Number", "ManaBox ID", "Scryfall ID"
]

VERSION_TEXT = "Version 2.80"

# ─── helpers ─────────────────────────────────────────
def id_to_names(ci):
    if pd.isna(ci) or not ci:
        return "Colorless"
    return ", ".join(COLOR_MAP.get(c, c) for c in str(ci).split(","))


def split_type_line(tline):
    if pd.isna(tline) or not tline:
        return "", ""
    if "—" in tline:
        left, right = tline.split("—", 1)
    elif " - " in tline:
        left, right = tline.split(" - ", 1)
    else:
        left, right = tline, ""
    return left.strip(), right.strip()


def get_bulk_meta():
    meta = requests.get("https://api.scryfall.com/bulk-data", timeout=30).json()
    entry = next(d for d in meta["data"] if d["type"] == "default_cards")
    return entry["updated_at"], entry["download_uri"]


def download_bulk(uri, stamp, log, bar):
    r = requests.get(uri, stream=True, timeout=60)
    r.raise_for_status()
    total = int(r.headers.get("Content-Length", 0))
    bar["maximum"] = total // 1024
    done = 0
    start = time.time()
    last_log = start
    log(f"ℹ️ Starting download (~{total//1024//1024} MB)...")
    with open(BULK_JSON, "wb") as f:
        for chunk in r.iter_content(1024 * 64):
            if chunk:
                f.write(chunk)
                done += len(chunk)
                bar["value"] = done // 1024
                bar.update()
                now = time.time()
                if now - last_log >= 5 and done > 0:
                    elapsed = now - start
                    rate = done / elapsed
                    rem = (total - done) / rate if rate > 0 else 0
                    log(f"   ⏳ {done//1024}/{total//1024} KB ({(done/total)*100:.1f}%), ETA {int(rem)}s")
                    last_log = now
    with open(BULK_META, "w") as f:
        f.write(stamp)
    bar["value"] = 0
    log("✅ Bulk DB downloaded.")


def load_lookup():
    if not os.path.exists(BULK_JSON):
        raise FileNotFoundError(
            f"Bulk DB '{BULK_JSON}' not found. Use the Download button to fetch it."
        )
    with open(BULK_JSON, "r", encoding="utf-8") as f:
        return {c["id"]: c for c in json.load(f)}


def fetch_random_image():
    try:
        d = requests.get("https://api.scryfall.com/cards/random", timeout=20).json()
        img = requests.get(d["image_uris"]["normal"], timeout=20).content
        return Image.open(io.BytesIO(img)).resize((IMG_W, IMG_H))
    except Exception:
        return Image.open(CARD_BACK).resize((IMG_W, IMG_H))


def autosize(ws):
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value else 0 for c in col) + 2
        ws.column_dimensions[col[0].column_letter].width = width

# ─── enrichment ──────────────────────────────────────
def enrich(path_in, path_out, log):
    df_raw = pd.read_csv(path_in)
    if "Scryfall ID" not in df_raw.columns:
        log("❌ 'Scryfall ID' column missing.")
        return
    df = df_raw.dropna(subset=["Scryfall ID"]).reset_index(drop=True)
    log(f"📄 {len(df)} rows loaded for enrichment.")

    try:
        lookup = load_lookup()
    except FileNotFoundError as e:
        log(f"❌ {e}")
        return

    rows = []
    for i, sid in enumerate(df["Scryfall ID"]):
        cd = lookup.get(str(sid), {})
        left, right = split_type_line(cd.get("type_line", ""))
        power = cd.get("power")
        toughness = cd.get("toughness")
        if (power is None or toughness is None) and cd.get("card_faces"):
            face = cd["card_faces"][0]
            power = power or face.get("power")
            toughness = toughness or face.get("toughness")
        rows.append({
            "Card Name": cd.get("name", ""),
            "Mana Value": cd.get("cmc", ""),
            "Power": power or "",
            "Toughness": toughness or "",
            "Color": id_to_names(",".join(cd.get("color_identity", []))),
            "Rarity": cd.get("rarity", "").capitalize(),
            "Card": left,
            "Type": right,
            "Set Name": cd.get("set_name", "")
        })
        if (i + 1) % 50 == 0:
            log(f"   ⏳ {i+1}/{len(df)} enriched")

    df_meta = pd.DataFrame(rows)
    df_combined = pd.concat([df_meta, df.reset_index(drop=True)], axis=1)
    df_clean = df_combined.rename(columns={"Collector number": "Collector Number"})
    df_clean = df_clean.loc[:, ~df_clean.columns.duplicated()]
    for col in COL_ORDER:
        if col not in df_clean.columns:
            df_clean[col] = ""
    df_clean = df_clean[COL_ORDER]
    df_clean['Power'] = pd.to_numeric(df_clean['Power'], errors='coerce')
    df_clean['Toughness'] = pd.to_numeric(df_clean['Toughness'], errors='coerce')

    with pd.ExcelWriter(path_out, engine="openpyxl") as w:
        df_clean.to_excel(w, "MtG Collection", index=False)
    wb = load_workbook(path_out)
    ws = wb["MtG Collection"]
    autosize(ws)
    ws.auto_filter.ref = ws.dimensions
    wb.save(path_out)
    log(f"✅ Completed. File saved to:\n{path_out}")

# ─── GUI ─────────────────────────────────────────────
class MTGGUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("MTG Collection Enricher")
        try:
            self.root.iconbitmap(ICON_FILE)
        except Exception:
            pass
        self.root.geometry(f"{WIN_W}x{WIN_H}")
        self.root.resizable(False, False)

        body = Frame(self.root)
        body.pack(pady=10)
        left = Frame(body, width=IMG_W, height=IMG_H)
        left.pack(side="left", padx=10)
        self.right = Frame(body, width=RIGHT_W)
        self.right.pack(side="right", fill="y", padx=10)

        # Card image
        self.img_label = Label(left)
        self.img_label.pack()

        # ManaBox file selector
        Label(self.right, text="Select your ManaBox export file").pack(pady=(6,0))
        frm_in = Frame(self.right)
        frm_in.pack(pady=2)
        self.file_var = StringVar()
        Entry(frm_in, textvariable=self.file_var, width=26).pack(side="left", padx=4)
        Button(frm_in, text="Browse", command=self.browse_input).pack(side="left")

        # Output file selector
        Label(self.right, text="Select output Excel file").pack(pady=(8,0))
        frm_out = Frame(self.right)
        frm_out.pack(pady=2)
        self.out_var = StringVar()
        Entry(frm_out, textvariable=self.out_var, width=26).pack(side="left", padx=4)
        Button(frm_out, text="Browse", command=self.browse_output).pack(side="left")

        # ScryFall DB controls
        Label(self.right, text="ScryFall Database Version:").pack(pady=(10,0))
        self.db_version_var = StringVar()
        Label(self.right, textvariable=self.db_version_var).pack()
        self.db_button = Button(self.right, text="Download DB", command=self.handle_db)
        self.db_button.pack(pady=4)

        # Run enrichment
        Button(
            self.right, text="RUN", bg="green", fg="white",
            width=14, command=self.run
        ).pack(pady=(10,8))

        # Progress bar
        self.progress = ttk.Progressbar(self.right, length=220, mode="determinate")
        self.progress.pack(pady=3)

        # Status / Log
        Label(self.right, text="Status / Log:").pack()
        self.log_box = scrolledtext.ScrolledText(self.right, width=34, height=12)
        self.log_box.pack(padx=4, pady=4)

        Label(self.right, text=VERSION_TEXT, fg="gray").pack(pady=(2,0))

        # Initialize DB status and start image cycling
        self.refresh_db_status()
        self.cycle_image()
        self.root.mainloop()

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    def browse_input(self):
        p = filedialog.askopenfilename(filetypes=[("CSV files","*.csv")])
        if p:
            self.file_var.set(p)
            base, _ = os.path.splitext(p)
            default_out = f"{base}_enriched.xlsx"
            self.out_var.set(default_out)

    def browse_output(self):
        initial = self.out_var.get() or ""
        p = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")], initialfile=os.path.basename(initial), initialdir=os.path.dirname(initial)
        )
        if p:
            self.out_var.set(p)

    def refresh_db_status(self):
        if os.path.exists(BULK_META):
            try:
                with open(BULK_META) as f:
                    stamp = f.read().strip()
                self.db_version_var.set(stamp)
                self.db_button.config(text="Update DB", state="disabled")
            except Exception:
                self.db_version_var.set("Unknown")
                self.db_button.config(text="Download DB", state="normal")
        else:
            self.db_version_var.set("Not present")
            self.db_button.config(text="Download DB", state="normal")

    def handle_db(self):
        def task():
            try:
                stamp, uri = get_bulk_meta()
            except Exception as e:
                self.log(f"❌ Cannot reach Scryfall: {e}")
                return
            local = open(BULK_META).read().strip() if os.path.exists(BULK_META) else None
            if local == stamp:
                self.log("✅ Bulk DB up-to-date.")
            else:
                self.db_button.config(state="disabled")
                download_bulk(uri, stamp, self.log, self.progress)
            self.refresh_db_status()
        threading.Thread(target=task, daemon=True).start()

    def run(self):
        inp = self.file_var.get()
        outp = self.out_var.get()
        if not inp or not outp:
            messagebox.showerror("Error","Choose both input and output files first.")
            return
        threading.Thread(target=lambda: [enrich(inp, outp, self.log), setattr(self.progress, 'value', 0)], daemon=True).start()

    def cycle_image(self):
        img = fetch_random_image()
        self.tk_img = ImageTk.PhotoImage(img)
        self.img_label.configure(image=self.tk_img)
        self.root.after(IMAGE_REFRESH, self.cycle_image)

if __name__ == "__main__":
    MTGGUI()

# To build an EXE with this icon, use:
#   pyinstaller --onefile --windowed --icon=mtg.ico mtg_collection_enricher.py
