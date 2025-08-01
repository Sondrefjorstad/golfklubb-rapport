from fastapi import FastAPI, UploadFile, File, Form
from starlette.middleware.sessions import SessionMiddleware
from fastapi.responses import HTMLResponse
from datetime import datetime
import pandas as pd
import io
import re
import calendar
import pdfplumber
import locale
import os
import shutil
import json

METADATA_FIL = "rapporter_metadata.json"

def last_metadata():
    if os.path.exists(METADATA_FIL):
        with open(METADATA_FIL, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except Exception:
                return []
    return []

def lagre_metadata(metadata):
    with open(METADATA_FIL, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Struktur: [{navn, sti, kilde, dato, bruker}]
lagrede_rapporter = []

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key="en_valgfri_lang_nok_streng")
try:
    locale.setlocale(locale.LC_ALL, 'nb_NO.UTF-8')
except locale.Error:
    # Railway/Heroku har ikke nb_NO.UTF-8, bruk systemets default
    locale.setlocale(locale.LC_ALL, '')

USERS = {
    "Sondre": "MoldeGK",
    "Trond": "MoldeGK",
    "Terje": "Accountor",
}

# ------------------------------------------
# Login
# ------------------------------------------
from fastapi import Request, Form, Response
from fastapi.responses import HTMLResponse, RedirectResponse

@app.get("/login")
def login_form():
    html = """
    <html>
    <head>
        <style>
            body {
                min-height: 100vh;
                margin: 0;
                padding: 0;
                background-image: linear-gradient(rgba(255,255,255,0.6), rgba(255,255,255,0.6)),
                url('https://moldegolf.no/files/images/archive/original/eikrembilde_premier-1.jpg');
                background-size: cover;
                background-position: center;
                font-family: Arial, sans-serif;
            }
            .login-container {
                background: rgba(255,255,255,0.95);
                border-radius: 16px;
                box-shadow: 0 6px 32px rgba(0,0,0,0.1);
                max-width: 370px;
                margin: 90px auto;
                padding: 36px 32px 32px 32px;
                display: flex;
                flex-direction: column;
                align-items: center;
            }
            h2 {
                text-align: center;
                margin-top: 0;
                font-size: 2.5rem;
            }
            label {
                margin-top: 18px;
                margin-bottom: 3px;
                display: block;
                text-align: left;
                width: 100%;
                font-size: 1.1rem;
            }
            input[type="text"], input[type="password"] {
                width: 100%;
                padding: 12px;
                margin-bottom: 12px;
                border: 1px solid #b3b3b3;
                border-radius: 6px;
                font-size: 1.05rem;
                box-sizing: border-box;
            }
            input[type="submit"] {
                background: #41a349;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 6px;
                padding: 13px;
                font-size: 1.12rem;
                width: 100%;
                margin-top: 5px;
                cursor: pointer;
                transition: background 0.2s;
            }
            input[type="submit"]:hover {
                background: #37863a;
            }
        </style>
    </head>
    <body>
        <div class="login-container">
            <h2>Logg inn</h2>
            <form method="post" action="/login">
                <label for="username">Brukernavn:</label>
                <input type="text" name="username" id="username" required>
                <label for="password">Passord:</label>
                <input type="password" name="password" id="password" required>
                <input type="submit" value="Logg inn">
            </form>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html)

@app.post("/login")
async def login(request: Request, username: str = Form(...), password: str = Form(...)):
    if USERS.get(username) == password:
        request.session["user"] = username
        return RedirectResponse("/", status_code=303)
    else:
        return HTMLResponse(content="<html><body>Feil brukernavn/passord. <a href='/login'>Prøv igjen</a></body></html>")

# ------------------------------------------
# Logout
# ------------------------------------------
@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login")

# ------------------------------------------
# Månedsvalgskjema
# ------------------------------------------
def vis_månedsvalgskjema(filnavn, kilde, filinnhold):
    NORSKE_MÅNEDER = [
        "Januar", "Februar", "Mars", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Desember"
    ]
    # Definer options FØR du bruker det i html!
    options = "".join([f"<option value='{i+1}'>{navn}</option>" for i, navn in enumerate(NORSKE_MÅNEDER)])
    html = f"""
    <h3>Finner ikke hvilken måned rapporten gjelder for <span style='color:#226622'>{kilde}</span>!</h3>
    <form method="post" action="/velg_rapportmåned_generell" enctype="multipart/form-data" id="månedskjema">
        <input type="hidden" name="kilde" value="{kilde}">
        <input type="hidden" name="filnavn" value="{filnavn}">
        <input type="hidden" name="filinnhold" value="{filinnhold.hex()}">
        <label for="måned">Velg måned:</label>
        <select name="måned" id="måned" required>
            {options}
        </select>
        <br><br>
        <button type="submit">Lagre rapport med valgt måned</button>
        <button type="button" onclick="lukkModal()">Avbryt</button>
    </form>
    """
    return HTMLResponse(content=html)

@app.post("/velg_rapportmåned_generell")
async def velg_rapportmåned_generell(
    request: Request,
    kilde: str = Form(...),
    filnavn: str = Form(...),
    filinnhold: str = Form(...),
    måned: int = Form(...)
):
    bruker = request.session.get("user", "Ukjent")
    contents = bytes.fromhex(filinnhold)
    NORSKE_MÅNEDER = [
        "Januar", "Februar", "Mars", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Desember"
    ]
    måned_navn = NORSKE_MÅNEDER[måned - 1].capitalize()

    amount = 0.0

    # Sjekk hvilken parser du skal bruke (generell logikk for Nayax)
    if kilde == "Nayax":
        with pdfplumber.open(io.BytesIO(contents)) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Finn beløp
        amount_match = re.search(r"Reimburse by Nayax:\s*kr\s*([\d\s.,]+)", text)
        if amount_match:
            amount_str = amount_match.group(1).replace(" ", "").replace(",", "")
            try:
                amount = float(amount_str)
            except Exception:
                amount = 0.0
        else:
            amount = 0.0

        if amount == 0.0:
            match = re.search(r"Billable Payments:\s*(?:\n|\r|\r\n|\s)*?(kr|NOK)\s*([\d\s.,]+)", text, re.IGNORECASE)
            if match:
                amount_str = match.group(2).replace(" ", "").replace(",", "")
                try:
                    amount = float(amount_str)
                except Exception:
                    amount = 0.0

        summer = pd.DataFrame([{
            "Varegruppe": "Kafe&kiosk",
            "Beløp inkl. MVA": amount,
            "Inntektskonto": 3000
        }])

        nøkkel = f"Nayax – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Nayax")
        opplastede_rapporter.setdefault("Nayax", {})[måned] = True

    # Her kan du legge inn tilsvarende for Vipps, Golfmore osv, hvis ønskelig!

    # Lagre fil fysisk
    filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filnavn}"
    filsti = os.path.join(UPLOAD_FOLDER, filename)
    with open(filsti, "wb") as f:
        f.write(contents)

    # Legg inn i metadata
    lagrede_rapporter.append({
        "rapportnavn": filnavn,
        "filsti": filsti,
        "kilde": kilde,
        "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "opplastet_av": bruker,
        "måned": måned_navn
    })
    lagre_metadata(lagrede_rapporter)

    from fastapi.responses import RedirectResponse

    return RedirectResponse("/", status_code=303)

# ------------------------------------------
# INITIERING
# ------------------------------------------
lagret_data = {}  # nøkkel = "Kilde – Måned", verdi = DataFrame
opplastede_kilder = set()
opplastede_rapporter = {}  # {"Vipps": {6: True, 7: True}, ...}
KILDER = ["Vipps", "Nets", "NyFaktura", "Golfmore", "Stripe", "Nayax", "Eagl"]

KONTO_TIL_VAREGRUPPE = {
    3000: "Kiosk&kafe",
    3001: "Golfutstyr",
    3021: "Sponsor",
    3023: "Golfbil",
    3120: "Greenfee",
    3119: "Turneringer",
    3121: "Medlemskontigent",
    3124: "Simulator",
    3123: "Drivingrange"
}

NORSKE_MÅNEDER = ["", "Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"]

def hent_norsk_måned(månedstall):
    try:
        return NORSKE_MÅNEDER[månedstall].capitalize()
    except Exception:
        return "Ukjent"

# ------------------------------------------
# LANDINGSSIDE
# ------------------------------------------
@app.get("/")
def main(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login")
    
    måned_navn_liste = [
        "Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", 
        "September", "Oktober", "November", "Desember"
    ]

    # Bygg oversiktstabell (Kilde x Måned)
    oversikt_html = "<h3>Opplastede rapporter per måned</h3><table border='1' style='margin:1 auto;'><tr><th>Kilde</th>"
    for mnd in måned_navn_liste:
        oversikt_html += f"<th>{mnd}</th>"
    oversikt_html += "</tr>"

    for kilde in KILDER:
        oversikt_html += f"<tr><td>{kilde}</td>"
        for mnd_nr in range(1, 13):
            check = "✔️" if opplastede_rapporter.get(kilde, {}).get(mnd_nr) else ""
            oversikt_html += f"<td style='text-align:center'>{check}</td>"
        oversikt_html += "</tr>"
    oversikt_html += "</table>"

    content = f"""
    <html>
    <head>
        <style>
            body {{
                display: flex; justify-content: center; align-items: center; min-height: 100vh;
                font-family: Arial, sans-serif; background-color: #f4f4f4;
                background-image: url('https://moldegolf.no/files/images/archive/original/eikrembilde_premier-3_5.jpg');
                background-size: cover; background-repeat: no-repeat; background-position: center;
            }}
            .upload-container {{
                background-color: rgba(255,255,255,0.95); padding: 40px; border-radius: 12px;
                box-shadow: 0 0 10px rgba(0,0,0,0.1); text-align: center; width: 800px;
            }}
            select, input[type="file"], button {{
                display: block; margin: 20px auto; font-size: 16px; padding: 10px; width: 100%;
            }}
            .drop-area {{
                border: 2px dashed #aaa; padding: 20px; margin-top: 20px; cursor: pointer; background-color: #f9f9f9;
            }}
            .drop-area.dragover {{ background-color: #e0e0e0; }}
            #filename-label {{ margin-top: 10px; font-size: 18px; color: #555; }}
            ul {{ list-style: none; padding: 0; margin-top: 30px; font-size: 20px; }}
        </style>
    </head>
    <body>
        <div class="upload-container">
            <img src="https://moldegolf.no/files/images/logo-header.png" alt="Molde GK logo" style="display:block; margin:0 auto 40px auto; max-height:80px;">
            <h2>Molde Golfklubb - Rapportopplasting</h2>
            <a href="/logout" style="display:inline-block; float:right; background:#eee; color:#226622;
                border-radius:5px; padding:8px 20px; text-decoration:none; margin-bottom:10px; font-weight: bold;">
                Logg ut
            </a>
            <form id="upload-form" enctype="multipart/form-data">
                <label for="kilde" style="text-align: left; display: block;">Velg kilde:</label>
                <select name="kilde" id="kilde-select">
                    {''.join([f'<option value="{k}">{k}</option>' for k in KILDER])}
                </select>
                <div class="drop-area" id="drop-area">Dra og slipp en Excel-fil her eller klikk for å velge</div>
                <input id="file-input" name="file" type="file" style="display: none;">
                <div id="filename-label"></div>
                <button type="button" onclick="submitForm(false)">Last opp</button>
                <button type="button" onclick="submitForm(true)">Se rapport</button>
            </form>
            {oversikt_html}
            <!-- Legg til etter {oversikt_html} i content-fstringen -->
            <div id="modal-overlay" style="display:none; position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(0,0,0,0.5); z-index:1000;">
                <div id="modal-content" style="background:#fff; border-radius:10px; max-width:400px; margin:10% auto; padding:30px; position:relative;">
                    <!-- Her settes skjemaet inn dynamisk -->
                </div>
            </div>
            <div style='margin-top:30px; text-align:center;'>
                <a href='/opplastede_rapporter' style='font-size:1.15em; color:#226622; text-decoration:underline;'>
                    Se alle opplastede rapporter og last ned
                </a>
            </div>
            <script>
                const dropArea = document.getElementById('drop-area');
                const fileInput = document.getElementById('file-input');
                const filenameLabel = document.getElementById('filename-label');
                dropArea.addEventListener('click', () => fileInput.click());
                dropArea.addEventListener('dragover', (e) => {{
                    e.preventDefault(); dropArea.classList.add('dragover');
                }});
                dropArea.addEventListener('dragleave', () => {{
                    dropArea.classList.remove('dragover');
                }});
                dropArea.addEventListener('drop', (e) => {{
                    e.preventDefault(); dropArea.classList.remove('dragover');
                    const files = e.dataTransfer.files; fileInput.files = files;
                    if (files.length > 0) {{
                        filenameLabel.textContent = "Valgt fil: " + files[0].name;
                    }}
                }});
                fileInput.addEventListener('change', () => {{
                    if (fileInput.files.length > 0) {{
                        filenameLabel.textContent = "Valgt fil: " + fileInput.files[0].name;
                    }}
                }});
                function submitForm(showReport) {{
                    const fileInputEl = document.getElementById("file-input");
                    const formData = new FormData();
                    const kilde = document.getElementById("kilde-select").value;
                    if (fileInputEl.files.length > 0) {{
                        formData.append("kilde", kilde); formData.append("file", fileInputEl.files[0]);
                        fetch("/uploadfile/", {{method: "POST", body: formData}})
                        .then(response => response.text())
                        .then(data => {{
                            // Hvis "Last opp" (showReport == false) og serveren sender skjema for måned, vis alltid skjemaet
                            if (!showReport && data.includes("<form")) {{
                                document.getElementById("modal-content").innerHTML = data;
                                document.getElementById("modal-overlay").style.display = "block";
                            }} else if (showReport) {{
                                document.open(); document.write(data); document.close();
                            }} else {{
                                location.reload();
                            }}
                        }});
                    }} else if (showReport) {{
                        window.location.href = "/rapportoversikt";
                    }}
                    else {{
                        alert("Vennligst velg en fil før du laster opp.");
                    }}
                }}
                function lukkModal() {{
                    document.getElementById("modal-overlay").style.display = "none";
                    }}
            </script>
        </div>
    </body>
    </html>
    """

    return HTMLResponse(content=content)

# ------------------------------------------
# PARSING
# ------------------------------------------

def parse_vipps(contents):
    try:
        df_raw = pd.read_excel(io.BytesIO(contents), sheet_name='Hovedsiden', header=None)
        periode_tekst = df_raw.iloc[3, 3] if isinstance(df_raw.iloc[3, 3], str) else ""
        match = re.search(r"\d{2}\.(\d{2})\.\d{4}", periode_tekst)
        måned = int(match.group(1)) if match else None
        måned_navn = hent_norsk_måned(måned)
        # LAGRE OPPLASTET KILDE/MÅNED
        opplastede_rapporter.setdefault("Vipps", {})[måned] = True
        df = pd.read_excel(io.BytesIO(contents), sheet_name='Hovedsiden')
        df.columns = ['Vare', 'Varegruppe', 'Antall', 'Beløp inkl. MVA', 'MVA', 'Beløp eks. MVA']
        df = df[['Varegruppe', 'Beløp inkl. MVA']].dropna()
        df['Beløp inkl. MVA'] = pd.to_numeric(df['Beløp inkl. MVA'], errors='coerce')
        df = df[df['Varegruppe'].apply(lambda x: isinstance(x, str))]
        df = df[~df['Varegruppe'].str.contains("Beløp", case=False)]
        df = df[df['Beløp inkl. MVA'] < 100000]
        summer = df.groupby('Varegruppe')['Beløp inkl. MVA'].sum().reset_index()
        summer["Inntektskonto"] = summer["Varegruppe"].map({v: k for k, v in KONTO_TIL_VAREGRUPPE.items()}).fillna("")
        nøkkel = f"Vipps – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Vipps")
        return HTMLResponse(content=f"<html><body><h3>Vipps-data lastet opp for {måned_navn}.</h3></body></html>")
    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Vipps-parser: {str(e)}</h3></body></html>")

def parse_nyfaktura(contents):
    try:
        # Les inn fil uten header for å hente måned og sjekke struktur
        raw = pd.read_excel(io.BytesIO(contents), header=None)
        # Hent måned fra celle B3 ("06-2025")
        periode_tekst = str(raw.iloc[2, 1])  # Gir for eksempel "06-2025"
        måned_nr = int(periode_tekst[:2]) if periode_tekst and len(periode_tekst) >= 2 else 0
        norske_måneder = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                          "august", "september", "oktober", "november", "desember"]
        måned_navn = norske_måneder[måned_nr].capitalize() if 1 <= måned_nr <= 12 else "Ukjent"
        
        # LAGRE OPPLASTET KILDE/MÅNED
        opplastede_rapporter.setdefault("NyFaktura", {})[måned_nr] = True

        # Les filen på nytt med header fra rad 4 (Excel rad 4 = header=3 i pandas)
        df = pd.read_excel(io.BytesIO(contents), header=3)

        # Hent alle rader med tall i både Konto og Sum
        data = df[["Konto", "Sum"]].dropna()
        data = data[data["Konto"].apply(lambda x: isinstance(x, (int, float)))]
        data["Inntektskonto"] = data["Konto"].astype(int)
        data["Beløp inkl. MVA"] = pd.to_numeric(data["Sum"], errors="coerce")

        # Mapping kontonummer til varegruppe
        konto_to_varegruppe = {
            3000: "Kiosk&kafe",
            3001: "Golfutstyr",
            3015: "Ukjent",
            3021: "Sponsor",
            3160: "Ukjent",
            3121: "Medlemskontigent",
            3900: "Ukjent"
        }
        data["Varegruppe"] = data["Inntektskonto"].map(konto_to_varegruppe).fillna("Ukjent")

        # Rydd til visningsformat
        visningstabell = data[["Varegruppe", "Beløp inkl. MVA", "Inntektskonto"]]

        # Lagre med kilde og måned i nøkkel
        nøkkel = f"NyFaktura – {måned_navn}"
        lagret_data[nøkkel] = visningstabell
        opplastede_kilder.add("NyFaktura")
        return HTMLResponse(content=f"<html><body><h3>NyFaktura-data lastet opp for {måned_navn}.</h3></body></html>")

    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i NyFaktura-parser: {str(e)}</h3></body></html>")

def parse_nets(contents):
    try:
        raw = pd.read_excel(io.BytesIO(contents), header=None)
        dato_fra_a2 = raw.iloc[1, 0]
        print("DEBUG: Innholdet i A2:", repr(dato_fra_a2))  # <-- Se hva som faktisk står her

        måned_nr = None
        if isinstance(dato_fra_a2, pd.Timestamp) or isinstance(dato_fra_a2, datetime):
            måned_nr = dato_fra_a2.month
        elif isinstance(dato_fra_a2, (float, int)):
            excel_start = pd.Timestamp('1899-12-30')
            parsed = excel_start + pd.to_timedelta(dato_fra_a2, unit="D")
            måned_nr = parsed.month
        elif isinstance(dato_fra_a2, str):
            parsed = pd.to_datetime(dato_fra_a2.strip(), dayfirst=True, errors='coerce')
            måned_nr = parsed.month if not pd.isnull(parsed) else None

        NORSKE_MÅNEDER = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                        "august", "september", "oktober", "november", "desember"]
        måned_navn = NORSKE_MÅNEDER[måned_nr].capitalize() if måned_nr else "Ukjent"

        # Så leses selve transaksjonsdataene med header
        df = pd.read_excel(io.BytesIO(contents), header=0)
        if "Order ID" not in df.columns or "Payment amount" not in df.columns:
            return HTMLResponse(content="<html><body><h3>Filen mangler nødvendige kolonner ('Order ID' og 'Payment amount').</h3></body></html>")
        
        # Ekstraher første bokstav i Order ID, og gjør om til store bokstaver
        df["Type"] = df["Order ID"].astype(str).str[0].str.upper()
        summer = df.groupby("Type")["Payment amount"].sum().reset_index()
        
        # Mapper til ønsket varegruppe og inntektskonto
        type_to_varegruppe = {"T": "Turnering", "N": "Greenfee"}
        varegruppe_to_konto = {"Turnering": 3119, "Greenfee": 3120}

        summer["Varegruppe"] = summer["Type"].map(type_to_varegruppe)
        summer = summer[summer["Varegruppe"].notnull()]  # Behold kun T/N

        summer["Inntektskonto"] = summer["Varegruppe"].map(varegruppe_to_konto)
        summer = summer[["Varegruppe", "Payment amount", "Inntektskonto"]]
        summer = summer.rename(columns={"Payment amount": "Beløp inkl. MVA"})
        
        # Lagre på riktig måte
        nøkkel = f"Nets – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Nets")
        opplastede_rapporter.setdefault("Nets", {})[måned_nr] = True

        return HTMLResponse(content=f"<html><body><h3>Nets-data lastet opp for {måned_navn}.</h3></body></html>")

    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Nets-parser: {str(e)}</h3></body></html>")

def parse_golfmore(contents):
       
    print("Golfmore-parser ble kalt")
    
    # Les filen uten header, så vi kan hente nøyaktige celler
    df = pd.read_excel(io.BytesIO(contents), header=None)

    # Finn dato fra B2 (rad 1, kolonne 1)
    dato_cell = df.iloc[1, 1]  # B2 er rad 1, kol 1 (0-indeksert)
    måned_nr = None

    if isinstance(dato_cell, pd.Timestamp):
        måned_nr = dato_cell.month
    else:
        # Prøv å konvertere string til dato
        try:
            dato = pd.to_datetime(str(dato_cell).strip(), dayfirst=True, errors="coerce")
            if not pd.isnull(dato):
                måned_nr = dato.month
        except Exception:
            måned_nr = None

    # Sett månedsnavn på norsk
    måned_navn = NORSKE_MÅNEDER[måned_nr].capitalize() if måned_nr else "Ukjent"

    # Finn sum fra D6 (rad 5, kolonne 3)
    sum_value = df.iloc[5, 3]  # D6 er rad 5, kol 3

    # Inntektskonto for "Simulator" - hent fra mappingen din, evt. hardkod 3124
    inntektskonto = 3123

    # Lag DataFrame likt som de andre parserne forventer
    summer = pd.DataFrame([{
        "Varegruppe": "Drivingrange",
        "Beløp inkl. MVA": sum_value,
        "Inntektskonto": inntektskonto
    }])

    # Lagre med nøkkel "Golfmore – Juni"
    nøkkel = f"Golfmore – {måned_navn}"
    lagret_data[nøkkel] = summer
    opplastede_kilder.add("Golfmore")
    opplastede_rapporter.setdefault("Golfmore", {})[måned_nr] = True

    return HTMLResponse(content=f"<html><body><h3>Golfmore-data lastet opp for {måned_navn}.</h3></body></html>")

def parse_nayax(contents, filnavn="nayax.pdf"):
    try:
        with pdfplumber.open(io.BytesIO(contents)) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        print("--- PDF TEXT ---")
        print(text)
        print("--- SLUTT PÅ PDF TEXT ---")

        period_match = re.search(
            r"Reimbursement Period:\s*([0-9]{2})/([0-9]{2})/([0-9]{4})\s*-\s*([0-9]{2})/([0-9]{2})/([0-9]{4})", text)
        if period_match:
            month = int(period_match.group(2))
            print("Fant måned:", month)
        else:
            month = None
            print("Fant ikke måned i rapport!")

        amount = 0.0
        match = re.search(r"Reimburse by Nayax:\s*(kr|NOK)\s*([\d\s.,]+)", text, re.IGNORECASE)
        if match:
            amount_str = match.group(2).replace(" ", "").replace(",", "")
            try:
                amount = float(amount_str)
            except Exception:
                amount = 0.0

        if amount == 0.0:
            match = re.search(r"Billable Payments:\s*(?:\n|\r|\r\n|\s)*?(kr|NOK)\s*([\d\s.,]+)", text, re.IGNORECASE)
            if match:
                amount_str = match.group(2).replace(" ", "").replace(",", "")
                try:
                    amount = float(amount_str)
                except Exception:
                    amount = 0.0

        print("SLUTTVERDI amount:", amount)

        # Hvis måned ikke ble funnet, la bruker velge måned (men print beløp uansett)
        if month is None or month < 1 or month > 12:
            print("Ingen måned - viser valgskjema!")
            return vis_månedsvalgskjema(
                filnavn=filnavn,
                kilde="Nayax",
                filinnhold=contents
            )

        måned_navn = NORSKE_MÅNEDER[month].capitalize() if month else "Ukjent"

        summer = pd.DataFrame([{
            "Varegruppe": "Kafe&kiosk",
            "Beløp inkl. MVA": amount,
            "Inntektskonto": 3000
        }])

        print("DATAFRAME SUMMER:", summer)

        nøkkel = f"Nayax – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Nayax")
        opplastede_rapporter.setdefault("Nayax", {})[month] = True

        return HTMLResponse(content=f"<html><body><h3>Nayax-data lastet opp for {måned_navn}. Beløp: {amount} kr.</h3></body></html>")

    except Exception as e:
        print("NAYAX PARSER FEIL:", e)
        return HTMLResponse(content=f"<html><body><h3>Feil i Nayax-parser: {str(e)}</h3></body></html>")

def parse_eagl(contents):
    print("Eagl: Parser kalt!")
    try:
        # Les filen uten header først, så vi kan finne måned og header
        df = pd.read_excel(io.BytesIO(contents), header=None)

        # Hent måned fra B2 (rad 1, kolonne 1)
        dato_str = str(df.iloc[1, 1])
        try:
            måned_nr = pd.to_datetime(dato_str).month
        except Exception:
            måned_nr = None

        måned_navn = NORSKE_MÅNEDER[måned_nr].capitalize() if måned_nr else "Ukjent"

        # Finn header-rad for "Price (NOK)"
        header_row = None
        for i in range(10):
            if "Price (NOK)" in df.iloc[i].values:
                header_row = i
                break

        if header_row is not None:
            df_data = pd.read_excel(io.BytesIO(contents), header=header_row)
            if "Price (NOK)" in df_data.columns:
                total_sum = df_data["Price (NOK)"].sum()
            else:
                total_sum = 0
        else:
            total_sum = 0

        # Lag DataFrame til visning
        summer = pd.DataFrame([{
            "Varegruppe": "Golfbil",
            "Beløp inkl. MVA": int(round(total_sum)),
            "Inntektskonto": 3023
        }])

        nøkkel = f"Eagl – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Eagl")
        opplastede_rapporter.setdefault("Eagl", {})[måned_nr] = True

        return HTMLResponse(content=f"<html><body><h3>Eagl-data lastet opp for {måned_navn}.</h3></body></html>")

    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Eagl-parser: {str(e)}</h3></body></html>")

def parse_stripe(contents, filename=""):

    print("Stripe: Parser kalt!")

    df = pd.read_csv(io.BytesIO(contents), header=None)

    # 1. Finn måned (tekst eller filnavn)
    måned_nr = None
    for i in range(min(5, len(df))):
        rad = str(df.iloc[i, 0])
        match = re.search(r"(\d{4})-(\d{2})-\d{2}", rad)
        if match:
            måned_nr = int(match.group(2))
            break
    if måned_nr is None and filename:
        match = re.search(r"_(\d{4})-(\d{2})-\d{2}_", filename)
        if match:
            måned_nr = int(match.group(2))
    NORSKE_MÅNEDER = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                      "august", "september", "oktober", "november", "desember"]
    måned_navn = NORSKE_MÅNEDER[måned_nr].capitalize() if måned_nr else "Ukjent"

    # 2. Finn beløp: Søk etter rad der kolonne 1 == "Account activity before fees"
    beløp = 0
    for i in range(len(df)):
        if str(df.iloc[i, 1]).strip() == "Account activity before fees":
            beløp_str = str(df.iloc[i, 2]).replace(",", "").replace(" ", "")
            try:
                beløp = int(float(beløp_str))
            except Exception:
                beløp = 0
            break

    varegruppe = "Drivingrange"
    inntektskonto = 3123

    summer = pd.DataFrame([{
        "Varegruppe": varegruppe,
        "Beløp inkl. MVA": beløp,
        "Inntektskonto": inntektskonto
    }])

    nøkkel = f"Stripe – {måned_navn}"
    lagret_data[nøkkel] = summer
    opplastede_kilder.add("Stripe")
    opplastede_rapporter.setdefault("Stripe", {})[måned_nr] = True

    print(måned_navn, varegruppe, beløp, inntektskonto)
    return HTMLResponse(content=f"<html><body><h3>Stripe-data lastet opp for {måned_navn}.</h3></body></html>")

# ------------------------------------------
# OPPLASTINGSENDPOINT
# ------------------------------------------
@app.post("/uploadfile/")
async def upload_file(kilde: str = Form(...), file: UploadFile = File(...), request: Request = None):
    bruker = request.session.get("user", "Ukjent")
    contents = await file.read()

    # Lagre fil fysisk
    filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
    filsti = os.path.join(UPLOAD_FOLDER, filename)
    with open(filsti, "wb") as f:
        f.write(contents)

    # Lagre metadata
    lagrede_rapporter.append({
        "rapportnavn": file.filename,
        "filsti": filsti,
        "kilde": kilde,
        "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "opplastet_av": bruker
    })
    lagre_metadata(lagrede_rapporter)

    if kilde == "Vipps":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_vipps(contents)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Vipps.</h3></body></html>")
    elif kilde == "NyFaktura":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_nyfaktura(contents)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for NyFaktura.</h3></body></html>")
    elif kilde == "Nets":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_nets(contents)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Nets.</h3></body></html>")
    elif kilde == "Golfmore":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_golfmore(contents)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Golfmore.</h3></body></html>")
    elif kilde == "Nayax":
        if file.filename.endswith(".pdf"):
            return parse_nayax(contents, filnavn=file.filename)
        else:
            return HTMLResponse(content="<html><body><h3>Kun PDF-filer støttes for Nayax.</h3></body></html>")
    elif kilde == "Eagl":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_eagl(contents)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Eagl.</h3></body></html>")
    elif kilde == "Stripe":
        if file.filename.endswith(".csv"):
            return parse_stripe(contents, filename=file.filename)
        else:
            return HTMLResponse(content="<html><body><h3>Kun CSV-filer støttes for Stripe.</h3></body></html>")    
    else:
        opplastede_kilder.add(kilde)
        return HTMLResponse(content=f"<html><body><h3>{kilde}-parser ikke implementert ennå.</h3></body></html>")

# ------------------------------------------
# Opplastede filer
# ------------------------------------------
from fastapi.responses import FileResponse

@app.get("/opplastede_rapporter")
def vis_opplastede_rapporter(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login")

    html = """
    <html><head>
    <style>
    body { font-family: Arial; background:#f5f7fa; padding:40px; }
    table { border-collapse:collapse; width:90%; margin:0 auto; }
    th, td { border:1px solid #bbb; padding:8px; text-align:left; }
    th { background:#e3ecef; }
    tr:nth-child(even) { background:#f9fbfc; }
    a { color:#296ab3; text-decoration:none; }
    </style>
    </head><body>
    <h2>Opplastede rapporter</h2>
    <table>
      <tr>
        <th>Rapportnavn</th>
        <th>Kilde</th>
        <th>Dato opplastet</th>
        <th>Opplastet av</th>
        <th>Last ned</th>
      </tr>
    """
    for rapport in reversed(lagrede_rapporter):
        link = f"/nedlast_fil?filsti={rapport['filsti']}"
        html += f"<tr><td>{rapport['rapportnavn']}</td><td>{rapport['kilde']}</td><td>{rapport['dato']}</td><td>{rapport['opplastet_av']}</td><td><a href='{link}'>Last ned</a></td></tr>"
    html += "</table></body></html>"
    return HTMLResponse(content=html)

# ------------------------------------------
# LASTE NED FIL
# ------------------------------------------
@app.get("/nedlast_fil")
def nedlast_fil(filsti: str, request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login")
    # Sjekk at filen faktisk finnes
    if not os.path.exists(filsti):
        return HTMLResponse(content="<h3>Filen finnes ikke.</h3>")
    return FileResponse(filsti, filename=os.path.basename(filsti), media_type="application/octet-stream")

# ------------------------------------------
# RAPPORTOVERSIKT
# ------------------------------------------
@app.get("/rapportoversikt")
def rapport_oversikt():
    # 1. Finn alle kilder og alle måneder som faktisk finnes
    kilder = ["Vipps", "NyFaktura", "Nets", "Golfmore", "Stripe", "Nayax", "Eagl"]
    alle_måneder = [m.capitalize() for m in NORSKE_MÅNEDER if m and any(f"{k} – {m.capitalize()}" in lagret_data for k in kilder)]
    if not alle_måneder:
        return HTMLResponse(content="<html><body><h2>Ingen rapporter er lastet opp ennå.</h2></body></html>")

    # 2. Start felles tabell med ny stil
    html = """
    <html>
    <head>
    <style>
    body {
        font-family: Arial;
        padding: 40px;
        background-color: #e6f2ff;
    }
    table {
        border-collapse: separate;
        border-spacing: 0;
        margin: 0 auto;
        background-color: #fff;
    }
    th, td {
        border: 1px solid #ddd;
        padding: 8px;
        vertical-align: top;
        white-space: nowrap;
    }
    th {
        background-color: #f2f2f2;
    }
    .månedrad {
        height: 48px;
        background: none !important;
    }
    .månedcelle {
        text-align: center;
        font-size: 1.25em;
        font-weight: bold;
        background: none !important;
        border: none !important;
        padding: 24px 0 8px 0;
    }
    .mellomrom {
        height: 15px;
        border: none !important;
        background-color: #e6f2ff;
    }
    </style>
    </head>
    <body>
    <h2>Rapportoversikt</h2>
    <table>
        <tr>
            <th>Måned</th>
            <th>Vipps</th>
            <th>NyFaktura</th>
            <th>Nets</th>
            <th>Golfmore</th>
            <th>Stripe</th>
            <th>Nayax</th>
            <th>Eagl</th>
        </tr>
    """

    # 3. Bygg rader måned for måned, med mellomrom mellom månedene
    for i, måned in enumerate(alle_måneder):
        # Legg inn mellomrom mellom månedene
        if i != 0:
            html += "<tr class='mellomrom'><td colspan='8'></td></tr>"

        html += f"<tr class='månedrad'>"
        html += f"<td class='månedcelle'>{måned}</td>"
        for kilde in kilder:
            nøkkel = f"{kilde} – {måned}"
            if nøkkel in lagret_data:
                df = lagret_data[nøkkel]
                # Lag minitabell for denne kilden/måneden
                html += "<td><table style='width: 220px;'>"
                html += "<tr><th>Varegruppe</th><th>Beløp inkl. MVA</th><th>Inntektskonto</th></tr>"
                for _, row in df.iterrows():
                    beløp = f"{int(row['Beløp inkl. MVA']):,}".replace(",", " ")
                    konto = row["Inntektskonto"] if "Inntektskonto" in row else ""
                    html += f"<tr><td>{row['Varegruppe']}</td><td>{beløp} kr</td><td>{konto}</td></tr>"

                total_belop = df["Beløp inkl. MVA"].sum()
                html += f"<tr style='font-weight:bold; background:#f6f6f6;'><td>Total</td><td>{f'{int(total_belop):,}'.replace(',', ' ')} kr</td><td></td></tr>"
                
                html += "</table></td>"
            else:
                html += "<td style='text-align:center; color:#bbb;'>–</td>"
        html += "</tr>"

    html += "</table></body></html>"
    return HTMLResponse(content=html)

