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

    # Bygg tabell med sticky "Kilde"-kolonne og scrollbare måneder
    oversikt_html = """
    <div class="oversikt-box">
      <h3 class="oversikt-title">Opplastede rapporter per måned</h3>
      <div class="oversikt-scroll">
      <table class="rapport-tabell">
        <thead>
        <tr>
          <th class='sticky-kilde'>Kilde</th>""" + "".join([f"<th>{mnd}</th>" for mnd in måned_navn_liste]) + "</tr></thead><tbody>"

    for kilde in KILDER:
        oversikt_html += f"<tr><td class='sticky-kilde'>{kilde}</td>"
        for mnd_nr in range(1, 13):
            check = "✔️" if opplastede_rapporter.get(kilde, {}).get(mnd_nr) else ""
            oversikt_html += f"<td>{check}</td>"
        oversikt_html += "</tr>"
    oversikt_html += "</tbody></table></div></div>"

    content = f"""
    <html>
    <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700&display=swap" rel="stylesheet">
        <style>
            html, body {{
                height: 100%; margin: 0; padding: 0;
            }}
            body {{
                font-family: 'Inter', Arial, sans-serif;
                background-image: url('https://moldegolf.no/files/images/archive/original/eikrembilde_premier-3_5.jpg');
                background-size: cover;
                background-repeat: no-repeat;
                background-position: center;
                min-height: 100vh;
            }}
            .sticky-logo {{
                position: sticky;
                top: 0;
                background: rgba(255,255,255,0.94);
                z-index: 1000;
                padding: 13px 0 11px 0;
                display: flex;
                justify-content: center;
                align-items: center;
                box-shadow: 0 3px 16px #b5dfc928;
            }}
            .sticky-logo img {{
                max-height: 62px;
            }}
            .upload-card {{
                background: rgba(255,255,255,0.98);
                padding: 44px 38px 32px 38px;
                border-radius: 19px;
                box-shadow: 0 12px 38px #24304228;
                max-width: 700px;
                width: 100%;
                margin: 32px auto 0 auto;
                display: flex;
                flex-direction: column;
                align-items: center;
                position: relative;
            }}
            .header-title {{
                font-size: 2.0rem;
                font-weight: 800;
                color: #244e31;
                letter-spacing: -0.5px;
                margin-bottom: 8px;
                text-align: center;
                width: 80%;
            }}
            .upload-desc {{
                color: #396546;
                font-size: 1.11em;
                margin-bottom: 20px;
                text-align: center;
                width: 100%;
            }}
            .logout-btn {{
                position: absolute;
                top: 26px;
                right: 38px;
                background: #e9f7ef;
                color: #338d56;
                border: none;
                border-radius: 7px;
                padding: 8px 18px;
                font-weight: bold;
                font-size: 1.02em;
                text-decoration: none;
                box-shadow: 0 2px 12px #338d5632;
                transition: background 0.12s;
            }}
            .logout-btn:hover {{ background: #c8eedc; }}
            .source-select-wrap {{
                margin: 0 auto 16px auto;
                width: 100%;
                display: flex;
                justify-content: center;
            }}
            .source-select {{
                background: #f6faf7;
                border: 2.5px solid #88c29d;
                border-radius: 15px;
                box-shadow: 0 2px 12px #276d4422;
                font-size: 1.2em;
                padding: 17px 12px;
                width: 100%;
                margin: 0 0 12px 0;
                appearance: none;
                outline: none;
                color: #234b2e;
                font-weight: 700;
                text-align: center;
                transition: border-color 0.18s;
                cursor: pointer;
            }}
            .source-select:focus, .source-select:hover {{
                border-color: #338d56;
                background: #f1fff4;
            }}
            label[for="kilde"] {{
                display: block; font-weight:700; text-align:center;
                margin-bottom: 5px; color:#32613d; font-size:1.14em;
            }}
            .drop-area {{
                border: 2.5px dashed #abd2bc;
                border-radius: 13px;
                padding: 36px 0 28px 0;
                margin-bottom: 16px;
                width: 100%;
                background: #f7fcfa;
                cursor: pointer;
                font-size: 1.1em;
                color: #284a33;
                text-align: center;
                transition: background 0.18s, border-color 0.16s;
                position: relative;
            }}
            .drop-area.dragover {{
                background: #e7f7ec;
                border-color: #3ca767;
            }}
            .drop-area svg {{
                width: 34px; height: 34px; color: #57bb7a; display: block; margin: 0 auto 8px auto;
            }}
            .filename-label {{
                margin-top: 10px; font-size: 1.09em; color: #2c5837; text-align: center;
            }}
            .main-btn {{
                background: #338d56;
                color: #fff;
                font-weight: 700;
                border: none;
                border-radius: 7px;
                padding: 13px 0;
                font-size: 1.12rem;
                width: 100%;
                margin-top: 12px;
                margin-bottom: 7px;
                cursor: pointer;
                box-shadow: 0 2px 10px #338d5612;
                transition: background 0.13s;
            }}
            .main-btn:hover {{
                background: #226c3b;
            }}
            .main-btn.secondary {{
                background:#fff;
                color:#338d56;
                border:1.4px solid #338d56;
            }}
            .main-btn.secondary:hover {{
                background: #e7f7ec;
            }}
            .oversikt-box {{
                margin-top: 35px;
                width: 100%;
                background: #e9f7ef;
                border-radius: 14px;
                box-shadow: 0 4px 18px #a1cfc911;
                padding: 20px 15px 14px 15px;
            }}
            .oversikt-title {{
                margin-bottom: 13px;
                font-size: 1.2em;
                color: #244e31;
                text-align: center;
                font-weight: 700;
            }}
            .oversikt-scroll {{
                width: 100%;
                overflow-x: auto;
            }}
            .rapport-tabell {{
                border-collapse: separate;
                width: 100%;
                background: #fff;
                margin: 0;
                font-size: 1.03em;
                border-radius: 13px;
                overflow: hidden;
                box-shadow: 0 2px 24px #4ab17a11;
                min-width: 850px;
            }}
            .rapport-tabell th, .rapport-tabell td {{
                padding: 8px 13px;
                text-align: center;
                border-bottom: 1px solid #ebf6f0;
                min-width: 74px;
            }}
            .rapport-tabell th {{
                background: #c9e8d8;
                font-weight: 700;
                font-size: 1.00em;
                color: #214f38;
                position: sticky;
                top: 0;
                z-index: 2;
            }}
            .sticky-kilde {{
                position: sticky;
                left: 0;
                background: #c9e8d8 !important;
                z-index: 3;
                box-shadow: 2px 0 6px #b7dcc62b;
                text-align: left !important;
                font-weight:700;
            }}
            .rapport-tabell tr:nth-child(even) {{ background: #f5fdfa; }}
            .rapport-tabell tr:hover td {{ background: #d8f5e3; }}
            .rapport-tabell td {{
                font-size: 1.05em;
                color: #25492f;
                font-weight: 500;
            }}
            .rapport-tabell td:empty {{ background: #f3f6f4; }}
            @media (max-width: 800px) {{
                .upload-card {{ padding: 17px 1vw 20px 1vw; max-width: 99vw; min-width: 0; }}
                .rapport-tabell th, .rapport-tabell td {{ font-size: 0.99em; padding: 6px 4px; min-width:58px;}}
            }}
            .se-alle-link {{
                margin-top: 22px;
                display: block;
                text-align: center;
                font-size: 1.13em;
                color: #338d56;
                text-decoration: underline;
                font-weight: 600;
            }}
            /* Snackbar / Toast style */
            #snackbar {{
                visibility: hidden;
                min-width: 280px;
                background-color: #2b9745;
                color: #fff;
                text-align: center;
                border-radius: 10px;
                padding: 20px 26px;
                position: fixed;
                left: 50%;
                bottom: 45px;
                font-size: 1.13em;
                transform: translateX(-50%);
                z-index: 2001;
                box-shadow: 0 6px 24px #338d5618;
                opacity: 0;
                transition: opacity 0.45s, bottom 0.32s;
            }}
            #snackbar.show {{
                visibility: visible;
                opacity: 1;
                bottom: 80px;
            }}
        </style>
    </head>
    <body>
        <div class="sticky-logo">
            <img src="https://moldegolf.no/files/images/logo-header.png" alt="Molde GK logo">
        </div>
        <div class="upload-card">
            <a href="/logout" class="logout-btn">Logg ut</a>
            <div class="header-title">Molde Golfklubb - Rapportopplasting</div>
            <div class="upload-desc">Dra og slipp Excel-fil, eller <span style="color:#338d56;font-weight:600;text-decoration:underline;cursor:pointer;" id="click-file-link">klikk for å velge rapport</span></div>
            <form id="upload-form" enctype="multipart/form-data" style="width:100%;">
                <div class="source-select-wrap">
                    <label for="kilde">Velg kilde:</label>
                </div>
                <select name="kilde" id="kilde-select" class="source-select">
                    {''.join([f'<option value="{k}">{k}</option>' for k in KILDER])}
                </select>
                <div class="drop-area" id="drop-area">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 17V3m0 0L5.5 9.5M12 3l6.5 6.5"/><rect width="18" height="18" x="3" y="3" rx="4"/></svg>
                    <span id="drop-msg">Slipp filen her, eller <span style="color:#338d56;font-weight:600;text-decoration:underline;cursor:pointer;" id="click-file-link-inside">klikk for å velge</span></span>
                </div>
                <input id="file-input" name="file" type="file" style="display: none;">
                <div id="filename-label" class="filename-label"></div>
                <button type="button" class="main-btn" onclick="submitForm(false)">Last opp</button>
                <button type="button" class="main-btn secondary" onclick="submitForm(true)">Se rapport</button>
            </form>
            {oversikt_html}
            <a href='/opplastede_rapporter' class="se-alle-link">
                Se alle opplastede rapporter og last ned
            </a>
        </div>
        <div id="modal-overlay" style="display:none; position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(0,0,0,0.4); z-index:1000;">
            <div id="modal-content" style="background:#fff; border-radius:14px; max-width:400px; margin:10% auto; padding:30px; position:relative;">
                <!-- Her settes skjemaet inn dynamisk -->
            </div>
        </div>
        <div id="snackbar">Fil er lastet opp!</div>
        <script>
        document.addEventListener("DOMContentLoaded", function() {{
            const dropArea = document.getElementById('drop-area');
            const fileInput = document.getElementById('file-input');
            const filenameLabel = document.getElementById('filename-label');

            // Klikk på drop-area åpner filvelger
            dropArea.addEventListener('click', function() {{
                fileInput.click();
            }});

            // Klikk på begge «klikk for å velge»-tekster åpner filvelger
            var clickFileLinks = [document.getElementById('click-file-link'), document.getElementById('click-file-link-inside')];
            clickFileLinks.forEach(function(el) {{
                if (el) {{
                    el.addEventListener('click', function(e) {{
                        e.stopPropagation();
                        fileInput.click();
                    }});
                }}
            }});

            // Drag and drop support
            dropArea.addEventListener('dragover', function(e) {{
                e.preventDefault();
                dropArea.classList.add('dragover');
            }});
            dropArea.addEventListener('dragleave', function(e) {{
                e.preventDefault();
                dropArea.classList.remove('dragover');
            }});
            dropArea.addEventListener('drop', function(e) {{
                e.preventDefault();
                dropArea.classList.remove('dragover');
                if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {{
                    fileInput.files = e.dataTransfer.files;
                    filenameLabel.textContent = "Valgt fil: " + e.dataTransfer.files[0].name;
                }}
            }});

            // Oppdater filnavn når bruker velger fil
            fileInput.addEventListener('change', function() {{
                if (fileInput.files.length > 0) {{
                    filenameLabel.textContent = "Valgt fil: " + fileInput.files[0].name;
                }}
            }});
        }});

        function showSnackbar(msg) {{
            var x = document.getElementById("snackbar");
            x.textContent = msg;
            x.className = "show";
            setTimeout(function() {{ x.className = x.className.replace("show", ""); }}, 3000);
        }}

        function submitForm(showReport) {{
            const fileInputEl = document.getElementById("file-input");
            const formData = new FormData();
            const kilde = document.getElementById("kilde-select").value;

            if (fileInputEl.files.length > 0) {{
                formData.append("kilde", kilde);
                formData.append("file", fileInputEl.files[0]);
                fetch("/uploadfile/", {{method: "POST", body: formData}})
                .then(response => response.text())
                .then(data => {{
                    if (!showReport && data.includes("<form")) {{
                        document.getElementById("modal-content").innerHTML = data;
                        document.getElementById("modal-overlay").style.display = "block";
                    }} else if (showReport) {{
                        document.open(); document.write(data); document.close();
                    }} else {{
                        showSnackbar("Fil er lastet opp!");
                        setTimeout(() => {{ location.reload(); }}, 1400);
                    }}
                }})
                .catch(err => {{
                    showSnackbar("Opplasting feilet!");
                }});
            }} else if (showReport) {{
                window.location.href = "/rapportoversikt";
            }}
            else {{
                showSnackbar("Vennligst velg en fil før du laster opp.");
            }}
        }}

        function lukkModal() {{
            document.getElementById("modal-overlay").style.display = "none";
        }}
        </script>
    </body>
    </html>
    """

    return HTMLResponse(content=content)

# ------------------------------------------
# PARSING
# ------------------------------------------

def parse_vipps(contents, file=None, bruker="Ukjent"):
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

        # ---- LAGRE FIL OG METADATA MED MÅNED ----
        if file is not None:
            # LAGRE FIL FYSISK
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
            filsti = os.path.join(UPLOAD_FOLDER, filename)
            with open(filsti, "wb") as fobj:
                fobj.write(contents)
            # LAGRE METADATA MED MÅNED
            lagrede_rapporter.append({
                "rapportnavn": file.filename,
                "filsti": filsti,
                "kilde": "Vipps",
                "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "opplastet_av": bruker,
                "måned": måned_navn       # <--- Denne linjen er nøkkelen!
            })
            lagre_metadata(lagrede_rapporter)

        return HTMLResponse(content=f"<html><body><h3>Vipps-data lastet opp for {måned_navn}.</h3></body></html>")
    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Vipps-parser: {str(e)}</h3></body></html>")

def parse_nyfaktura(contents, file=None, bruker="Ukjent"):
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

        # ---- LAGRE FIL OG METADATA MED MÅNED ----
        if file is not None:
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
            filsti = os.path.join(UPLOAD_FOLDER, filename)
            with open(filsti, "wb") as fobj:
                fobj.write(contents)
            lagrede_rapporter.append({
                "rapportnavn": file.filename,
                "filsti": filsti,
                "kilde": "NyFaktura",
                "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "opplastet_av": bruker,
                "måned": måned_navn   # <- Dette gjør at månedsoversikten alltid blir korrekt!
            })
            lagre_metadata(lagrede_rapporter)

        return HTMLResponse(content=f"<html><body><h3>NyFaktura-data lastet opp for {måned_navn}.</h3></body></html>")

    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i NyFaktura-parser: {str(e)}</h3></body></html>")

def parse_nets(contents, file=None, bruker="Ukjent"):
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

        # ---- LAGRE FIL OG METADATA MED MÅNED ----
        if file is not None:
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
            filsti = os.path.join(UPLOAD_FOLDER, filename)
            with open(filsti, "wb") as fobj:
                fobj.write(contents)
            lagrede_rapporter.append({
                "rapportnavn": file.filename,
                "filsti": filsti,
                "kilde": "Nets",
                "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "opplastet_av": bruker,
                "måned": måned_navn  # <-- Nøkkelen for korrekt visning!
            })
            lagre_metadata(lagrede_rapporter)

        return HTMLResponse(content=f"<html><body><h3>Nets-data lastet opp for {måned_navn}.</h3></body></html>")

    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Nets-parser: {str(e)}</h3></body></html>")

def parse_golfmore(contents, file=None, bruker="Ukjent"):
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
    NORSKE_MÅNEDER = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                      "august", "september", "oktober", "november", "desember"]
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

    # ---- LAGRE FIL OG METADATA MED MÅNED ----
    if file is not None:
        filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
        filsti = os.path.join(UPLOAD_FOLDER, filename)
        with open(filsti, "wb") as fobj:
            fobj.write(contents)
        lagrede_rapporter.append({
            "rapportnavn": file.filename,
            "filsti": filsti,
            "kilde": "Golfmore",
            "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "opplastet_av": bruker,
            "måned": måned_navn  # <-- Nøkkelen for korrekt visning!
        })
        lagre_metadata(lagrede_rapporter)

    return HTMLResponse(content=f"<html><body><h3>Golfmore-data lastet opp for {måned_navn}.</h3></body></html>")

def parse_nayax(contents, file=None, bruker="Ukjent", filnavn="nayax.pdf"):
    try:
        with pdfplumber.open(io.BytesIO(contents)) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        period_match = re.search(
            r"Reimbursement Period:\s*([0-9]{2})/([0-9]{2})/([0-9]{4})\s*-\s*([0-9]{2})/([0-9]{2})/([0-9]{4})", text)
        if period_match:
            month = int(period_match.group(2))
        else:
            month = None

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

        # Hvis måned ikke ble funnet, la bruker velge måned (men print beløp uansett)
        if month is None or month < 1 or month > 12:
            return vis_månedsvalgskjema(
                filnavn=filnavn,
                kilde="Nayax",
                filinnhold=contents
            )

        NORSKE_MÅNEDER = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                          "august", "september", "oktober", "november", "desember"]
        måned_navn = NORSKE_MÅNEDER[month].capitalize() if month else "Ukjent"

        summer = pd.DataFrame([{
            "Varegruppe": "Kafe&kiosk",
            "Beløp inkl. MVA": amount,
            "Inntektskonto": 3000
        }])

        nøkkel = f"Nayax – {måned_navn}"
        lagret_data[nøkkel] = summer
        opplastede_kilder.add("Nayax")
        opplastede_rapporter.setdefault("Nayax", {})[month] = True

        # ---- LAGRE FIL OG METADATA MED MÅNED ----
        if file is not None:
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
            filsti = os.path.join(UPLOAD_FOLDER, filename)
            with open(filsti, "wb") as fobj:
                fobj.write(contents)
            lagrede_rapporter.append({
                "rapportnavn": file.filename,
                "filsti": filsti,
                "kilde": "Nayax",
                "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "opplastet_av": bruker,
                "måned": måned_navn  # <-- Viktig!
            })
            lagre_metadata(lagrede_rapporter)

        return HTMLResponse(content=f"<html><body><h3>Nayax-data lastet opp for {måned_navn}. Beløp: {amount} kr.</h3></body></html>")

    except Exception as e:
        print("NAYAX PARSER FEIL:", e)
        return HTMLResponse(content=f"<html><body><h3>Feil i Nayax-parser: {str(e)}</h3></body></html>")

def parse_eagl(contents, file=None, bruker="Ukjent"):
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

        NORSKE_MÅNEDER = ["", "januar", "februar", "mars", "april", "mai", "juni", "juli",
                        "august", "september", "oktober", "november", "desember"]
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

        # ---- LAGRE FIL OG METADATA MED MÅNED ----
        if file is not None:
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
            filsti = os.path.join(UPLOAD_FOLDER, filename)
            with open(filsti, "wb") as fobj:
                fobj.write(contents)
            lagrede_rapporter.append({
                "rapportnavn": file.filename,
                "filsti": filsti,
                "kilde": "Eagl",
                "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "opplastet_av": bruker,
                "måned": måned_navn
            })
            lagre_metadata(lagrede_rapporter)

        return HTMLResponse(content=f"<html><body><h3>Eagl-data lastet opp for {måned_navn}.</h3></body></html>")
    
    except Exception as e:
        return HTMLResponse(content=f"<html><body><h3>Feil i Eagl-parser: {str(e)}</h3></body></html>")

def parse_stripe(contents, file=None, bruker="Ukjent"):
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
    if måned_nr is None and file is not None:
        match = re.search(r"_(\d{4})-(\d{2})-\d{2}_", file.filename)
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

    # ---- LAGRE FIL OG METADATA MED MÅNED ----
    if file is not None:
        filename = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}"
        filsti = os.path.join(UPLOAD_FOLDER, filename)
        with open(filsti, "wb") as fobj:
            fobj.write(contents)
        lagrede_rapporter.append({
            "rapportnavn": file.filename,
            "filsti": filsti,
            "kilde": "Stripe",
            "dato": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "opplastet_av": bruker,
            "måned": måned_navn  # <-- Lagrer måneden her!
        })
        lagre_metadata(lagrede_rapporter)

    print(måned_navn, varegruppe, beløp, inntektskonto)
    return HTMLResponse(content=f"<html><body><h3>Stripe-data lastet opp for {måned_navn}.</h3></body></html>")

# ------------------------------------------
# OPPLASTINGSENDPOINT
# ------------------------------------------
@app.post("/uploadfile/")
async def upload_file(kilde: str = Form(...), file: UploadFile = File(...), request: Request = None):
    bruker = request.session.get("user", "Ukjent")
    contents = await file.read()

    if kilde == "Vipps":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_vipps(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Vipps.</h3></body></html>")
    elif kilde == "NyFaktura":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_nyfaktura(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for NyFaktura.</h3></body></html>")
    elif kilde == "Nets":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_nets(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Nets.</h3></body></html>")
    elif kilde == "Golfmore":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_golfmore(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Golfmore.</h3></body></html>")
    elif kilde == "Nayax":
        if file.filename.endswith(".pdf"):
            return parse_nayax(contents, file=file, bruker=bruker, filnavn=file.filename)
        else:
            return HTMLResponse(content="<html><body><h3>Kun PDF-filer støttes for Nayax.</h3></body></html>")
    elif kilde == "Eagl":
        if file.filename.endswith(".xlsx") or file.filename.endswith(".xls"):
            return parse_eagl(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun Excel-filer støttes for Eagl.</h3></body></html>")
    elif kilde == "Stripe":
        if file.filename.endswith(".csv"):
            return parse_stripe(contents, file=file, bruker=bruker)
        else:
            return HTMLResponse(content="<html><body><h3>Kun CSV-filer støttes for Stripe.</h3></body></html>")
    else:
        opplastede_kilder.add(kilde)
        return HTMLResponse(content=f"<html><body><h3>{kilde}-parser ikke implementert ennå.</h3></body></html>")

# ------------------------------------------
# Opplastede filer
# ------------------------------------------
@app.get("/opplastede_rapporter")
def vis_opplastede_rapporter(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login")

    html = """
    <html><head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body { 
        font-family: 'Inter', Arial, sans-serif; 
        background: #1e2632;
        padding: 40px;
        min-height: 100vh;
    }
    h2 {
        text-align: center;
        color: #eaf0fc;
        letter-spacing: 1px;
        margin-bottom: 38px;
        margin-top: 18px;
        font-size: 2.2em;
    }
    table { 
        border-collapse: separate;
        border-spacing: 0;
        width: 94%;
        margin: 0 auto;
        background: #fff;
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 6px 42px #222d4455;
    }
    th, td {
        border: none;
        padding: 14px 15px;
        text-align: left;
        font-size: 1.04em;
        white-space: nowrap;
    }
    th {
        background: #29435b;
        color: #eaf0fc;
        font-weight: 700;
        font-size: 1.10em;
    }
    tr:nth-child(even) { background: #f2f6fb; }
    tr:nth-child(odd)  { background: #e4eaf3; }
    tr:hover td { background: #c2d8f4; }
    td {
        color: #243050;
        font-weight: 500;
    }
    a {
        color: #3571c7;
        text-decoration: none;
        font-weight: 600;
        transition: color 0.15s;
    }
    a:hover { color: #193454; }
    .fjern-btn {
        border: none;
        background: none;
        color: #d62424;
        font-size: 1.2em;
        cursor: pointer;
        margin-left: 14px;
        vertical-align: middle;
        transition: color 0.18s;
    }
    .fjern-btn:hover { color: #a90000; }
    </style>
    </head><body>
    <h2>Opplastede rapporter</h2>
    <table>
      <tr>
        <th>Rapportnavn</th>
        <th>Kilde</th>
        <th>Måned</th>
        <th>Dato opplastet</th>
        <th>Opplastet av</th>
        <th>Last ned</th>
      </tr>
    """

    for rapport in reversed(lagrede_rapporter):
        link = f"/nedlast_fil?filsti={rapport['filsti']}"
        # Bruk "måned" direkte hvis det finnes (nyere rapporter)
        måned = rapport.get("måned")
        if not måned:
            # For eldre rapporter: slå opp via lagret_data
            måned = ""
            for nøkkel in lagret_data:
                if nøkkel.startswith(f"{rapport['kilde']} – "):
                    # Siden du bare har én fil per nøkkel, denne matcher
                    måned = nøkkel.split(" – ", 1)[1]
                    break

        slett_knapp = f"""
        <button class='fjern-btn' title='Slett rapport' onclick="slettRapport('{rapport['rapportnavn']}', '{rapport['kilde']}')">
          &#10060;
        </button>
        """

        html += f"""
        <tr>
            <td>{rapport['rapportnavn']}</td>
            <td>{rapport['kilde']}</td>
            <td>{måned or ""}</td>
            <td>{rapport['dato']}</td>
            <td>{rapport['opplastet_av']}</td>
            <td>
                <a href='{link}'>Last ned</a>{slett_knapp}
            </td>
        </tr>
        """

    html += """
    </table>
    <script>
    function slettRapport(rapportnavn, kilde) {
        if (!confirm('Er du sikker på at du vil slette denne rapporten?')) return;
        fetch('/slett_rapport', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ rapportnavn: rapportnavn, kilde: kilde })
        })
        .then(res => {
            if (res.ok) location.reload();
            else alert('Noe gikk galt – rapporten ble ikke fjernet.');
        });
    }
    </script>
    </body></html>
    """
    return HTMLResponse(content=html)

# ------------------------------------------
# SLETTE EN NED FIL
# ------------------------------------------

from fastapi import Request
from fastapi.responses import JSONResponse

@app.post("/slett_rapport")
async def slett_rapport(request: Request):
    data = await request.json()
    rapportnavn = data["rapportnavn"]
    kilde = data["kilde"]

    # Slett fra lagrede_rapporter
    global lagrede_rapporter
    lagrede_rapporter = [r for r in lagrede_rapporter if not (r["rapportnavn"] == rapportnavn and r["kilde"] == kilde)]

    # Slett fra lagret_data for denne kilden (og evt. måned hvis unikt)
    for nøkkel in list(lagret_data):
        if nøkkel.startswith(f"{kilde} – "):
            lagret_data.pop(nøkkel)

    return JSONResponse(content={"status": "ok"})

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
    kilder = ["Vipps", "NyFaktura", "Nets", "Golfmore", "Stripe", "Nayax", "Eagl"]
    kilde_ikoner = {
        "Vipps": "💳", "NyFaktura": "🧾", "Nets": "🏦", "Golfmore": "⛳",
        "Stripe": "⛳", "Nayax": "🧃", "Eagl": "🚗"
    }
    alle_måneder = [m.capitalize() for m in NORSKE_MÅNEDER if m and any(f"{k} – {m.capitalize()}" in lagret_data for k in kilder)]
    if not alle_måneder:
        return HTMLResponse(content="<html><body><h2>Ingen rapporter er lastet opp ennå.</h2></body></html>")
    
    html = """
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body {
        font-family: 'Inter', Arial, sans-serif;
        margin: 0;
        background: #eaf2fb;
    }
    .header {
        position: sticky;
        top: 0; z-index: 100;
        background: #fafdffdd;
        padding: 28px 0 12px 0;
        margin-bottom: 12px;
        border-bottom: 1px solid #c7e3fc;
        box-shadow: 0 2px 16px #b8d5f74a;
    }
    .header-title {
        font-size: 2.1em;
        font-weight: 800;
        color: #285689;
        margin-left: 36px;
        letter-spacing: -0.5px;
        white-space: nowrap;
    }
    .info {
        margin: 0 36px 8px 36px;
        font-size: 1em;
        color: #3c5775;
        white-space: nowrap;
    }
    .rapport-filter {
        margin: 10px 36px 30px 36px;
        display: flex;
        gap: 16px;
        align-items: center;
        white-space: nowrap;
    }
    .kort-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(410px, 1fr));  /* bredere kort */
        gap: 32px;
        margin: 0 30px 42px 30px;
    }
    .kort {
        background: #fff;
        border-radius: 22px;
        box-shadow: 0 4px 30px #b8d5f72e;
        padding: 24px 22px 22px 22px;
        display: flex;
        flex-direction: column;
        transition: box-shadow 0.2s;
        position: relative;
        min-width: 390px;
        max-width: 900px;
    }
    .kort:hover {
        box-shadow: 0 6px 36px #b8d5f756;
        border: 1.5px solid #8bc3ef2a;
    }
    .kort-header {
        font-size: 1.15em;
        font-weight: 700;
        color: #275682;
        display: flex;
        align-items: center;
        gap: 9px;
        margin-bottom: 6px;
        white-space: nowrap;
    }
    .kort-måned {
        background: #e3effc;
        font-weight: 600;
        font-size: 1.02em;
        color: #2d5678;
        border-radius: 12px;
        padding: 2px 14px;
        margin-right: 6px;
        white-space: nowrap;
    }
    .kort-ikon {
        font-size: 1.4em;
        margin-right: 6px;
    }
    .data-tbl {
        width: 100%;
        border-collapse: collapse;
        font-size: 1em;
        margin-top: 8px;
    }
    .data-tbl th, .data-tbl td {
        white-space: nowrap;
    }
    .data-tbl th {
        background: #f4faff;
        color: #3666a7;
        font-weight: 700;
        border: none;
        padding: 5px 0;
    }
    .data-tbl td {
        border-bottom: 1px solid #e5eefd;
        padding: 4px 0;
        color: #374a5f;
    }
    .varegruppe-ukjent {
        color: #c35532; font-style: italic;
    }
    .totalrad {
        font-weight: 800;
        color: #214364;
        background: #f6fafd;
        border-top: 2px solid #d3e3fc;
    }
    .sum-tall {
        font-size: 1.09em;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    .endring {
        font-size: 0.93em;
        font-weight: 500;
        color: #008c53;
        margin-left: 8px;
    }
    /* Tooltip styling */
    .tooltip {
        border-bottom: 1px dotted #2d5678;
        cursor: help;
        position: relative;
    }
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 190px;
        background: #24466e;
        color: #fff;
        text-align: left;
        border-radius: 8px;
        padding: 8px;
        position: absolute;
        z-index: 1;
        top: 120%;
        left: 40%;
        margin-left: -40px;
        opacity: 0;
        transition: opacity 0.2s;
        font-size: 0.98em;
    }
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
    @media (max-width: 800px) {
        .kort-grid { grid-template-columns: 1fr; margin: 0 2vw; }
        .header, .rapport-filter, .info { margin-left: 10px; margin-right: 10px; }
    }
    </style>
    </head>
    <body>
      <div class="header">
        <span class="header-title">Regnskapsrapport <span style='font-weight:400;font-size:0.85em;'>Molde Golfklubb</span></span>
        <div class="info">Automatisk oppsummert per måned og inntektskilde. Hold over <span class="tooltip">kolonner<span class="tooltiptext">Du kan holde musepekeren over noen kolonner for forklaring.</span></span> for mer info.</div>
      </div>
      <div class="rapport-filter">
        <label for="måned">Måned:</label>
        <select id="måned" onchange="filterMåned()">
            <option value='alle'>Alle</option>""" + "".join([f"<option value='{m}'>{m}</option>" for m in alle_måneder]) + """
        </select>
      </div>
      <div class="kort-grid">
    """
    for måned in alle_måneder:
        for kilde in kilder:
            nøkkel = f"{kilde} – {måned}"
            if nøkkel in lagret_data:
                df = lagret_data[nøkkel]
                total_belop = df["Beløp inkl. MVA"].sum()
                df["Andel %"] = (df["Beløp inkl. MVA"] / total_belop * 100).round(1)
                html += f"""
                <div class="kort" data-måned="{måned}">
                    <div class="kort-header">
                        <span class="kort-ikon">{kilde_ikoner.get(kilde, '')}</span>
                        <span>{kilde}</span>
                        <span class="kort-måned">{måned}</span>
                    </div>
                    <table class="data-tbl">
                        <tr>
                            <th>Varegruppe</th>
                            <th>Beløp inkl. MVA</th>
                            <th class="tooltip">Inntektskonto
                                <span class="tooltiptext">Hvilken regnskapskonto inntekten er bokført på.</span>
                            </th>
                            <th class="tooltip">% andel
                                <span class="tooltiptext">Andel av totalen for denne kilden/måneden.</span>
                            </th>
                        </tr>
                """
                for _, row in df.iterrows():
                    beløp = f"{int(row['Beløp inkl. MVA']):,}".replace(",", " ")
                    konto = row["Inntektskonto"] if "Inntektskonto" in row else ""
                    varegruppe = row["Varegruppe"]
                    if varegruppe.lower() == "ukjent":
                        varegruppe_html = f"<span class='varegruppe-ukjent'>{varegruppe}</span>"
                    else:
                        varegruppe_html = varegruppe
                    andel = f"{row['Andel %']}%"
                    html += f"<tr><td>{varegruppe_html}</td><td>{beløp} kr</td><td>{konto}</td><td>{andel}</td></tr>"
                html += f"""<tr class='totalrad'><td>Total</td>
                            <td class="sum-tall">{f'{int(total_belop):,}'.replace(',', ' ')} kr</td>
                            <td></td><td>100%</td></tr>
                    </table>
                </div>
                """
    html += "</div>"  # kort-grid

    html += """
    <script>
    function filterMåned() {
        var val = document.getElementById('måned').value;
        var kort = document.querySelectorAll('.kort');
        kort.forEach(function(k) {
            k.style.display = (val === 'alle' || k.getAttribute('data-måned') === val) ? '' : 'none';
        });
    }
    </script>
    """

    html += "</body></html>"
    return HTMLResponse(content=html)
