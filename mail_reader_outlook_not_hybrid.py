import email
import json
import sys
from imaplib import IMAP4_SSL
import tempfile
import email.utils
import pdfplumber
from datetime import datetime
import html2text
from anthropic import Anthropic
from tkcalendar import DateEntry
import tkinter.font as tkFont
import tkinter
from tkinter import ttk, filedialog
import customtkinter
import os
import msal
from dotenv import load_dotenv
import re
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.styles import Font, PatternFill

# Outlook login data
load_dotenv()
CLIENT_ID = os.getenv("APPLICATION_ID")
USERNAME = os.getenv("OUTLOOK_USER")
CACHE_FILE = "token_cache.bin"
SCOPES = ["https://outlook.office.com/IMAP.AccessAsUser.All"]

client = Anthropic(api_key=os.getenv("claude_api_key"))

# Get token to access Outlook mail
def get_token(verbose=False):
    # Cache to store token
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    # Create outlook object to make connection
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/common",
        token_cache=cache,
    )

    # See if an account is logged in locally, try to get access token with refresh token or with logged in account
    accounts = app.get_accounts(username=USERNAME)
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    # Manual login if needed
    if not result:
        if verbose:
            print("Geen geldig token in cache, eenmalige login nodig")

        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception(f"Device flow mislukt: {flow.get('error')}: {flow.get('error_description', flow)}")

        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    # Save cache
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", "Kon geen token krijgen"))

    return result["access_token"]


# Connecting with mail server
def open_connection(verbose=False):
    token = get_token(verbose)
    if verbose: print(f"Connecting to Outlook as {USERNAME}")
    connection = IMAP4_SSL("outlook.office365.com")
    auth_string = f"user={USERNAME}\1auth=Bearer {token}\1\1"
    connection.authenticate("XOAUTH2", lambda x: auth_string.encode())
    return connection

def clean_text(msg):
    msg = re.sub(r'\n\s*\n+', '\n', msg) #1 enter in plaats van meerdere
    msg = re.sub(r'[ \t]+', ' ', msg) #spatie voor tab
    return msg.strip()

def html_to_text(html):
    converter = html2text.HTML2Text()
    converter.ignore_links = True
    converter.ignore_images = True
    converter.body_width = 0
    return converter.handle(html)


def email_to_text(msg):
    date_tuple = email.utils.parsedate_tz(msg.get("Date"))
    if date_tuple:
        mail_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime("%d-%m-%Y")
    else:
        mail_date = ""
    complete_mail = f"Verzendatum: {mail_date} - "

    plain_text_found = False
    html_fallback = None

    parts = msg.walk() if msg.is_multipart() else [msg]

    for part in parts:
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition", ""))

        if content_type == "text/plain" and "attachment" not in content_disposition:
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                complete_mail += payload.decode(charset, errors="ignore")
                plain_text_found = True

        elif content_type == "text/html" and "attachment" not in content_disposition:
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                html_fallback = html_to_text(payload.decode(charset, errors="ignore"))

        elif content_type == "application/pdf":
            payload = part.get_payload(decode=True)
            if payload:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(payload)
                    tmp_filepath = tmp_file.name
                try:
                    with pdfplumber.open(tmp_filepath) as pdf:
                        pdf_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
                    complete_mail += "Text from PDF:" + pdf_text
                finally:
                    os.remove(tmp_filepath)

    if not plain_text_found and html_fallback:
        complete_mail += html_fallback

    return clean_text(complete_mail)

# Fetch emails and turn into strings
def get_all_email_data(conn, map_name, since_date):
    print("Map zoeken met naam: ", map_name)
    status, _ = conn.select(f'"{map_name}"')
    if status != "OK":
        raise ValueError(f"Kon map niet openen: {map_name}")
    criterium = f'(SINCE "{since_date}")'
    status, data = conn.search(None, criterium)
    if status != "OK":
        raise RuntimeError(f"Kon mails niet ophalen met startdatum:  {since_date}")
    mail_ids  = data[0].split()

    messages= []
    for mail_id in mail_ids :
        status, msg_data = conn.fetch(mail_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        messages.append(email_to_text(msg))
    return messages

def initialize_excel_sheet(excel_path, map):
    wb = load_workbook(excel_path)
    timestamp = datetime.now().strftime('%d-%m _ %H-%M')
    title = f"{map} - {timestamp}"
    ws = wb.create_sheet(title=title)

    ws.cell(row=1, column=1).value = "Boekingsdatum"
    ws.cell(row=1, column=2).value = "Datum"
    ws.cell(row=1, column=3).value = "Tour"
    ws.cell(row=1, column=4).value = "Passagier"
    ws.cell(row=1, column=5).value = "Bestemming"
    ws.cell(row=1, column=6).value = "Prijs"
    ws.cell(row=1, column=7).value = "Fee"
    ws.cell(row=1, column=8).value = ""
    ws.cell(row=1, column=9).value = "Voorgesteld alternatief"
    ws.cell(row=1, column=10).value = "Missed/ Earned Savings"
    ws.cell(row=1, column=11).value = "Prijs Excel"
    ws.cell(row=1, column=12).value = "PNR"
    ws.cell(row=1, column=13).value = "Airline"

    highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    row = 2
    ws.cell(row=row, column=1).value = "Vluchten"
    ws.cell(row=row, column=1).fill = highlight
    ws.cell(row=row, column=1).font = Font(bold=True)

    row += 2
    ws.cell(row=row, column=1).value = "Trein/bus"
    ws.cell(row=row, column=1).fill = highlight
    ws.cell(row=row, column=1).font = Font(bold=True)

    row += 2
    ws.cell(row=row, column=1).value = "Hotels"
    ws.cell(row=row, column=1).fill = highlight
    ws.cell(row=row, column=1).font = Font(bold=True)

    row += 2
    ws.cell(row=row, column=1).value = "Refunds"
    ws.cell(row=row, column=1).fill = highlight
    ws.cell(row=row, column=1).font = Font(bold=True)

    wb.save(excel_path)
    return title

def extract_flight_data(email_text):
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8000,
        temperature=0,
        messages=[
            {"role": "user", "content": f"""Analyseer deze email (bijlagen starten met 'Text from PDF:') en extraheer reis-/boekingsgegevens als JSON.
    
    Formatten per type:
    
    [{{
        "type": "vlucht of trein/bus",
        "boekingsdatum": "dd/mm/jjjj (meestal de datum waarop de mail gestuurd is)",
        "datums": [
            {{"datum": "dd/mm/jjjj (datum heenreis)"}},
            {{"datum": "dd/mm/jjjj (datum terugreis, indien aanwezig)"}}
        ],
        "passagiers": [
            {{"naam": "voornaam achternaam"}}
        ],
        "bestemming": [
            {{"vlucht": "Stad van vertrek - Stad van aankomst"}},
            {{"vlucht": "Terugvlucht (indien aanwezig)"}}
        ],
        "prijs": "totale eindprijs (bijv. 123.45)",
        "PNR": "boekingscode",
        "airline": "naam van de luchtvaartmaatschappij"
    }}]
    
    Hotels, als het leeg is moet het op dezelfde manier als bij vluchten:
    [
    {{
        "type": "hotel",
        "boekingsdatum": "",
        "datum": "Incheck datum (dd/mm) - uitcheck datum (dd/mm)",
        "passagiers": [
            {{"naam": "voornaam achternaam"}}
        ],
        "bestemming": "Naam hotel, Stad",
        "prijs": "",
        "PNR": "",
        "airline": ""
    }}]
    
    Deze vorm voor refunds, als het leeg is moet het op dezelfde manier als bij vluchten:
    [{{
        "type": "refund",
        "boekingsdatum": "dd/mm/jjjj",
        "datum": "dd/mm/jjjj",
        "passagiers": [
            {{"naam": "voornaam achternaam"}}
        ],
        "bestemming": "Enkel heenvlucht (vb Brussel - Amsterdam)",
        "prijs": vb -123.45",
        "PNR": "",
        "airline": ""
    }}]
    
    Regels:
    **Namen & Titels:**
    - Verwijder titels: Mr, Mrs, Ms → niet in naam
    - Format: "Voornaam Achternaam" (hoofdletters aan begin, NOOIT drukletters)
    - LOT format "ACHTERNAAM VOORNAAM Mr" → draai om naar "Voornaam Achternaam"
    
    **Plaatsnamen:**
    - Altijd Nederlands en voluit (vertaal indien nodig)
    - Alleen stadsnaam, geen luchthavennaam
    
    **Vluchten:**
    - Tussenstops samenvoegen: Amsterdam - Warschau - Wroclaw → Amsterdam - Wroclaw
    
    **Maatschappijen:**
    - Verkort: "LOT Airlines" → "LOT", "KLM Airlines" → "KLM", "TAP Air Portugal" → "TAP", "Expedia TAAP" → "Expedia", "Booking.com" → "Booking"
    - Altijd hoofdletter eerste letter
    - NMBS: PNR = DNR van ticket
    
    **Prijzen:**
    - Euro: alleen cijfer (123.45)
    - Andere valuta: cijfer + code (179.99 PLN)
    - Expedia: neem bedrag bij "Betaald aan Expedia", anders hoogste bedrag
    
    **Hotels:**
    - Passagier ontbreekt EN "Company" of "Pieter Smit" staat vermeld → passagier: "Company of Pieter Smit [BE/NL] [aantal]x"
    -Naam bij expedia het bedrag waar bij staat betaald aan expedia (dit is niet altijd het hoogste bedrag)
    
    **Refunds (KLM):**
    - Vaak alleen: boekingsdatum, PNR, mogelijk naam
    - Rest velden leeg laten
    
    **Lege velden:**
    - Leeg string "", NOOIT "N/A" of null
    
    **Output:**
    - Alleen pure JSON, geen ```json tags (Dit zorgt ervoor dat heel het programma crasht en is extreem belangrijk!!!), geen tekst eromheen
    - Exact formaat zoals voorbeelden
    EMAIL:
    \"\"\"
    {email_text}
    \"\"\"
    """}
        ]
    )
    return response.content[0].text

def extracted_flightdata_to_excel(ws, row, item):
    first_destination = True
    bestemmingen = item.get("bestemming",[])
    datums = item.get("datums", [])
    passagiers = item.get("passagiers", [])

    current_row = row

    for dest_index, dest in enumerate(bestemmingen):
        first_passenger = True

        datum = datums[dest_index].get("datum", "") if dest_index < len(datums) else ""

        for passenger in passagiers:

            ws.insert_rows(current_row)

            if first_passenger and first_destination:
                try:
                    date_obj = datetime.strptime(item.get("boekingsdatum", ""), "%d/%m/%Y").date()
                    cell = ws.cell(row=current_row, column=1, value=date_obj)
                    cell.number_format = "DD/MM/YYYY"
                except (ValueError, AttributeError):
                    ws.cell(row=current_row, column=1).value = item.get("boekingsdatum", "")

                try:
                    date_obj = datetime.strptime(datum, "%d/%m/%Y").date()
                    cell = ws.cell(row=current_row, column=2, value=date_obj)
                    cell.number_format = "DD/MM/YYYY"
                except (ValueError, AttributeError):
                    ws.cell(row=current_row, column=2).value = datum

                ws.cell(row=current_row, column=3).value = ""
                ws.cell(row=current_row, column=4).value = passenger.get("naam", "")
                ws.cell(row=current_row, column=5).value = dest.get("vlucht", "")
                ws.cell(row=current_row, column=6).value = ""

                try:
                    price = float(item["prijs"])
                    cell = ws.cell(row=current_row, column=11, value=price)
                    cell.number_format = "#,##0.00"
                except (ValueError, KeyError):
                    ws.cell(row=current_row, column=11).value = item.get("prijs", "")

                ws.cell(row=current_row, column=12).value = item.get("PNR", "")
                ws.cell(row=current_row, column=13).value = item.get("airline", "")

            else:
                try:
                    date_obj = datetime.strptime(datum, "%d/%m/%Y").date()
                    cell = ws.cell(row=current_row, column=2, value=date_obj)
                    cell.number_format = "DD/MM/YYYY"
                except (ValueError, KeyError):
                    ws.cell(row=current_row, column=2).value = datum

                ws.cell(row=current_row, column=4).value = passenger.get("naam", "")
                ws.cell(row=current_row, column=5).value = dest.get("vlucht", "")

            current_row += 1
            first_passenger = False

        first_destination = False

def extracted_data_to_excel(ws, row, item):
    passagiers = item.get("passagiers", [])
    first_passenger = True
    current_row = row

    for passenger in passagiers:
        ws.insert_rows(current_row)
        if first_passenger:
            try:
                date_obj = datetime.strptime(item["boekingsdatum"], "%d/%m/%Y").date()
                cell = ws.cell(row=current_row, column=1, value=date_obj)
                cell.number_format = "DD/MM/YYYY"
            except (ValueError, KeyError):
                ws.cell(row=current_row, column=1).value = item.get("boekingsdatum", "")
            try:
                date_obj = datetime.strptime(item["datum"], "%d/%m/%Y").date()
                cell = ws.cell(row=current_row, column=2, value=date_obj)
                cell.number_format = "DD/MM/YYYY"
            except (ValueError, KeyError):
                ws.cell(row=current_row, column=2).value = item.get("datum", "")
            ws.cell(row=current_row, column=3).value = ""
            ws.cell(row=current_row, column=5).value = item.get("bestemming", "")
            ws.cell(row=current_row, column=6).value = ""
            try:
                price = float(item["prijs"])
                cell = ws.cell(row=current_row, column=11, value=price)
                cell.number_format = "#,##0.00"
            except (ValueError, KeyError):
                ws.cell(row=current_row, column=11).value = item.get("prijs", "")
            ws.cell(row=current_row, column=12).value = item.get("PNR", "")
            ws.cell(row=current_row, column=13).value = item.get("airline", "")
        ws.cell(row=current_row, column=4).value = passenger.get("naam", "")
        current_row += 1
        first_passenger = False

def append_item_to_excel(item, excel_path, sheet_name):
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]

    ticket_type = item.get("type","")

    section_names = {
        "vlucht": "Vluchten",
        "trein/bus": "Trein/bus",
        "hotel": "Hotels"
    }
    section_name = section_names.get(ticket_type, "Refunds")

    section_start = 0
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == section_name:
            section_start = row
            break

    if section_start == 0:
        print(f"⚠️ WAARSCHUWING: Sectie '{section_name}' niet gevonden! Item overgeslagen.")
        return

    next_section_row = None
    for row in range(section_start + 1, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if cell_value in ["Vluchten", "Trein/bus", "Hotels", "Refunds"]:
            next_section_row = row
            break

    if next_section_row:
        insert_row = next_section_row - 1
    else:
        insert_row = ws.max_row + 1

    if ticket_type == "vlucht" or ticket_type == "trein/bus":
        extracted_flightdata_to_excel(ws, insert_row, item)
    else:
        extracted_data_to_excel(ws, insert_row, item)

    format_excel_cells(ws)
    wb.save(excel_path)

def format_excel_cells(ws):
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.number_format = numbers.FORMAT_DATE_DDMMYY

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=11, max_col=11):
        for cell in row:
            cell.number_format = "#,##0.00"

# Starting mail extraction
def start(date_entry, map_var, excel_path, conn):
    date_str = date_entry.get_date().strftime("%d-%b-%Y")
    map_name = map_var.get()
    excel_file = excel_path.get()
    mails = get_all_email_data(conn, map_name, date_str)
    amount_mails = len(mails)
    print(amount_mails, " mails gevonden")

    #with open("test_mails_output.txt", "w", encoding="utf-8") as f:
    #    for i, mail in enumerate(mails, start=1):
    #        f.write(f"\n{'=' * 60}\nMail {i}/{len(mails)} (lengte: {len(mail)} tekens)\n{'=' * 60}\n")
    #        f.write(mail + "\n")
    #print(f"{len(mails)} mails weggeschreven naar test_mails_output.txt")

    # Excel blad maken
    try:
        sheet_name = initialize_excel_sheet(excel_file, map_name)
    except Exception as e:
        print(e)
        return

    errors = 0
    number_of_handled_mails = 0
    for m in mails:
        try:
            json_string = extract_flight_data(m)
            print(json_string)
            parsed = json.loads(json_string)
            if isinstance(parsed, list):
                if len(parsed) > 0:
                    item = parsed[0]
                else:
                    print("⚠️ Lege lijst ontvangen van AI, email overgeslagen")
                    errors += 1
                    continue
            else:
                item = parsed

            append_item_to_excel(item, excel_file, sheet_name)

            number_of_handled_mails += 1

        except json.JSONDecodeError as e:
            print("JSON kon niet gelezen worden:", e)
            print(json_string)
            errors += 1
            continue

        except Exception as e:
            print(f"Onverwachte error: {e}")
            errors += 1
            continue
    print(f'{number_of_handled_mails} van de {amount_mails} verwerkt en in Excel gezet')

def logout(conn, app):
    conn.logout()
    if os.path.exists("token_cache.bin"):
        os.remove("token_cache.bin")
    app.destroy()
    sys.exit()

# Browsing files to select Excel file
def browse_file(excel_path, file_label, run_btn):
    filepath = filedialog.askopenfilename(
        title="Selecteer Excel-bestand",
        filetypes=[("Excel bestanden", "*.xlsx")],
    )
    if filepath:
        excel_path.set(filepath)
        file_label.configure(text=f"Geselecteerd: {os.path.basename(filepath)}", text_color="green")
        run_btn.configure(state=tkinter.NORMAL)

def get_mailboxes(conn):
    status, mailboxes = conn.list()
    if status == "OK":
        unparsed_names =  [mailbox.decode() for mailbox in mailboxes]
        parsed_names = []
        for unparsed_name in unparsed_names:
            match = re.match(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)', unparsed_name)
            if not match:
                parsed_names += [unparsed_name]
            else:
                name = match.group("name").strip()
                if name.startswith('"') and name.endswith('"'):
                    name = name[1:-1]
                parsed_names += [name]
        return parsed_names
    return []

def build_ui(mailboxes, conn):
    # System setting
    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")

    # App frame
    app = customtkinter.CTk()
    app.geometry("720x480")
    app.title("Excel assistent")

    # UI Elements
    date_title = customtkinter.CTkLabel(app, text="Vul de datum van de eerste mail in")
    date_title.pack(pady=(10,0))

    date = tkinter.StringVar()
    date_entry = DateEntry(app, date_pattern='dd/mm/yyyy')
    date_entry.configure(font=tkFont.Font(size=14))
    date_entry.pack(padx=10, pady=(10,20))

    map_title = customtkinter.CTkLabel(app, text="Vul de naam van de map in (vb. NL)")
    map_title.pack()

    map = tkinter.StringVar()
    cb = ttk.Combobox(app, values=mailboxes, textvariable=map, width=25, height=40)
    cb.configure(font=tkFont.Font(size=14) )
    cb.pack(pady=(0,10))

    def filter_mailboxes(event):
        typed = map.get().lower()
        if typed == "":
            cb["values"] = mailboxes
        else:
            cb["values"] = [m for m in mailboxes if typed in m.lower()]

    cb.bind("<KeyRelease>", filter_mailboxes)
    excel_path = tkinter.StringVar()
    file_label = customtkinter.CTkLabel(app, text="Geen bestand geselecteerd", text_color="red")
    run_btn = customtkinter.CTkButton(app, text="Start", command=lambda: start(date_entry, map, excel_path, conn))
    browse_btn = customtkinter.CTkButton(app, text="Kies Excel-bestand", command=lambda: browse_file(excel_path, file_label, run_btn))
    browse_btn.pack(pady=(20,0))
    file_label.pack()

    run_btn.pack(pady=(20,10))
    run_btn.configure(state=tkinter.DISABLED)

    logout_btn = customtkinter.CTkButton(app, text="Log uit", command=lambda: logout(conn, app))
    logout_btn.pack(pady=(20,10))

    finish_label = customtkinter.CTkLabel(app, text = "")
    finish_label.pack()

    progress_label = customtkinter.CTkLabel(app, text = "")
    progress_label.pack()

    return app

def main():
    conn = open_connection(False)
    mailboxes = get_mailboxes(conn)
    app = build_ui(mailboxes, conn)
    app.mainloop()

if __name__ == "__main__":
    main()