import email
import sys
from imaplib import IMAP4_SSL
import tempfile
import email.utils
import pdfplumber
from datetime import datetime
import html2text
from tkcalendar import DateEntry
import tkinter.font as tkFont
import tkinter
from tkinter import ttk, filedialog
import customtkinter
import os
import msal
from dotenv import load_dotenv
import re

# Outlook login data
load_dotenv()
CLIENT_ID = os.getenv("APPLICATION_ID")
USERNAME = os.getenv("OUTLOOK_USER")
CACHE_FILE = "token_cache.bin"
SCOPES = ["https://outlook.office.com/IMAP.AccessAsUser.All"]

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
    parts = msg.walk() if msg.is_multipart() else [msg]

    for part in parts:
        content_type = part.get_content_type()
        content_disposition = str(part.get("Content-Disposition", ""))

        if content_type == "text/plain" and "attachment" not in content_disposition:
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                complete_mail += payload.decode(charset, errors="ignore")

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

# Starting mail extraction
def start(date_entry, map_var, excel_path, conn):
    date_str = date_entry.get_date().strftime("%d-%b-%Y")
    map_name = map_var.get()
    excel_file = excel_path.get()
    mails = get_all_email_data(conn, map_name, date_str)
    print(len(mails), " mails gevonden")

    #with open("test_mails_output.txt", "w", encoding="utf-8") as f:
    #    for i, mail in enumerate(mails, start=1):
    #        f.write(f"\n{'=' * 60}\nMail {i}/{len(mails)} (lengte: {len(mail)} tekens)\n{'=' * 60}\n")
    #        f.write(mail + "\n")
    #print(f"{len(mails)} mails weggeschreven naar test_mails_output.txt")


def logout(conn, app):
    conn.logout()
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