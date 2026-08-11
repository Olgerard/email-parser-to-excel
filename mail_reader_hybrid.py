from imaplib import IMAP4_SSL
from tkcalendar import DateEntry
import tkinter.font as tkFont
import tkinter
from tkinter import ttk, filedialog
import customtkinter
import os
import imaplib
import msal
from dotenv import load_dotenv

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

# Fetch emails
def emai_to_text():
    print("email")

# Starting mail extraction
def start(date_entry, map_var, excel_path):
    date_str = date_entry.get_date().strftime("%d-%b-%Y")
    map_name = map_var.get()
    excel_file = excel_path.get()
    print(date_str, map_name, excel_file)
    emai_to_text()

def logout():
    print("logout")

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
        return [mailbox.decode() for mailbox in mailboxes]
    return []

def build_ui(mailboxes):
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
    run_btn = customtkinter.CTkButton(app, text="Start", command=lambda: start(date_entry, map, excel_path))
    browse_btn = customtkinter.CTkButton(app, text="Kies Excel-bestand", command=lambda: browse_file(excel_path, file_label, run_btn))
    browse_btn.pack(pady=(20,0))
    file_label.pack()

    run_btn.pack(pady=(20,10))
    run_btn.configure(state=tkinter.DISABLED)

    logout_btn = customtkinter.CTkButton(app, text="Log uit", command=logout)
    logout_btn.pack(pady=(20,10))

    finish_label = customtkinter.CTkLabel(app, text = "")
    finish_label.pack()

    progress_label = customtkinter.CTkLabel(app, text = "")
    progress_label.pack()

    return app

def main():
    conn = open_connection(False)
    mailboxes = get_mailboxes(conn)
    app = build_ui(mailboxes)
    app.mainloop()

if __name__ == "__main__":
    main()