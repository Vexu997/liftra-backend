import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, simpledialog
import os
from docx import Document
from docx.shared import Pt

# Folder, w którym będą zapisywane i wczytywane pliki
FOLDER_PATH = r'G:\Mój dysk\Projekt LIFTRA'

# --- Funkcje pomocnicze ---

def extract_columns(text):
    lines = text.strip().split('\n')
    extracted = []
    for line in lines:
        parts = line.split('\t')
        if len(parts) >= 5:
            extracted.append((parts[0], parts[2], parts[4], ""))
    return extracted

def show_input_screen():
    checklist_frame.grid_forget()
    input_frame.grid(row=0, column=0, sticky="nsew")
    global checklist_items
    checklist_items = []
    input_text.delete("1.0", tk.END)

def show_checklist_screen():
    input_frame.grid_forget()
    checklist_frame.grid(row=0, column=0, sticky="nsew")

def generate_and_display_checklist(items_with_states=None, from_input_text=False):
    # Ustaw pole nazwy pliku na puste przy generowaniu nowej listy
    file_name_var.set("")
    # Reszta funkcji
    for item in checklist_tree.get_children():
        checklist_tree.delete(item)
    global checklist_items
    if items_with_states is not None:
        # Dane w formacie: (is_checked, col1, col3, col5, uwaga)
        checklist_items = [item[1:5] for item in items_with_states]
        data_to_display = [(item[1], item[2], item[3], item[4], item[0]) for item in items_with_states]
    elif from_input_text:
        checklist_items = extract_columns(input_text.get("1.0", tk.END))
        data_to_display = [(col1, col3, col5, uwaga, False) for col1, col3, col5, uwaga in checklist_items]
    elif checklist_items:
        data_to_display = [(col1, col3, col5, uwaga, False) for col1, col3, col5, uwaga in checklist_items]
    else:
        data_to_display = []

    for col1, col3, col5, uwaga, is_checked in data_to_display:
        tag = 'checked' if is_checked else 'unchecked'
        values = (col1, col3, col5, uwaga)
        checklist_tree.insert("", "end", values=values, tags=(tag,))
        if is_checked:
            checklist_tree.tag_configure('checked', background='lightgreen')
        else:
            checklist_tree.tag_configure('unchecked', background='white')

    if not checklist_frame.winfo_ismapped():
        show_checklist_screen()
def generate_from_input():
    generate_and_display_checklist(from_input_text=True)

def show_context_menu(event):
    item_id = checklist_tree.identify_row(event.y)
    # Resetowanie stanu menu
    context_menu.entryconfig("Edytuj", state="disabled")
    context_menu.entryconfig("Usuń", state="disabled")
    context_menu.entryconfig("Dodaj/Edytuj uwagę", state="disabled")
    context_menu.entryconfig("Dodaj", command=add_item)

    if item_id:
        checklist_tree.selection_set(item_id)
        context_menu.entryconfig("Edytuj", state="normal", command=lambda: edit_item(item_id))
        context_menu.entryconfig("Usuń", state="normal", command=lambda: delete_item(item_id))
        context_menu.entryconfig("Dodaj/Edytuj uwagę", state="normal", command=lambda: add_edit_uwaga(item_id))
    else:
        checklist_tree.selection_remove(checklist_tree.selection())

    try:
        context_menu.tk_popup(event.x_root, event.y_root)
    finally:
        context_menu.grab_release()

def edit_item(item_id):
    if item_id:
        current_values = checklist_tree.item(item_id, 'values')
        if not current_values or len(current_values) < 4:
            print("Błąd: Nie można pobrać danych do edycji.")
            return
        col1, col3, col5, uwaga = current_values
        new_col1, new_col3, new_col5 = ask_for_edit_item(col1, col3, col5)
        if new_col1 is None or new_col5 is None:
            return
        current_tags = checklist_tree.item(item_id, 'tags')
        checklist_tree.item(item_id, values=(new_col1, new_col3, new_col5, uwaga), tags=current_tags)
        # Aktualizacja listy w pamięci
        current_items_with_states = get_all_items_with_states_from_treeview()
        global checklist_items
        checklist_items = [item[:4] for item in current_items_with_states]

def delete_item(item_id):
    if item_id:
        checklist_tree.delete(item_id)
        current_items_with_states = get_all_items_with_states_from_treeview()
        global checklist_items
        checklist_items = [item[:4] for item in current_items_with_states]

def ask_for_new_item():
    dialog = tk.Toplevel(root)
    dialog.title("Dodaj nowy element")
    dialog.grab_set()

    ttk.Label(dialog, text="Detal (*):").grid(row=0, column=0, padx=5, pady=5, sticky='e')
    entry_det = ttk.Entry(dialog, width=30)
    entry_det.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(dialog, text="PR:").grid(row=1, column=0, padx=5, pady=5)
    entry_pr = ttk.Entry(dialog, width=30)
    entry_pr.grid(row=1, column=1, padx=5, pady=5)

    ttk.Label(dialog, text="Ilość (*):").grid(row=2, column=0, padx=5, pady=5)
    entry_ilosc = ttk.Entry(dialog, width=30)
    entry_ilosc.grid(row=2, column=1, padx=5, pady=5)

    result = {}

    def on_ok():
        det = entry_det.get().strip()
        pr = entry_pr.get().strip()
        ilosc = entry_ilosc.get().strip()
        if not det:
            messagebox.showerror("Błąd", "Pole 'Detal' jest wymagane.")
            return
        if not ilosc:
            messagebox.showerror("Błąd", "Pole 'Ilość' jest wymagane.")
            return
        result['det'] = det
        result['pr'] = pr
        result['ilosc'] = ilosc
        dialog.destroy()

    def on_cancel():
        result.clear()
        dialog.destroy()

    btn_frame = ttk.Frame(dialog)
    btn_frame.grid(row=3, column=0, columnspan=2, pady=10)

    ttk.Button(btn_frame, text="OK", command=on_ok).pack(side='left', padx=5)
    ttk.Button(btn_frame, text="Anuluj", command=on_cancel).pack(side='left', padx=5)

    entry_det.focus()
    dialog.wait_window()

    if result:
        return result['det'], result['pr'], result['ilosc']
    else:
        return None, None, None

def add_item():
    det, pr, ilosc = ask_for_new_item()
    if det is None or ilosc is None:
        return
    new_uwaga = ""
    checklist_tree.insert("", "end", values=(det, pr, ilosc, new_uwaga), tags=('unchecked',))
    checklist_tree.tag_configure('unchecked', background='white')
    current_items_with_states = get_all_items_with_states_from_treeview()
    global checklist_items
    checklist_items = [item[:4] for item in current_items_with_states]

def add_edit_uwaga(item_id):
    if item_id:
        current_values = checklist_tree.item(item_id, 'values')
        if not current_values or len(current_values) < 4:
            print("Błąd: Nie można pobrać danych do edycji uwagi.")
            return
        current_uwaga = current_values[3]
        new_uwaga = simpledialog.askstring("Uwaga", "Wprowadź uwagę:", initialvalue=current_uwaga, parent=root)
        if new_uwaga is None:
            return
        col1, col3, col5, _ = current_values
        current_tags = checklist_tree.item(item_id, 'tags')
        checklist_tree.item(item_id, values=(col1, col3, col5, new_uwaga), tags=current_tags)
        current_items_with_states = get_all_items_with_states_from_treeview()
        global checklist_items
        checklist_items = [item[:4] for item in current_items_with_states]

def toggle_row_state(event):
    item_id = checklist_tree.identify_row(event.y)
    if item_id:
        current_tags = checklist_tree.item(item_id, 'tags')
        is_checked = 'checked' in current_tags
        if is_checked:
            checklist_tree.item(item_id, tags=('unchecked',))
            checklist_tree.tag_configure('unchecked', background='white')
        else:
            checklist_tree.item(item_id, tags=('checked',))
            checklist_tree.tag_configure('checked', background='lightgreen')

def get_all_items_with_states_from_treeview():
    items_data = []
    for item_id in checklist_tree.get_children():
        values = checklist_tree.item(item_id, 'values')
        tags = checklist_tree.item(item_id, 'tags')
        is_checked = 'checked' in tags
        if len(values) >= 4:
            items_data.append((is_checked, values[0], values[1], values[2], values[3]))
        elif len(values) == 3:
            items_data.append((is_checked, values[0], values[1], values[2], ""))
    return items_data

def is_checklist_complete():
    # Lista jest ukończona, gdy WSZYSTKIE elementy są zaznaczone
    for item in get_all_items_with_states_from_treeview():
        if not item[0]:
            return False
    return True

def save_checklist():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaInMemoryUpload
    import json

    SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
    CREDENTIALS_FILE = 'credentials.json'
    TOKEN_FILE = 'token.json'
    DRIVE_FOLDER_ID = '17gKONL0gLBx7Wvd4Cx3cBVEhHVdLQP2_'

    # Pobierz nazwę pliku
    base_name_input = file_name_var.get().strip()
    if not base_name_input:
        messagebox.showerror("Błąd", "Proszę wpisać nazwę pliku do zapisania.")
        return

    # Pobierz dane z Treeview
    items = get_all_items_with_states_from_treeview()
    if not items:
        messagebox.showwarning("Brak danych", "Brak pozycji do zapisania.")
        return

    is_complete = is_checklist_complete()

    # Obsługa prefixu [DONE]
    def determine_final_filename(name_input, is_done):
        name = name_input.strip()
        if name.lower().endswith('.xlsx'):
            name = name[:-5]
        has_prefix = name.startswith("[DONE]")
        clean = name[6:] if has_prefix else name
        return (f"[DONE]{clean}.xlsx", f"{clean}.xlsx") if is_done else (f"{clean}.xlsx", f"[DONE]{clean}.xlsx")

    final_filename, to_delete_filename = determine_final_filename(base_name_input, is_complete)
    sheet_title = final_filename.replace(".xlsx", "")

    try:
        # Autoryzacja
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        sheets_service = build('sheets', 'v4', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        # Usuń istniejący arkusz o nazwie `to_delete_filename`
        query = f"name='{to_delete_filename.replace('.xlsx', '')}' and trashed=false"
        results = drive_service.files().list(q=query, spaces='drive', fields="files(id, name)").execute()
        for file in results.get('files', []):
            drive_service.files().delete(fileId=file['id']).execute()

        # Utwórz nowy arkusz
        spreadsheet_body = {
            'properties': {'title': sheet_title}
        }
        sheet = sheets_service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
        sheet_id = sheet.get('spreadsheetId')

        # Przenieś arkusz do folderu
        drive_service.files().update(
            fileId=sheet_id,
            addParents=DRIVE_FOLDER_ID,
            removeParents='root',
            fields='id, parents'
        ).execute()

        # Dane do wpisania
        values = [["X", "Detal", "PR", "Ilość", "Uwagi"]]
        for is_checked, col1, col3, col5, uwaga in items:
            checkbox = "[X]" if is_checked else "[]"
            values.append([checkbox, col1, col3, col5, uwaga])

        data = {'values': values}

        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range='A1',
            valueInputOption='RAW',
            body=data
        ).execute()

        # Aktualizacja pola z nazwą
        file_name_var.set(final_filename.replace(".xlsx", ""))
        messagebox.showinfo("Sukces", f"Lista została zapisana do arkusza Google:\n{sheet_title}")

    except HttpError as error:
        messagebox.showerror("Błąd", f"Błąd podczas zapisu do Google Sheets: {error}")
    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")

def load_checklist():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.metadata.readonly']
    TOKEN_FILE = 'token.json'
    FOLDER_ID = '17gKONL0gLBx7Wvd4Cx3cBVEhHVdLQP2_'

    try:
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)

        # Pobierz pliki z folderu
        results = drive_service.files().list(
            q=f"'{FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
            fields="files(id, name)", orderBy="name"
        ).execute()
        files = results.get('files', [])

        if not files:
            messagebox.showinfo("Brak plików", "Brak arkuszy w folderze Google Drive.")
            return

        # Stwórz okno wyboru
        selection_window = tk.Toplevel(root)
        selection_window.title("Wybierz arkusz")
        selection_window.geometry("400x300")

        listbox = tk.Listbox(selection_window, font=('TkDefaultFont', 12))
        listbox.pack(fill='both', expand=True, padx=10, pady=10)

        for f in files:
            listbox.insert(tk.END, f['name'])

        def on_select():
            index = listbox.curselection()
            if not index:
                return
            selected_file = files[index[0]]
            file_id = selected_file['id']
            file_name = selected_file['name']

            try:
                result = sheets_service.spreadsheets().values().get(
                    spreadsheetId=file_id,
                    range='A2:E'
                ).execute()
                rows = result.get('values', [])

                items_with_states = []
                for row in rows:
                    checkbox_str = row[0].strip() if len(row) > 0 else ""
                    col1 = row[1] if len(row) > 1 else ""
                    col3 = row[2] if len(row) > 2 else ""
                    col5 = row[3] if len(row) > 3 else ""
                    uwaga = row[4] if len(row) > 4 else ""
                    is_checked = checkbox_str == "[X]"
                    items_with_states.append((is_checked, col1, col3, col5, uwaga))

                generate_and_display_checklist(items_with_states=items_with_states)
                file_name_var.set(file_name.replace(".xlsx", "").replace("[DONE]", ""))
                selection_window.destroy()

            except Exception as e:
                messagebox.showerror("Błąd", f"Błąd podczas wczytywania arkusza:\n{e}")

        tk.Button(selection_window, text="Wczytaj wybrany", command=on_select).pack(pady=10)

    except HttpError as e:
        messagebox.showerror("Błąd API", f"Błąd Google API: {e}")
    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")

# --- Inicjalizacja GUI ---

root = tk.Tk()
root.title("Lista Produktów")
root.geometry("950x600")
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

# Globalna lista tekstowa
checklist_items = []

# --- Ramka dla ekranu wprowadzania ---
input_frame = ttk.Frame(root)
input_frame.grid_rowconfigure(0, weight=1)
input_frame.grid_columnconfigure(0, weight=1)

input_text = scrolledtext.ScrolledText(input_frame, width=70, height=25)
input_text.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

generate_btn = ttk.Button(input_frame, text="Generuj listę kontrolną", command=generate_from_input)
generate_btn.grid(row=1, column=0, pady=5)

load_checklist_btn = ttk.Button(input_frame, text="Wczytaj listę kontrolną", command=load_checklist)
load_checklist_btn.grid(row=2, column=0, pady=5)

# --- Ramka dla listy kontrolnej ---
checklist_frame = ttk.Frame(root)
checklist_frame.grid_rowconfigure(0, weight=0)  # na górze pole do nazwy pliku
checklist_frame.grid_rowconfigure(1, weight=1)  # tabela
checklist_frame.grid_columnconfigure(0, weight=1)

# Dodajemy pole do nazwy pliku (nad tabelą)
top_subframe = ttk.Frame(checklist_frame)
top_subframe.grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

ttk.Label(top_subframe, text="POR/").pack(side='left', padx=5)
file_name_var = tk.StringVar()
file_name_entry = ttk.Entry(top_subframe, textvariable=file_name_var, width=50)
file_name_entry.pack(side='left', padx=5)

columns = ('col1', 'col3', 'col5', 'col_uwaga')
checklist_tree = ttk.Treeview(checklist_frame, columns=columns, show='headings', selectmode='browse')
checklist_tree.heading('col1', text='Detal')
checklist_tree.heading('col3', text='PR')
checklist_tree.heading('col5', text='Ilość')
checklist_tree.heading('col_uwaga', text='Uwagi')
checklist_tree.column('col1', width=150, minwidth=100, anchor='w')
checklist_tree.column('col3', width=200, minwidth=150, anchor='w')
checklist_tree.column('col5', width=250, minwidth=200, anchor='w')
checklist_tree.column('col_uwaga', width=300, minwidth=200, anchor='w')

# Paski przewijania
tree_vscroll = ttk.Scrollbar(checklist_frame, orient="vertical", command=checklist_tree.yview)
tree_hscroll = ttk.Scrollbar(checklist_frame, orient="horizontal", command=checklist_tree.xview)
checklist_tree.configure(yscrollcommand=tree_vscroll.set, xscrollcommand=tree_hscroll.set)

checklist_tree.grid(row=1, column=0, sticky="nsew")
tree_vscroll.grid(row=1, column=1, sticky="ns")
tree_hscroll.grid(row=2, column=0, sticky="ew")

checklist_tree.bind("<Double-1>", toggle_row_state)
checklist_tree.bind("<Button-3>", show_context_menu)

# Przyciski
back_btn = ttk.Button(checklist_frame, text="Wróć", command=show_input_screen)
back_btn.grid(row=3, column=0, pady=5, sticky="ew")
save_btn = ttk.Button(checklist_frame, text="Zapisz listę kontrolną", command=save_checklist)
save_btn.grid(row=4, column=0, pady=5, sticky="ew")

# --- Menu kontekstowe ---
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Edytuj")
context_menu.add_command(label="Usuń")
context_menu.add_command(label="Dodaj/Edytuj uwagę")
context_menu.add_separator()
context_menu.add_command(label="Dodaj", command=add_item)

# --- Uruchomienie ---
show_input_screen()
root.mainloop()