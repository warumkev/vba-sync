import os
import win32com.client
import time
import pywintypes
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

#
# --- HELFERFUNKTIONEN ---
#

def get_component_info(vb_component_type):
    """Gibt den Unterordner und die Dateiendung für einen VBA-Komponententyp zurück."""
    if vb_component_type == 1:  # vbext_ct_StdModule
        return "Modules", ".bas"
    elif vb_component_type == 2:  # vbext_ct_ClassModule
        return "ClassModules", ".cls"
    elif vb_component_type == 3:  # vbext_ct_MSForm
        return "UserForms", ".frm"
    elif vb_component_type == 100: # vbext_ct_Document
        return "Sheets", ".cls"
    else:
        return "Misc", ".txt"

def _clean_vba_code_string(code_string):
    """Entfernt Header/Attribute aus einem Code-String für das Update via CodeModule."""
    lines = code_string.splitlines()
    start_index = 0
    for i, line in enumerate(lines):
        stripped_line = line.strip()
        if not stripped_line.startswith(('Attribute ', 'VERSION ', 'BEGIN', 'END', 'MultiUse ')):
            start_index = i
            break
    clean_lines = lines[start_index:]
    return "\r\n".join(clean_lines).strip()

def clean_exported_file(file_path):
    """
    Liest eine exportierte VBA-Datei mit der korrekten Windows-Kodierung (cp1252)
    und entfernt alle Header- und Attribut-Zeilen.
    """
    try:
        # Excel exportiert mit einer Windows-Codepage, typischerweise cp1252
        with open(file_path, 'r', encoding='cp1252') as f:
            lines = f.readlines()

        start_index = 0
        for i, line in enumerate(lines):
            stripped_line = line.strip()
            if not stripped_line.startswith(('Attribute ', 'VERSION ', 'BEGIN', 'END', 'MultiUse ')):
                start_index = i
                break
        
        clean_lines = lines[start_index:]
        clean_code = "".join(clean_lines).strip()

        # Schreibe die bereinigte Datei immer als UTF-8 für maximale Kompatibilität
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(clean_code)
        
        print(f"Datei '{os.path.basename(file_path)}' wurde bereinigt.")
    except UnicodeDecodeError as ude:
        print(f"Fehler bei der Kodierung der Datei {os.path.basename(file_path)}: {ude}. Versuchen Sie, die Datei in VBA neu zu speichern.")
    except Exception as e:
        print(f"Fehler beim Bereinigen der Datei {os.path.basename(file_path)}: {e}")

def _get_excel_app(file_name, abs_office_file):
    """Verbindet sich mit einer laufenden Excel-Instanz oder startet eine neue."""
    excel = None
    workbook = None
    we_started_excel = False
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except pywintypes.com_error:
        excel = win32com.client.Dispatch("Excel.Application")
        we_started_excel = True
    
    if we_started_excel:
        excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks(file_name)
    except pywintypes.com_error:
        if we_started_excel:
            workbook = excel.Workbooks.Open(abs_office_file)
        else:
            raise FileNotFoundError(f"Die Datei '{file_name}' ist in der laufenden Excel-Instanz nicht geöffnet.")
            
    return excel, workbook, we_started_excel

#
# --- KERNLOGIK (PULL, PUSH, WATCH) ---
#

def extract_vba(office_file, output_dir):
    """Extrahiert VBA-Code in die entsprechenden Unterordner."""
    excel, workbook, we_started_excel = None, None, False
    abs_office_file = os.path.abspath(office_file)
    abs_output_dir = os.path.abspath(output_dir)
    file_name = os.path.basename(abs_office_file)
    os.makedirs(abs_output_dir, exist_ok=True)

    try:
        excel, workbook, we_started_excel = _get_excel_app(file_name, abs_office_file)

        if not workbook.HasVBProject:
            print("Warnung: Die Arbeitsmappe enthält kein VBA-Projekt.")
            return
            
        vb_project = workbook.VBProject
        for component in vb_project.VBComponents:
            subdir, file_ext = get_component_info(component.Type)
            
            component_dir = os.path.join(abs_output_dir, subdir)
            os.makedirs(component_dir, exist_ok=True)
            
            export_path = os.path.join(component_dir, component.Name + file_ext)
            
            component.Export(export_path)
            print(f"Modul '{component.Name}' nach '{os.path.relpath(export_path)}' exportiert.")
            
            if file_ext == ".bas":
                clean_exported_file(export_path)

    finally:
        if workbook and we_started_excel: workbook.Close(SaveChanges=False)
        if excel and we_started_excel: excel.Quit()
        elif excel: print("Die verbundene Excel-Instanz bleibt geöffnet.")

def push_vba(source_dir, office_file):
    """Synchronisiert alle lokalen Dateien aus den Unterordnern mit dem VBA-Projekt."""
    excel, workbook, we_started_excel = None, None, False
    abs_office_file = os.path.abspath(office_file)
    abs_source_dir = os.path.abspath(source_dir)
    file_name = os.path.basename(abs_office_file)

    try:
        excel, workbook, we_started_excel = _get_excel_app(file_name, abs_office_file)
        vb_project = workbook.VBProject
        
        local_files = {}
        for root, _, files in os.walk(abs_source_dir):
            for file in files:
                module_name = os.path.splitext(file)[0]
                local_files[module_name.lower()] = os.path.join(root, file)
        
        project_modules = {comp.Name.lower() for comp in vb_project.VBComponents}

        for module_name_lower, file_path in local_files.items():
            with open(file_path, 'r', encoding='utf-8') as f:
                file_content = f.read()

            try:
                component = next(c for c in vb_project.VBComponents if c.Name.lower() == module_name_lower)
                clean_code = _clean_vba_code_string(file_content)
                if component.CodeModule.CountOfLines > 0:
                    component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                if clean_code:
                    component.CodeModule.AddFromString(clean_code)
                print(f"Modul '{component.Name}' erfolgreich aktualisiert.")
            except StopIteration:
                vb_project.VBComponents.Import(file_path)
                print(f"Neues Modul aus '{os.path.basename(file_path)}' importiert.")

        modules_to_remove = [c for c in vb_project.VBComponents if c.Name.lower() in project_modules - set(local_files.keys()) and c.Type != 100]
        for component in modules_to_remove:
            vb_project.VBComponents.Remove(component)
            print(f"Veraltetes Modul '{component.Name}' entfernt.")

        if we_started_excel: workbook.Save()
        else: print("Arbeitsmappe ist sichtbar geöffnet. Bitte manuell speichern.")
    finally:
        if workbook and we_started_excel: workbook.Close(SaveChanges=False)
        if excel and we_started_excel: excel.Quit()
        elif excel: print("Die verbundene Excel-Instanz bleibt geöffnet.")

def push_single_file(source_file_path, office_file):
    excel, workbook, we_started_excel = None, None, False
    try:
        abs_office_file = os.path.abspath(office_file)
        file_name = os.path.basename(abs_office_file)
        excel, workbook, we_started_excel = _get_excel_app(file_name, abs_office_file)
        
        vb_project = workbook.VBProject
        module_name = os.path.splitext(os.path.basename(source_file_path))[0]
        
        with open(source_file_path, 'r', encoding='utf-8') as f:
            file_content = f.read()

        try:
            component = next(c for c in vb_project.VBComponents if c.Name.lower() == module_name.lower())
            clean_code = _clean_vba_code_string(file_content)
            if component.CodeModule.CountOfLines > 0:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            if clean_code:
                component.CodeModule.AddFromString(clean_code)
            print(f"Modul '{component.Name}' erfolgreich aktualisiert.")
        except StopIteration:
            vb_project.VBComponents.Import(source_file_path)
            print(f"Neues Modul aus '{os.path.basename(source_file_path)}' importiert.")
        
        if we_started_excel: workbook.Save()
    finally:
        if workbook and we_started_excel: workbook.Close(SaveChanges=False)
        if excel and we_started_excel: excel.Quit()

def delete_module(module_name, office_file):
    excel, workbook, we_started_excel = None, None, False
    try:
        abs_office_file = os.path.abspath(office_file)
        file_name = os.path.basename(abs_office_file)
        excel, workbook, we_started_excel = _get_excel_app(file_name, abs_office_file)
        
        vb_project = workbook.VBProject
        try:
            component = next(c for c in vb_project.VBComponents if c.Name.lower() == module_name.lower())
            if component.Type != 100:
                vb_project.VBComponents.Remove(component)
                print(f"Modul '{component.Name}' erfolgreich entfernt.")
            else:
                print(f"Dokument-Modul '{component.Name}' kann nicht entfernt werden.")
        except StopIteration:
            print(f"Modul '{module_name}' nicht im Projekt gefunden, nichts zu löschen.")
            
        if we_started_excel: workbook.Save()
    finally:
        if workbook and we_started_excel: workbook.Close(SaveChanges=False)
        if excel and we_started_excel: excel.Quit()

class VbaChangeHandler(FileSystemEventHandler):
    def __init__(self, source_dir, office_file):
        self.source_dir = source_dir
        self.office_file = office_file
        print(f"👀 Watching for changes in '{os.path.abspath(self.source_dir)}'...")

    def on_modified(self, event):
        if not event.is_directory:
            print(f"\n🔄 File '{event.src_path}' modified. Triggering update...")
            push_single_file(event.src_path, self.office_file)
            print(f"\n👀 Watching for changes again...")

    def on_created(self, event):
        if not event.is_directory:
            print(f"\n✨ File '{event.src_path}' created. Triggering import...")
            push_single_file(event.src_path, self.office_file)
            print(f"\n👀 Watching for changes again...")

    def on_deleted(self, event):
        if not event.is_directory:
            print(f"\n🗑️ File '{event.src_path}' deleted. Triggering removal...")
            module_name = os.path.splitext(os.path.basename(event.src_path))[0]
            delete_module(module_name, self.office_file)
            print(f"\n👀 Watching for changes again...")

def start_watching(source_dir, office_file):
    """Starts the file system watcher."""
    event_handler = VbaChangeHandler(source_dir, office_file)
    observer = Observer()
    observer.schedule(event_handler, path=source_dir, recursive=True)
    observer.start()
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n🛑 Watcher stopped by user.")
    observer.join()