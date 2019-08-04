from pywinauto import application, keyboard
import sys
import os
import platform
import datetime

# Declare default variables
program_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
path_folder = "C:\\backups"
file_extension = "Outlook Data File (.pst)"
backup_name = "first"
options_save = "Do not export duplicate items"
password = "Some pasdd"


def item_selection(dlg, title_name, cont_type, keys_botton):
    """
    Select needed element in a new window
    """
    import_elem = dlg.child_window(title=title_name, control_type=cont_type)
    import_elem.click_input()
    import_elem.type_keys(keys_botton)


def creation_date(path_to_file):
    """
    Try to get the date that a file was created, falling back to when it was
    last modified if that isn't possible.
    """
    if platform.system() == 'Windows':
        return os.path.getmtime(path_to_file)
    else:
        stat = os.stat(path_to_file)
        try:
            return stat.st_birthtime
        except AttributeError:
            # We're probably on Linux. No easy way to get creation dates here,
            # so we'll settle for when its content was last modified.
            return stat.st_mtime


def create_arhive(dlg):
    """
    Creating backup file and retern result
    """
    dlg.type_keys("%FOI")
    item_selection(dlg, "Export to a file", "ListItem", "%N")

    item_selection(dlg, file_extension, "ListItem", "%N")

    keyboard.send_keys("{UP}")
    keyboard.send_keys("%n")

    dlg.Browse.click()
    path = dlg.child_window(title="All locations", control_type="SplitButton")
    path.click_input()
    keyboard.send_keys(path_folder + "{ENTER}")
    item_selection(dlg, "File name:", "Edit", backup_name + "{ENTER}")

    item_selection(dlg, options_save, "RadioButton", "%n")
    dlg.Finish.click()

    item_selection(dlg, "Password:", "Edit", password)
    if dlg.child_window(title="Verify Password:", control_type="Edit").exists():
        item_selection(dlg, "Verify Password:", "Edit", password)
        dlg.OK.click()
    else:
        dlg.OK.click()
        item_selection(dlg, "Password:", "Edit", password)
        dlg.OK.click()

    # check exsist arhive file
    date_update_file = datetime.datetime.fromtimestamp(creation_date(f"{path_folder}\\{backup_name}.pst")).strftime('%Y-%m-%d')
    date_now = datetime.datetime.today().strftime('%Y-%m-%d')
    if date_update_file == date_now:
        print("Well done!")
    else:
        print("Can`t create arhive")


def main():
    try:
        if os.path.isdir(path_folder):
            pass
        else:
            os.mkdir(path_folder)
        app = application.Application(backend="uia")
        app.start(program_path, timeout=10)
        dlg = app.window(title_re=".*- Outlook*")
        dlg.wait('visible')
        if dlg.child_window(title="Upload profile", control_type="Text").exists():
            print("There are no accounts in Outlook")
        else:
            create_arhive(dlg)
    except:
        print("Unexpected error:", sys.exc_info()[0])


if __name__ == "__main__":
    main()
# dlg.print_control_identifiers()
