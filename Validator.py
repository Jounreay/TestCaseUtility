import getpass as gp
import pathlib as path
import os
import win32com.client
import colorama
import Config as c
from importlib import reload


colorama.init(convert=True)
username = str(gp.getuser())
notepad_path = c.APP_PATHS['notepad']
datagrip_path = c.APP_PATHS['datagrip']
txt_format = c.JIRA_FORMAT['txt_format']
full_file = c.APP_PATHS, c.JIRA_FORMAT


##//TODO: Maybe add something here to read the file? It's either that or restart to reload the Config object.
def main(notepad_path, datagrip_path):
    reload(c)
    for key, value in c.APP_PATHS.items():
        if value != '':
            print(
                colorama.Fore.WHITE + 'Found ' + colorama.Fore.CYAN + '{0}'.format(key) + colorama.Fore.WHITE + ' at ''{0}'''.format(value))
            break
        else:
            ##//TODO: This evaluates the first key, then moves on. It needs to have the flexiblity, like a switch statement, to evaluate AS is.
            if 'notepad' in key:
                print(colorama.Fore.WHITE + 'Validating {0} path.'.format(key))
                notepad_validation(key, notepad_path)
                continue
            elif 'datagrip' in key:
                print(colorama.Fore.WHITE + 'Validating {0} path.'.format(key))
                datagrip_validation(key, datagrip_path)
            print('App paths {0} have been created!'.format(key))
            break


def notepad_validation(key, notepad_path):
    if notepad_path != '':
        print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.CYAN + 'NotePad++ path' + colorama.Fore.WHITE + ' at ''{0}'''.format(notepad_path))
        return main(notepad_path, datagrip_path)
    else:
        for root, dirs, files in os.walk('C:\\', topdown=True):
            for exe in path.Path(root).rglob('notepad++.exe'):
                if exe:
                    print(colorama.Fore.WHITE + 'Notepad++ path found at "{0}"'.format(exe))
                    os.path.isfile(path.Path(exe))
                    print(colorama.Fore.GREEN + 'Writing path to file...')
                    config_writer(key, str(exe))
                    print("Done!")
                    main(notepad_path, datagrip_path)

            ##//TODO: Finding that reload module seems to work when calling main(). Just reload the config module.


def datagrip_validation(key, datagrip_path):
    if datagrip_path != '':
        print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.LIGHTMAGENTA_EX + 'DataGrip path' + colorama.Fore.WHITE
              + ' at {0}'.format(c.APP_PATHS['datagrip']))
    else:
        print('DataGrippath configuration file not found, creating.')
        for root, dirs, files in os.walk(
                'C:\\Users\\%s\\AppData\\Roaming\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\'
                % username,
                topdown=True):
            for exe in path.Path(root).rglob("Data*.lnk"):
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut('{0}'.format(exe))
                config_writer(key, shortcut)
                print(colorama.Fore.GREEN + 'Done.')
                break


## //This took way too long to write. I think i finally got the syntax...again. pretty simple. Replaces the literal '' if there is an empty path, and replaces it with the
##// validators findings.
def config_writer(key, new_path):
    with open('Config.py', 'r') as r:
        config = r.read()
        newfile = config.replace("'{0}':r''".format(key), "'{0}':r'{1}'".format(key, new_path))
        with open('Config.py', 'w') as f:
            f.write(newfile)
            f.close()
            r.close()
            reload(c)


if __name__ == '__main__':
    main(notepad_path=notepad_path, datagrip_path=datagrip_path)
