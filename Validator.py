import getpass as gp
import pathlib as path
import os
import win32com.client
import colorama
import Config as c

colorama.init(convert=True)
username = str(gp.getuser())
notepad_path = c.CONFIG['notepad']
datagrip_path = c.CONFIG['datagrip']
txt_format = c.CONFIG['txt_format']
full_file = c.CONFIG


def main(notepad_path, datagrip_path):
    app_paths = [notepad_path, datagrip_path]
    for f in app_paths:
        if f != '':
            print(
                colorama.Fore.WHITE + 'Found ' + colorama.Fore.CYAN + 'app path' + colorama.Fore.WHITE + ' at ''{0}'''.format(
                    f))
            break
        else:
            print(colorama.Fore.WHITE + 'Validating exe path.')
            notepad_validation(notepad_path)
            print(notepad_path)
            datagrip_validation(datagrip_path)
            print('App paths {0} have been created!'.format(app_paths))
            break


def notepad_validation(notepad_path):
    if notepad_path != '':
        print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.CYAN + 'NotePad++ path' + colorama.Fore.WHITE + ' at ''{0}'''.format(notepad_path))
        return main(notepad_path, datagrip_path)
    else:
        for root, dirs, files in os.walk('C:\\', topdown=True):
            for exe in path.Path(root).rglob('notepad++.exe'):
                print(colorama.Fore.WHITE + 'Notepad++ path not found, creating.')
                os.path.isfile(path.Path(exe))
                print(exe)
                if exe:
                    print('Found {0}'.format(exe))
                    print(colorama.Fore.GREEN + 'Writing path to file...')
                    config_writer(str(exe))
                    main(exe, datagrip_path)
                break


def datagrip_validation(datagrip_path):
    if datagrip_path != '':
        print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.LIGHTMAGENTA_EX + 'DataGrip path' + colorama.Fore.WHITE
              + ' at {0}'.format(c.CONFIG['datagrip']))
    else:
        print('DataGrippath configuration file not found, creating.')
        for root, dirs, files in os.walk(
                'C:\\Users\\%s\\AppData\\Roaming\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\'
                % username,
                topdown=True):
            for exe in path.Path(root).rglob("Data*.lnk"):
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut('%s' % exe)
                c.CONFIG['datagrip'] = str(shortcut.Targetpath)
                print(colorama.Fore.GREEN + 'Done.')
                break


## //This took way too long to write. I think i finally got the syntax...again. pretty simple. Replaces the literal '' if there is an empty path, and replaces it with the
##// validators findings.
def config_writer(new_path):
    with open('Config.py', 'r+') as f:
        config = f.read()
        newfile = config.replace("''", "'{0}'".format(new_path))
        print(newfile)
        f.write(newfile)
        f.close()


if __name__ == '__main__':
    main(notepad_path=notepad_path, datagrip_path=datagrip_path)
