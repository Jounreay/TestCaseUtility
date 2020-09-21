import getpass as gp
from pathlib import Path
import pathlib as path
import subprocess as sp
import os
import win32com.client
import colorama
import time

colorama.init(convert=True)
username = str(gp.getuser())
userdir = r'C:\Users\%s\Desktop\Tasks' % username


def initial():
    print(colorama.Fore.GREEN + 'Hello! ')
    print(colorama.Fore.WHITE + 'Starting application.')
    try:
        for f in open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\ntpdpath.txt' % username, "r"):
            path.Path(f).rglob("*grip.exe")
            print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.CYAN + 'notepad ++ ' + colorama.Fore.WHITE
                  + ' path.')
    except FileNotFoundError:
        print(colorama.Fore.WHITE + 'Notepad++path configuration file not found, creating.')
        open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\ntpdpath.txt' % username, "x")

        for root, dirs, files in os.walk('C:\\Program Files\\', topdown=True):

            for exe in path.Path(root).rglob('notepad++.exe'):
                ntpdpath = open(str('C:\\Users\\%s\\Desktop\\TestCaseUtility\\ntpdpath.txt' % username), "w")
                print(colorama.Fore.GREEN + 'Writing path to file...')
                time.sleep(5)
                ntpdpath.write(str(exe))
                ntpdpath.close()
        with open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\ntpdpath.txt' % username, "r") as f:
            for line in f:
                notepadpath = line
                print(colorama.Fore.WHITE + '%s written.' % notepadpath)
                print(colorama.Fore.GREEN +'Done.')
    finally:

        try:
            for f in open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\datagrippath.txt' % username, "r"):
                datagripconfig = f
                print(colorama.Fore.WHITE + 'Found ' + colorama.Fore.LIGHTMAGENTA_EX + 'datagrip' + colorama.Fore.WHITE
                      + ' path.')
        except FileNotFoundError:
            open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\datagrippath.txt' % username, 'x')
            print('DataGrippath configuration file not found, creating.')
            for root, dirs, files in os.walk(
                'C:\\Users\\%s\\AppData\\Roaming\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\'
                    % username,
                    topdown=True):
                for exe in path.Path(root).rglob("Data*.lnk"):
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut('%s' % exe)
                    datagripexe = open(
                        str('C:\\Users\\%s\\Desktop\\TestCaseUtility\\datagrippath.txt' % username), "w")
                    print(colorama.Fore.WHITE + 'Writing path to file...')
                    time.sleep(5)
                    datagripexe.write(str(shortcut.Targetpath))
                    datagripexe.close()
            print(colorama.Fore.GREEN + 'Done.')

    with open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\ntpdpath.txt' % username, "r") as ntpdfile:
        ntpdconfig = ntpdfile.read()
    initial.ntpdplus = str(r'%s' % ntpdconfig)
    with open('C:\\Users\\%s\\Desktop\\TestCaseUtility\\datagrippath.txt' % username, "r") as dtagrip:
        dtgripconfig = dtagrip.read()
    initial.datagrip = str(r'%s' % dtgripconfig)

    try:
        os.chdir(userdir)
    except FileNotFoundError:
        os.mkdir(userdir)
    else:
        user_choice()


def user_choice():
    print(colorama.Fore.LIGHTWHITE_EX + '')
    method_choice = input('Create task?'' (c), Open Task? (o) or Exit? (e):  ')
    if method_choice == 'c':
        test_case_creator_input()
    if method_choice == 'o':
        test_case_open()
    if method_choice == 'e':
        raise SystemExit
    else:
        user_choice()


def test_case_creator_input():
    tcci = test_case_creator_input
    try:
        os.chdir(userdir)
        tcci.root_path = userdir
        tcci.task_type = input('Task type to create (MEDRXF, AP, EXT, etc.): ')
        tcci.task_number = input('Task number to create: ')
        tcci.testcase_directory = tcci.task_type + "-" + tcci.task_number
        tcci.txt_file = tcci.testcase_directory + ".txt"
        tcci.sql_file = tcci.testcase_directory + ".sql"
        tcci.dir_path = str(tcci.root_path + "\\" + tcci.testcase_directory)
        tcci.complt_path \
            = str(tcci.root_path + "\\" + tcci.testcase_directory + "\\" + tcci.txt_file)
        tcci.sqlcomplt_path = str(tcci.root_path + "\\" + tcci.testcase_directory + "\\" + tcci.sql_file)
        tcci.path = Path(tcci.root_path, tcci.dir_path)
    finally:
        create_file_function()


def create_file_function():
    #This pulls input, so the variables can be saved and passed to method that catches exceptions. Renaming methods here
    #So this is a little more readable.
    tcci = test_case_creator_input
    cff = create_file_function
    #try catch here to make sure to account for a file that already exists, rerouting it to the exception catch method.
    try:
        os.mkdir(tcci.dir_path)
        os.chdir(tcci.path)
    except FileExistsError:
        reroute = input("File exists, open? (y/n): ")
        if reroute == 'y':
            test_case_creator_reroute()
        if reroute == 'n':
            print(colorama.Fore.WHITE + "File not opened. Returning to main screen.")
            user_choice()
    else:
#The creation of the txt file, sql file, and jira formatting that will go in the text file.
            print(colorama.Fore.WHITE + "Directory '% s' created" % tcci.testcase_directory)
            cff.name_of_file = tcci.txt_file
            cff.file_path = tcci.complt_path
            cff.sql_file = tcci.root_path + "\\" + tcci.testcase_directory + "\\" + tcci.testcase_directory + ".sql"
    try:
        with open(r'C:\Users\%s\Desktop\TestCaseUtility\jiraformat.txt' % username) as f:
            with open(cff.file_path, "w") as new_file:
                for line in f:
                    new_file.write(line)
                new_file.close()
    except FileNotFoundError:
            print('Test case format configuration not found, please place jiraformat.txt within the '
                  'TestCaseUtility directory.')
            user_choice()
    else:
        open(cff.sql_file, "x")
        print(colorama.Fore.WHITE + "Task file '% s.txt'created." % tcci.testcase_directory)
        print("SQL file '% s.sql' created." % tcci.testcase_directory)
        print(colorama.Fore.LIGHTWHITE_EX + '')
        openfile = input('Open text file (t), sql file (s), both (b) or neither (n):  ')
        if openfile == 't':
            print(colorama.Fore.WHITE + 'Task file created and opened. Returning to main screen.')
            sp.Popen([initial.ntpdplus, cff.file_path])
            user_choice()
        if openfile == 's':
            print(colorama.Fore.WHITE + 'Task files created and opened. Returning to main screen.')
            sp.Popen([initial.datagrip, cff.sql_file])
            user_choice()
        if openfile == 'b':
            print(colorama.Fore.WHITE + 'Test case files created and opened. Returning to main screen.')
            sp.Popen([initial.datagrip, cff.sql_file])
            sp.Popen([initial.ntpdplus, cff.file_path])
            user_choice()
        if openfile == 'n':
            print(colorama.Fore.WHITE + 'Test case files created but not opened. Returning to main screen.')
            user_choice()


#the exception catch when a user attempts to create a test case that already exists, opening it instead.
def test_case_creator_reroute():
    tcci = test_case_creator_input
    os.chdir(tcci.path)
    complete_path = tcci.path
    sql_file = str(tcci.sql_file)
    txt_file = str(tcci.txt_file)
    print(colorama.Fore.LIGHTWHITE_EX + '')
    try:
        open_file = input('Open text file (t), sql file (s), both (b) or neither (n): ')
        if open_file == 't':
            for txt_file in Path(complete_path).rglob(str(txt_file)):
                sp.Popen([initial.ntpdplus, txt_file])
                print(colorama.Fore.WHITE + 'Text file opened. Returning to main screen.')
                user_choice()
        if open_file == 's':
            for sql_file in Path(complete_path).rglob(str(sql_file)):
                sp.Popen([initial.datagrip, sql_file])
                print(colorama.Fore.WHITE + 'Sql file opened. Returning to main screen.')
                user_choice()
        if open_file == 'b':
            for txt_file in Path(complete_path).rglob(str(txt_file)):
                sp.Popen([initial.ntpdplus, txt_file])
            for sql_file in Path(complete_path).rglob(str(sql_file)):
                sp.Popen([initial.datagrip, sql_file])
                print(colorama.Fore.WHITE + 'Files opened. Returning to main screen.')
                user_choice()
        if open_file == 'n':
            print(colorama.Fore.WHITE + 'Files not opened. Returning to main screen.')
            user_choice()
    except FileNotFoundError:
        print(colorama.Fore.RED + 'File(s) not found, returning to main screen.')
        user_choice()
    else:
        user_choice()


#pretty self explanatory, it opens test case files created in a unified format.
def test_case_open():
    tco = test_case_open
    os.chdir(userdir)
    tco.root_path = userdir
    tco.task_type = input('Task type to open (MEDRXF, AP, EXT, etc.): ')
    tco.task_number = input('Task number to open: ')
    tco.testcase_directory = tco.task_type + "-" + tco.task_number
    tco.txt_file = tco.testcase_directory + ".txt"
    tco.sql_file = tco.testcase_directory + ".sql"
    tco.complt_path \
        = str(tco.root_path + "\\" + tco.testcase_directory + "\\" + tco.txt_file)
    tco.sqlcomplt_path = str(tco.root_path + "\\" + tco.testcase_directory + "\\" + tco.sql_file)
    tco.path = Path(tco.root_path, tco.testcase_directory)
    try:
        os.chdir(tco.path)
    except FileNotFoundError:
        reroute = input('File not found, create? (y/n): ')
        if reroute == 'y':
            print(colorama.Fore.WHITE + "Rerouting to create file...")
            #Here is where it reroutes to create a dir, text file, and sql file that doesn't exist when a user
            #tries to open one.
            test_case_open_reoute()
        if reroute == 'n':
            print(colorama.Fore.WHITE + 'File not opened. Returning to main screen.')
            user_choice()
    else:
        open_function()


def open_function():
    tco = test_case_open
    print(colorama.Fore.WHITE + 'Found \'%s\'' % tco.complt_path)
    print(colorama.Fore.WHITE + 'Found \'%s\'' % tco.sqlcomplt_path)
    print(colorama.Fore.LIGHTWHITE_EX + '')
    try:
        open_file = input('Open text file (t), sql file (s), both (b) or neither (n): ')
        if open_file == 't':
            for tco.txt_file in Path(tco.path).rglob(str(tco.txt_file)):
                sp.Popen([initial.ntpdplus, tco.txt_file])
                print(colorama.Fore.WHITE + 'File opened. Returning to main screen.')
                user_choice()
        if open_file == 's':
            for tco.sql_file in Path(tco.path).rglob(str(tco.sql_file)):
                sp.Popen([initial.datagrip, tco.sql_file])
                print(colorama.Fore.WHITE + 'Sql file opened. Returning to main screen.')
                user_choice()
        if open_file == 'b':
            for tco.txt_file in Path(tco.path).rglob(str(tco.txt_file)):
                sp.Popen([initial.ntpdplus, tco.txt_file])
            for tco.sql_file in Path(tco.path).rglob(str(tco.sql_file)):
                sp.Popen([initial.datagrip, tco.sql_file])
                print(colorama.Fore.WHITE + 'Files opened. Returning to main screen.')
                user_choice()
        if open_file == 'n':
            print(colorama.Fore.WHITE + 'Files not opened. Returning to main screen.')
            user_choice()
    except FileNotFoundError:
        print(colorama.Fore.RED + 'File(s) not found, returning to main screen.')
        user_choice()
    else:
        user_choice()


#Here is where the open method calls during exception handling; A file that doesn't exist.
def test_case_open_reoute():
    tco = test_case_open
    path = tco.complt_path
    directory = tco.testcase_directory
    dir_path = userdir + '\\' + directory
    os.mkdir(dir_path)
    os.chdir(dir_path)
    print("Directory '% s' created" % directory)
    try:
        with open(r'C:\Users\%s\Desktop\TestCaseUtility\jiraformat.txt' % username) as f:
            with open(path, "w") as new_file:
                for line in f:
                    new_file.write(line)
                new_file.close()
    except FileNotFoundError:
            print(colorama.Fore.RED + 'Test case format configuration not found, please place jiraformat.txt within the '
                  'TestCaseUtility directory.')
            user_choice()
    else:
        open(tco.sqlcomplt_path, "x")
        print(colorama.Fore.WHITE + "Task file '% s.txt' created." % directory)
        print(colorama.Fore.WHITE + "SQL file '% s.sql' created." % directory)
        print(colorama.Fore.LIGHTWHITE_EX + '' )
        openfile = input('Open text file (t), sql file (s), both (b) or neither (n): ')
        if openfile == 't':
            print(colorama.Fore.WHITE + 'Task file created and opened. Returning to main screen.')
            sp.Popen([initial.ntpdplus, path])
            user_choice()
        if openfile == 's':
            print(colorama.Fore.WHITE + 'Task files created and opened. Returning to main screen.')
            sp.Popen([initial.datagrip, tco.sqlcomplt_path])
            user_choice()
        if openfile == 'b':
            print(colorama.Fore.WHITE + 'Test case files created and opened. Returning to main screen.')
            sp.Popen([initial.datagrip, tco.sqlcomplt_path])
            sp.Popen([initial.ntpdplus, path])
            user_choice()
        if openfile == 'n':
            print(colorama.Fore.WHITE + 'Test case files created but not opened. Returning to main screen.')
            user_choice()


if __name__ == '__main__':
    initial()
