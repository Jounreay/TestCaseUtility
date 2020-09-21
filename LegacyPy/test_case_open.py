import os

import test_case_creator as tcc

import user_choice as uc

import subprocess as sp
import sys


def test_case_open():

    #root = r'C:\Users\jreid\Desktop\Tasks'
    root = r'C:\Users\Joshua Reid\Desktop\Tasks'
    task_type = input('Task type to open?: ')
    task_number = input('Task number to open?: ')
    directory = task_type + "-" + task_number
    txt_file = directory + ".txt"
    #ntpdplus = r'C:\Program Files\Notepad++\notepad++.exe'
    ntpdplus = r'C:\Program Files (x86)\Notepad++\notepad++.exe'
    path = os.path.join(root, directory)
    open_file = input('Open this file? ')

    if open_file == 'y':

        with os.scandir(path) as cp:

            try:
                if txt_file in cp:
                    sp.Popen([ntpdplus, txt_file])
            except FileNotFoundError:
                reroute_tcc = input('File does not exist, create?:  ')
                if reroute_tcc == 'y':
                    tcc.test_case_creator()

            finally:
                if txt_file in cp:
                    sp.Popen([ntpdplus, txt_file])

    else:

        uc.user_choice()


if __name__ == '__main__':
    test_case_open()

