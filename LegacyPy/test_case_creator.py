import os

import user_choice as uc

#import test_case_creator as tcc

#import test_case_open as tco



#import signal#


def test_case_creator():

    TaskType = input(str('Task Type to Create?: '))
    TaskNumber = input('Task Number to Create?: ')
    directory = TaskType + "-" + TaskNumber
    parent_dir = r'C:\Users\Joshua Reid\Desktop\Tasks'

    path = os.path.join(parent_dir, directory)
    os.mkdir(path)
    print("Directory '% s' created" % directory)

    #Test case files are named and created using a corresponding file path
    name_of_file = directory
    file_path = path + "\\" + directory + ".txt"
    testheader = "*{color:#14892c}TESTING ENVIRONMENT:{color}*\n"

    newline = "\n"

    stepstaken = "*{color:#333333}Steps Taken:{color}*\n"

    Test1 = "*{color:#333333}Test 1:{color}*\n"

    Test2 = "*{color:#333333}Test 2:{color}*\n"

    Test3 = "*{color:#333333}Test 3:{color}*\n"
    new_file = open(file_path, "w")
    new_file.write(testheader)
    new_file.write(newline)
    new_file.write(stepstaken)
    new_file.write(newline)
    new_file.write(Test1)
    new_file.write(newline)
    new_file.write(Test2)
    new_file.write(newline)
    new_file.write(Test3)
    new_file.close()
    print("Task file '% s' created" % directory)

    #Newly created test files are opened in notepad++
    openfile = input('Open file?: ')
    #ntpdplus = r'C:\Program Files\Notepad++\notepad++.exe'
    ntpdplus = r'C:\Program Files (x86)\Notepad++\notepad++.exe'
    if openfile == 'y':
        import subprocess as sp

        sp.Popen([ntpdplus, file_path])
        sp.Popen.kill(sp.Popen([ntpdplus, file_path]))

    if openfile == 'n':
        uc.user_choice()

    else:
        uc.user_choice()


if __name__ == '__main__':
    test_case_creator()

# I would like this to be a somewhat complicated script. First it has a script that opens something like a home screen.
# Matt mentioned tk.enter. Worth looking into.

