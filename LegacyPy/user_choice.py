import test_case_creator as tcc
import test_case_open as tco


def user_choice():

    method_choice = input(str('Create task? (c) or Open Task? (o):  '))

    if method_choice == 'c':
        tcc.test_case_creator()

    if method_choice == 'o':
        tco.test_case_open()

    else:
        import sys

        sys.exit(True)


if __name__ == '__main__':
    user_choice()