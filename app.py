import os
import generic_comparer

try:
    generic_comparer.GenericComparer.startComparing()
    print("Thanks for using pdf highlighter program.")
    os.system('pause')
except(IOError, ValueError, EOFError, PermissionError) as e:
    print('oops! error occurred.')
    print(e)
    os.system('pause')
