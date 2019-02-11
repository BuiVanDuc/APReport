import os

current_path = os.path.normpath(os.path.join(os.path.realpath(__file__), '../../'))

print(current_path)