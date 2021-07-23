import os


def remove_images():
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
    files = [f for f in files if '.png' in f]

    for f in files:
        os.remove(f)


def get_file_name(file, name_list):
    for name in name_list:
        if file in name:
            return name


