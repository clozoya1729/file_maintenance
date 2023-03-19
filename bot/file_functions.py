import os
import subprocess

from bot.string_functions import match_substring
from config import ignoredDirectories, ignoredSubstrings, searchedExtensions


def remove_empty_directory(directory):
    try:
        os.rmdir(directory)
        print("Removed: {}".format(directory))
        return True
    except OSError as e:
        return False


def brute_remove_all_empty_directories(path):
    run = True
    while (run):
        for root, directories, files in os.walk(path):
            run = remove_empty_directory(root)
            if run:
                break
    print("Removed all empty directories.")


def brute_rename_all_paths(path):
    for root, directories, files in os.walk(path):
        try:
            ignoreDirectory = match_substring(root, ignoredDirectories)
            if (ignoreDirectory == False):
                update_directories(root, directories)
                for filename in files:
                    ignoreFile = match_substring(filename, ignoredSubstrings) \
                                 and not match_substring(filename, searchedExtensions)
                    if (ignoreFile == False):
                        update_filename(root, filename)
        except Exception as e:
            pass
    print("Renamed all paths.")


def cleanup_filename(filename):
    newFilename = ' '.join(filename.split('-'))
    newFilename = ' '.join(newFilename.split())
    newFilename = newFilename.lower()
    newFilename = newFilename.replace(' ', '_')
    newFilename = newFilename.replace('-', '_')
    return newFilename


def update_filename(root, filename):
    newFilename = cleanup_filename(filename)
    oldPath = os.path.join(root, filename)
    newPath = os.path.join(root, newFilename)
    os.rename(oldPath, newPath)


def update_directories(root, directories):
    for directory in directories:
        newName = cleanup_filename(directory)
        oldPath = os.path.join(root, directory)
        newPath = os.path.join(root, newName)
        os.rename(oldPath, newPath)


def show_in_explorer(path):
    subprocess.Popen(r'explorer /select,"{}"'.format(path))
