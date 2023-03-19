from bot.duplicates import check_for_duplicates
from bot.file_functions import brute_remove_all_empty_directories, brute_rename_all_paths
from config import processPath
import os
from bot.format_ppt import brute_format_all_powerpoints


def run():
    root = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(root, processPath)
    print("Running on: {}".format(path))
    #brute_remove_all_empty_directories(path)
    #brute_rename_all_paths(path)
    brute_format_all_powerpoints(path)
    #check_for_duplicates(path, '{}/duplicates.txt'.format(path))
    print("Finished")

run()