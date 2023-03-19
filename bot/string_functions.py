def match_substring(name, substringList):
    nameLowercase = name.lower()
    for substring in substringList:
        if (substring in nameLowercase):
            return True
    return False



