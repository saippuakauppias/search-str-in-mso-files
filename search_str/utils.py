import os


def fileslist(dir, filetypes=None):
    result = []
    dir = os.path.abspath(dir)
    for filename in os.listdir(dir):
        filepath = os.path.join(dir, filename)

        if os.path.isdir(filepath):
            result += fileslist(filepath, filetypes)
        elif os.path.isfile(filepath):
            result.append(filepath)

    if filetypes:
        result = filter(lambda x: x.endswith(filetypes), result)
    return result
