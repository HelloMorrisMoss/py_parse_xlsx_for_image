"""WIP: to convert a zip file into a dictionary of directories as dictionaries of files."""

import os
import pprint
import zipfile as zip



pp = pprint.PrettyPrinter()


def zlistdir(zfile):
    """eturn a list of directories from zip file."""
    return list(filter(lambda f: f.startswith("subdir"), zfile.namelist()))


def zip_listdir(zip_file, target_dir):

    file_names = zip_file.namelist()

    if not target_dir.endswith("/"):
        target_dir += "/"

    if target_dir == "/":
        target_dir = ""

    result = [file_name
              for file_name in file_names
              if file_name.startswith(target_dir) and
              "/" not in file_name[len(target_dir):]
             ]

    return result

def zdir(zipfile):
    """Return a list of directories at the top level of a zip file."""
    z = zipfile
    # remove the filename
    dirs = list(set([os.path.dirname(x) for x in z.namelist()]))
    # print('zdir:dirs', dirs)

    # unique set of directories
    topdirs = set([os.path.split(x) for x in dirs])
    # print('topdirs', topdirs)
    cleaned_dir = list(filter(lambda x: x[0] == '', topdirs))
    root_dirs = [x[1] for x in cleaned_dir]
    # re_cleaned = list(filter(lambda x: x[0] != '', root_dirs))
    return root_dirs



def is_dir(zipinfo):
    # is_dir = lambda zinfo: zipinfo.filename.endswith('/')
    return zipinfo.filename.endswith('/')

#  with zip.ZipFile(xl_file_path) as xlf:


def zip_to_dict(zip_file, path, d):
    # great little recursive dir to dict:
    # https://stackoverflow.com/a/46415935/10941169
    split_path = path.split('/')
    name = split_path[len(split_path)]

    if path.endswith('/'):
        if name not in d['dirs']:
            d['dirs'][name] = {'dirs':{},'files':[]}
        for x in os.listdir(path):
            zip_to_dict(os.path.join(path, x), d['dirs'][name])
    else:
        d['files'].append(name)
    return d

# todo: try starting this from the original again (below), see if os.path stuff can work on some parts
# probably need a replacement for listdir? but join, isdir, basename etc may work even on zip file paths
# zdir is giving the top level directories, can probably build from there, pulling apart the filenames() list
# for sub directory paths etc

# def path_to_dict(path, d):
#
#     name = os.path.basename(path)
#
#     if os.path.isdir(path):
#         if name not in d['dirs']:
#             d['dirs'][name] = {'dirs':{},'files':[]}
#         for x in os.listdir(path):
#             path_to_dict(os.path.join(path,x), d['dirs'][name])
#     else:
#         d['files'].append(name)
#     return d
#
#
# mydict = path_to_dict('.', d = {'dirs':{},'files':[]})
#
# pp.pprint(mydict)

# # the original for reference

# import os
# import pprint
#
# pp = pprint.PrettyPrinter()
#
# def path_to_dict(path, d):
#
#     name = os.path.basename(path)
#
#     if os.path.isdir(path):
#         if name not in d['dirs']:
#             d['dirs'][name] = {'dirs':{},'files':[]}
#         for x in os.listdir(path):
#             path_to_dict(os.path.join(path,x), d['dirs'][name])
#     else:
#         d['files'].append(name)
#     return d
#
#
# mydict = path_to_dict('.', d = {'dirs':{},'files':[]})
#
# pp.pprint(mydict)


# mydict = zip_to_dict('.', d = {'dirs': {}, 'files': []})

# pp.pprint(mydict)

xl_file_path = r'Operator Lookup Special Instructions v2.1.xlsx'

# open the excel file as a zip archive
with zip.ZipFile(xl_file_path) as xlf:
    # inf <- these are the information objects for the files in the zip (xlsx) file
    # for inf in xlf.infolist():
    #     xl_part_name = inf.filename
    #     print(inf, is_dir(inf))
    # xlf.

    # print(zlistdir.namelist())
    # print(zip_listdir(xlf, '.'))
    # print('zlistdir', zlistdir(xlf))
    # print(zdir(xlf))

    pass

    # names = xlf.namelist()
    #
    # # adding zip_file to the beginning so that the root directory is obvious
    # prepend = ['zip_file/' + name for name in names]
    # dirs = [os.path.dirname(x) + '/' for x in prepend]
    # print(names)
    # print(dirs)
    #
    # xls = set(filter(lambda x: x.endswith('/'), dirs))
    # print('xls:', xls)
    #
    # split_paths = [os.path.split(path) for path in xls]
    # print(split_paths)
    #
    # # tldir = zdir(xlf)
    # # print('tldr', tldir)
    # # for dr in tldir:
    # #     zip_listdir(xlf, dr)
    # # # print(zip_listdir(xlf, 'xl'))
    # # pp.pprint(xlf.namelist())