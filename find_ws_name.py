from zipfile import ZipFile as zip
from xml.dom import minidom as md


# import lxml # requires _winreg


def find_ws_name(file_path, id_num):
    """Return the worksheet name with the id number."""
    # open the excel file as a zip archive
    with zip(file_path) as xlf:
        #   TODO: PUT ERROR HANDLING HERE FOR missing file
        # inf <- these (below) are info objects of files in the zip (xlsx) file
        for inf in xlf.infolist():

            # the file name in the zip
            xl_part_name = inf.filename

            # if it's an xml file
            if xl_part_name[-3:] == 'xml' and 'workbook' in xl_part_name:
                print(xl_part_name)
                #   TODO: PUT ERROR HANDLING HERE FOR BAD XML FILES, JUST IN CASE
                # open the xml file
                with xlf.open(xl_part_name) as xmlf:
                    # parse that file
                    xml_obj = md.parse(xmlf)

                    # find the sheet tags, put them in a list
                    shets = xml_obj.getElementsByTagName("sheet")

                    # go through that list
                    for sht in shets:
                        # split the text version of the result, there's also a bunch of clean up going on here
                        txt = sht.toxml().encode('ascii', 'ignore').replace("\"", "").replace("/>", "").split()
                        # print('text', txt)

                        # the sheet id number should be here
                        sh_num = int(txt[3].split("=")[1])

                        # if it's the one we're looking for, we're done!
                        if sh_num == id_num:
                            return txt[1].split("=")[1]

                    # if somehow we didn't find it, return false
                    return False


if __name__ == '__main__':
    print('main module, run test')
    xl_file = r'20190514 pre-floor scale numbers.xlsx'
    print('result: ', find_ws_name(xl_file, 1))
    # print(find_ws_name(xl_file, 43595.59475694445, 'Sheet2'))
