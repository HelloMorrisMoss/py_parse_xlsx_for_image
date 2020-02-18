"""Ugly, WET code to get an image from an excel file given the sheet name and row/col address."""
# todo: cleanup; lots of development #print() lines left; plus it was changed to a class partway through and
# todo: then essential turned into a glorified function from there
# todo: function-alize things, I'm looking at you get target from .xml.rel file


import os
import zipfile as zf
from io import BytesIO
from xml.etree import cElementTree as et
from py_xl_image_extract.get_xml_namespace import getns as gns
from shutil import copyfileobj as imgsv


class GetWsXml:
    """Return the xml for a worksheet by its name."""
    def __init__(self, file_path, sheet_name, row, col):
        self.file_path = file_path
        self.find_sht = sheet_name
        self.zip_file = self.open_xlsx(self.file_path)
        self.row_number = row - 1  # for some reason we're picking up an extra 1 here but not for column
        self.col_number = col   # column seems fine
        # print('init zip_file', self.zip_file)

    def open_xlsx(self, fpath):
        """Return the zipfile object for the zip file."""
        zfl = zf.ZipFile(fpath, 'r')
        print('open xlsx, zfl', zfl)
        return zfl


    def findws(self): #, xl_file_path, sheet_name):
        """This is basically "main" if this were a module."""
        xl_file_path = self.file_path
        sheet_name = self.find_sht

        # get the 'rId#' for the sheet we're starting from
        ws_id = self.get_ws_id(xl_file_path, sheet_name)
        print('findws id:', ws_id)

        # find the sheet#.xml with that id
        ws_xml_path = self.get_ws_xml_path(xl_file_path, ws_id)
        print('findws xml path', ws_xml_path)

        # get the drawing xml for that sheet
        drawing_xml_path = self.get_ws_drawings_xml_path(self.zip_file, ws_xml_path + '.rels')
        print('drawing xml', drawing_xml_path)

        drawing_rid, new_drawing_xml = self.get_drawings_rid(self.zip_file, drawing_xml_path,
                                                             self.row_number, self.col_number)
        print('drawing rid', drawing_rid, new_drawing_xml)
        if drawing_rid and new_drawing_xml:
            print('drawing rid', drawing_rid)

            img_path = self.get_image_path(self.zip_file, new_drawing_xml, drawing_rid)
            print(img_path)

            image = self.get_image(self.zip_file, img_path)
            print(image)

            return_this = image
            return return_this
        return False

    def get_ws_id(self, fp, ws_name):
        """Get the rId# for the sheetname."""
        # open the excel file as a zip archive
        with zf.ZipFile(fp) as xlf:
            # inf <- these are the information objects for the files in the zip (xlsx) file
            for inf in xlf.infolist():
                xl_part_name = inf.filename
                # print(xl_part_name)
                # if the file is a drawing.xml
                if xl_part_name[-3:] == 'xml' and 'workbook' in xl_part_name:
                    # read the xml file as a string
                    xl_part_file = xlf.read(xl_part_name)

                    # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                    xml_obj = et.fromstring(xl_part_file)

                    # get the namespaces dictionary from the xml string
                    namespaces = gns(xl_part_file)
                    # print('namespaces', namespaces)

                    sheets_element = xml_obj.find('.//sheets', namespaces)#[0].findall('sheet')
                    # print('sheets_elelment', sheets_element)

                    sheets = sheets_element.findall('.//sheet', namespaces)
                    # print('sheets', sheets)

                    for sheet in sheets:
                        attribs = sheet.attrib
                        # print(sheet, attribs, attribs['name'])
                        rid_key = '{lb}{r}{rb}id'.format(r=namespaces['r'], lb=r'{', rb=r'}')
                        if attribs['name'] == ws_name:
                            # print('found it', attribs)
                            return attribs[rid_key]

    def get_ws_xml_path(self, fp, id_num):
        """Get the path for the xml file for the sheet."""
        # open the excel file as a zip archive
        with zf.ZipFile(fp) as xlf:
            # inf <- these are the information objects for the files in the zip (xlsx) file
            for inf in xlf.infolist():
                xl_part_name = inf.filename
                # print(xl_part_name)
                # if the file is a drawing.xml
                if 'workbook.xml.rels' in xl_part_name:
                    # read the xml file as a string
                    xl_part_file = xlf.read(xl_part_name)

                    # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                    xml_obj = et.fromstring(xl_part_file)

                    # get the namespaces dictionary from the xml string
                    namespaces = gns(xl_part_file)
                    # print(namespaces)

                    # rel_target = xml_obj.findall('Relationship Target', namespaces)
                    # # print('sheets_elelment', sheets_element)
                    for rel_target in xml_obj:
                        attribs = rel_target.attrib
                        if attribs['Id'] == id_num:
                            # print('found it', attribs, attribs['Id'])
                            return attribs['Target']

    def get_ws_drawings_xml_path(self, zfile, sh_xml_path):
        """Get the path for the drawings#.xml for the sheet."""
        # split up the path and remake it for the _rels folder & rel files
        split_path = os.path.split(sh_xml_path)
        # print('split path', split_path)
        new_path = split_path[0] + '/_rels/' + split_path[-1]
        # print('new path', new_path)

        for inf in zfile.infolist():
            xl_part_name = inf.filename
            if new_path in xl_part_name:
                # read the xml file as a string
                xl_part_file = zfile.read(xl_part_name)

                # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                xml_obj = et.fromstring(xl_part_file)

                # find the target and return it
                for rel_target in xml_obj:
                    # print('relt', rel_target)
                    attribs = rel_target.attrib
                    return attribs['Target']

    def get_drawings_rid(self, zfile, sh_xml_path, row_num, col_num):
        """Get the rId# for the image."""
        # split up the path and remake it for the _rels folder & rel files
        # split_path = os.path.split(sh_xml_path)
        # print('split path', split_path)
        new_path = sh_xml_path[2:]# + '/_rels/' + split_path[-1]
        # print('new path', new_path)

        for inf in zfile.infolist():
            xl_part_name = inf.filename
            if new_path in xl_part_name:
                # read the xml file as a string
                xl_part_file = zfile.read(xl_part_name)

                # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                xml_obj = et.fromstring(xl_part_file)

                # get the namespaces dictionary from the xml string
                namespaces = gns(xl_part_file)
                # print('namespaces', namespaces)

                alt_content = xml_obj.findall('.//xdr:twoCellAnchor', namespaces)
                print('alt content', alt_content)
                print('looking for sheet address row col:', row_num, col_num)
                print(type(row_num))

                for cont in alt_content:
                    from_addr = cont.findall('xdr:from', namespaces)
                    for addr in from_addr:
                        rnum = int(addr.find('xdr:row', namespaces).text)
                        cnum = int(addr.find('xdr:col', namespaces).text)
                        print('rnum and cnum', rnum, cnum, type(rnum))
                        if rnum == row_num and cnum == col_num:
                            print('found the pic element')
                            rid = cont.find(".//a:blip", namespaces).attrib['{lb}{ns}{rb}embed'.format(ns=namespaces['r'], lb=r'{', rb=r'}')]
                            print('image rid:', rid)
                            return rid, new_path
        return False, False

    def get_image_path(self, zfile, sh_xml_path, rid):
        """Get the path for the image file."""
        # split up the path and remake it for the _rels folder & rel files
        split_path = os.path.split(sh_xml_path)
        # print('split path', split_path)
        new_path = split_path[0] + '/_rels/' + split_path[-1]
        # print('new path', new_path)

        for inf in zfile.infolist():
            xl_part_name = inf.filename
            if new_path in xl_part_name:
                # read the xml file as a string
                xl_part_file = zfile.read(xl_part_name)

                # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                xml_obj = et.fromstring(xl_part_file)

                # find the target and return it
                for rel_target in xml_obj:
                    # print('relt', rel_target)
                    attribs = rel_target.attrib
                    if attribs['Id'] == rid:
                        print(attribs)
                        return attribs['Target']

    def get_image(self, zfile, img_path):
        """Get a bytes object version of the image."""
        # split up the path and remake it for the _rels folder & rel files
        # split_path = os.path.split(img_path)
        # print('split path', split_path)
        new_path = img_path[2:]
        print('new path', new_path)

        for inf in zfile.infolist():
            xl_part_name = inf.filename
            # print(xl_part_name)
            if new_path in xl_part_name:
                print('found the image', xl_part_name)
                # read the xml file as a string
                xl_part_file = zfile.read(xl_part_name)
                return BytesIO(xl_part_file)


def bymage(byte_obj, new_file):
    """Write a byte object of an image file as an image file."""
    with open(new_file, 'wb') as file:
        imgsv(byte_obj, file, length=131072)

def find_the_img(xlsx_filepath, ws_name, row_num, col_num, save_file):
    bymage_obj = GetWsXml(xlsx_filepath, ws_name, row_num, col_num).findws()
    if bymage_obj:
        bymage(bymage_obj, save_file)
    else:
        print("Couldn't find the image?.")


if __name__ == '__main__':
    print('testing')
    xl_filepath = r'Operator Lookup Special Instructions v2.1.xlsx'
    sheet_finder = GetWsXml(xl_filepath, 'Pic Lookup', 7, 1).findws()
    # sht = sheet_finder.findws(xl_filepath, 'Pic Lookup')
    print(sheet_finder)
    bymage(sheet_finder, 'test_image2.png')

    # with open('./test_image.jpeg', 'wb') as file:
    #     imgsv(sheet_finder, file, length=131072)
