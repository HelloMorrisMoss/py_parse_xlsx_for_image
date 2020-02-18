"""Searching through xml for images anchor addresses."""

import zipfile as zip
# from xml.etree import cElementTree as ElementTree
from xml.etree import cElementTree as et
from py_xl_image_extract.get_xml_namespace import getns


# import lxml # requires _winreg
class xl_img_finder():
    def __init__(self, file_path=''):
        # self.xl_file_path = r'20190514 pre-floor scale numbers.xlsx'
        self.xl_file_path = r'Operator Lookup Special Instructions v2.1.xlsx'
        self.folder_dict = {}

    # moved to get_xml_namespace module
    # def ucd(self, input):
    #     """Convert a string or unicode object, to a unicode version."""
    #     print(type(input), input)
    #     try:
    #         input = input.decode('utf-8')
    #     except AttributeError:
    #         pass
    #     return input

    def getns_xml(self, xml):
        """Get a dictionary of namespaces from string or unicode."""
        # MOVED THIS TO ITS OWN MODULE get_xml_namespace
        # Thanks to Davide Brunato on stackoverflow for his answer on this bit.
        # https://stackoverflow.com/a/37409050/10941169
        # xml = self.ucd(xml)
        # namespaces = dict([
        #     node for _, node in et.iterparse(
        #         StringIO(self.ucd(xml)), events=['start-ns']
        #         # StringIO(str(xml), events=['start-ns']
        #     )
        # ])
        namespaces = getns(xml)
        return namespaces

    def get_xml_file(self, file_name):
        """Return an xml file from an excel file."""
        print('not implemented yet')

    def open(self, xlsx_file_path):
        """Unzip and return the files from inside an excel file."""
        print('not implemented yet')
        # class xl_folder(folder):
        #     def __init__(self):
        #         self._folder_dict = {}
        #
        #     def add_folder(self, new_folder, folder_name):
        #         self._folder_dict[folder_name] = new_folder
        #
        # xl = xl_folder

        xml_dic = {}
        image_dic = {}

        with zip.ZipFile(xlsx_file_path) as xlf:
            # inf <- these are the information objects for the files in the zip (xlsx) file
            for inf in xlf.infolist():
                xl_part_name = inf.filename
                # if the file is a drawing.xml
                if xl_part_name[-3:] == 'xml' and 'drawing' in xl_part_name:
                    # open it
                    with xlf.open(xl_part_name) as xmlf:
                        # read the xml file as a string
                        xl_part_file = xlf.read(xl_part_name)

                        # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                        xml_obj = et.fromstring(xl_part_file)

                        # get the namespaces dictionary from the xml string
                        namespaces = self.getns_xml(xl_part_file)

    def img_dict(self, xl_file_path):
        """Generate a dictionary of images with their ids' and locations."""
        # xl_file_path = r'20190514 pre-floor scale numbers.xlsx'

        # empty dictionary to populate
        lst_dcts = {}

        # open the excel file as a zip archive
        with zip.ZipFile(xl_file_path) as xlf:
            # inf <- these are the information objects for the files in the zip (xlsx) file
            for inf in xlf.infolist():
                xl_part_name = inf.filename

                # if the file is a drawing.xml
                if xl_part_name[-3:] == 'xml' and 'drawing' in xl_part_name:

                    # open it
                    # with xlf.open(xl_part_name) as xmlf:
                    if True:
                        # read the xml file as a string
                        xl_part_file = xlf.read(xl_part_name)

                        # read the xml file into an xml object as a string, fromstring returns the ROOT ELEMENT
                        xml_obj = et.fromstring(xl_part_file)

                        # get the namespaces dictionary from the xml string
                        namespaces = self.getns_xml(xl_part_file)

                        # .// before the element name will make findall search sub elements
                        # etree needs to be provided the namespaces dictionary

                        # find all cell anchor elements, these contain
                        anchors = xml_obj.findall('.//xdr:twoCellAnchor', namespaces)

                        for anchor in anchors:
                            col = anchor.find('.//xdr:col', namespaces).text    # finding col number
                            row = anchor.find('.//xdr:row', namespaces).text    # finding row number
                            pic_id = anchor.find('.//xdr:cNvPr', namespaces).attrib['id']    # finding pic id
                            lst_dcts[pic_id] = [xl_part_name, row, col]

        return lst_dcts

    def test(self):
        return self.img_dict(self.xl_file_path)


fndr = xl_img_finder()
tdc = fndr.test()
print('img id cell dict', tdc)
