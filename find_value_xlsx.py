import pylightxlnew as xl
from py_xl_image_extract.xml_img import xl_img_finder as imf
# from pylightxlnew import database as db
from py_xl_image_extract.workbook_for_which_sheet_xml import find_the_img




# import xlrd # _winreg


def find_value(db, value_to_find, sheet_name=''):
    """Return the first address in the xlsx file where the value matches."""

    # create in memory database from file
    # db = xl.readxl(file_path)

    # list of worksheet names
    if sheet_name == '':
        # if no sheet provided, check them all
        wsheets = db.ws_names
    elif sheet_name not in db.ws_names:
        # if the sheet name isn't present, we'll look in all the sheets
        wsheets = db.ws_names
    else:
        # if there is a sheet provided, put it in a list of only itself for below
        wsheets = [sheet_name]

    # print(wsheets, type(wsheets[0]), db.ws_names)
    # print(sheet_name==db.ws_names[1])

    # go through the sheets
    for sh_name in wsheets:

        # print('sheet: ', sh_name)

        # get this worksheet from the database
        ws = db.ws(sh_name)

        # row number (base 0) and the row from the worksheet
        for rn, rw in enumerate(ws.rows):
            # print(rw)

            # column number (base 0) and the column from the row, so the cell
            for cn, clm in enumerate(rw):
                # print('address: ', rn, cn, 'cel value:', clm)

                # if the value matches, return the value
                if clm == value_to_find:
                    # since excel is index base 1 and python base 0, add 1 to the row and column values
                    ret_row = rn
                    ret_col = cn + 1
                    return sh_name, ret_row, ret_col

    # didn't find it, return False
    return False, False, False

def get_tabrow(db, tabcode, sheet_name = ''):
    """Return the first address in the xlsx file where the value matches."""

    # get the first row headers
    headers = db.ws(sheet_name).row(1)
    headers[0] = 'tabcode'

    trow = db.ws(sheet_name).keyrow(key=tabcode)

    # print(trow)
    row_dict = dict(zip(headers, trow))
    return row_dict


if __name__ == '__main__':
    print('find_value_xlsx main module, running test')
    # xl_file = r'20190514 pre-floor scale numbers.xlsx'
    xl_file = r'Operator Lookup Special Instructions v2.1.xlsx'
    # print(find_value(xl_file, 43595.5947569444, 'Sheet1'))
    # print(find_value(xl_file, 43595.59475694445, 'Sheet2'))
    # print(find_value(xl_file, 'T30324', 'Tabcode Lookup'))
    # 'T30324', 'Tabcode Lookup'
    # create in memory database from file
    db = xl.readxl(xl_file)
    # print('db', db)

    # gett the values from the sheet, most of these will eventually be used on screen
    # tcode = 'T11118'    # pic01
    tcode = 'T30313'    # pic02
    # tcode = 'T30324'    # pic03
    rd = get_tabrow(db, tcode, 'Tabcode Lookup')

    # todo: this has stuff which needs to show on the screen
    print('dictionary of items from instruction sheet', rd)
    pic_nm = rd['Picture Reference']

    # todo: find pic location by the name in the dict, then match with get picture
    # find the row with the correct image
    # pic_row = db.ws('Pic Lookup').keyrow(key=pic_nm)
    # print('picture name and row', pic_nm, pic_row)

    # find the location of the image
    pctrw = find_value(db, pic_nm, 'Pic Lookup')
    print('picture row on lookup sheet', pic_nm, pctrw)

    # my image finding class
    fnd = imf()
    # dictionary of images, keys = ids and values = list: [drawing#.xml, row#, col#]
    img_dct = fnd.img_dict(xl_file)
    print('img dict', img_dct)

    # # create a list of sheets with 'drawings'; I don't remember why I wanted this?
    # img_sht_names = set([x[0] for x in img_dct.values()])
    # print('"list" of sheets with images', img_sht_names)
    rw = pctrw[1]
    cl = pctrw[2]
    print('row col', rw, cl)
    find_the_img(xl_file, 'Pic Lookup', rw, cl, 'test_image3.png')

