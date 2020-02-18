import pylightxlnew as xl
from py_xl_image_extract.xlsx_reader import readPic

xl_file = r'20190514 pre-floor scale numbers.xlsx'


# ##### this is working, but may not give picture info
# read the file into an in memory database
db = xl.readxl(xl_file)


# print(db.ws_names) # available ws names
# sh1nm = db.ws_names[0]
#
# sh1 = db.ws('Sheet1')
# a1 = sh1.address("A1")  # the column letter must be in CAPITAL LETTERS

# print(sh1nm, sh1)

# for rw in sh1.rows:
#     print(rw)

# print(a1)
#
# print([readPic(xl_file, 1)]) # gives addresses of pictures in sheet, no info about them

wsheets = db.ws_names

for indx, sh_name in enumerate(wsheets):
    ws = db.ws(sh_name)
    print('file {f} index {i}'.format(f=xl_file,i=indx))
    for rw in ws.rows:
        print(rw)
    if indx != 0:
        img =[]
        try:
           img.append(readPic(xl_file, indx))
           print(img)
           # ws.address(img)
        except ValueError:
            pass
        # if img:
        img_add = img[0][0:2]
        print(img_add)
        img_obj = ws.address(img_add)
        print(img_obj)
# print(wsheets)