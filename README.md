# py_parse_xlsx_for_image
Using pure python, extract an image from an excel file by its address.

Initially for a project which had to run in Jython (for Ignition Scada). Needed to get instructions, including an image, from an Excel xlsx workbook.

As Jython cannot use anything that's another language wrapped in python, it had to be pure python. To find the worksheet data [pylightxl](https://github.com/PydPiper/pylightxl) worked very well. But, I couldn't find anything for pulling the image based on where it was in the workbook. So, I dug around in the underlying xml and built it.

As it is now, it's a bit of a mess and not ready to be used elsewhere. But, it's a start.

Things currently start out in find_value_xlsx.py, where it finds the [Excel Picture Lookup](https://exceloffthegrid.com/automatically-change-picture) name and then the image location from there, using pylightxl. Now, having an address, goes through the interconnected web of xml and xml.rels files inside an xlsx to eventually get and save the image somewhere.
