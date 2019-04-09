#!/usr/bin/env python

import googlemaps
import xlrd     # https://www.geeksforgeeks.org/reading-excel-file-using-python/
import xlwt     # https://www.programering.com/a/MTMyQDNwATU.html

gmaps = googlemaps.Client(key = 'INSERT A VALID KEY HERE')
fileLoc = ("/path/to/directory/file.xlsx")
wb = xlrd.open_workbook(fileLoc)
wb2 = xlwt.Workbook(encoding = 'ascii')
wbSheet = wb2.add_sheet('My Worksheet')
sheet = wb.sheet_by_index(0)

for i in range(1,250):
    try:
        value = (sheet.cell_value(i,2) + ' ' + sheet.cell_value(i,3) + ' ' + sheet.cell_value(i,4) + ' ' + sheet.cell_value(i,5) + ' ' + sheet.cell_value(i,6))
        geocode_result = gmaps.geocode(value)
        print(value)
        wbSheet.write(i,1,geocode_result[0]["geometry"]["location"]["lat"])
        wbSheet.write(i,2,geocode_result[0]["geometry"]["location"]["lng"])
    except:
        v = 'invalid address'
        print(v)
        wbSheet.write(i,1,v)
        wbSheet.write(i,2,v)

wb2.save("/path/to/directory/new_file.xlsx")




