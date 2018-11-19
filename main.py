#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 19 18:07:02 2018

@author: harinair97
"""

#importing the required packages
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
#import cv2
import xlrd




def create_certi(name):
    #calling the required template
    img = Image.open("/home/harinair97/Desktop/automatic certificate generator/img.jpg")
    draw = ImageDraw.Draw(img)


    path="/home/harinair97/anaconda3/lib/python3.6/site-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans-Oblique.ttf"
    selectFont = ImageFont.truetype(path, size = 60)

    draw.text( (500,525),name,(0,0,0), font=selectFont)

    img.save( 'certificate\\'+name+'.pdf', "PDF", resolution=100.0)

    #cv2.imwrite("certi1.pdf",img)









if __name__ == "__main__":
   
    # Read data from an excel sheet from row 2
    xl_path="/home/harinair97/Desktop/automatic certificate generator/Gate pass list (19_10, 10am).xlsx"
    Book = xlrd.open_workbook(xl_path)
    WorkSheet = Book.sheet_by_name('Sheet1')
    
    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1
        
        name = WorkSheet.cell_value(row,0)
        
        
        # Make certificate and check if it was successful
        filename = create_certi(name)
        
        