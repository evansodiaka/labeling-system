from tkinter import *
import tkinter as tkk
from tabulate import tabulate 
import numpy as np
import pandas as pd

from barcode import EAN13
from barcode.writer import ImageWriter
from io import BytesIO
import string
import re
from docx import Document
from docx.shared import Cm, Pt

import os
import sys

class Table:
    
   
    def __init__(self,root):
    
        sTag="lv"
        mTag=100
        eTag="a5"
        quote= "{}-{}-{}".format(sTag,mTag,eTag)

        for i in range(12):
            for j in range(25):
                e = Entry(root, width=10,  relief='solid', font=('Arial',10,'bold'))
                e.grid( ipady=10,ipadx=10, row=j,column=i)
                #+e.insert(END,"text/")
                e.insert(END,quote)
          
             
        #         #e.pack(RIGHT,fill=Y)
        # text=str("bhg\nkkk")
        # list1 =list(text)
        # text =''.join(list1)
        # x3 = np.array((25,12),str)
        # i=0
       
       
        # tableDict1= [[quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote],
        #         [quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote]]
        # ##
        tableDict= ((quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote),
                (quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote,quote))
        
        tableD=[]
        for i in range (25):
            tableD.append([])

        limit=130
        
        for row in tableD:
            i=0
            while i < 12:
                if  mTag <limit:
                    quote = "{}-{}-{}".format(sTag,mTag,eTag)
                    row.append(quote)
                    mTag+=1
                    i+=1
                elif mTag == limit or mTag > limit:
                    quote="      "+""
                    row.append(quote)
                    i+=1
                
        # Make sure to pass the number as string
        number = '5901234123457'
  
        # Now, let's create an object of EAN13 class and 
        # pass the number with the ImageWriter() as the 
        # writer
        #rv = BytesIO()
        #EAN13(str(100000902922), writer=ImageWriter()).write(rv)
        my_code = EAN13(number, writer=ImageWriter())
        #print(my_code)
        # Our barcode is ready. Let's save it.
        my_code.save("new_code1")        

        #print(tableD)
        print(tabulate(tableD, tablefmt='fancy_grid'))

        #create new document
        document = Document()

        #set up font
        font= document.styles['Normal'].font
        font.name = 'Arial'
        font.size = Pt(8)

        

        #set up margins
        sections = document.sections
        for section in sections:
            section.top_margin= Cm(1.25)
            section.bottom_margin = Cm(1.25)
            section.left_margin = Cm(1.75)
            section.right_margin = Cm(1.75)


        table = document.add_table(rows =1, cols =12)
        table.style= 'LightGrid'
       

        for quote1,quote2,quote3,quote4,quote5,quote6,quote7,quote8,quote9,quote10,quote11,quote12 in tableD:
            row_cells = table.add_row().cells
            row_cells[0].text =quote1
            row_cells[1].text =quote2
            row_cells[2].text =quote3
            row_cells[3].text =quote4
            row_cells[4].text =quote5
            row_cells[5].text =quote6
            row_cells[6].text =quote7
            row_cells[7].text =quote8
            row_cells[8].text =quote9
            row_cells[9].text =quote10
            row_cells[10].text =quote11
            row_cells[11].text =quote12

        set_column_width(table.columns[9],Cm(1.75))
        # set_column_width(table.columns[3],Cm(80.75))
        
        document.save('demo.docx')

      
    
    #def export 


def set_column_width(column,width):
    for cell in column.cells:
        cell.width = width
        print("ddfgfd")

def closes_creen():
   root.destroy()

def create_grid():
    gridSize=12
    grid = [[0 for _ in range (gridSize) for _ in range(gridSize)]]
    print_grid(grid)

def print_grid(grid):
        for row in grid:
            print(row)     

      

root = tkk.Tk()
root.geometry("563x750")
root.title("Lankstar network sheets")
t=Table(root)
print(root.grid_size()) # tuple  columns , rows 


root.mainloop()