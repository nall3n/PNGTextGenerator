# Os imports
import os
from os import defpath, listdir, path
from os.path import isfile, join
import tkinter

# Excel Reading import
import xlrd

# GUI Tkinter imports
from tkinter import * 
from tkinter import filedialog
from tkinter import colorchooser

#Img prossesing imports
from typing import Counter
from PIL import Image, ImageDraw, ImageFont 

# For wraping the text arround
import textwrap

import configparser


class Png_Generator():

    def __init__(self, iImg_width = 2048, iImg_height = 1366, iFont_size = 200, sOutput_Path = '', sFont_path = '',
                sFont = '',
                iY_offset = 180
                ) -> None:
        super().__init__()

        self.img_width = int(iImg_width)
        self.img_height = int(iImg_height)
        self.font_size = int(iFont_size)
        self.y_offset = int(iY_offset)
        self.default_font = sFont

        self.tFg_color = (255,255,255)
        self.tBg_color = (73, 109, 137)
        
        


        self.output_path = sOutput_Path
        self.font_path = sFont_path

    # PNG GENERATOR!!!!
    def create_png(
        self,
        text = '', 
        background_color = (73, 109, 137), 
        font_color='white',
        font = 'Helvetica-Bold.ttf',
        ):
        """
        Generates a png 2048x1366 with a text centerd in the middel 
        Takes in text, background color, font color, font and font size
        and outputs a .png file
        """
        text_offset = 180

        # If font not in font dir use standard font Helvetica
        if font not in os.listdir('Fonts'):
            font = self.default_font
        # Font and size
        font_path = os.path.join('Fonts', font)

        # Text and colors
        img_text = text
        background_color = background_color
        text_color = font_color
        #===================================================================

        # Create canvas with color
        img = Image.new('RGB', (self.img_width, self.img_height) , color = background_color)
        d = ImageDraw.Draw(img)
        # Set font and size
        font = ImageFont.truetype(font_path, self.font_size)


        # Split up img text into lines if more than set amount of caracters
        lines = textwrap.wrap(img_text, width=20)


        # Get size of text with font 
        text_size = font.getsize( text=img_text)

        # Set y start pos for text (text_size[1] / 2)
        y_text = self.img_height/2 - self.y_offset


        # Check if lines are to big for img res
        for line in lines:
            # Defines som vars for later use
            last_line = None
            last_width = 0
            
            # Get width and heigt of line
            width, height = font.getsize(line)
            
            # If last line is set get the width and height of it 
            if last_line:
                last_width, last_height = font.getsize(last_line)

            # Set last line last
            last_line = line

            # If last line width is greater than current line 
            # Dont scale anny other liens beacouse then the big lien will be to big
            if last_width < width:
                while width > self.img_width: # Run until line width is less than img width    
                    self.font_size = self.font_size - 50
                    font = ImageFont.truetype(font_path, self.font_size)
                    width, height = font.getsize(text=line)
                
                y_text = self.img_height/2 - self.y_offset # Reset the img Y pos (text_size[1] / 2)

        if len(lines) > 2:
            y_text = y_text - font.getsize(lines[0])[1] - 50
            
        # Adds all lines to img
        for line in lines:
            width, height = font.getsize(line)
            
            d.text((self.img_width/2 - (width / 2), y_text), line, fill=text_color, font=font )
            y_text += height + 50

        # Saves image as text name
        img.save(os.path.join('output', img_text + '_app.png'))


    # Standard colors

    # Tk funktions 
    #=================================================================================
    def bg_pick_color(self):
        color = colorchooser.askcolor(title ="Choose color")
        
        self.tBg_color = color[1]
        self.background_color.config(bg=color[1])
        self.background_color.config(text=color[1])
        self.preview_lbl.config(bg=color[1])

    def font_pick_color(self,):
        color = colorchooser.askcolor(title ="Choose color")

        self.tFg_color = color[1]
        self.font_color.config(bg=color[1])
        self.font_color.config(text=color[1])
        self.preview_lbl.config(fg=color[1])
    
    # function to gater inputs and then generate the PNG
    # button is clicked
    def generate_checker(self):
        res = self.txt.get()
        
        if not self.font_listbox.curselection():
            font = self.default_font
        else:
            font = self.font_listbox.get(self.font_listbox.curselection())


        self.create_png(    
            text=res, 
            background_color=self.tBg_color, 
            font_color=self.tFg_color, 
            font=font
        )
        self.lbl.configure(text=res)

    # Read in file excel or csv
    # Take input from file and if color/font is not set in file take settings from program
    def generate_from_file(self):
        file_path = filedialog.askopenfilename()

        wb = xlrd.open_workbook(file_path)
        import_file = wb.sheet_by_index(0)


        temp_text = ''
        temp_bg = ''
        temp_fg = ''
        temp_font = ''

        # Create a image from each row in file
        for i in range(import_file.nrows):
            temp_text = import_file.cell_value(i, 0)
            temp_bg = import_file.cell_value(i, 1)
            temp_fg = import_file.cell_value(i, 2)
            temp_font = import_file.cell_value(i, 3)
            
            #test.configure(text=temp_bg)
        

            if temp_text == None:
                return

            if temp_bg == None or temp_bg == '':
                temp_bg = self.tBg_color
            if temp_fg == None or temp_fg == '':
                temp_fg = self.tFg_color
            if temp_font == None or temp_font == '':
                if not self.font_listbox.curselection():
                    temp_font = self.default_font
                else: 
                    temp_font = self.font_listbox.get(self.font_listbox.curselection())

            self.create_png(
                
                text=temp_text,
                background_color=temp_bg,
                font_color=temp_fg,
                font=temp_font
            )

    #=================================================================================

    def tk_setup(self):

        # Tk layout and contents
        #=================================================================================
        tk_font_size = 12

        root = Tk()
        root.title('PNG GENERATOR!!!!')
        root.geometry('600x500')

        #==========================================
        # adding a label to the root window
        self.lbl = Label(
            root, 
            text = "Enter Name", 
            font=('Times', tk_font_size)
        )
        self.lbl.grid(column=1, row=0, pady=15,)
        #========================================== 
        # adding Entry Field
        self.txt = Entry(
            root, 
            width=20, 
            font=('Times', tk_font_size)
        )
        self.txt.grid(column =2, row =0,pady=15,)
        #==========================================
        self.background_color = Label(
            root,
            text='#496d89',
            font=('Times', tk_font_size),
            relief = SOLID,
            bg='#496d89',
            padx=10, 
            pady=10,
            width=20
        )
        self.background_color.grid(column=1,row=1)
        bg_button = Button(
            root, 
            text = "Choose Background Color",
            command = self.bg_pick_color,
            padx=10,
            pady=10,
            font=('Times', tk_font_size),
            bg='#4a7a8c'
            )
        bg_button.grid(column=2, row=1)
        #==========================================
        self.font_color = Label(
            root,
            text='#FFFFFF',
            font=('Times', tk_font_size),
            relief = SOLID,
            bg='#FFFFFF',
            padx=10, 
            pady=10,
            width=20
        )
        self.font_color.grid(column=1,row=2)
        font_button = Button(
            root, 
            text = "Choose Font Color",
            command = self.font_pick_color,
            padx=10,
            pady=10,
            font=('Times', tk_font_size),
            bg='#4a7a8c'
            )
        font_button.grid(column=2, row=2)
        #==========================================
        self.font_listbox = Listbox(root,
            selectmode=SINGLE,
        )
        font_list = os.listdir('Fonts')
        index = 1
        for font in font_list:
            self.font_listbox.insert(index, font)
            index += 1
        self.font_listbox.grid(column=1, row=3)
        listbox_lbl = Label(root, text = "Leave blank to use Helvetica-Bold", font=('Times', tk_font_size))
        listbox_lbl.grid(column=2, row=3)
        #==========================================
        # Buton to generate png
        btn = Button(
            root, 
            text = "GENERATE", 
            fg = "black",
            font=('Times', tk_font_size),
            padx=10,
            pady=10,
            command=self.generate_checker
        )
        btn.grid(column=1, row=5)

        import_btn = Button(
            root,
            text = 'Load from file',
            fg = 'black',
            font=('Times', tk_font_size),
            padx=10,
            pady=10,
            command=self.generate_from_file
        )
        import_btn.grid(column=2, row=5)

        test = Label(
            fg='black',
            text='If no color or font is entered in the import file. The settings set here will be used'
        )
        test.grid(
            column=1, 
            row=6, 
            columnspan = 2,    
            padx=10,
            pady=15,
            sticky = tkinter.W+tkinter.E
        )

        self.preview_lbl = Label(
            root,
            text='COLOR Preview',
            font=('Times', 20),
            fg='#FFFFFF',
            bg='#496d89'
        )
        self.preview_lbl.grid(
            column = 1, 
            row = 10, 
            columnspan = 2,
            rowspan= 3,
            pady=10, 
            ipady=10,
            sticky = tkinter.W+tkinter.E,
        )

        # Columns to center content
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(3, weight=1)
        
        # Execute Tkinter
        root.mainloop()


#====================================
# Start all the stuff

if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('config.ini')

    default = config['DEFAULT']
    path = config['PATHS']

    png = Png_Generator(
        default['width'], 
        default['height'], 
        default['font_size'],
        path['output'],
        path['fonts'],
        default['font'],
        default['y_offset']
        )

    png.tk_setup()