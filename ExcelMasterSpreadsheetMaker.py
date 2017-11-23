# -*- coding: utf-8 -*-
"""
Created on Thu Nov 10 10:16:54 2016

@author: pawel.cwiek
"""
import tkinter.tix as tk
import tkinter.ttk as ttk
import tkinter.filedialog
import make_master_excel_file

import logging
#logging.basicConfig(level=logging.INFO)
import os
log_path = os.path.join(os.getcwd(),'log.log')
print(log_path)
logg_config = dict([('level',logging.WARNING)])
logg_config.update(dict([('filename',log_path),('filemode','w')]))
logging.basicConfig(**logg_config)


def find_files(dir_path,extension):
    # returns list of filepaths including only xml files (no subfolders)
    import os

    files = [fn for fn in next(os.walk(dir_path))[2] if fn.count(extension)]

    return [os.path.join(dir_path,fn) for fn in files]


class PixelLabel(ttk.Frame):
    def __init__(self,master, w, h=20, *args, **kwargs):
        '''
        creates label inside frame,
        then frame is set NOT to adjust to child(label) size
        and the label keeps extending inside frame to fill it all,
        whatever long text inside it is
        '''
        ttk.Frame.__init__(self, master, width=w, height=h)
        self.grid_columnconfigure(0, weight=1)
        self.grid_propagate(False) # don't shrink
        self.label = ttk.Label(self, *args, **kwargs)
        self.label.grid(sticky='nswe')

    def resize(self,parent,*other_childs):
        '''
        resizes label to take rest of the width from parent
        that other childs are not using
        '''
        parent.update()
        new_width = parent.winfo_width()

        for widget in other_childs:
            widget.update()
            new_width -= widget.winfo_width()

        self.configure(width = new_width)

class MyGui(object):
    def __init__(self,root):
        self.errors = None
        self.source_files = None
        
        self.main_frame = ttk.Frame(root)
        self.main_frame.grid()

        intro_txt = 'This little app merges data from multiple MS Excel files into one master spreadsheet.'
        self.introlbl = ttk.Label(self.main_frame, justify="left", anchor="n", padding=(10, 2, 10, 6), text=intro_txt)
        self.introlbl.grid(row=0)

        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.grid(row=1)

        self.controls_frame = ttk.Frame(self.top_frame)
        self.controls_frame.grid(row=1,column=0,sticky='e')

        self.dir_path = tk.StringVar()
        self.folder_lbl = ttk.Label(self.controls_frame, justify="left", anchor="w", text='Source files path:')
        self.folder_lbl.grid(row=1,column=0,sticky='e')
        self.path_entry = ttk.Entry(self.controls_frame, textvariable=self.dir_path,width=50)
        self.path_entry.grid(row=1,column=1,sticky='we', padx=(5, 5))
        self.browse_btn = ttk.Button(self.controls_frame, command = self.browse_dir, text = 'Browse')
        self.browse_btn.grid(row=1,column=2)

        self.masterfile_path = tk.StringVar()
        self.folder_lbl = ttk.Label(self.controls_frame, justify="left", anchor="w", text='Master spreadsheet location:')
        self.folder_lbl.grid(row=2,column=0,sticky='e')
        self.path_entry = ttk.Entry(self.controls_frame, textvariable=self.masterfile_path,width=50)
        self.path_entry.grid(row=2,column=1,sticky='we', padx=(5, 5))
        self.browse_btn = ttk.Button(self.controls_frame, command = self.browse_file, text = 'Browse')
        self.browse_btn.grid(row=2,column=2)

        self.maxrows_lbl = ttk.Label(self.controls_frame, justify="left", anchor="e", text='Max row number:')
        self.maxrows_spiner = tk.Spinbox(self.controls_frame, from_=1, to=1000, increment=10,width=5)
        self.maxrows_lbl.grid(row=3,column=0,sticky='e')
        self.maxrows_spiner.grid(row=3,column=1,sticky='w', padx=(5, 5))

        self.maxcols_lbl = ttk.Label(self.controls_frame, justify="left", anchor="e", text='Max column number:')
        self.maxcols_spiner = tk.Spinbox(self.controls_frame, from_=1, to=1000, increment=10,width=5)
        self.maxcols_lbl.grid(row=4,column=0,sticky='e')
        self.maxcols_spiner.grid(row=4,column=1,sticky='w', padx=(5, 5))

        self.navi_frame = ttk.Frame(self.top_frame)
        self.navi_frame.grid(row=2, padx=(10, 10),pady=(20,10),sticky='s')

        self.calc_btn = ttk.Button(self.navi_frame, text = 'Merge data', command = self.gui_calc)
        self.calc_btn.grid(row=1,column=1,sticky='nsew')
        self.close_btn = ttk.Button(self.navi_frame, command = lambda: root.destroy(), text = 'Close')
        self.close_btn.grid(row=1,column=3,sticky='nsew')
        self.help_btn = ttk.Button(self.navi_frame, text = 'Help', command = self.show_help)
        self.help_btn.grid(row=1,column=2,sticky='nsew')

        self.botom_frame = ttk.Frame(self.main_frame)
        self.botom_frame.grid(row=2)

        status_txt = "Prepare your master file first (read Help for more info)"
        self.status_lbl = PixelLabel(self.botom_frame, 1, borderwidth=1, relief='sunken', background='#D9D9D9', text = status_txt)
        self.status_lbl.grid(row=3,sticky='w',column=0)
        # self.status_lbl.grid(row=3,sticky='nsew',columnspan=1,wraplenght='4i')

        self.brag = ttk.Label(self.botom_frame, text = "about", borderwidth=1, relief='sunken', background='#D9D9D9',width=6)
        self.brag.grid(row=3,column=1,sticky='e')

        self.status_lbl.resize(root,self.brag)

        # root.tk.call("load", "", "Tix")
        self.baloon = tk.Balloon()
        self.baloon.bind_widget(self.brag, balloonmsg='ver1.0 created in Warsaw by pawel.cwiek@arup.com')

    def browse_dir(self):
        options = {}
        options['initialdir'] = self.dir_path.get()
        options['mustexist'] = True
        options['title'] = 'Location of folder with Excel files You want merged:'
        new_path = tkinter.filedialog.askdirectory(**options)
        if new_path != '': self.dir_path.set(new_path)

    def browse_file(self):
        options = {}
        if self.masterfile_path.get() == '':
            options['initialdir'] = self.dir_path.get()
        else:
            options['initialdir'] = self.masterfile_path.get()
        options['filetypes'] = [('Excel new files', '.xlsx')]
        options['title'] = 'Master spreadsheet location:'
        new_path = tkinter.filedialog.askopenfilename(**options)
        if new_path != '': self.masterfile_path.set(new_path)

    def merge_to_master(self):
        master_wb = self.masterfile_path.get().replace('\\','/')
        self.source_files = find_files(self.dir_path.get().replace('\\','/'),'.xlsx')
        cell_range = (int(self.maxrows_spiner.get()),int(self.maxcols_spiner.get()))
        self.errors = make_master_excel_file.merge_data(master_wb,self.source_files,cell_range)           

    def finish_calc(self):
        if self.calc_thread.is_alive():
            root.after(500,master.finish_calc)
        else:
            self.progress_bar.stop()
            self.progress_bar.destroy()

            if self.errors:
                txt = 'Calc finished with {} errors - look for info in cell comments. Files processed: {}.'.format(self.errors,len(self.source_files))
            else:
                txt = 'Calc finished succesfully. Files processed: {}. '.format(len(self.source_files))
            self.status_lbl.label.configure(text=txt)
            root.update()
            self.status_lbl.resize(root,self.brag)
            
            self.calc_btn.config(state="normal")


    def gui_calc(self):
        import threading
        
        self.calc_btn.config(state="disabled")

        self.progress_bar = ttk.Progressbar(self.botom_frame,orient='horizontal',mode='indeterminate')
        self.progress_bar.grid(row=1)
        self.progress_bar.start()

        self.calc_thread = threading.Thread(target=self.merge_to_master)
        self.calc_thread.deamon = True
        self.calc_thread.start()
        root.after(500,master.finish_calc)

    def show_help(self):
        info_text = '''
        The tool will SUM all the number type cells (can be formula calculated) from source files and paste sum into master.
        It will NOT:
            * fill in cells that are LOCKED in master (read below how to prepare master file)
        Max row and column number - type maximum range across all the tabs/sheets in Excel file to be used for processing
        NOTE: All source and master spreadsheets must have the SAME LAYOUT!

        1. Prepare Your master spreadsheet, possibly copy one of the source files
        2. Edit master spreadsheet:
            * for each spreadsheet tab:
                a) select all you want filled
                b) right-click and select "Format Cells"
                c) go to last tab and unselect "Locked"
            * if there are number cells in master file that you want to leave untouched by the tool, leave this cell as "Locked"

        The tool will put comment for each summed cell. Look for errors and success messages in master file.
        app for internal use only!
        Created by Pawel Cwiek @ Arup Warsaw
        '''
        toplevel = tkinter.Toplevel()
        toplevel.title='Help'
        label1 = tkinter.Label(toplevel, text=info_text, height=0, width=100, justify="left")
        label1.pack()

if __name__ == '__main__':

    root = tk.Tk()
    root.wm_title("Excel Master Spreadsheet Maker")
    root.resizable(0,0)
    master = MyGui(root)

    root.mainloop()
