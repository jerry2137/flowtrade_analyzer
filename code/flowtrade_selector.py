from datetime import timedelta
import os
import sys
import json
import warnings
warnings.filterwarnings('ignore', category=FutureWarning)

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import tkcalendar
from babel.dates import get_month_names
import pandastable

import pandas as pd


CODE_FOLDER = os.path.dirname(os.path.realpath(sys.argv[0]))
CONFIG_PATH = f'{CODE_FOLDER}/config.json'
FIRST_DATE = pd.to_datetime('2022/12/30')

TIPS = {
    'record': 'Select the trade record Excel file',
    'record_sheet': 'Trade record sheet name',
}

class ToolTip(object):
    def __init__(self):
        self.tipwindow = None

    def showtip(self, event):
        if type(event.widget) is myRowHeader:
            id = event.widget.find_withtag("current")[0]
            self.text = event.widget.itemcget(id, "text")
        else:
            self.text = TIPS[event.widget.tag]
        
        x = event.x + event.widget.winfo_rootx() + 10
        y = event.y + event.widget.winfo_rooty() + 10

        self.tipwindow = tk.Toplevel(event.widget)
        self.tipwindow.attributes('-topmost', True)
        self.tipwindow.wm_overrideredirect(True)
        self.tipwindow.wm_geometry("+%d+%d" % (x, y))
        tk.Label(
            self.tipwindow,
            text=self.text,
            justify='left',
            background="#ffffe0",
            relief='solid',
            borderwidth=1,
            font=("tahoma", "8", "normal")
        ).pack(ipadx=1)

    def hidetip(self, event):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None

class YearMonthDateEntry(tkcalendar.DateEntry):
    def __init__(self, master=None, root=None, cycle='day', name='start', **kw):
        self.root = root
        self.cycle = cycle
        self.name = name
        if self.cycle not in ['day', 'month', 'year']:
            raise ValueError('cycle must be one of "day", "month", "year"')
        if self.name not in ['start', 'end']:
            raise ValueError('name must be one of "start", "end"')

        super().__init__(master, date_pattern='yyyy/mm/dd', state='readonly', **kw)

        self.bind('<FocusIn>', self._my_set_text)
        self.bind('<FocusOut>', self._my_set_text)
        self.bind('<<DateEntrySelected>>', self.root.update_date)

    def _my_set_text(self, event=None):
        offset = 0
        if event:
            widget_config = str(event.widget).split('.!')
            if widget_config[-1].startswith('button'):
                # previous
                if widget_config[-1] == 'button':
                    offset = -1
                # next
                elif widget_config[-1] == 'button2':
                    offset = 1
                # month
                if widget_config[-2] == 'frame':
                    name2month = {number:name for name, number in get_month_names('wide', locale='en').items()}
                    new_month = (name2month[self._calendar._header_month.cget('text')] - 1 + offset) % 12 + 1
                    if self.name == 'end':
                        next_month = self._date.replace(month=new_month, day=28) + timedelta(days=4)
                        self._date = next_month - timedelta(days=next_month.day)
                    else:
                        self._date = self._date.replace(month=new_month)
                # year
                elif widget_config[-2] == 'frame2':
                    self._date = self._date.replace(year=int(self._calendar._header_year.cget('text')) + offset)
        date = pd.to_datetime(self.format_date(self._date))

        if self.cycle == 'day':
            txt = date.strftime('%Y/%m/%d')
        if self.cycle == 'month':
            txt = date.strftime('%Y/%m')
        elif self.cycle == 'year':
            txt = date.strftime('%Y')
        self._set_text(txt)
        self.root.update_date(0)

    def set_cycle_day(self):
        self.cycle = 'day'
        self._my_set_text()

        self._calendar._cal_frame.pack(fill="both", expand=True, padx=2, pady=2)
        self._calendar._l_month.pack(side='left', fill="y")
        self._calendar._header_month.pack(side='left', padx=4)
        self._calendar._r_month.pack(side='left', fill="y")
        
        self._calendar._l_year.unbind('<ButtonRelease>')
        self._calendar._r_year.unbind('<ButtonRelease>')
        self._calendar._l_month.unbind('<ButtonRelease>')
        self._calendar._r_month.unbind('<ButtonRelease>')

    def set_cycle_month(self):
        self.cycle = 'month'
        if self.name == 'start':
            self._date = self._date.replace(day=1)
        elif self.name == 'end':
            next_month = self._date.replace(day=28) + timedelta(days=4)
            self._date = next_month - timedelta(days=next_month.day)
        self._my_set_text()

        self._calendar._cal_frame.pack_forget()
        self._calendar._l_month.pack(side='left', fill="y")
        self._calendar._header_month.pack(side='left', padx=4)
        self._calendar._r_month.pack(side='left', fill="y")

        self._calendar._l_year.bind('<ButtonRelease>', self._my_set_text)
        self._calendar._r_year.bind('<ButtonRelease>', self._my_set_text)
        self._calendar._l_month.bind('<ButtonRelease>', self._my_set_text)
        self._calendar._r_month.bind('<ButtonRelease>', self._my_set_text)
    
    def set_cycle_year(self):
        self.cycle = 'year'
        if self.name == 'start':
            self._date = self._date.replace(month=1, day=1)
        elif self.name == 'end':
            self._date = self._date.replace(month=12, day=31)
        self._my_set_text()

        self._calendar._cal_frame.pack_forget()
        self._calendar._l_month.pack_forget()
        self._calendar._header_month.pack_forget()
        self._calendar._r_month.pack_forget()

        self._calendar._l_year.bind('<ButtonRelease>', self._my_set_text)
        self._calendar._r_year.bind('<ButtonRelease>', self._my_set_text)

    def configure(self, cycle=None, cnf={}, **kw):
        super().configure(cnf, **kw)
        if not cycle:
            return
        if cycle == 'same':
            cycle = self.cycle
        if cycle == 'day':
            self.set_cycle_day()
        elif cycle == 'month':
            self.set_cycle_month()
        elif cycle == 'year':
            self.set_cycle_year()
        else:
            raise ValueError('cycle must be one of "day", "month", "year", "same"')

    def my_get_date(self):
        return self._date

class ListboxFrame(ttk.Frame):
    def __init__(self, root, parent, name):
        self.root = root
        self.name = name

        super().__init__(parent)
        ttk.Label(self, text=name).grid(row=0, column=0)
        tk.Button(self, text='select all', command=self.select_all).grid(row=0, column=1)
        tk.Button(self, text='clear all', command=self.clear_all).grid(row=0, column=2)

        self.frame_box = ttk.Frame(self)
        self.frame_box.grid(row=1, columnspan=3)
        self.vscrollbar = ttk.Scrollbar(self.frame_box)
        self.listbox = tk.Listbox(
            self.frame_box,
            selectmode='multiple',
            yscrollcommand=self.vscrollbar.set,
            exportselection=False,
        )
        self.vscrollbar.config(command=self.listbox.yview)
        self.vscrollbar.pack(side='right', fill='y')
        self.listbox.pack()

        tk.Button(self, text='summary on this column', command=self.summary).grid(row=2, columnspan=3)

        self.listbox.bind('<<ListboxSelect>>', self.update_root_df)

    def update_options(self, options):
        self.listbox.delete(0, 'end')
        for option in options:
            self.listbox.insert('end', option)
            self.listbox.select_set('end')

    def select_all(self):
        for i in range(len(self.listbox.get(0, 'end'))):
            self.listbox.select_set(i)
        self.update_root_df()

    def clear_all(self):
        for i in range(len(self.listbox.get(0, 'end'))):
            self.listbox.select_clear(i)
        self.update_root_df()
    
    def summary(self):
        self.root.update_df(summary_column=self.name)

    def update_root_df(self, *_):
        self.root.update_df(column=self.name, values=[self.listbox.get(i) for i in self.listbox.curselection()])

class myRowHeader(pandastable.RowHeader):
    def __init__(self, parent=None, table=None):
        self.max_col_width = 0
        super().__init__(parent=parent, table=table)

    def redraw(self):
        max_column_width = 200

        self.height = self.table.rowheight * self.table.rows+10
        self.configure(scrollregion=(0,0, self.width, self.height))
        self.delete('rowheader','text', 'rect')

        visiblerows = self.table.visiblerows
        if not visiblerows:
            return
        scale = self.table.getScale()
        index = self.model.df.index

        if isinstance(index, pd.core.indexes.multi.MultiIndex):
            cols_all = [pd.Series(i).astype('object').astype(str).replace('nan','') for i in list(zip(*index.values))]
            l = [c.str.len().max() for c in cols_all]

            nl = [len(n) if n is not None else 0 for n in index.names]
            #pick higher of index names and row data
            l = [max(l_element, nl_element) for l_element, nl_element in zip(l, nl)]
            widths = [min(i*scale+6, max_column_width) for i in l]
            cols = [pd.Series(i).astype('object').astype(str).replace('nan','') for i in list(zip(*index.values[visiblerows]))]
            
            def cumsum(l):
                total = 0
                for x in l:
                    yield total
                    total += x
            xpos = list(cumsum(widths))
            # xpos = [0] + list(np.cumsum(widths))[:-1]
        else:
            l = index.astype('str').str.len().max()
            widths = [min(l*scale+6, max_column_width)]
            cols = [index[visiblerows].fillna('').astype('object').astype('str')]
            xpos = [0]
        
        w = sum(widths)
        self.widths = widths

        w = max(w, 50)
        if self.width != w:
            self.config(width=w)
            self.width = w

        for i in range(len(cols)):
            x = xpos[i]
            for j in range(len(cols[i])):
                text = cols[i][j]
                x1, y1, x2, y2 = self.table.getCellCoords(j+visiblerows[0], 0)

                # skip if all indices are the same as the previous row
                if j != 0:
                    skip = True
                    now_index = i
                    while now_index >= 0:
                        if cols[now_index][j] != cols[now_index][j-1]:
                            skip = False
                            break
                        now_index -= 1
                    if skip:
                        continue

                rect_id = self.create_rectangle(x, y1, w-1, y2, outline='white', width=1, tag='rowheader', fill=self.bgcolor)
                text_id = self.create_text(x+5, y1+self.table.rowheight/2, text=text, fill='black', font=self.table.thefont, tag='text', anchor='w')
                last_text = text if text != 'Total' else ''

                toolTip = ToolTip()
                self.tag_bind(text_id, '<Enter>', toolTip.showtip)
                self.tag_bind(text_id, '<Leave>', toolTip.hidetip)

        self.config(bg=self.bgcolor)

class Root(tk.Tk):
    def __init__(self):
        super().__init__()
        # self.wm_attributes('-topmost', True)
        self.resizable(width=False, height=False)

        style = ttk.Style()
        style.configure('TFrame', background='light gray', relief='ridge')
        style.configure('Heading.TLabel', font=('Arial', 25), background='SystemButtonFace')
        style.configure('Heading2.TLabel', font=('Arial', 15), background='SystemButtonFace')
        style.configure('TLabel', background='light gray')

        self.entries_boxes = {}
        self.list_boxes = {}
        self.table_condition = {}
        self.summary_columns = []
        self.columns = [
                'Account',
                'customer2',
                'Trade Dt',
                'Security',
                'count',
                'ID',
                'Qty (M)',
                'Profit',
                'Notional USD',
        ]

        self.title('Capital Bond Report Generator')

        ttk.Label(self, text='Capital Bond Report Generator', style='Heading.TLabel').grid(ipady=20, ipadx=10)

        # Select file
        self.file_frame = ttk.Frame(self, width=700, height=100)
        self.file_frame.grid()
        self.file_frame.rowconfigure((0,1,2), weight=1)
        self.file_frame.columnconfigure(0, weight=1)
        self.file_frame.grid_propagate(False)

        ttk.Label(self.file_frame, text='Select file', font=('Arial', 15)).grid()

        ## Read trade record
        self.record_frame = ttk.Frame(self.file_frame)
        self.record_frame.grid()

        record_file_box = ttk.Entry(self.record_frame, width=60)
        self.entries_boxes['record'] = record_file_box
        record_file_box.grid(column=0, row=0)

        tk.Button(self.record_frame, text='browse', command=self.select_file).grid(column=1, row=0)

        record_sheet_box = ttk.Entry(self.record_frame, width=25)
        self.entries_boxes['record_sheet'] = record_sheet_box
        record_sheet_box.grid(column=2, row=0)

        for name, box in self.entries_boxes.items():
            box.tag = name
            self.add_hint(box)
            box.bind('<FocusOut>', self.add_hint)
            box.bind('<FocusIn>', self.remove_hint)
            toolTip = ToolTip()
            box.bind('<Enter>', toolTip.showtip)
            box.bind('<Leave>', toolTip.hidetip)

        tk.Button(self.file_frame, text='select', command=self.read_file).grid()

        # Set conditions
        self.condition_frame = ttk.Frame(self, width=700, height=100)
        self.condition_frame.grid()

        ttk.Label(self.condition_frame, text='Set conditions', font=('Arial', 15)).grid(row=0, columnspan=6)

        # Select time range
        self.time_frame = ttk.Frame(self.condition_frame)
        self.time_frame.grid(row=1, column=0)
        
        self.interval_frame = ttk.Frame(self.time_frame)
        self.interval_frame.grid()
        ttk.Label(self.interval_frame, text='time interval').grid(row=0, column=0)
        self.interval_box = ttk.Combobox(self.interval_frame, values=['day', 'month', 'year'], state='readonly', width=10)
        self.interval_box.grid(row=0, column=1)
        self.interval_box.bind('<<ComboboxSelected>>', self.change_interval)
        self.interval_box.current(0)

        self.start_date_frame = ttk.Frame(self.time_frame)
        self.start_date_frame.grid()
        ttk.Label(self.start_date_frame, text='start date').grid(row=0, column=0)
        self.start_date_box = YearMonthDateEntry(self.start_date_frame, name='start', root=self)
        self.start_date_box.grid(row=0, column=1)

        self.end_date_frame = ttk.Frame(self.time_frame)
        self.end_date_frame.grid()
        ttk.Label(self.end_date_frame, text='end date').grid(row=0, column=0)
        self.end_date_box = YearMonthDateEntry(self.end_date_frame, name='end', root=self)
        self.end_date_box.grid(row=0, column=1)

        self.total_var = tk.BooleanVar()
        ttk.Checkbutton(self.time_frame, text='show total', variable=self.total_var).grid()

        for i, column_name in enumerate(['Trade Dt', 'Account', 'customer2', 'Security', 'count']):
            self.list_boxes[column_name] = ListboxFrame(self, self.condition_frame, column_name)
            self.list_boxes[column_name].grid(row=1, column=i+1)

        # read config
        self.read_config()

        ttk.Label(self, text='Edit Info', style='Heading2.TLabel').grid()
        self.table_frame = ttk.Frame(self)
        self.table_frame.grid()

        self.table = pandastable.Table(
            self.table_frame,
            dataframe=pd.DataFrame(columns=self.columns),
            showstatusbar=True,
            width=800,
            height=350
        )
        self.table.editable = False

        # Actions on closing the window
        self.protocol('WM_DELETE_WINDOW', self.on_closing)
        # Start the window
        self.mainloop()

    def select_file(self):
        entry_box = self.entries_boxes['record']
        folderpath = filedialog.askopenfilename(title=TIPS['record'], filetypes=[('Excel', 'xls'), ('Excel', 'xlsx'), ('Excel', 'xlsm')])
        self.update()
        if folderpath:
            entry_box.delete(0, 'end')
            entry_box.insert('end', folderpath)
            entry_box.configure(foreground='black')
    
    def add_hint(self, event):
        entry_box = event.widget if type(event) is tk.Event else event
        if entry_box.get() == '':
            entry_box.insert('end', TIPS[entry_box.tag])
            entry_box.configure(foreground='gray')

    def remove_hint(self, event):
        entry_box = event.widget if type(event) is tk.Event else event
        if entry_box.get() == TIPS[entry_box.tag]:
            entry_box.delete(0, 'end')
            entry_box.configure(foreground='black')
    
    def add_total(self, df, columns=[]):
        total = df.sum(numeric_only=True)
        if not columns:
            return total
        
        df = df.groupby([columns[0]]).apply(self.add_total, columns=columns[1:], include_groups=False)
    
        if isinstance(df.index, pd.MultiIndex):
            if len(df.index.levels[0]) == 1:
                return df
            df.loc[tuple(['Total']*len(columns)), :] = total
            df = df.reindex(['Total'] + list(df.index.levels[0][:-1]), level=0)
        else:
            if len(df.index) == 1:
                return df
            df.loc['Total', :] = total
            df = df.reindex(['Total'] + list(df.index[:-1]))

        return df

    def update_df(self, column='', values=[], summary_column=''):
        if column:
            self.table_condition[column] = values
        df = self.df_record

        for now_column, now_values in self.table_condition.items():
            df = df.loc[df[now_column].astype(str).isin(now_values)]

        if not summary_column:
            pass
        elif summary_column in self.summary_columns:
            self.summary_columns.remove(summary_column)
        else:
            self.summary_columns.append(summary_column)
        if self.summary_columns:
            df['COUNT'] = 1
            if self.total_var.get():
                df = self.add_total(df, self.summary_columns)
            else:
                df = df.groupby(self.summary_columns).sum(numeric_only=True)
        
        model = pandastable.TableModel(dataframe=df)
        self.table.updateModel(model)
        
        if hasattr(self.table, 'rowheader'):
            self.table.rowheader.config(width=45)
            
        self.table.show()

        self.table.rowheader.grid_forget()
        self.table.rowheader = myRowHeader(parent=self.table_frame, table=self.table)
        self.table.rowheader.grid(row=1,column=0,rowspan=1,sticky='news')
        
        self.table.showIndex()
        self.table.redraw()

    def read_file(self):
        try:
            record_path = self.entries_boxes['record'].get()
            record_sheet = self.entries_boxes['record_sheet'].get()
            self.df_record = pd.read_excel(
                record_path,
                sheet_name=record_sheet,
                usecols=self.columns,
            )
            self.table_condition = {}
            self.summary_columns = []
            self.df_record['Trade Dt'] = self.df_record['Trade Dt'].dt.date

            for column in self.df_record.columns:
                self.table_condition[column] = self.df_record[column].unique().astype(str)

            for column_name, list_box in self.list_boxes.items():
                list_box.update_options(self.df_record[column_name].unique().astype(str))

            self.start_date_box.set_date(self.df_record['Trade Dt'].iloc[0])
            self.start_date_box.configure(cycle='same')
            self.end_date_box.set_date(self.df_record['Trade Dt'].iloc[-1])
            self.end_date_box.configure(cycle='same')
            
        except Exception as e:
            messagebox.showerror('Error', e)

    def update_date(self, event):
        start_date = self.start_date_box.my_get_date()
        end_date = self.end_date_box.my_get_date()
        
        option_dates = []
        for date in pd.date_range(start_date, end_date):
            date_str = str(date.date())
            
            if date_str in self.df_record['Trade Dt'].unique().astype(str):
                option_dates.append(date_str)

        self.list_boxes['Trade Dt'].update_options(option_dates)
        self.list_boxes['Trade Dt'].update_root_df(0)

    def change_interval(self, event):
        interval = event.widget.get()
        self.start_date_box.configure(cycle=interval)
        self.end_date_box.configure(cycle=interval)

    def read_config(self):
        if not os.path.isfile(CONFIG_PATH):
            return
        with open(CONFIG_PATH, 'r', encoding='utf-8') as file:
            config_dict = json.load(file)
        for name, entry_box in self.entries_boxes.items():
            if config_dict[name]:
                entry_box.delete(0, 'end')
                entry_box.insert('end', config_dict[name])
                entry_box.configure(foreground='black')

    def on_closing(self):
        config_dict = {}
        for name, entry_box in self.entries_boxes.items():
            if entry_box.get() != TIPS[name]:
                config_dict[name] = entry_box.get()
            else:
                config_dict[name] = ''
        with open(CONFIG_PATH, 'w') as file:
            json.dump(config_dict, file)
        
        self.destroy()


if __name__ == '__main__':
    root = Root()