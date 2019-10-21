import openpyxl
import win32print
import kivy
from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.checkbox import CheckBox
kivy.require("1.10.1")

wb = openpyxl.load_workbook('Printers.xlsx')
sheet = wb['Printers']
# sheet = wb['Test']


def getcell(cell):
    if cell is not None:
        if sheet[cell].value is not None:
            return sheet[cell].value.strip()
        else:
            return ""
    else:
        return ""


locations = []
printers = []
for row in range(2, sheet.max_row + 1):
    printer = getcell('A' + str(row))
    path = getcell('B' + str(row))
    location = getcell('C' + str(row))
    if location not in locations:
        locations.append(location)
    printers.append({"Printer": printer, "Path": path, "Location": location})

printers = sorted(printers, key=lambda i: (i['Printer'], i['Location'], i["Path"]))
locations.sort()
    # if printer is not None or printer != "":
    #     print(printer)
        # win32print.AddPrinterConnection(printer)
        # win32print.DeletePrinterConnection(printer)
        # win32print.SetDefaultPrinter(printer)


class MainPage(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        # self.cols = 1
        self.orientation = "vertical"
        self.padding = '15sp'
        self.loc_title = GridLayout()
        self.loc_title.cols = len(locations)
        self.add_widget(self.loc_title)
        self.loc_grid = GridLayout()
        self.loc_grid.cols = len(locations)
        self.loc_grid.size_hint_y = 20
        self.loc_grid.id = "loc_grid"
        self.add_widget(self.loc_grid)


        for loc in locations:
            loc_id = loc.replace(" ", "_")
            title_grid = BoxLayout()
            title_grid.size_hint_y = 2
            title_grid.add_widget(Label(text=loc, font_size='20sp'))
            loc_check = CheckBox()
            loc_check.bind(active=self.tg_box)
            loc_check.id = loc_id
            title_grid.add_widget(loc_check)
            self.loc_title.add_widget(title_grid)

            self.loc_box = BoxLayout()
            self.loc_box.orientation = "vertical"
            self.loc_box.id = loc_id
            self.loc_grid.add_widget(self.loc_box)
            for p in printers:
                if p["Location"] == loc:
                    print_grid = GridLayout()
                    print_grid.cols = 2
                    lb_id = p["Location"].replace(" ", "_")
                    print_grid.add_widget(Label(text=p["Printer"]))
                    check = CheckBox()
                    check.id = p["Path"]
                    print_grid.add_widget(check)
                    self.loc_box.add_widget(print_grid)

        print(self.ids)

        btnSelectAll = Button()
        btnSelectAll.text = "Select All"
        btnSelectAll.bind(on_press=self.sa_button)
        self.add_widget(btnSelectAll)

        btnSelectNone = Button()
        btnSelectNone.text = "Select None"
        btnSelectNone.bind(on_press=self.sn_button)
        self.add_widget(btnSelectNone)

        btnInstall = Button()
        btnInstall.text = "Install"
        btnInstall.bind(on_press=self.install_button)
        self.add_widget(btnInstall)

        btnRemove = Button()
        btnRemove.text = "Remove"
        btnRemove.bind(on_press=self.remove_button)
        self.add_widget(btnRemove)

        self.statusbar = Label(text="")
        self.add_widget(self.statusbar)

    def sa_button(self, instance):
        for lg_child in self.loc_grid.children:
            for pg in lg_child.children:
                pg.children[0].active = True

    def sn_button(self, instance):
        for lg_child in self.loc_grid.children:
            for pg in lg_child.children:
                pg.children[0].active = False

    def process_checkboxes(self, ir):
        for lg_child in self.loc_grid.children:
            for pg in lg_child.children:
                txt = pg.children[1].text
                cb_id = pg.children[0].id
                if txt not in locations and txt is not None and cb_id is not None:
                    if pg.children[0].active:
                        try:
                            if ir:
                                print("Installing {}".format(txt))
                                self.statusbar.text = "Installing {}".format(txt)
                                win32print.AddPrinterConnection(cb_id)
                                print("Installation of {} Complete".format(txt))
                                self.statusbar.text = "Installation of {} Complete".format(txt)
                            else:
                                try:
                                    print("Removing {}".format(txt))
                                    self.statusbar.text = "Removing {}".format(txt)
                                    win32print.DeletePrinterConnection(cb_id)
                                    print("Removal of {} Complete".format(txt))
                                    self.statusbar.text = "Removal of {} Complete".format(txt)
                                except Exception as e:
                                    if str(e) == "(1801, 'DeletePrinterConnection', 'The printer name is invalid.')":
                                        self.statusbar.text = "{} Not Currently Installed".format(txt)
                                    else:
                                        self.statusbar.text = str(e)
                                        print(e)
                        except Exception as e:
                            print(e)

    def install_button(self, instance):
        self.process_checkboxes(True)

    def remove_button(self, instance):
        self.process_checkboxes(False)

    def tg_box(self, cb, value):
        for cb_child in cb.parent.parent.parent.children:
            if cb_child.id == "loc_grid":
                for lg_child in cb_child.children:
                    # Location Grid Columns
                    # print(lg_child.id)
                    if lg_child.id == cb.id:
                        for pgc in lg_child.children:
                            # print(pgc)
                            for cbc in pgc.children:
                                if value:
                                    cbc.active = True
                                else:
                                    cbc.active = False


class PrinterManager(App):
    def build(self):
        self.title = "Printer Management"
        return MainPage()


if __name__ == '__main__':
    PrinterManager().run()