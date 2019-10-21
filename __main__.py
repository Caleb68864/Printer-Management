import openpyxl
import win32print
import wx
from wx import CheckListBox
from wx import StaticText
from wx import BoxSizer
import FrmMain

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


class MyFrame(wx.Frame):
    def Select(self, sel_type, instance=None):
        if sel_type == "Location":
            if instance is not None:
                col = instance.GetEventObject().GetName()
                loc_id = col.replace(" ", "_")
                loc_id = loc_id.replace("col_", "")
                loc_list = self.FindWindowByName(loc_id, self)
                for i in range(0, len(loc_list.Items)):
                    if loc_list.IsChecked(i):
                        loc_list.Check(i, False)
                    else:
                        loc_list.Check(i, True)
        else:
            for loc in locations:
                checked = []
                loc_id = loc.replace(" ", "_")
                loc_list = self.FindWindowByName(loc_id, self)
                if sel_type == "All":
                    for p in printers:
                        if loc == p['Location']:
                            checked.append(p['Path'])
                    loc_list.SetCheckedStrings(checked)
                else:
                    for i in range(0, len(loc_list.Items)):
                        loc_list.Check(i, False)

    def updateStatus(self, txt):
        print(txt)
        self.sbStatus.PushStatusText(txt)
        self.sbStatus.Refresh()
        self.sbStatus.Update()

    def ir_Printer(self, ir):
        for loc in locations:
            checked = []
            loc_id = loc.replace(" ", "_")
            loc_list = self.FindWindowByName(loc_id, self)
            for i in loc_list.GetCheckedItems():
                item = loc_list.Items[i]
                # print(i)
                try:
                    status = ""
                    end_status = ""
                    if ir == "Install":
                        status = "Installing {}".format(item)
                        self.updateStatus(status)
                        win32print.AddPrinterConnection(item)
                        status = "Installation of {} Complete".format(item)
                        self.updateStatus(status)
                        end_status = "Installation Complete"
                    else:
                        status = "Removing {}".format(item)
                        self.updateStatus(status)
                        win32print.DeletePrinterConnection(item)
                        status = "Removal of {} Complete".format(item)
                        self.updateStatus(status)
                        end_status = "Removal Complete"
                    loc_list.Check(i, False)
                    loc_list.Refresh()
                    loc_list.Update()
                    self.Refresh()
                    self.Update()
                except Exception as e:
                    self.updateStatus(str(e))
        self.updateStatus(end_status)

    def btnSelectAll_Click(self, instance):
        self.Select("All")

    def btnSelectNone_Click(self, instance):
        self.Select("None")

    def btnSelectLoc_Click(self, instance):
        self.Select("Location", instance=instance)

    def btnInstall_Click(self, instance):
        self.ir_Printer("Install")

    def btnRemove_Click(self, instance):
        self.ir_Printer("Remove")

    def __init__(self, parent):
        FrmMain.FrmMain.__init__(self, parent)
        bsLocs = self.btnDestroy.GetContainingSizer()
        self.btnDestroy.Destroy()
        for loc in locations:
            loc_id = loc.replace(" ", "_")
            loc_box = BoxSizer()
            loc_box.SetOrientation(wx.VERTICAL)
            cbl_lbl = StaticText(self)
            cbl_lbl.Label = str(loc)
            cbl_lbl.SetName("col_{}".format(loc_id))
            cbl_lbl.Bind( wx.EVT_LEFT_UP, self.btnSelectLoc_Click )
            cbl_loc = CheckListBox(self)
            cbl_loc.SetName(loc_id)
            loc_box.Add(cbl_lbl, 0, wx.ALL, 5)
            loc_box.Add(cbl_loc, 0, wx.ALL, 5)
            bsLocs.Add(loc_box, 0, wx.ALL, 5)
        for p in printers:
            lid = p["Location"].replace(" ", "_")
            loc_list = self.FindWindowByName(lid, self)
            index = len(loc_list.Items) - 1
            if index < 0:
                index = 0

            loc_list.InsertItems([p["Path"]], index)
            # loc_list.SetString(index, p["Printer"])

        self.Layout()
        self.Fit()
        self.Show(True)


if __name__ == '__main__':
    app = wx.App(False)
    frame = MyFrame(None)
    app.MainLoop()
