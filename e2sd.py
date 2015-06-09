#Boa:Frame:MainFrame
# -*- coding: cp1250 -*-

import wx, os.path, e_max, win32api, sys
#import ConfigParser

def create(parent):
    return MainFrame(parent)

[wxID_MAINFRAME, wxID_MAINFRAMEBUTTON1, wxID_MAINFRAMECONVERTBUTTON, 
 wxID_MAINFRAMEDXFCHECKBOX, wxID_MAINFRAMEDXFFILENAME, 
 wxID_MAINFRAMEFOLDERBUTTON, wxID_MAINFRAMEPANEL1, wxID_MAINFRAMEPOINTS, 
 wxID_MAINFRAMEPOINTSIN3D, wxID_MAINFRAMESOLIDCHECKBOX, 
 wxID_MAINFRAMESTATICTEXT1, wxID_MAINFRAMESTATICTEXT2, wxID_MAINFRAMETCHTEXT, 
] = [wx.NewId() for _init_ctrls in range(13)]

class FileDropTarget(wx.FileDropTarget):
    """ This object implements Drop Target functionality for Files """
    def __init__(self, obj):
        """ Initialize the Drop Target, passing in the Object Reference to
        indicate what should receive the dropped files """
        # Initialize the wxFileDropTarget Object
        wx.FileDropTarget.__init__(self)
        # Store the Object Reference for dropped files
        self.obj = obj

    def OnDropFiles(self, x, y, filenames):
        """ Implement File Drop """
        # For Demo purposes, this function appends a list of the files dropped at the end of the widget's text
        # Move Insertion Point to the end of the widget's text
        #self.obj.SetInsertionPointEnd()
        # append a list of the file names dropped
        #self.obj.WriteText("%d file(s) dropped at %d, %d:\n" % (len(filenames), x, y))
        #for file in filenames:
        if len(filenames) == 1 and filenames[0][-3:].lower() == "tch":
            self.obj.Value = filenames[0]
        elif len(filenames) > 1 and filenames[0][-3:].lower() == "tch":
            self.obj.Value = filenames[0]
        #self.obj.WriteText('\n')
    #def OnDragOver(self, x, y, filenames): 
    #   self.obj2.SetBackgroundColour("Green") 
    #   self.obj2.Refresh()
    #   #self.obj.WriteText('hehe')


class MainFrame(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_MAINFRAME, name=u'MainFrame',
              parent=prnt, pos=wx.Point(430, 267), size=wx.Size(713, 314),
              style=wx.DEFAULT_FRAME_STYLE, title=u'E2SD')
        self.SetClientSize(wx.Size(705, 287))

        self.panel1 = wx.Panel(id=wxID_MAINFRAMEPANEL1, name='panel1',
              parent=self, pos=wx.Point(0, 0), size=wx.Size(705, 287),
              style=wx.TAB_TRAVERSAL)
        self.panel1.Enable(True)

        self.dxfcheckbox = wx.CheckBox(id=wxID_MAINFRAMEDXFCHECKBOX,
              label=u'Output to dxf file', name=u'dxfcheckbox',
              parent=self.panel1, pos=wx.Point(32, 160), size=wx.Size(112, 13),
              style=0)
        self.dxfcheckbox.SetValue(True)
        self.dxfcheckbox.SetToolTipString(u'toggle converting to dxf')
        self.dxfcheckbox.Bind(wx.EVT_CHECKBOX, self.OnDxfcheckboxCheckbox,
              id=wxID_MAINFRAMEDXFCHECKBOX)

        self.solidcheckbox = wx.CheckBox(id=wxID_MAINFRAMESOLIDCHECKBOX,
              label=u'Output to SolidWorks Sketch', name=u'solidcheckbox',
              parent=self.panel1, pos=wx.Point(32, 224), size=wx.Size(176, 13),
              style=0)
        self.solidcheckbox.SetValue(False)
        self.solidcheckbox.SetToolTipString(u'toggle convert to solidworks script')

        self.staticText1 = wx.StaticText(id=wxID_MAINFRAMESTATICTEXT1,
              label=u'Drop file onto program window or paste link to file or use Open button',
              name='staticText1', parent=self.panel1, pos=wx.Point(88, 40),
              size=wx.Size(504, 19), style=0)
        self.staticText1.SetFont(wx.Font(12, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))

        self.tchtext = wx.TextCtrl(id=wxID_MAINFRAMETCHTEXT, name=u'tchtext',
              parent=self.panel1, pos=wx.Point(16, 72), size=wx.Size(584, 21),
              style=0, value=u'')
        self.tchtext.Bind(wx.EVT_TEXT, self.OnTchtextText,
              id=wxID_MAINFRAMETCHTEXT)

        self.button1 = wx.Button(id=wxID_MAINFRAMEBUTTON1, label=u'Open',
              name='button1', parent=self.panel1, pos=wx.Point(616, 72),
              size=wx.Size(75, 23), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_MAINFRAMEBUTTON1)

        self.staticText2 = wx.StaticText(id=wxID_MAINFRAMESTATICTEXT2,
              label=u'Convert E-max file (*.tch) to DXF file or to Solidworks sketch',
              name='staticText2', parent=self.panel1, pos=wx.Point(64, 8),
              size=wx.Size(569, 25), style=0)
        self.staticText2.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))

        self.dxffilename = wx.TextCtrl(id=wxID_MAINFRAMEDXFFILENAME,
              name=u'dxffilename', parent=self.panel1, pos=wx.Point(32, 184),
              size=wx.Size(568, 21), style=0, value=u'')
        self.dxffilename.SetToolTipString(u'Enter dxf filename')
        self.dxffilename.Bind(wx.EVT_TEXT, self.OnDxffilenameText,
              id=wxID_MAINFRAMEDXFFILENAME)

        self.convertbutton = wx.Button(id=wxID_MAINFRAMECONVERTBUTTON,
              label=u'CONVERT', name=u'convertbutton', parent=self.panel1,
              pos=wx.Point(32, 104), size=wx.Size(568, 39), style=0)
        self.convertbutton.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, u'Tahoma'))
        self.convertbutton.SetToolTipString(u'Start converting tch to required format')
        self.convertbutton.Bind(wx.EVT_BUTTON, self.OnConvertbuttonButton,
              id=wxID_MAINFRAMECONVERTBUTTON)

        self.folderbutton = wx.Button(id=wxID_MAINFRAMEFOLDERBUTTON,
              label=u'Show Folder', name=u'folderbutton', parent=self.panel1,
              pos=wx.Point(616, 184), size=wx.Size(80, 23), style=0)
        self.folderbutton.Enable(False)
        self.folderbutton.Bind(wx.EVT_BUTTON, self.OnFolderbuttonButton,
              id=wxID_MAINFRAMEFOLDERBUTTON)

        self.points = wx.CheckBox(id=wxID_MAINFRAMEPOINTS,
              label=u'Include points entered while creating other objects',
              name=u'points', parent=self.panel1, pos=wx.Point(32, 256),
              size=wx.Size(264, 13), style=0)
        self.points.SetValue(True)

        self.pointsin3d = wx.CheckBox(id=wxID_MAINFRAMEPOINTSIN3D,
              label=u'3D points', name=u'pointsin3d', parent=self.panel1,
              pos=wx.Point(144, 160), size=wx.Size(70, 13), style=0)
        self.pointsin3d.SetValue(False)

    def __init__(self, parent):
        self._init_ctrls(parent)
        ARGS = sys.argv
        if len(ARGS) == 2 and os.path.isfile(ARGS[1]) and ARGS[1][-3:].lower() == "tch":
            self.tchtext.Value = ARGS[1]
        dt1 = FileDropTarget(self.tchtext)
        self.panel1.SetDropTarget(dt1)
        self.tcholdtext = ""
        #odczytanie ustawien z ostatniego uruchomienia
        #CONFIG = ConfigParser.ConfigParser()


    def OnButton1Button(self, event):
        wildcard = "E-max Files (*.tch)|*.tch"
        dialog = wx.FileDialog ( self, message="Choose a tch file", wildcard = wildcard )
        if dialog.ShowModal() == wx.ID_OK:
            self.tchtext.Value = dialog.GetPath()
        dialog.Destroy()
        event.Skip()

    def OnTchtextText(self, event):
        if self.tchtext.Value[-3:].lower() == "tch" and os.path.isfile(self.tchtext.Value) and self.dxfcheckbox.Value == True:
            if self.dxffilename.Value.strip() == "" or not self.dxffilename.Value[-3:].lower() == "dxf" or not self.tcholdtext == "":
                self.dxffilename.Value = self.tchtext.Value[:-3]+"dxf"
                self.tcholdtext = self.tchtext.Value
        event.Skip()

    def OnDxfcheckboxCheckbox(self, event):
        if self.dxfcheckbox.Value == False:
            self.dxffilename.Enabled = False
        else:
            self.dxffilename.Enabled = True
            if os.path.isfile(self.tchtext.Value):
                self.dxffilename.Value = self.dxffilename.Value = self.tchtext.Value[:-3]+"dxf"
        event.Skip()

    def OnConvertbuttonButton(self, event):
        if os.path.isfile(self.tchtext.Value):
            E = e_max.tch_file(self.tchtext.Value)
            E.get_objects()
            SUCCESS = False
            if self.dxfcheckbox.Value == True and len(self.dxffilename.Value) > 4:
                #konwersja do dxf
                try:
                    E.converttodxf(self.dxffilename.Value, self.points.Value, self.pointsin3d.Value)
                    SUCCESS = True
                except:
                    dlg = wx.MessageDialog(self,
                        "Error while saving to dxf",
                        "Error", wx.OK)
                    result = dlg.ShowModal()
            if self.solidcheckbox.Value:
                #konwersja do solida
                #try:
                E.solidsketch(self.points.Value, self.pointsin3d.Value)
                SUCCESS = True
                #except:
                #    dlg = wx.MessageDialog(self,
                #        "Error while comunicating with SolidWorks API",
                #        "Error", wx.OK)
                #    result = dlg.ShowModal()
            if SUCCESS:
                #self.tchtext.Value = ""
                #self.dxffilename.Value = ""
                dlg = wx.MessageDialog(self,
                    "Convert done!",
                    "Message", wx.OK)
                result = dlg.ShowModal()
                dlg.Destroy()
                dlg.Destroy()
		
        event.Skip()

    def OnFolderbuttonButton(self, event):
        if os.path.isfile(self.dxffilename.Value):
            win32api.WinExec('explorer /select,'+self.dxffilename.Value.replace('/','\\'))
        elif os.path.isdir(os.path.dirname(self.dxffilename.Value)):
            win32api.WinExec('explorer '+os.path.dirname(self.dxffilename.Value).replace('/','\\'))
        event.Skip()

    def OnDxffilenameText(self, event):
        if os.path.isdir(os.path.dirname(self.dxffilename.Value)) or os.path.isfile(self.dxffilename.Value):
            self.folderbutton.Enabled = True
        else:
            self.folderbutton.Enabled = False    
        event.Skip()


if __name__ == '__main__':
    app = wx.PySimpleApp()
    frame = create(None)
    frame.Show()

    app.MainLoop()
