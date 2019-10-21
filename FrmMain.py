# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Oct 26 2018)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class FrmMain
###########################################################################

class FrmMain ( wx.Frame ):

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Printer Management", pos = wx.DefaultPosition, size = wx.Size( 475,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

		self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

		bsMain = wx.BoxSizer( wx.VERTICAL )

		bsLocs = wx.BoxSizer( wx.HORIZONTAL )

		self.btnDestroy = wx.Button( self, wx.ID_ANY, u"To Be Removed", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsLocs.Add( self.btnDestroy, 0, wx.ALL, 5 )


		bsMain.Add( bsLocs, 1, wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL, 5 )

		bsButtons = wx.BoxSizer( wx.HORIZONTAL )

		self.btnSelectAll = wx.Button( self, wx.ID_ANY, u"Select All", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsButtons.Add( self.btnSelectAll, 0, wx.ALL, 5 )

		self.btnSelectNone = wx.Button( self, wx.ID_ANY, u"Select None", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsButtons.Add( self.btnSelectNone, 0, wx.ALL, 5 )

		self.btnInstall = wx.Button( self, wx.ID_ANY, u"Install", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsButtons.Add( self.btnInstall, 0, wx.ALL, 5 )

		self.btnRemove = wx.Button( self, wx.ID_ANY, u"Remove", wx.DefaultPosition, wx.DefaultSize, 0 )
		bsButtons.Add( self.btnRemove, 0, wx.ALL, 5 )


		bsMain.Add( bsButtons, 0, wx.ALIGN_CENTER_HORIZONTAL, 5 )


		self.SetSizer( bsMain )
		self.Layout()
		self.sbStatus = self.CreateStatusBar( 1, wx.STB_SIZEGRIP, wx.ID_ANY )

		self.Centre( wx.BOTH )

		# Connect Events
		self.btnSelectAll.Bind( wx.EVT_BUTTON, self.btnSelectAll_Click )
		self.btnSelectNone.Bind( wx.EVT_BUTTON, self.btnSelectNone_Click )
		self.btnInstall.Bind( wx.EVT_BUTTON, self.btnInstall_Click )
		self.btnRemove.Bind( wx.EVT_BUTTON, self.btnRemove_Click )

	def __del__( self ):
		pass


	# Virtual event handlers, overide them in your derived class
	def btnSelectAll_Click( self, event ):
		event.Skip()

	def btnSelectNone_Click( self, event ):
		event.Skip()

	def btnInstall_Click( self, event ):
		event.Skip()

	def btnRemove_Click( self, event ):
		event.Skip()


