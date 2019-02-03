# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Sep 12 2010)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.richtext


###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1(wx.Frame):

	def __init__(self, parent):
		wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"HCE RRAP TOOL v2.0", pos=wx.DefaultPosition,
						  size=wx.Size(500, 600), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

		self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

		fgSizer1 = wx.FlexGridSizer(2, 2, 0, 0)
		fgSizer1.SetFlexibleDirection(wx.BOTH)
		fgSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

		gbSizer1 = wx.GridBagSizer(0, 0)
		gbSizer1.SetFlexibleDirection(wx.BOTH)
		gbSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

		self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"Test Tool Report", wx.DefaultPosition, wx.DefaultSize, 0)
		self.m_staticText1.Wrap(-1)
		gbSizer1.Add(self.m_staticText1, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 8)

		inputChoiceChoices = [u"ICCSolutions", u"Galitt", u"UL"]
		self.inputChoice = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, inputChoiceChoices, 0)
		self.inputChoice.SetSelection(1)
		gbSizer1.Add(self.inputChoice, wx.GBPosition(0, 1), wx.GBSpan(1, 1), wx.ALL, 6)

		self.xmlButton = wx.Button(self, wx.ID_ANY, u"Generate XML Analyze Result", wx.Point(-1, -1), wx.Size(220, -1),
								   0)
		self.xmlButton.SetDefault()
		gbSizer1.Add(self.xmlButton, wx.GBPosition(2, 2), wx.GBSpan(1, 1), wx.ALL, 5)

		self.executeLog = wx.richtext.RichTextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
												   wx.Size(250, 450),
												   0 | wx.VSCROLL | wx.HSCROLL | wx.NO_BORDER | wx.WANTS_CHARS)
		gbSizer1.Add(self.executeLog, wx.GBPosition(1, 0), wx.GBSpan(1, 3), wx.EXPAND | wx.ALL, 5)

		self.pdfButton = wx.Button(self, wx.ID_ANY, u"Generate PDF Analyze Result", wx.Point(-1, -1), wx.Size(220, -1),
								   0)
		self.pdfButton.SetDefault()
		gbSizer1.Add(self.pdfButton, wx.GBPosition(3, 2), wx.GBSpan(1, 1), wx.ALL, 5)

		gbSizer1.Add((0, 310), wx.GBPosition(1, 3), wx.GBSpan(1, 1), wx.EXPAND, 5)

		self.inputPath = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*",
										   wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
		gbSizer1.Add(self.inputPath, wx.GBPosition(0, 2), wx.GBSpan(1, 2), wx.ALL, 1)

		fgSizer1.Add(gbSizer1, 1, wx.EXPAND, 5)

		self.SetSizer(fgSizer1)
		self.Layout()
		self.m_menubar1 = wx.MenuBar(0)
		self.m_menu1 = wx.Menu()
		self.menuManual = wx.MenuItem(self.m_menu1, wx.ID_ANY, u"Manual", wx.EmptyString, wx.ITEM_NORMAL)
		self.m_menu1.Append(self.menuManual)

		self.menuAbout = wx.MenuItem(self.m_menu1, wx.ID_ANY, u"About", wx.EmptyString, wx.ITEM_NORMAL)
		self.m_menu1.Append(self.menuAbout)

		self.m_menubar1.Append(self.m_menu1, u"Help")

		self.SetMenuBar(self.m_menubar1)

		self.Centre(wx.BOTH)

		# Connect Events
		self.xmlButton.Bind(wx.EVT_BUTTON, self.genXmlReport)
		self.pdfButton.Bind(wx.EVT_BUTTON, self.genPdfReport)

	def __del__(self):
		pass

	# Virtual event handlers, overide them in your derived class
	def genXmlReport(self, event):
		self.genCompareReport('xml')

	def genPdfReport(self, event):
		self.genCompareReport('pdf')

	def genCompareReport(self, format):
		pass



