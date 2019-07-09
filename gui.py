# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Sep 12 2010)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
#from wx.tools.img2py import img2py
import visa

# use img2py transfer image to py
#ret = img2py('C:\\Users\\user\\Desktop\\hce_rrap\\visa.png', 'C:\\Users\\user\\Desktop\\hce_rrap\\visa.py', append=True)

class MyFrame1(wx.Frame):

	def __init__(self, parent):
		wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"RRAP_HCE_182a_R0", pos=wx.DefaultPosition,
						  size=wx.Size(-1, -1), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

		self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)

		fgSizer1 = wx.FlexGridSizer(2, 2, 0, 0)
		fgSizer1.SetFlexibleDirection(wx.BOTH)
		fgSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

		gbSizer1 = wx.GridBagSizer(0, 0)
		gbSizer1.SetFlexibleDirection(wx.BOTH)
		gbSizer1.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

		self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"Test Tool Vendor", wx.DefaultPosition, wx.DefaultSize, 0)
		self.m_staticText1.Wrap(-1)
		gbSizer1.Add(self.m_staticText1, wx.GBPosition(1, 0), wx.GBSpan(1, 1), wx.ALL, 8)

		self.m_staticText11 = wx.StaticText(self, wx.ID_ANY, u"Test Tool Report path", wx.DefaultPosition,
											wx.DefaultSize, 0)
		self.m_staticText11.Wrap(-1)
		gbSizer1.Add(self.m_staticText11, wx.GBPosition(2, 0), wx.GBSpan(1, 1), wx.ALL, 8)

		inputChoiceChoices = [u"ICCSolutions", u"Galitt", u"UL"]
		self.inputChoice = wx.Choice(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, inputChoiceChoices, 0)
		self.inputChoice.SetSelection(1)
		gbSizer1.Add(self.inputChoice, wx.GBPosition(1, 2), wx.GBSpan(1, 1), wx.ALL, 6)

		self.xmlButton = wx.Button(self, wx.ID_ANY, u"Generate XML Analyze Result", wx.Point(-1, -1), wx.Size(220, -1),
								   0)
		self.xmlButton.SetDefault()
		gbSizer1.Add(self.xmlButton, wx.GBPosition(4, 2), wx.GBSpan(1, 1), wx.ALL, 5)

		self.pdfButton = wx.Button(self, wx.ID_ANY, u"Generate PDF Analyze Result", wx.Point(-1, -1), wx.Size(220, -1),
								   0)
		self.pdfButton.SetDefault()
		gbSizer1.Add(self.pdfButton, wx.GBPosition(5, 2), wx.GBSpan(1, 1), wx.ALL, 5)

		self.executeLog = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(500, 400), style=wx.TE_MULTILINE )
		gbSizer1.Add(self.executeLog, wx.GBPosition(3, 0), wx.GBSpan(1, 5), wx.ALL, 5)

		gbSizer1.Add((0, 310), wx.GBPosition(3, 5), wx.GBSpan(1, 1), wx.EXPAND, 5)

		self.inputPath = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select the test report", u"*.*",
										   wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
		gbSizer1.Add(self.inputPath, wx.GBPosition(2, 2), wx.GBSpan(1, 2), 0, 1)

		self.m_bitmap3 = wx.StaticBitmap(self, wx.ID_ANY, visa.visa.getBitmap(),
										 wx.DefaultPosition, wx.DefaultSize, 0)
		gbSizer1.Add(self.m_bitmap3, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 5)

		# self.m_staticText6 = wx.StaticText(self, wx.ID_ANY, u"Developed by", wx.DefaultPosition, wx.DefaultSize,
		# 								   0)
		# self.m_staticText6.Wrap(-1)
		# self.m_staticText6.SetFont(wx.Font(wx.NORMAL_FONT.GetPointSize(), 70, 90, 92, False, wx.EmptyString))
		# self.m_staticText6.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_GRAYTEXT))
		# gbSizer1.Add(self.m_staticText6, wx.GBPosition(4, 0), wx.GBSpan(1, 1), wx.ALIGN_BOTTOM|wx.ALL, 5)
		#
		# #image = wx.Image(u"applus.png", wx.BITMAP_TYPE_ANY)
		# image = applus.applus.getBitmap().ConvertToImage()
		# image.Rescale(image.GetWidth()/2.5, image.GetHeight()/2.5)
		# self.m_bitmap4 = wx.StaticBitmap(self, wx.ID_ANY, image.ConvertToBitmap(),
		# 								 wx.DefaultPosition, wx.DefaultSize, 0)
		#
		# gbSizer1.Add((0, 5), wx.GBPosition(6, 6), wx.GBSpan(1, 1), wx.EXPAND, 5)
		#
		# gbSizer1.Add(self.m_bitmap4, wx.GBPosition(5, 0), wx.GBSpan(1, 1), wx.ALL, 5)

		fgSizer1.Add(gbSizer1, 1, wx.EXPAND, 5)

		self.SetSizer(fgSizer1)
		self.Layout()
		fgSizer1.Fit(self)

		self.Centre(wx.BOTH)

		# Connect Events
		self.xmlButton.Bind(wx.EVT_BUTTON, self.genXmlReport)
		self.pdfButton.Bind(wx.EVT_BUTTON, self.genPdfReport)

	def __del__(self):
		pass

	# Virtual event handlers, override them in your derived class
	def genXmlReport(self, event):
		self.genCompareReport('xml')

	def genPdfReport(self, event):
		self.genCompareReport('pdf')

	def genCompareReport(self, format):
		pass

