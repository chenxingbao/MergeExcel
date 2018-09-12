#-*- coding:utf-8 -*-
import wx
import xlrd
import xlwt
import wx.grid as gridlib
import os

class MainFrame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self, None, -1, "Excel Merge", size=(960, 540))
		# 创建菜单
		self.CreateMenuBar()
		self.mainSizer = wx.BoxSizer(wx.VERTICAL)
		self.panel = wx.Panel(self)
		self.mainSizer.Add(self.panel, proportion = 1, flag = wx.ALL|wx.EXPAND, border = 1)
		self.SetSizer(self.mainSizer)
		#self.cols = []
		self.chooseCols = []
		self.data =[]
		self.paths = []
		self.Centre()
		
	def CreateBody(self, cols, paths):
		vbox = wx.BoxSizer(wx.VERTICAL)
		
		self.sb = wx.StaticBox(self.panel, label = "choose need fields:")
		boxsizer = wx.StaticBoxSizer(self.sb, wx.VERTICAL)
		
		for col in cols:
			boxsizer.Add(wx.CheckBox(self.panel, label = col), flag = wx.ALL|wx.EXPAND, border = 8)
		self.Bind(wx.EVT_CHECKBOX,self.onChecked) 
		vbox.Add(boxsizer, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 10)
		
		btn = wx.Button(self.panel, label = "Select")
		# print(self.chooseCols)
		btn.Bind(wx.EVT_BUTTON, self.showTable)
		vbox.Add(btn, flag = wx.ALL, border = 10)
		
		self.grid = wx.grid.Grid(self.panel, -1)
		# 创建表格100行，10列
		# self.grid.CreateGrid(100,len(cols))
		

		vbox.Add(self.grid, proportion = 1, flag = wx.EXPAND|wx.ALL, border = 10)
		# vbox.AddGrowableRow(2)
		
		self.panel.SetSizerAndFit(vbox)
		self.panel.Refresh()
		'''
		menuBar = wx.MenuBar()
		menu = wx.Menu()
		self.menu_open = menu.Append(-1, u'打开(O)\tCtrl+O')
		self.Bind(wx.EVT_MENU, self.onOpen, self.menu_open)
		self.menu_import = menu.Append(-1, u'导出(I)\tCtrl+I')
		self.Bind(wx.EVT_MENU, self.onImport, self.menu_import)
		menu.AppendSeparator()
		self.menu_quit = menu.Append(-1, u'退出(Q)\tCtrl+Q') #快捷键
		self.Bind(wx.EVT_MENU, self.onQuit, self.menu_quit)
		menuBar.Append(menu, u'文件(F)')
		self.SetMenuBar(menuBar)
		'''
	# 菜单数据
	def menuData(self):
		return [("&File",
						(("Open\tCtrl+O","Open Excel Files",self.onOpen),
						 ("Import\tCtrl+I","Import Merged Excel File",self.onImport),
						 ("","",""),            #分隔线
						 ("Quit\tCtrl+Q","Quit",self.onQuit),
						)
				 ),
				 ("&About",
						(("About Me","About Me",self.aboutInfo),
						)
				 )
				]
	# 创建菜单栏
	def CreateMenuBar(self):
		menuBar = wx.MenuBar()
		for eachMenuData in self.menuData():
			menuLabel = eachMenuData[0]
			menuItems = eachMenuData[1]
			menuBar.Append(self.CreateMenu(menuItems),menuLabel)
		self.SetMenuBar(menuBar)
	# 创建菜单
	def CreateMenu(self, menuItem):
		menu = wx.Menu()
		for eachItem in menuItem:
			if len(eachItem) == 2:
				label = menuItem[0]
				subMenu = self.CreateMenu(menuItem[1])
				menu.AppendMenu(wx.NewId(), label, subMenu)
			else:
				# print(eachItem)
				self.CreateMenuItem(menu,*eachItem)
		return menu
	# 创建菜单项
	def CreateMenuItem(self, menu, label, status, handler, kind = wx.ITEM_NORMAL):
		if not label:
			menu.AppendSeparator()
			return
		menuItem = menu.Append(-1, label, status, kind)
		self.Bind(wx.EVT_MENU, handler, menuItem)
	# 打开文件
	def onOpen(self, e):
		print("onOpen")
		self.reSize()
		filesFilter = "Excel files(*.xls;*.xlsx,*.xl*)|*.xl*"
		fd = wx.FileDialog(self, "Open Excel Files ...", wildcard = filesFilter, style = wx.FD_MULTIPLE)
		openResult = fd.ShowModal()
		if openResult != wx.ID_OK:
			return
		self.paths = fd.GetPaths()
		# print(paths)
		fd.Destroy()
		cols = self.getCols(self.paths)
		self.CreateBody(cols, self.paths)
	# 导出Excel
	def onImport(self, e):
		print("onImport")
		filesFilter = "Excel files(*.xls)|*.xls|" "Excel files(*.xlsx)|*.xlsx" 
		fd = wx.FileDialog(self, "Save Excel Files ...", wildcard = filesFilter, style = wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT)
		if fd.ShowModal() == wx.ID_OK:
			self.saveFile(fd.GetPath(), fd.GetFilename(), self.data, self.chooseCols)
		else:
			wx.MessageBox("Save File failed...", "Error", wx.OK |wx.ICON_ERROR)
		fd.Destroy()
		
	def aboutInfo(self, e):
		wx.MessageBox("About Me ...", "Info", wx.OK)
	# 退出
	def onQuit(self, e):
		self.Close();
	# 读取文件
	def getCols(self, paths):
		# print (paths)
		cols = []
		if paths:
			for file in paths:
				data = xlrd.open_workbook(file)
				sheet_names = data.sheet_names()
				print(sheet_names)
				for sheet_name in sheet_names:
					sheet = data.sheet_by_name(sheet_name)
					print(sheet.row_values(0))
					cols = cols + sheet.row_values(0)
			print(cols)
		tmp = list(set(cols))
		return tmp
	# 保存文件
	def saveFile(self, path, filename, data, fields):
		print(data)
		print(fields)
		print(path)
		# print(filename)
		if path:
			wtFile = xlwt.Workbook()
			outTable = wtFile.add_sheet(filename)
			i = 0
			for field in fields:
				outTable.write(0, i, field)
				i = i + 1
			# i = 0
			j = 1
			for row in data:
				for rl in range(len(row)):
					outTable.write(j, rl, row[rl])
				j = j + 1
			filename = path
			print(filename)
			wtFile.save(filename)
	# 复选框事件
	def onChecked(self, e):
		cb = e.GetEventObject()
		if cb.IsChecked():
			self.chooseCols.append(cb.GetLabel())
		else:
			self.chooseCols.remove(cb.GetLabel())
	#显示表格数据
	def showTable(self,event):
		self.data = []
		rowcn = -1

		for c in range(len(self.chooseCols)):
			self.grid.SetColLabelValue(c, self.chooseCols[c])
		if self.paths:
			for file in self.paths:
				data = xlrd.open_workbook(file)
				sheet_names = data.sheet_names()
				# print(sheet_names)
				for sheet_name in sheet_names:
					sheet = data.sheet_by_name(sheet_name)
					rows = sheet.row_values(0)
					for row in sheet.get_rows():
						# print(row)
						tmp = []
						if rowcn > -1:
							for c0 in range(len(self.chooseCols)):
								tmp = tmp + [row[rows.index(self.chooseCols[c0])].value]
							self.data.extend([tmp])
						rowcn = rowcn + 1
		table = LineupTable(self.data, self.chooseCols)
		self.grid.SetTable(table, True)
		self.grid.AutoSize()
		self.grid.Refresh()
		self.Refresh()
		# print(self.data)
	# 重置窗体
	def reSize(self):
		# self.panel.Destroy()
		self.sb.Destroy()
		# self.panel = wx.Panel(self)
		# self.mainSizer.Add(self.panel, proportion = 1, flag = wx.ALL|wx.EXPAND, border = 1)
		self.chooseCols = []
		self.data =[]
		self.paths = []
		# self.grid.ForceRefresh()
		# self.grid.AutoSize()

class LineupTable(wx.grid.GridTableBase):
    def __init__(self,data,fields):
        wx.grid.GridTableBase.__init__(self)
        self.data=data
        self.fields=fields
##        self.dataTypes = [wx.grid.GRID_VALUE_STRING,
##                          wx.grid.GRID_VALUE_STRING,
##                          #gridlib.GRID_VALUE_CHOICE + ':only in a million years!,wish list,minor,normal,major,critical',
##                          #gridlib.GRID_VALUE_NUMBER + ':1,5',
##                          #gridlib.GRID_VALUE_CHOICE + ':all,MSW,GTK,other',
##                          #gridlib.GRID_VALUE_BOOL,
##                          #gridlib.GRID_VALUE_BOOL,
##                          #gridlib.GRID_VALUE_BOOL,
##                          wx.grid.GRID_VALUE_FLOAT + ':6,2',
##                          ]

        #---Grid cell attributes

        self.odd = wx.grid.GridCellAttr()
        self.odd.SetBackgroundColour("grey")
        self.odd.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))

        self.even = wx.grid.GridCellAttr()
        self.even.SetBackgroundColour("white")
        self.even.SetFont(wx.Font(8, wx.SWISS, wx.NORMAL, wx.BOLD))

    #---Mandatory constructors for grid

    def GetNumberRows(self):
       # if len(self.data)<10:
           # rowcounts=10
        #else:
            #rowcounts=len(self.data)
        return len(self.data)

    def GetNumberCols(self):
        return len(self.fields)

    def GetColLabelValue(self, col):
        return self.fields[col]

    def IsEmptyCell(self, row, col):
        if self.data[row][col] == "" or self.data[row][col] is None:
            return True
        else:
            return False

#    def GetValue(self, row, col):
#        value = self.data[row][col]
#        if value is not None:
#            return value
#        else:
#            return ''
#
#    def SetValue(self, row, col, value):
#        #print col
#        def innerSetValue(row, col, value):
#            try:
#                self.data[row][col] = value
#            except IndexError:
#                # add a new row
#                self.data.append([''] * self.GetNumberCols())
#                innerSetValue(row, col, value)
#
#                # tell the grid we've added a row
#                msg = gridlib.GridTableMessage(self,            # The table
#                        gridlib.GRIDTABLE_NOTIFY_ROWS_APPENDED, # what we did to it
#                        1                                       # how many
#                        )
#
#                self.GetView().ProcessTableMessage(msg)
#        innerSetValue(row, col, value)
        
    def GetValue(self, row, col):
        #id = self.fields[col]
        return self.data[row][col]
    
    def SetValue(self, row, col, value):
        #id = self.fields[col]
        self.data[row][col] = value
    
    #--------------------------------------------------
    # Some optional methods
    
    # Called when the grid needs to display column labels
#    def GetColLabelValue(self, col):
#        #id = self.fields[col]
#        return self.fields[col][0]
   

    def GetAttr(self, row, col, kind):
        attr = [self.even, self.odd][row % 2]
        attr.IncRef()
        return attr

    def SortColumn(self, col,ascordesc):
        """
        col -> sort the data based on the column indexed by col
        """
        name = self.fields[col]
        _data = []

        for row in self.data:
            #print row
            #rowname, entry = row
            
            _data.append((row[col], row))

        _data.sort(reverse=ascordesc)
        self.data = []

        for sortvalue, row in _data:
            self.data.append(row)

    def AppendRow(self, row):#增加行
        #print 'append'
        entry = []

        for name in self.fields:
            entry.append('A')

        self.data.append(tuple(entry ))
        return True

    def MoveColumn(self,frm,to):
        grid = self.GetView()

        if grid:
            # Move the identifiers
            old = self.fields[frm]
            del self.fields[frm]

            if to > frm:
                self.fields.insert(to-1,old)
            else:
                self.fields.insert(to,old)
            
            print self.fields

            # Notify the grid
            grid.BeginBatch()
           
            msg = wx.grid.GridTableMessage(
                    self, wx.grid.GRIDTABLE_NOTIFY_COLS_INSERTED, to, 1
                    )
            grid.ProcessTableMessage(msg)

            msg = wx.grid.GridTableMessage(
                    self, wx.grid.GRIDTABLE_NOTIFY_COLS_DELETED, frm, 1
                    )
            grid.ProcessTableMessage(msg)
            
            grid.EndBatch()
           

    # Move the row
    def MoveRow(self,frm,to):
        grid = self.GetView()

        if grid:
            # Move the rowLabels and data rows
            oldLabel = self.rowLabels[frm]
            oldData = self.data[frm]
            del self.rowLabels[frm]
            del self.data[frm]

            if to > frm:
                self.rowLabels.insert(to-1,oldLabel)
                self.data.insert(to-1,oldData)
            else:
                self.rowLabels.insert(to,oldLabel)
                self.data.insert(to,oldData)

            # Notify the grid
            grid.BeginBatch()

            msg = wx.grid.GridTableMessage(
                    self, wx.grid.GRIDTABLE_NOTIFY_ROWS_INSERTED, to, 1
                    )
            grid.ProcessTableMessage(msg)

            msg = wx.grid.GridTableMessage(
                    self, wx.grid.GRIDTABLE_NOTIFY_ROWS_DELETED, frm, 1
                    )
            grid.ProcessTableMessage(msg)
    
            grid.EndBatch()
		
def main():
#设置了主窗口的初始大小960x540 800x450 640x360
	root = wx.App()
	frame = MainFrame()
	frame.Show(True)
	root.MainLoop()
 
 
if __name__ == "__main__":
	main()

