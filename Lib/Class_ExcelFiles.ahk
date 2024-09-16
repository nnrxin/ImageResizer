;操作excel文件的类，xlsx、xls
class ExcelFiles
{
	;创建
	__New() {
		;创建com
		this.workbooks := Map()
		this.workbook := ""
		this.worksheet := ""
		try this.app := ComObject("Excel.application")
		try	this.app.Visible := False ; 不可见
		;try this.app.DisplayAlerts = 0 ; 警告和消息的处理的方式
		;wdAlertsAll	-1	显示所有的消息框和警告，并将错误返回宏。
		;wdAlertsMessageBox	-2	仅显示消息框，同时捕获错误并将其返回宏。
		;wdAlertsNone	0	不显示任何警告或消息框。 如果运行宏时遇到消息框，则选择默认值并继续运行宏。
		return this
	}

	;销毁
	__Delete() => this.Quit()

	;退出
	Quit() {
		this.CloseAll()
		try this.app.Quit()
	}

	;打开文件
	Open(path, readOnly := false) {
		;检测path路径的workbook是否存在
		if this.SetWorkbook(path)
			return this.workbook
		;新建
		try this.workbook := this.app.Workbooks.Open(path,, readOnly)
		catch
			return
		return this.workbooks[path] := this.workbook
	}

	;设置Workbook,返回this.workbook
	SetWorkbook(path := "") {
		if !path
			this.workbook := this.workbook ? this.workbook : this.app.Workbooks.Count ? this.app.ActiveWorkbook : ""
		else if this.workbooks.Has(path) {
			try this.workbooks[path].FullName
			catch {
				this.workbooks.Delete(path)
				this.workbook := ""
			} else
				this.workbook := this.workbooks[path]
		}
		else
			this.workbook := ""
		return this.workbook
	}

	;关闭文档,默认不保存
	Close(path := "", SaveChanges := 0) {
		if !this.SetWorkbook(path)
			return
		this.workbook.Close(SaveChanges)
		for path, workbook in this.workbooks {
			if workbook = this.workbook {
				this.workbooks.Delete(path)
				break
			}
		}
	}

	;关闭全部文档,默认不保存
	CloseAll(SaveChanges := 0) {
		for path, workbook in this.workbooks {
			try workbook.Close(SaveChanges)
			this.workbooks.Delete(path)
		}
	}

	;保存文档
	Save(path := "") {
		if !this.SetWorkbook(path)
			return
		try this.workbook.Save()
	}

	;保存全部文档
	SaveAll() {
		for k, workbook in this.workbooks
			try workbook.Save()
	}

	;设置Worksheet,返回this.Worksheet
	SetWorksheet(worksheetName := 1, path := "") {
		if !this.SetWorkbook(path)
			this.worksheet := ""
		else if worksheetName is Integer {
			if worksheetName > 0 and worksheetName <= this.workbook.Worksheets.Count
				this.worksheet := this.workbook.Worksheets.Item(worksheetName)
			else
				this.worksheet := this.workbook.ActiveSheet
		} else if worksheetName {
			this.worksheet := this.workbook.ActiveSheet
			loop this.workbook.worksheets.Count {
				if worksheetName = this.workbook.Worksheets.Item(A_Index).Name {
					this.worksheet := this.workbook.Worksheets.Item(A_Index)
					break
				}
			}
		}
		else
			this.worksheet := this.worksheet ? this.worksheet : this.workbook.ActiveSheet
		return this.worksheet
	}

	;返回单元格对象
	Cells(rowI := 1, columnI := 1) => this.worksheet.Cells(Integer(rowI), Integer(columnI))

	;返回行对象
	Rows(rowI := 1) => this.worksheet.Rows(Integer(rowI))

	/*
	;插入图片到单元格
	AddPictureToCell(picPath, row := 1, column := 1, pictureName := "")
	{
		if !FileExist(picPath)
			return -1
		c := this.worksheet.Cells(row, column)
		p := this.worksheet.Shapes.AddPicture(picPath, 0, 1, 0, 0, -1, -1)  ;插入图片
		this.PictureMatchToTheCell(p, c, "WA", "Safe_" pictureName)   ;图片固定到指定单元格
		return p
	}

	;导出照片
	ExportPicture(picName, picPath, scale := 1)
	{
		pic := this.worksheet.Shapes(picName)
		if (scale != 1)   ;需要缩放
		{
			pic.LockAspectRatio := 1
			picWidthBefore := pic.Width
			pic.Width := picWidthBefore * scale
		}
		pic.Copy
		c := this.worksheet.ChartObjects.Add(0, 0, pic.Width, pic.Height)   ;创建空白表
		if (scale != 1)   ;缩放复原
		{
			pic.Width := picWidthBefore
		}
		c.Chart.Paste
		c.Chart.Export(picPath)   ;导出
		c.Delete
	}

	;所有照片全部压缩
	AllPictureCompress(scale := 1, SBSetProgress := false, GuiWidth := 600)
	{
		shapeNames := []
		Loop this.worksheet.Shapes.Count
		{
			pic := this.worksheet.Shapes(A_index)
			if (pic.Type = 13)   ;只添加照片
				shapeNames.Push(pic.Name)
		}
		if SBSetProgress   ;进度条状态栏
		{
			;oneTenthWidth := (GuiWidth - 0)//10
			;SB_SetParts(oneTenthWidth * 4, oneTenthWidth, oneTenthWidth * 5)
			;MaxI := shapeNames.Length
			;SB_SetProgress(0, 3, "show Range0-" MaxI)
		}
		Loop shapeNames.Length
		{
			this.PictureCompress(shapeNames[A_index], scale)
			if SBSetProgress   ;进度条状态栏
			{
				;SB_SetText("图片压缩中..." shapeNames[A_index], 1)
				;SB_SetText(A_index "/" MaxI, 2)
				;SB_SetProgress(A_index, 3)
			}
		}
		if SBSetProgress   ;进度条状态栏
		{
			;SB_SetProgress(0, 3, "hide")
			;SB_SetParts()
		}
	}

	;压缩照片
	PictureCompress(picName, scale := 1)
	{
		pic := this.worksheet.Shapes(picName)
		picNameBefore := pic.Name
		picLeftBefore := pic.Left
		picTopBefore := pic.Top
		picWidthBefore := pic.Width
		picHeightBefore := pic.Height
		picPlacementBefore := pic.Placement
		if (scale != 1)   ;需要缩放
		{
			pic.LockAspectRatio := 0
			pic.Width := picWidthBefore * scale
			pic.Height := picHeightBefore * scale
		}
		pic.Copy
		this.worksheet.PasteSpecial(1)
		newPic := this.excelApp.Selection   ;获取焦点对象(不是pic对象,但是能重命名)
		pic.Delete
		newPic.Name := picNameBefore
		newPic := this.worksheet.Shapes(newPic.Name)   ;从新获取准确的pic对象
		newPic.Left := picLeftBefore
		newPic.Top := picTopBefore
		newPic.Width := picWidthBefore
		newPic.LockAspectRatio := 0
		newPic.Height := picHeightBefore
		newPic.LockAspectRatio := 1
		newPic.Placement := picPlacementBefore

	}

	;调整sheet中所有的图片,使其匹配其所在单元格
	AllPictureMatchToCell(resize := "WA", rename := "SameWithPicSafe", mod := "weak", SBSetProgress := false, GuiWidth := 600)
	{
		shapeI := []
		Loop this.worksheet.Shapes.Count
		{
			pic := this.worksheet.Shapes(A_index)
			if (pic.Type = 13)   ;只添加照片
			{
				shapeI.Push(A_index)
				pic.Placement := 3    ;图片自由活动
			}
		}
		if SBSetProgress   ;进度条状态栏
		{
			;oneTenthWidth := (GuiWidth - 0)//10
			;SB_SetParts(oneTenthWidth * 4, oneTenthWidth, oneTenthWidth * 5)
			;MaxI := shapeI.Length
			;SB_SetProgress(0, 3, "show Range0-" MaxI)
		}
		Loop shapeI.Length
		{
			shape := this.worksheet.Shapes(shapeI[A_Index])
			this.PictureMatchToCell(shape, resize, rename, mod)
			if SBSetProgress   ;进度条状态栏
			{
				;SB_SetText("图片整理中..." shape.Name, 1)
				;SB_SetText(A_index "/" MaxI, 2)
				;SB_SetProgress(A_index, 3)
			}
		}
		if SBSetProgress   ;进度条状态栏
		{
			;SB_SetProgress(0, 3, "hide")
			;SB_SetParts()
		}
	}

	;调整一张图片,使其匹配其所在单元格
	PictureMatchToCell(shape, resize := "WA", rename := "SameWithPicSafe", mod := "weak")
	{
		range := this.PictureLocation(shape)
		cell := this.worksheet.Cells(range.TN, range.LN)
		if (range.WN = 1 and range.HN = 1) or (mod = "power")
		{
			this.PictureMatchToTheCell(shape, cell, resize, rename)
		}
	}

	;调整一张图片,使其匹配到固定的单元格
	PictureMatchToTheCell(shape, cell, resize := "WA", rename := "")
	{
		shape.Left := cell.Left + 1
		shape.Top := cell.Top + 1
		;尺寸
		shape.Placement := 3    ;图片自由活动
		Switch resize
		{
		Case "W":
			shape.LockAspectRatio := 1
			shape.Width := cell.Width - 2
		Case "WA":
			shape.LockAspectRatio := 1
			shape.Width := cell.Width - 2
			if (cell.Height < shape.Height + 2)
				this.worksheet.Rows(cell.Row).RowHeight := shape.Height + 2   ;调整行高适应图片大小
		Case "H":
			shape.LockAspectRatio := 1
			shape.Height := cell.Height - 2
		Case "HA":
			shape.LockAspectRatio := 1
			shape.Height := cell.Height - 2
			if (cell.Width < shape.Width + 2)
				this.worksheet.Columns(cell.Column).ColumnWidth := shape.Width + 2   ;调整列宽适应图片大小
		Case "WH", "HW":
			shape.LockAspectRatio := 0
			shape.Width := cell.Width - 2
			shape.Height := cell.Height - 2
		Default:
		}
		shape.Placement := 1    ;大小位置随单元格变化
		;重命名
		Switch rename
		{
		Case "SameWithPic":
			cell.Value := shape.Name
		Case "SameWithPicSafe":
			if (cell.Value = "")
				cell.Value := shape.Name
		Case "SameWithCell":
			if (cell.Value != "")
				shape.Name := cell.Value
		Case "":
		Default:
			if (SubStr(rename, 1, 5) = "Safe_")   ;安全模式
			{
				if (cell.Value = "") and (rename != "Safe_")
				{
					shape.Name := SubStr(rename, 6)
					cell.Value := shape.Name
				}
			}
			else
			{
				shape.Name := rename
				cell.Value := shape.Name
			}
		}
	}

	;根据照片L T W H信息,返回其是否位于某个单元格内
	PictureLocation(shape)
	{
		l := shape.Left         ;左
		t := shape.Top          ;上
		r := l + shape.Width    ;右
		b := t + shape.Height   ;下
		;先搜索左右
		Loop 10000
		{
			if (this.worksheet.Cells(1, A_index).Left > l)   ;确定了左边界
			{
				columnLN := A_index - 1
				rangeL := this.worksheet.Cells(1, columnLN).Left
				loop 10000
				{
					if (this.worksheet.Cells(1, columnLN + A_index).Left > r)   ;确定了右边界
					{
						columnWN := A_index
						rangeW := this.worksheet.Cells(1, columnLN + A_index).Left - rangeL
						break 2
					}
				}
				break
			}
		}
		;再搜索上下
		Loop 10000
		{
			if (this.worksheet.Cells(A_index, 1).Top > t)   ;确定了上边界
			{
				rowTN := A_index - 1
				rangeT := this.worksheet.Cells(rowTN, 1).Top
				loop 10000
				{
					if (this.worksheet.Cells(rowTN + A_index, 1).Top > b)   ;确定了下边界
					{
						rowHN := A_index
						rangeH := this.worksheet.Cells(rowTN + A_index, 1).Top - rangeT
						break 2
					}
				}
				break
			}
		}
		return {LN:columnLN, WN:columnWN, TN:rowTN, HN:rowHN, Left:rangeL, Width:rangeW, Top:rangeT, Heigth:rangeH}
	}
	*/
}