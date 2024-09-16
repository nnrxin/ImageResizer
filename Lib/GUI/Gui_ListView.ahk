/*
	Name: Gui_ListView.ahk
	Version 0.1 (2024-08-28)
	Created: 2024-08-28
	Author: nnrxin

	Description:
	为原生表格视图对象Gui.ListView增加了一些有用的方法

    Methods:
    LV.GetColTitleIndex(title)                                 => 获取LV中标题所在第一个列数,未找到时返回0
	LV.AdjustColumnsWidth(hiddenColumns := "", maxW := 160)    => 根据LV的总宽度自动调各列宽度
	LV.ModifyByColumnName(rowI, columnName, value)             => 修改指定标题下某行的内容, 成功不返回值, 失败返回值
	LV.LoadFileIcon(rowI, filePath, hIcons := "")              => 将文件显示的ICON设置到LV的图像列表中并返回图像序号

/**


/**
 * function: 获取LV中标题所在第一个列数,未找到时返回0
 * @param title ; 标题名称
 * @returns ; 返回所在列数
 */
Gui.ListView.Prototype.GetColTitleIndex := _LV_GetColTitleIndex
_LV_GetColTitleIndex(this, title) {
	Loop this.GetCount("Col")
		if title = this.GetText(0, A_index)
			return A_index
	return 0
}


/**
 * function: 根据LV的总宽度自动调各列宽度
 * @param hiddenColumns ; 隐藏列序号数组
 * @param maxW  ; 最大列宽
 */
Gui.ListView.Prototype.AdjustColumnsWidth := _LV_AdjustColumnsWidth
_LV_AdjustColumnsWidth(this, hiddenColumns := "", maxW := 160) {
	;确定隐藏列
	isHidden := Map()
	if IsObject(hiddenColumns) {
		for i, ColumnI in hiddenColumns
			isHidden[ColumnI] := true
	}
	;确定LV总列宽
	this.GetPos(,, &w)
	LVW := w - 21    ;21带侧标边条  4不带侧边条
	;获取各列内容适配的宽度
	colWs := []
	colWTotal := 0
	overWidthColumn := Map()
	overWidthTotal := 0
	Loop this.GetCount("Column") {
		if isHidden.Has(A_Index)
			this.ModifyCol(A_Index, 0)
		else
			this.ModifyCol(A_Index, "AutoHdr")
		colW := SendMessage(0x101D, A_Index - 1, 0, this)    ; 0x101D 是 LVM_GETCOLUMNWIDTH.
		colWs.Push(colW)
		colWTotal += colW
		if colW > maxW
			overWidthTotal += overWidthColumn[A_Index] := colW - maxW
	}
	;计算出最佳列宽
	if colWTotal > LVW {
		if colWTotal - overWidthTotal >= LVW {
			for columnI, v in overWidthColumn
				this.ModifyCol(columnI, maxW)
		} else {
			spaceW := LVW - (colWTotal - overWidthTotal)
			for columnI, v in overWidthColumn
				this.ModifyCol(columnI, maxW + spaceW * v // overWidthTotal)
		}
	}
}


/**
 * function: 修改指定标题下某行的内容, 成功不返回值, 失败返回值
 * @param rowI ; 行号
 * @param columnName ; 列名称
 * @param value ; 要修改的值
 * @return ; 成功不返回值, 失败返回值
 */
Gui.ListView.Prototype.ModifyByColumnName := _LV_ModifyByColumnName
_LV_ModifyByColumnName(this, rowI, columnName, value) {
	;确定列标题序号
	Loop this.GetCount("Column") {
		if this.GetText(0 , A_Index) = columnName {
			columnI := A_index
			break
		}
	}
	;未找到指定名称列标题时返回-1
	if !IsSet(columnI)
		return -1
	;修改LV内容,失败返回-1
	try this.Modify(rowI, "Col" columnI, value)
	catch
		return 1
}


/**
 * function: 将文件显示的ICON设置到LV的图像列表中并返回图像序号
 * @param filePath ; 文件路径
 * @param hIcons ; 一个用来储存图标句柄的数组或者空值, 图标数组序号与返回的图像序号一致
 * @return 图像序号
 */
Gui.ListView.Prototype.LoadFileIcon := _LV_LoadFileIcon
_LV_LoadFileIcon(this, filePath, hIcons := "")
{
	; 创建图像列表, 这样 ListView 才可以显示图标:
	static ImageListID1 := IL_Create(10)
	static ImageListID2 := IL_Create(10, 10, true)  ; 搭配小图标列表的大图标列表.
	; 缓存图标:
	static IconMap := Map()
	; 计算 SHFILEINFO 结构需要的缓冲大小:
	static sfi_size := A_PtrSize + 688
	static sfi := Buffer(sfi_size)
	; 已设置的LV.Hwnd列表:
	static LVs := Map()
	; 将图像列表到设置LV:
	if !LVs.Has(this.Hwnd) {
		this.SetImageList(ImageListID1, 1)    ;小图标
		this.SetImageList(ImageListID2, 0)    ;大图标
		LVs[this.Hwnd] := true
	}
    ; 建立唯一的扩展 ID 以避免变量名中的非法字符, 例如破折号. 这种使用唯一 ID 的方法也会执行地更好, 因为在数组中查找项目不需要进行搜索循环.
    SplitPath(filePath,,, &FileExt)  ; 获取文件扩展名.
    if FileExt ~= "^(EXE|ICO|ANI|CUR)$" {
        ExtID := FileExt  ; 特殊 ID 作为占位符.
        IconNumber := 0  ; 将其标记为未找到, 以便这些类型可以有一个唯一的图标.
    } else {  ; 其他的扩展名/文件类型, 计算它们的唯一 ID.
        ExtID := 0  ; 进行初始化来处理比其他更短的扩展名.
        Loop 7 {    ; 限制扩展名为 7 个字符, 这样之后计算的结果才能存放到 64 位值.
            ExtChar := SubStr(FileExt, A_Index, 1)
            if not ExtChar  ; 没有更多字符了.
                break
            ExtID := ExtID | (Ord(ExtChar) << (8 * (A_Index - 1))) ; 把每个字符与不同的位置进行运算来得到唯一ID
        }
        ; 检查此文件扩展名的图标是否已经在图像列表中, 可以避免多次调用并极大提高性能, 尤其对于包含数以百计文件的文件夹而言:
        IconNumber := IconMap.Has(ExtID) ? IconMap[ExtID] : 0
    }
    if not IconNumber {  ; 此扩展名还没有相应的图标, 所以进行加载.
        ; 取与此文件扩展名关联的高质量小图标:    0x101 是 SHGFI_ICON+SHGFI_SMALLICON
        if not DllCall("Shell32\SHGetFileInfoW", "Str", filePath, "Uint", 0, "Ptr", sfi, "UInt", sfi_size, "UInt", 0x101)
            IconNumber := 9999999  ; 把它设置到范围外来显示空图标.
        else { ; 成功加载图标.
            ; 从结构中提取 hIcon 成员:
            hIcon := NumGet(sfi, 0, "Ptr")
            ; 直接添加 HICON 到小图标和大图标列表.下面加上 1 来把返回的索引从基于零转换到基于一:
            IconNumber := DllCall("ImageList_ReplaceIcon", "Ptr", ImageListID1, "Int", -1, "Ptr", hIcon) + 1
            DllCall("ImageList_ReplaceIcon", "Ptr", ImageListID2, "Int", -1, "Ptr", hIcon)
			; 是否需要储存个备份供外部使用,不需要就销毁缓存图标
			if IsObject(hIcons)
				hIcons.Push(hIcon)
			else
            	DllCall("DestroyIcon", "Ptr", hIcon) ; 现在已经把它复制到图像列表, 所以应销毁原来的缓存图标来节省内存并提升加载性能
            IconMap[ExtID] := IconNumber
        }
    }
    return IconNumber
}


/**
 * function: 把2维数组加载到LV里
 * @param Array2D {array} ; 2维数组
 * @param ColumnTitles {array} ; 列标题
 * @param FormatStrs {map}; 每列的格式
 * @return
 */
Gui.ListView.Prototype.LoadArray2D := _LV_LoadArray2D
_LV_LoadArray2D(this, Array2D, ColumnTitles := "", FormatStrs := "") {
	this.Delete()
	while this.GetCount("Column")
		this.DeleteCol(1)
	loop Array2D[1].Length
		this.InsertCol(999, , ColumnTitles ? ColumnTitles[A_Index] : "T" A_Index) ; 插入列到末尾
	for i, row in Array2D {
		if FormatStrs
			for k, FormatStr in FormatStrs
				row[k] := Format(FormatStr, row[k])
		this.Add("", row*)
	}
}


/**
 * function: 把2维数组加载到LV里
 * @document https://learn.microsoft.com/zh-cn/office/client-developer/access/desktop-database-reference/recordset-object-ado-reference
 * @param rs {object} ; 数据库查询的记录
 * @param ColumnTitles {array} ; 列标题
 * @param FormatStrs {map}; 每列的格式
 * @return
 */
Gui.ListView.Prototype.LoadADORecordset := _LV_LoadADORecordset
_LV_LoadADORecordset(this, rs, FormatStrs := "") {
	this.Delete()
	while this.GetCount("Column")
		this.DeleteCol(1)
	FieldIndex := Map()
	loop rs.Fields.Count {
		this.InsertCol(999,, FieldName := rs.Fields.Item(A_Index-1).Name)
		FieldIndex[FieldName] := A_Index
	}	
	while !rs.EOF {
		row := []
		loop rs.Fields.Count
			row.Push(rs.Fields.Item(A_Index-1).Value)
		if FormatStrs
			for k, FormatStr in FormatStrs
				row[FieldIndex[k]] := Format(FormatStr, row[FieldIndex[k]])
		this.Add("", row*)
		rs.MoveNext
	}
}

