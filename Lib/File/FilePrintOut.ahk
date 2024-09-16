/* 函数:打印文件(直接打印)
 * param fileFullPath 文件绝对路径
 * param copies 打印份数
 * param printerName 打印机名称(不检测是否正确)
 * returns result 成功返回1,失败返回0
 */
FilePrintOut(fileFullPath, copies := 1, printerName := "") {
	if !FileExist(fileFullPath)
		return
	SplitPath fileFullPath, , , &ext
	Switch ext
	{
	Case "xls", "xlsx":
		try {
			excel := ComObject("Excel.application")
			workbook := excel.Workbooks.Open(fileFullPath,, 1)
			workbook.PrintOut(,, copies,, printerName,,1)    ;逐份直接打印
			excel.Quit()
		} catch {
			try excel.Quit()
			return
		}
	Case "doc", "docx":
		try {
			word := ComObject("Word.application")
			if printerName {
				lastPrinterName := word.ActivePrinter
				word.ActivePrinter := printerName                ;设置打印机
				word.PrintOut(,,,,,,, copies,,,, 1, fileFullPath)    ;逐份直接打印
				word.ActivePrinter := lastPrinterName            ;设置打印机
			} else
				word.PrintOut(,,,,,,, copies,,,, 1, fileFullPath)    ;逐份直接打印
			word.Quit()
		} catch {
			try word.Quit()
			return
		}
	Default:
		return
	}
	return
}