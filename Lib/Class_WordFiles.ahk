;操作word文件的类，docx、doc
class WordFiles
{
	;创建
	__New() {
		;创建com
		this.document := ""
		this.documents := Map()
		try this.app := ComObject("Word.application")
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
		;检测path路径的document是否存在
		if this.SetDocument(path)
			return this.document
		;新建
		try this.document := this.app.documents.Open(path,, readOnly)
		catch
			return
		return this.documents[path] := this.document
	}

	;设置document,返回this.document
	SetDocument(path := "") {
		if !path
			this.document := this.document ? this.document : this.app.Documents.Count ? this.app.ActiveDocument : ""
		else if this.documents.Has(path) {
			try this.documents[path].FullName
			catch {
				this.documents.Delete(path)
				this.document := ""
			} else
				this.document := this.documents[path]
		}
		else
			this.document := ""
		return this.document
	}

	;关闭文档,默认不保存
	Close(path := "", SaveChanges := 0) {
		if !this.SetDocument(path)
			return
		this.document.Close(SaveChanges)
		for path, document in this.documents {
			if document = this.document {
				this.documents.Delete(path)
				break
			}
		}
	}

	;关闭全部文档,默认不保存
	CloseAll(SaveChanges := 0) {
		for path, document in this.documents {
			try document.Close(SaveChanges)
			this.documents.Delete(path)
		}
	}

	;保存文档
	Save(path := "") {
		if !this.SetDocument(path)
			return
		try this.document.Save()
	}

	;保存全部文档
	SaveAll() {
		for k, document in this.documents
			try document.Save()
	}

	;打印
	PrintOut(copies := 1, printerName := "", path := "") {
		if !this.SetDocument(path)
			return
		if printerName {
			lastPrinterName := this.app.ActivePrinter
			this.app.ActivePrinter := printerName           ;设置打印机
			this.document.PrintOut(,,,,,,, copies,,,, 1)    ;逐份打印
			this.app.ActivePrinter := lastPrinterName       ;设置打印机
		}
		else
			this.document.PrintOut(,,,,,,, copies,,,, 1)    ;逐份打印
	}

	;在单元格下方插入新行
	InsertRowsBelow(rowI := 1, tableI := 1, path := "") {
		this.document.Tables.Item(tableI).Rows(rowI).Select()
		this.app.Selection.InsertRowsBelow()
	}

	;修改单元格
	CellText(text, rowI := 1, columnI := 1, replace := false, tableI := 1, path := "") {
		if replace
			this.document.Tables.Item(tableI).Cell(rowI, columnI).Range.Delete()
		this.document.Tables.Item(tableI).Cell(rowI, columnI).Range.InsertAfter(text)
	}
}