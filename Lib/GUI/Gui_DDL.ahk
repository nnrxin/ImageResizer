/*
	Name: Gui_DDL.ahk
	Version 0.1 (2024-08-28)
	Created: 2024-08-28
	Author: nnrxin

	Description:
	为原生下拉表对象Gui.DDL增加了一些有用的方法

    Methods:
    DDL.Update(items := [""], textChoosen := "")  => 更新DDL控件内容和当前选定
*/


/**
 * function : 更新DDL控件内容和当前选定
 * @param items ; 可以为Array或者Map()
 * @param textChoosen ; 选择一项
 * @returns
 */
Gui.DDL.Prototype.Update := _DDL_Update
_DDL_Update(this, items := [""], textChoosen := "") {
	this.Delete()
	for i, item in items {
		this.Add([item])
		if item = textChoosen
			this.Choose(i)
	}
    if this.Value = 0
        try this.Value := 1
}