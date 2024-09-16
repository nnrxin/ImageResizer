/*
	Name: Gui_Edit.ahk
	Version 0.1 (2024-08-28)
	Created: 2024-08-28
	Author: nnrxin

	Description:
	为原生下拉表对象Gui.Edit增加了一些有用的方法

    Methods:
    ED.Append(appendStr := "", timeStamp := false)  => 更新DDL控件内容和当前选定
*/


/**
 * function: 增加文本到edit控件末尾，且下拉条拉倒末尾
 * @param str 字符串
 * @param autoTrim 修剪多余空格
 * @returns str 首行字符串
 */
Gui.Edit.Prototype.Append := _ED_Append
_ED_Append(this, appendStr := "", timeStamp := false)
{
	if timeStamp
		appendStr := A_Hour ":" A_Min ":" A_Sec " " appendStr
	SendMessage(0x00B1, -2, -1, this)                         ;光标移至末尾
	SendMessage(0x00C2, False, StrPtr(appendStr), this)       ;edit新增
	;SendMessage(0x00B7, 0, 0, this)                          ;一些老系统中可能需要
}