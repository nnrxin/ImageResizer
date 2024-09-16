/*
	Name: Gui_StatusBar.ahk
	Version 0.1 (2024-08-28)
	Created: 2024-08-28
	Author: nnrxin

	Description:
	为原生状态栏对象Gui.StatusBar增加了一些有用的方法

    Methods:
    SB.SetTextWithAutoEmpty(newText, second := 10, partNumber := 1)  => 设定状态栏文字后在设定时间内清空文字 
*/


/**
 * function : 设定状态栏文字后在设定时间内清空文字
 * @param newText ; 新的文本
 * @param second ; 时间
 * @param partNumber ; 状态栏号
 * @returns
 */
Gui.StatusBar.Prototype.SetTextWithAutoEmpty := _SB_SetTextWithAutoEmpty
_SB_SetTextWithAutoEmpty(this, newText, second := 10, partNumber := 1) {
	SetTimer SB_OnUpdate%partNumber%, 0
	this.SetText(newText, partNumber)
	SetTimer SB_OnUpdate%partNumber%, - second * 1000

	SB_OnUpdate1() => this.SetText("", 1)
	SB_OnUpdate2() => this.SetText("", 2)
	SB_OnUpdate3() => this.SetText("", 3)
	SB_OnUpdate4() => this.SetText("", 4)
	SB_OnUpdate5() => this.SetText("", 5)
	SB_OnUpdate6() => this.SetText("", 6)
	SB_OnUpdate7() => this.SetText("", 7)
	SB_OnUpdate8() => this.SetText("", 8)
	SB_OnUpdate9() => this.SetText("", 9)
	SB_OnUpdate10() => this.SetText("", 10)
}