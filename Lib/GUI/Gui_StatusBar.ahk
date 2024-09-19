/*
	Name: Gui_StatusBar.ahk
	Version 0.1 (2024-08-28)
	Created: 2024-08-28
	Author: nnrxin

	Description:
	为原生状态栏对象Gui.StatusBar增加了一些有用的方法

    Methods:
    SB.SetTextWithAutoEmpty(newText, second := 10, partNumber := 1)  => 设定状态栏文字后在设定时间内清空文字 
	SB.AttachProgress(segment := 1, Options := "", Text := "") => 绑定进度条到状态栏
*/


/**
 * function : 设定状态栏文字后在设定时间内清空文字
 * @param newText ; 新的文本
 * @param second ; 时间
 * @param segment ; 状态栏位置
 * @returns
 */
Gui.StatusBar.Prototype.SetTextWithAutoEmpty := _SB_SetTextWithAutoEmpty
_SB_SetTextWithAutoEmpty(this, newText, second := 10, segment := 1) {
	SetTimer SB_OnUpdate%segment%, 0
	this.SetText(newText, segment)
	SetTimer SB_OnUpdate%segment%, - second * 1000

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


/**
 * function : 绑定进度条到状态栏
 * @param segment ; 状态栏位置
 * @param Options ; 进度条控件的选项
 * @param Text ; 进度条控件的文本
 * @returns Prog 返回一个progress对象
 * @remarks 1.需要在Gui.show()之后调用
 *          2.Gui尺寸调整后,需要调用Prog.Size()来调整进度条位置和大小
 */
Gui.StatusBar.Prototype.AttachProgress := _SB_AttachProgress
_SB_AttachProgress(thisSB, segment := 1, Options := "", Text := "") {
    thisGui := thisSB.Gui
    Prog := thisGui.Add("Progress", Options, Text)
    SB_Parts := SendMessage(0x406, 0, 0, , thisSB.Hwnd) ;获取状态栏的栏数
    If segment > SB_Parts
        throw ValueError("FAIL: Wrong Segment Count")
    Prog.segment := segment
    Prog.StatusBar := thisSB
    Prog.Size := Size
    Prog.Size()
    return Prog
    ;从新定位尺寸
    static Size(thisCtrl) {
        static RECT := Buffer(16, 0)     ; RECT = 4*4 Bytes / 4 Byte <=> Int
        SendMessage(0x40a, thisCtrl.segment-1, RECT, , thisCtrl.StatusBar.Hwnd) ; 获取第seg个状态栏的尺寸
        n1 := NumGet(RECT, 0, "int"), n2 := NumGet(RECT, 4, "int") ,n3 := NumGet(RECT, 8, "int"), n4 := NumGet(RECT, 12, "int")
        ControlGetPos(&xb, &yb,,, thisCtrl.StatusBar.Hwnd)
        ControlMove(xb+n1, yb+n2, n3-n1, n4-n2, thisCtrl.hwnd) ; 设置进度条位置和尺寸
    }
}