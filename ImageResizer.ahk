;=======================================================================================================================
; by:nnrxin
; email:nnrxin@163.com
;=======================================================================================================================

;基础参数设置
#Requires AutoHotkey v2.0
#NoTrayIcon               ;无托盘图标
#SingleInstance Ignore    ;不能双开
KeyHistory 0
ListLines 0
SendMode "Input"
SetWinDelay 0
SetControlDelay 0
ProcessSetPriority "H"


;必要的函数库
#Include <_BasicLibs_>
#Include <File\Path>
#Include <Image\ImagePut\ImagePut>

;APP基本信息
APP_NAME      := "IR"                    ; APP名称
APP_NAME_FULL := "ImageResizer"          ; APP全称
APP_NAME_CN   := "图片尺寸调整工具IR"     ; APP中文名称
APP_VERSION   := "v2.0.2"                 ; 当前版本


;APP保存信息(ini文件存储在同目录下)
INI := IniSaved(A_ScriptDir "\" APP_NAME "_config.ini")       ;创建配置ini类
;全局参数
G := {}

;=================================
;↓↓↓↓↓↓↓↓↓  MainGUI 构建 ↓↓↓↓↓↓↓↓↓
;=================================

;创建主GUI
MainGui := Gui("+Resize +MinSize700x500", APP_NAME_CN " " APP_VERSION)   ;GUI可修改尺寸
MainGui.Show("hide w10 h10")
MainGui.GetClientPos(,, &MainGuiWidth, &MainGuiHeight)
MainGui.MarginX := MainGui.MarginY := 0
MainGui.SetFont("s9", "微软雅黑")
;MainGui.BackColor := 0xCCE8CF   ;护眼蓝色
GroupWidth := 125 ; 右侧框架宽度

;增加Guitooltip
MainGui.Tips := GuiCtrlTips(MainGui)

;列表框
LV := MainGui.Add("ListView", "Section xm+5 ym+5 w" MainGuiWidth-GroupWidth-17 " h" MainGuiHeight-30 " AW AH", ["文件名","路径","大小","状态"])
LV.ModifyCol(3, "Right"), LV.ModifyCol(4, "Center")
;列表加载文件
filesInLV := Map()
LV.LoadFilesAndDirs := LV_LoadFilesAndDirs
LV_LoadFilesAndDirs(this, pathArray) {
	static exts := ["bmp","dib","rle","jpg","jpeg","jpe","jfif","gif","emf","wmf","tif","tiff","png","ico","heic","hif","webp","avif","avifs","pdf","svg"]
	this.Opt("-Redraw")
	files := []
	for _, path in pathArray {
		if DirExist(path) {
			Loop Files, path "\*.*", "FR"
				files.Push({path: A_LoopFileFullPath, midPath: Path_Relative(A_LoopFileFullPath, Path_Dir(path))})
			continue
		}
		files.Push({path: path, midPath: ""})
	}
	for i, file in files {
		if filesInLV.Has(file.path)
			continue
		SplitPath file.path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive
		if !exts.IndexOf(OutExtension)
			continue
		f := filesInLV[file.path] := {path: file.path, name: OutFileName, sizeKB: Format("{:.1f} KB", FileGetSize(file.path)/1024), status: "未处理", midPath: file.midPath}
		this.Add("Icon" this.LoadFileIcon(f.path), f.name, f.path, f.sizeKB, f.status)
	}
	this.AdjustColumnsWidth()
	this.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
}




;保存方式 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section x+7 ym w" GroupWidth " h100 AX", "保存方式")
MainGui.SetFont("cDefault norm", "微软雅黑")
RD1 := MainGui.Add("Radio", "xp+10 yp+22 AXP Group", "覆盖源文件")
RD1.Value := INI.Init(RD1, "save", "RD1", 0)
RD1.OnEvent("Click", RD_Click)
RD2 := MainGui.Add("Radio", "xp y+5 AXP", "保存为新文件")
RD2.Value := INI.Init(RD2, "save", "RD2", 1)
RD2.OnEvent("Click", RD_Click)
RD_Click(*) {
	if RD1.Value = 1 {
		DDLextension.Visible := false
		DDLextension2.Visible := true
	} else {
		DDLextension.Visible := true
		DDLextension2.Visible := false
	}
}

MainGui.Add("Text", "xs+10 yp+22 h26 w40 +0x200 AXP", "保存为:")
DDLextension := MainGui.Add("DDL", "x+0 yp w65 AXP", ["原格式",".bmp",".dib",".gif",".heic",".hif",".jpg",".jpeg",".jpe",".jfif",".png",".rle",".tif",".tiff"])
DDLextension.Value := INI.Init(DDLextension, "save", "extension", 1)
DDLextension2 := MainGui.Add("DDL", "xp yp wp AXP choose1 Disabled", ["原格式"]) ; 假的控件用于假装替换



;尺寸缩放 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section xs y+10 w" GroupWidth " h135 AXP", "尺寸缩放")
MainGui.SetFont("cDefault norm", "微软雅黑")

MainGui.Add("Text", "xs+10 yp+22 h26 w30 +0x200 AXP", "缩放:")
DDLdimensionMod := MainGui.Add("DDL", "x+0 yp w75 AXP", ["不缩放", "像素", "百分比"])
DDLdimensionMod.Value := INI.Init(DDLdimensionMod, "dimension", "dimensionMod", 1)
DDLdimensionMod.OnEvent("Change", DDLdimensionMod_Change)
DDLdimensionMod_Change(*) {
	i := DDLdimensionMod.Value
	EDwidth1.Visible := EDheight1.Visible := EDwidth2.Visible := EDheight2.Visible := EDwidth3.Visible := EDheight3.Visible := false
	EDwidth%i%.Visible := EDheight%i%.Visible := true
	TXdimUnit1.Value := TXdimUnit2.Value := (i = 2) ? "像素" : (i = 3) ? "%" : ""
	CBkeepAspectRatio.Enabled := (i = 2) ? true : false
}

MainGui.Add("Text", "xs+10 y+5 h22 w30 +0x200 AXP", "宽度:")
EDwidth1 := MainGui.Add("Edit", "x+0 yp hp w45 AXP Number Limit5 Center Disabled", "不变") ; 无动作
EDwidth2 := MainGui.Add("Edit", "xp yp hp wp AXP Number Limit5 Center") ; 像素
EDwidth2.Value := INI.Init(EDwidth2, "dimension", "width2")
MainGui.Tips.SetTip(EDwidth2, "留空时将根据纵横比自动计算")
EDwidth3 := MainGui.Add("Edit", "xp yp hp wp AXP Number Limit4 Center") ; 百分比
EDwidth3.Value := INI.Init(EDwidth3, "dimension", "width3")
EDwidth3.OnEvent("Change", (*) => EDheight3.Value := EDwidth3.Value)
TXdimUnit1 := MainGui.Add("Text", "x+0 yp hp w30 AXP +0x200 center", "像素")

MainGui.Add("Text", "xs+10 y+5 hp w30 +0x200 AXP", "高度:")
EDheight1 := MainGui.Add("Edit", "x+0 yp hp w45 AXP Number Limit5 Center Disabled", "不变") ; 无动作
EDheight2 := MainGui.Add("Edit", "xp yp hp wp AXP Number Limit5 Center") ; 像素
EDheight2.Value := INI.Init(EDheight2, "dimension", "heigth2")
MainGui.Tips.SetTip(EDheight2, "留空时将根据纵横比自动计算")
EDheight3 := MainGui.Add("Edit", "xp yp hp wp AXP Number Limit4 Center") ; 百分比
EDheight3.Value := INI.Init(EDheight3, "dimension", "heigth3")
EDheight3.OnEvent("Change", (*) => EDwidth3.Value := EDheight3.Value)
TXdimUnit2 := MainGui.Add("Text", "x+0 yp hp w30 AXP +0x200 center", "像素")

CBkeepAspectRatio := MainGui.Add("Checkbox", "xs+10 y+5 hp AXP", "保持图片不变形")
CBkeepAspectRatio.Value := INI.Init(CBkeepAspectRatio, "dimension", "keepAspectRatio", 1)
MainGui.Tips.SetTip(CBkeepAspectRatio, "必要时会裁剪图片,防止缩放后图片变形")



;画质调整 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section xs y+10 w" GroupWidth " h57 AXP", "画质调整")
MainGui.SetFont("cDefault norm", "微软雅黑")

MainGui.Add("Text", "xs+10 yp+22 h22 w55 +0x200 AXP", "指定画质:")
EDquailty := MainGui.Add("Edit", "x+0 yp hp w50 AXP Number Limit3 Center")
EDquailty.Value := INI.Init(EDquailty, "quality", "level", "")
EDquailty.OnEvent("Change", EDquailty_Change)
EDquailty_Change(*) {
	if EDquailty.Value and EDquailty.Value > 100 {
		EDquailty.Value := 100
		EDquailty.Focus()
	}
}
MainGui.Tips.SetTip(EDquailty, "JPEG质量等级`n范围:0 - 100`n推荐设置为不低于80`n留空时保持原图片质量等级")



;执行 Group
MainGui.SetFont("c0070DE bold", "微软雅黑")
MainGui.Add("GroupBox", "Section xs y+12 w" GroupWidth " h133 AXP", "执行")
MainGui.SetFont("cDefault norm", "微软雅黑")

BTclear := MainGui.Add("Button", "xs+10 yp+22 h27 w105 AXP", "移除所有项")
BTclear.OnEvent("Click", BTclear_Click)
BTclear_Click(thisCtrl, info) {
	LV.Opt("-Redraw")
	LV.Delete()
	filesInLV.Clear()
	LV.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
}

BTremoveFinished := MainGui.Add("Button", "xp y+5 hp wp AXP", "移除已完成")
BTremoveFinished.OnEvent("Click", BTremoveFinished_Click)
BTremoveFinished_Click(thisCtrl, info) {
	LV.Opt("-Redraw")
	deleteRows := []
	Loop LV.GetCount() {
		if LV.GetText(A_Index, 4) != "已完成"
			continue
		filesInLV.Delete(LV.GetText(A_Index, 2))
		deleteRows.Push(A_Index)
	}
	loop deleteRows.Length
		LV.Delete(deleteRows.Pop())
	LV.Opt("+Redraw")
	EnableBottons(LV.GetCount()) ; 控制按钮
}

BTstart := MainGui.Add("Button", "xp y+5 h40 wp AXP", "调整图片")
BTstart.OnEvent("Click", BTstart_Click)
BTstart_Click(thisCtrl, info) {
	;覆盖原文件时执行前提醒
	MainGui.Opt("+OwnDialogs")
	if RD1.Value and MsgBox("图片调整完将覆盖原文件,是否继续？",, 68) = "No"
		return
	EnableBottons(false) ; 禁用按钮
	;开始执行
	dirName := APP_NAME_FULL "_" A_Now
	loop LV.GetCount() {
		;确认目标文件名
	    file := filesInLV[LV.GetText(A_Index, 2)]
		if RD1.Value ; 覆盖原文件
			tarPath := file.path
		else { ; 另存为新文件
			tarPath := A_ScriptDir "\" dirName "\" file.midPath || file.name
			if DDLextension.Value != 1 
				tarPath := Path_RenameExt(tarPath, DDLextension.Text) ; 格式转换
		}
		;确认尺寸调整模式
		requireDim := (DDLdimensionMod.Value = 2) ? [EDwidth2.Value, EDheight2.Value] : (DDLdimensionMod.Value = 3) ? EDwidth3.Value / 100 : ""
		if ImageCropAndScale(file.path, tarPath, requireDim, CBkeepAspectRatio.Value, EDquailty.Value)
			file.status := "处理失败"
		else
			file.status := "已完成"
		LV.Modify(A_Index, "Vis Focus Col4", file.status) ; 可见 焦点 选中 列4修改
	}
	LV.AdjustColumnsWidth()
	EnableBottons(true) ; 启用按钮
}
/**
 * 将图片缩放到指定尺寸,缩放前会根据指定尺寸的宽高比裁剪(居中)图片
 * @wiki https://github.com/iseahound/ImagePut/wiki/Crop,-Scale,-&-Other-Flags#crop
 * @param srcPath 源图片路径
 * @param tarPath 目标图片路径
 * @param requireDim [w, h] 目标图片尺寸 
 * @param doCrop 执行裁剪
 * @param quality JPEG图片质量 0 - 100
 * @returns {Error} 失败时返回错误对象, 成功返回0
 */
ImageCropAndScale(srcPath, tarPath, requireDim, doCrop := true, quality := "") {
	if requireDim is Array {
		if !requireDim[1] && !requireDim[2] ; 尺寸都为空或0时, 不做缩放或裁剪处理
			image := {image: srcPath}
		else if !requireDim[1] ; 未指定宽度时, 缩放到指定宽度, 高度根据纵横比自动计算
			image := {image: srcPath, scale: ["auto", requireDim[2]]}
		else if !requireDim[2] ; 未指定高度时, 缩放到指定高度, 宽度根据纵横比自动计算
			image := {image: srcPath, scale: [requireDim[1], "auto"]}
		else { ; 宽高都指定时
			requireAspectRatio := requireDim[1] / requireDim[2]
			dim := ImagePut.Dimensions(srcPath), aspectRatio := dim[1] / dim[2]
			if !doCrop || aspectRatio = requireAspectRatio ; "不裁剪"或宽高比相同时, 按目标尺寸缩放
				image := {image: srcPath, scale: [requireDim[1], requireDim[2]]}
			else if aspectRatio > requireAspectRatio { ; 图片宽了,先裁剪到目标纵横比,再缩放到目标尺寸
				cropOnOneSide := Floor((dim[1] - dim[2] * requireAspectRatio) / 2) ; 单边多出来宽度
				crop := [-cropOnOneSide, 0, -cropOnOneSide, dim[2]] ; 剪裁crop: [x, y, w, h] / [-left, -top, -right, -bottom]
				image := {image: srcPath, crop: crop, scale: [requireDim[1], requireDim[2]]}
			} else { ; 图片窄了,先裁剪到目标纵横比,再缩放到目标尺寸
				cropOnOneSide := Floor((dim[2] - dim[1] / requireAspectRatio) / 2) ; 单边多出来高度
				crop := [0, -cropOnOneSide, dim[1], -cropOnOneSide]
				image := {image: srcPath, crop: crop, scale: [requireDim[1], requireDim[2]]}
			}
		}
	} else if IsNumber(requireDim)
		image := {image: srcPath, scale: requireDim}
	else
		image := {image: srcPath}
	
    try ImagePutFile(image, tarPath, quality)
	catch Error
	    return Error
	return 0
}
;启用/禁用按钮函数
EnableBottons(condition) {
	BTstart.Enabled := BTclear.Enabled := BTremoveFinished.Enabled := condition ? true : false
}




;状态栏
SB := MainGui.Add("StatusBar",, "")
SB.SetFont("bold")
SB.SetText("将图片文件或文件夹拖入窗口中")
;带自动清空的状态栏文字函数
MainGui_SB_SetTextWithAutoEmpty(newText, second := 10, partNumber := 1) {
	SB.SetTextWithAutoEmpty(newText, second, partNumber)
}


;GUI菜单
MainGui.OnEvent("ContextMenu", MainGui_ContextMenu)
MainGui_ContextMenu(GuiObj, GuiCtrlObj, Item, IsRightClick, X, Y) {
	;右键某控件上
	if IsRightClick and GuiCtrlObj and GuiCtrlObj.HasMethod("ContextMenu")
		GuiCtrlObj.ContextMenu(Item, X, Y)
}

;GUI文件拖放
MainGui.OnEvent("DropFiles", MainGui_DropFiles)
MainGui_DropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y) {
	MainGui.Opt("+Disabled")
	;主界面上的拖动
	Switch GuiCtrlObj {
		case LV:
			LV.LoadFilesAndDirs(FileArray)
	}

	MainGui.Opt("-Disabled")
}

;改变GUI尺寸时调整控件
MainGui.OnEvent("Size", MainGui_Size)
MainGui_Size(thisGui, MinMax, W, H) {
	LV.AdjustColumnsWidth()
}

;GUI关闭
MainGui.OnEvent("Close", MainGui_Close)
MainGui_Close(*) {
	ExitApp
}

;退出APP前运行
OnExit DoBeforeExit
DoBeforeExit(*) {
	MainGui.Hide()
	INI.SaveAll()            ;用户配置保存到ini文件
}


;Gui初始化
LV.LoadFilesAndDirs(Path_InArgs()) ; 拖拽文件到程序图标上启动
DDLdimensionMod_Change()           ; 尺寸缩放相关设置初始化
RD_Click()                         ; 保存方式相关设置初始化
EnableBottons(LV.GetCount())       ; 按钮初始化


;GUI显示
dpiRate := 96 / A_ScreenDPI
MainGui.Show("hide Center w" SysGet(16) * dpiRate " h" SysGet(17) * dpiRate)
guiSizeRate := 0.2 * dpiRate
MainGui.Show("Center w" SysGet(16) * guiSizeRate " h" SysGet(17) * guiSizeRate)

;=========================
return    ;自动运行段结束 |
;=========================