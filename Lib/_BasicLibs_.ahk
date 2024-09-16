; nnrxin库的基础库集合
; _BasicLibs_.ahk

;字符串/数组/关联数组的增强
#Include Extension\Array.ahk ; 数组增强
#Include Extension\Map.ahk ; 关联数组增强
#Include Extension\String.ahk ; 字符串增强

;GUI及控件的增强
#Include GUI\GuiCtrlTips.ahk ; Gui控件的tooltip提示
#Include GUI\Gui_Resizer.ahk ; Gui尺寸调整
#Include GUI\Gui_StatusBar.ahk ; 状态栏增强
#Include GUI\Gui_DDL.ahk ; 下拉框增强
#Include GUI\Gui_ListView.ahk ; 图标视图增强
#Include GUI\Gui_Edit.ahk ; 编辑框增强
#Include GUI\Class_LVICE_XXS.ahk ; LV单元格内编辑
#Include GUI\Class_LV_Colors.ahk ; LV单元格颜色

;标记语言
#Include Markups\Class_IniSaved.ahk ; ini简单存储类

;字符串相关
#Include String\StringFuctions.ahk ; 字符串增强函数集合
#Include String\GetTimeYYYYMMDD.ahk ; 获取固定格式的日期