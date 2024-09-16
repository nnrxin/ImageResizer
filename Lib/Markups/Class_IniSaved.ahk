; =================================================================================
; Function: InI easy save for Gui
; requires AHK version : 2.0+
; =================================================================================
class IniSaved {
	__New(filePath) {
		SplitPath filePath, &name, &dir, &ext, &name_no_ext, &drive
		if dir and !DirExist(dir)
			DirCreate dir
		if !FileExist(filePath)
		{
			FileAppend "", filePath
		}
		this.filePath := filePath
		this.items := Map()
	}

	Init(obj, section := "", key := "", default := "", objKey := "Value") {
		this.items[obj] := {section:section, key:key, objKey:objKey}
		return IniRead(this.filePath, section, key, default)
	}

	Read(section := "", key := "", default := "") {
		return IniRead(this.filePath, section, key, default)
	}

	Save(obj) {
		item := this.items[obj]
		IniWrite(obj.%item.objKey%, this.filePath, item.section, item.key)
	}

	SaveAll() {
		for obj, item in this.items {
			IniWrite(obj.%item.objKey%, this.filePath, item.section, item.key)
		}
	}
}