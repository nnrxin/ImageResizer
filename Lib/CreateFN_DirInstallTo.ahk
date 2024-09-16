; 安装源文件夹内所有文件到目标文件内
CreateFN_DirInstallTo(srcDir, functionName := "DirInstallTo")
{
	if not InStr(FileExist(srcDir), "D")
		return
	pathLength := StrLen(srcDir)
	files := []
	dirs := Map()
	;获取源路径内全部文件信息
	Loop Files, srcDir "\*", "FDR"
	{
		subPath := SubStr(A_LoopFilePath, pathLength + 1)
		if InStr(FileExist(A_LoopFilePath), "D")
			dirs[subPath] := true
		else
			files.Push(subPath)
	}
	;先只保留最长路径的目录
	match := Map()
	for k1 in dirs
	{
		for k2 in dirs
		{
			if (k1 != k2) and Instr("\/" k1, "\/" k2)    ;k1包含k2时标记k2
				match[k2] := true
		}
	}
	for k in match
		dirs.Delete(k)

	;开始
	t := ';安装文件函数`r`n' functionName '(targetPath, overwrite := 0)`r`n{`r`n`ttry`r`n`t{`r`n'

	;先创建文件夹
	t .= '`t`t;创建文件夹`r`n'
	for dirPath in dirs
		t .= '`t`tDirCreate(targetPath "' dirPath '")`r`n'

	;再安装文件
	t .= '`t`t;安装文件`r`n'
	for i, subFilePath in files
		t .= '`t`tFileInstall("' srcDir . subFilePath '", targetPath "' subFilePath '", overwrite)`r`n'

	;结尾
	t .= "`t}`r`n}"

	return t
}



