; 文件复制增强,目标文件夹不存在时创建
; 成功返回真,失败返回假
FileCopyEx(srcPath, tarPath, overwrite := false) {
	;源文件不存在时跳过
	if not (AttributeString := FileExist(srcPath)) or InStr(AttributeString, "F")
		return false
	;目标文件夹不存在时创建
	SplitPath tarPath, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive
	tarDir := OutExtension ? OutDir : tarPath
	if not DirExist(tarDir) {
		try
			DirCreate(tarDir)
		catch
			return false
	}
	;修改时间相同时跳过
	if FileExist(tarPath) and FileGetTime(srcPath) = FileGetTime(tarPath)
		return true    ;实际效果为真
	;复制文件
	try
		FileCopy(srcPath, tarPath, overwrite)
	catch
		return false
	;成功则返回真
	return true
}

