/* 函数:从文件名中获取文件所在文件夹路径,路径不存在时启用default
 * param filePath 文件路径
 * param verify 是否验证
 * param default 默认返回值
 * returns OutDir 所在的文件夹路径
 */
Path_Dir(filePath, verify := true, default := "")
{
	SplitPath filePath, , &OutDir
	if verify and !DirExist(OutDir)
		return default
	else
		return OutDir
}


/* 函数:返回含后缀的文件名
 * param path 文件/文件夹的路径
 * returns OutFileName 含后缀的文件名
 */
Path_FileName(path)
{
	SplitPath Path, &OutFileName
    return OutFileName
}


/* 函数:返回不含后缀的文件名
 * param path 文件/文件夹的路径
 * returns OutExtension 不含后缀的文件名
 */
Path_FileNameNoExt(path)
{
	SplitPath path,,,, &OutNameNoExt
    return OutNameNoExt
}


/* 函数:返回文件后缀
 * param path 文件/文件夹的路径
 * returns OutExtension 返回文件后缀(类似 txt doc)
 */
Path_Extension(path)
{
	SplitPath path,,, &OutExtension
    return OutExtension
}


/* 函数:返回文件或文件夹的完整路径
 * param path 文件/文件夹的路径
 * returns fullPath 返回文件/文件夹的完整路径
 */
Path_Full(path)
{
    Loop Files, path, "FD"  ; 包括文件和目录.
        return A_LoopFileFullPath
	return path
}


/* 函数:文件重命名后缀后返回新路径
 * param path 文件/文件夹的路径
 * param newExt 新后缀
 * param verify 是否验证路径
 * returns newpath 返回新路径
 */
Path_RenameExt(path, newExt, verify := false)
{
	SplitPath path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	if !verify
		return OutDir "\" OutNameNoExt . (OutExtension ? "." newExt : "")
	else if DirExist(path)
		return path
	else
		return OutDir "\" OutNameNoExt . (OutExtension ? "." newExt : "")
}


/* 函数:文件重命名后返回新路径
 * param path 文件/文件夹的路径
 * param newName 新名字
 * param verify 是否验证路径
 * returns newpath 返回新路径
 */
Path_Rename(path, newName, verify := false)
{
	SplitPath path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	if !verify
		return OutDir "\" newName . (OutExtension ? "." OutExtension : "")
	else if !FileExist(path)
		return ""
	else
		return OutDir "\" newName . (OutExtension and !DirExist(path) ? "." OutExtension : "")
}


/* 函数:文件名后面增加字符
 * param path 文件/文件夹的路径
 * param appendStr 增加的字符
 * param verify 是否验证路径
 * returns newpath 返回新路径
 */
Path_NameAppend(path, appendStr := "", verify := false)
{
	SplitPath path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	if !verify
		return OutDir "\" OutNameNoExt . appendStr . (OutExtension ? "." OutExtension : "")
	else if !FileExist(path)
		return ""
	else if OutExtension and DirExist(path)    ;文件名中带 . 的特殊情况
		return OutDir "\" OutFileName . appendStr
	else
		return OutDir "\" OutNameNoExt . appendStr . (OutExtension ? "." OutExtension : "")
}


/* 函数:尝试获取简单的相对路径(不带..),不成功则返回原路径
 * param path 文件/文件夹的路径
 * param appendStr 增加的字符
 * param verify 是否验证路径
 * returns newpath 返回新路径
 */
Path_Relative(path, basePath)
{
	return (InStr(path, basePath) = 1) ? SubStr(path, StrLen(basePath) + 2) : path
}


/* 函数:获取A_Args中包含的路径
 * returns argPaths 路径数组
 * 局限:路径中不能含有连续两个以上空格
 */
Path_InArgs()
{
	argPaths := []
	argPath := ""
	subPath := ""
	for i, arg in A_Args  ; 对每个参数 (或拖放到脚本上的文件) 进行循环:
	{
		if !FileExist(arg)
		{
			subPath .= subPath ? " " arg : arg
			if FileExist(subPath)
				argPath := subPath
			else
				continue
		}
		else
			argPath := arg
		subPath := ""
		argPaths.push(argPath)
	}
	return argPaths
}


/* 函数:路径合法化(长度限制,非法字符去除)
 * param str 字符串
 * param limit 长度限制
 * returns path 路径字符串
 */
Path_Legalize(str, limit := 255)
{
	path := ""
	str := RegExReplace(str, '[/\*\?"<>\|`r`n]')   ;去除非法字符/*?"<>|   不包括\:
	SplitPath str, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	limit := OutExtension ? limit - StrLen(OutExtension) - 2 : limit - 1    ;过长路径将在末尾添加~
	len := 0
	Loop Parse, OutDir "\" OutNameNoExt
	{
		len += (Ord(A_LoopField) > 0xFF) ? 2 : 1
		if (len > limit)
		{
			path .= "~"
			break
		}
		path .= A_LoopField
	}
	return OutExtension ? path "." OutExtension : path
}