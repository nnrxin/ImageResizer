/*
	Name: StringFuctions.ahk
	Version 0.1 (2024-08-31)
	Created: 2024-08-31
	Author: nnrxin

	Description:
	一些有用的字符串函数集合

	StrReplaceKeys(str, aMap, isPath := false)                              => 替换字符串中字符[key]为item[key]的值,然后进行一些处理
	StrOr(strs*)                                                            => 返回输入字符串里第一个非空字符串
    Str2Path(str, limit := 255, illegalReplace := "")                       => 字符串转换成路径
	StrLeft(str, Needle, Occurrence:=1, CaseSense:=false, StartingPos:=1)   => 根据字符截取字符串前半段
	StrIf(condition, strIfTrue, strIfFalse := "")                           => 根据条件真假返回字符串
	StrIfInStr(Haystack, Needle, strIfTrue, strIfFalse := "")               => 根据字符串中是否含有对应字符返回字符串
	StrFirstLine(str, autoTrim := true)                                     => 获取字符串的首行
	StrExtractLines(str, en := true, doTrim := true, ratio := 0.7)          => 从字符串中提取出指定行(中文或英文)
*/


/**
 * function:替换字符串中字符[key]为item[key]的值,然后进行一些处理
 * @param str 字符串
 * @param aMap 关联数组
 * @param isPath 是否路径
 * @returns str 第一个非空字符串
 */
StrReplaceKeys(str, aMap, isPath := false)
{
	for i, match in RegExMatchAll(str, "U)\[.*]")
	{
		key := Trim(match[0], "[]")
		replaceText := aMap.Has(key) ? aMap[key] : ""
		replaceText := isPath ? RegExReplace(replaceText, '[\\/:\*\?"<>\|]') : replaceText    ;文件名则替换掉\/:*?"<>|
		str := StrReplace(str, match[0], replaceText)
	}
	str := isPath ? RegExReplace(str, "(?<=.)\\{2,}", "\") : str    ;替换文件名中的 \\ 为 \

	;通过函数进行修改 %func(str,,1,,2)%
	for i, match in RegExMatchAll(str, "sU)%[^%]+\(.*\)%")
	{
		p := InStr(match[0],"(")
		str := StrReplace(str, match[0], %SubStr(match[0],2,p-2)%(StrSplit(SubStr(match[0],p+1,-2), ",,")*))
	}

	str := isPath ? RegExReplace(str, '[`r`n]') : str    ;文件名则替换掉回车
	return str

	/**
	 * Returns all RegExMatch results in an array: [RegExMatchInfo1, RegExMatchInfo2, ...]
	 * param Haystack The string whose content is searched.
	 * param NeedleRegEx The RegEx pattern to search for.
	 * param StartingPosition If StartingPos is omitted, it defaults to 1 (the beginning of Haystack).
	 * returns {Array}
	 */
	RegExMatchAll(Haystack, NeedleRegEx, StartingPosition := 1) {
		out := []
		While StartingPosition := RegExMatch(Haystack, NeedleRegEx, &OutputVar, StartingPosition) {
			out.Push(OutputVar), StartingPosition += OutputVar[0] ? StrLen(OutputVar[0]) : 1
		}
		return out
	}
}


/**
 * function:返回输入字符串里第一个非空字符串
 * @param str 字符串(可变参数)
 * @returns str 第一个非空字符串
 */
StrOr(strs*) {
	for i, str in strs {
		if str != ""
			return str
	}
}


/**
 * function: 字符串转换成路径
 * @param str 字符串
 * @param limit 长度限制
 * @param illegalReplace 非法文件名替换
 * @returns path 路径字符串
 */
Str2Path(str, limit := 255, illegalReplace := "") {
	path := ""
	str := RegExReplace(str, '[/\*\?"<>\|`r`n]', illegalReplace)   ;非法字符串替换掉/*?"<>|   不包括\:
	SplitPath str, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt
	limit := OutExtension ? limit - StrLen(OutExtension) - 2 : limit - 1    ;过长路径将在末尾添加~
	len := 0
	Loop Parse, OutDir "\" OutNameNoExt {
		len += (Ord(A_LoopField) > 0xFF) ? 2 : 1
		if (len > limit) {
			path .= "~"
			break
		}
		path .= A_LoopField
	}
	return OutExtension ? path "." OutExtension : path
}


/**
 * function:根据字符截取字符串前半段
 * @param str 字符串
 * @param Needle 要搜索的字符串
 * @param CaseSense 下列值之一(如果省略, 默认为 0): "On" 或 1(True): 搜索区分大小写. "Off" 或 0(False): 字母 A-Z 被视为与其小写字母相同.
 * @param StartingPos 省略 StartingPos 来搜索整个字符串. 否则, 指定开始搜索的位置, 其中 1 是第一个字符, 2 是第二个字符, 以此类推. 负值从 Haystack 的末尾开始计算, 所以 -1 是最后一个字符, -2 倒数第二个, 以此类推.
 * @param Occurrence 如果省略 Occurrence, 它默认为 Haystack 中 Needle 的第一个匹配. 如果 StartingPos 为负数, 搜索将从右向左进行; 否则将从左向右进行.
 * @returns str 前半部分字符串
 */
StrLeft(str, Needle, Occurrence := 1, CaseSense := false, StartingPos := 1) {
	if str and (p := InStr(str, Needle, CaseSense, StartingPos, Occurrence))
		return SubStr(str, 1, p-1)
	else
		return str
}


/**
 * function: 根据条件真假返回字符串
 * @param condition 条件
 * @param strIfTrue 如果为真返回的
 * @param strIfFalse 如果为假返回的
 * @returns
 */
StrIf(condition, strIfTrue, strIfFalse := "")
{
	if condition
		return strIfTrue
	else
		return strIfFalse
}


/**
 * function: 根据字符串中是否含有对应字符返回字符串
 * @param Haystack 搜索的字符
 * @param Needle 包含的字符
 * @param strIfTrue 如果为真返回的
 * @param strIfFalse 如果为假返回的
 * @returns
 */
StrIfInStr(Haystack, Needle, strIfTrue, strIfFalse := "")
{
	if InStr(Haystack, Needle)
		return strIfTrue
	else
		return strIfFalse
}


/**
 * function: 获取字符串的首行
 * @param str 字符串
 * @param autoTrim 修剪多余空格
 * @returns str 首行字符串
 */
StrFirstLine(str, autoTrim := true)
{
	Loop parse, str, "`n", "`r"  ; 在 `r 之前指定 `n, 这样可以同时支持对 Windows 和 Unix 文件的解析.
		return autoTrim ? Trim(A_LoopField) : A_LoopField
}


/**
 * function: 从字符串中提取出指定行(中文或英文)
 * @param str 字符串
 * @param en 真时提取英文,假时提取中文
 * @param t 是否去除多余的制表符和空格
 * @param ratio 英文所占比率
 * @returns outStr 输出字符串
 */
StrExtractLines(str, en := true, doTrim := true, ratio := 0.7)
{
	outStr := ""
	Loop parse, str, "`n", "`r"  ;在 `r 之前指定 `n, 这样可以同时支持对 Windows 和 Unix 文件的解析.
	{
		rowStr := doTrim ? Trim(A_LoopField) : A_LoopField
		if not maxCount := StrLen(rowStr)
		{
			outStr .= doTrim ? "" : "`n"
			continue
		}
		enCount := 0
		maxCount := 0.000001
		Loop parse, rowStr
		{
			n := Ord(A_LoopField)
			enCount += (n >= 56 and n <= 89 or n >= 97 and n <= 122) ? 1 : 0
			maxCount += (n >= 56 and n <= 89 or n >= 97 and n <= 122 or n > 127) ? 1 : 0
		}
		outStr .= (en and enCount/maxCount>ratio or !en and enCount/maxCount<ratio) ? rowStr "`n" : ""
	}
	return SubStr(outStr,1,-1) 
}