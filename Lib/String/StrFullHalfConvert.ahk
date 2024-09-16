/* 函数:全角半角字符互相转换函数
 * param str 字符串
 * param toFull 真时为全角,假时为半角
 * returns newStr 转化后的字符串
 */
StrFullHalfConvert(str, toFull := false)
{
	newStr := ""
	Loop parse, str
	{
		n := Ord(A_LoopField)
		if toFull
			newStr .= (n = 32) ? Chr(12288) : (n >= 33 and n <= 126) ? Chr(n + 65248) : A_LoopField
		else
			newStr .= (n = 12288) ? Chr(32) : (n >= 65281 and n <= 65374) ? Chr(n - 65248) : A_LoopField
	}
	return newStr
}

/*
全角字符unicode编码从 65281 ~ 65374 （十六进制 0xFF01 ~ 0xFF5E）
半角字符unicode编码从 33 ~ 126 （十六进制 0x21~ 0x7E）
空格比较特殊,全角为 12288（0x3000）,半角为 32 （0x20）
而且除空格外,全角/半角按unicode编码排序在顺序上是对应的
所以可以直接通过用+-法来处理非空格数据,对空格单独处理
*/