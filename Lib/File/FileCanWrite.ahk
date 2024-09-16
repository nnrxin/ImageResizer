/* 函数:文件打开状态
 * param path 路径
 * return
 */
FileCanWrite(path)
{
	if !FileExist(path)
		return 0
	try f := FileOpen(path, "rw")
	catch
		return 0
	f.Close()
	return 1
}