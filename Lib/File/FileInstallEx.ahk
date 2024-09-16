/* 函数:安装文件增强
 * param Source 源路径
 * param Dest 目标路径
 * param Overwrite 覆盖
 * return
 */
FileInstallEx(Source, Dest, Overwrite := false)
{
	if Overwrite or !FileExist(Dest)
	{
		FileInstall Source, Dest, Overwrite
		return
	}
}