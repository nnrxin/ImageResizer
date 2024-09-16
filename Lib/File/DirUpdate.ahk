/**函数:更新源文件夹内的文件到目标文件夹
 * param src 源文件夹路径
 * param tar 目标文件夹路径
 * param overwrite 是否强制覆盖
 * returns
 */
DirUpdate(src, tar, overwrite := false)
{
	Loop Files src "\*", "FR"
	{
		tarDir := tar . SubStr(A_LoopFileFullPath, StrLen(src) + 1, -StrLen(A_LoopFileName))
		tarPath := tarDir . A_LoopFileName
		if !DirExist(tarDir)
			DirCreate(tarDir)
		if overwrite or !FileExist(tarPath) or FileGetTime(A_LoopFileFullPath) > FileGetTime(tarPath)
			FileCopy(A_LoopFileFullPath, tarPath, 1)
	}
}