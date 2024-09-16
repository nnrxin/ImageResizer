/**
 * 返回当前系统CPU线程数
 * 参考：http://msdn.com/library/ms724509(vs.85,en-us)
 * 已知问题：
 * @returns {Integer | Number} 
 */
CPUThreads() {
    try {
        static hModule := DllCall("LoadLibrary", "Str", "ntdll.dll", "Ptr")
        buf := Buffer(0, 0) ,size := 0
        , DllCall("ntdll\NtQuerySystemInformation", "Int", 0x8, "Ptr", buf, "UInt", 0, "UInt*", &size)
        , buf := Buffer(size, 0)
        if DllCall("ntdll\NtQuerySystemInformation", "Int", 0x8, "Ptr", buf, "UInt", size, "UInt*", 0) != 0
          return 1
        CPU_COUNT := size // 48
    }
    catch {
        return 1
    }
    return CPU_COUNT
}