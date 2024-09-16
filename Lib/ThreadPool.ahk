/************************************************************************
 * @description 线程池类
 * @author nnrxin
 * @date 2024-09-06
 * @version 0.0.0
 * 无锁版本
 ***********************************************************************/
class ThreadPool {
    __New(maxThreads) {
        this.maxThreads := maxThreads
        this.tasks := []
        this.runningThreads := Map()
    }

    __Delete() {
        this.tasks.Length := 0
        this.StopAllTasks()
    }

    /**
     * function AddTask 添加任务到线程池
     * @param task.taskScript      任务脚本
     * @param task.taskFuncName    执行函数名称
     * @param task.taskFuncParams  执行函数参数
     * @param task.StartFunc       任务开始的回调函数 => StartFunc(thd, startFuncParams*)
     * @param task.startFuncParams 执行参数
     * @param task.FinishFunc      任务结束的回调函数 => FinishFunc(thd, result)
     * @param task.ErrorFunc       处理错误函数 => ErrorFunc(thd, exception)
     * @param task.ExitFunc        退出线程时的回调函数 => ExitFunc(thd)
     */
    AddTask(task) {
        this.tasks.Push({
            taskScript      : task.taskScript,
            taskFuncName    : task.taskFuncName,
            taskFuncParams  : task.HasProp("taskFuncParams") ? task.taskFuncParams : "",
            StartFunc       : task.HasMethod("StartFunc") ? task.StartFunc : unset,
            startFuncParams : task.HasProp("startFuncParams") ? task.startFuncParams : "",
            FinishFunc      : task.HasMethod("FinishFunc") ? task.FinishFunc : ((*) => ),
            ErrorFunc       : task.HasMethod("ErrorFunc") ? task.ErrorFunc : ((*) => ),
            ExitFunc        : task.HasMethod("ExitFunc") ? task.ExitFunc : unset,
        })
        this.StartThreads()
    }

    ; 启动线程
    StartThreads() {
        while (this.runningThreads.Count < this.maxThreads && this.tasks.Length > 0) {
            task := this.tasks.RemoveAt(1)
            thd := Worker("Persistent`n" task.taskScript)
            thd.task := task
            this.runningThreads[thd.ThreadID] := thd
            if task.HasMethod("StartFunc")
                task.StartFunc.Call(task.startFuncParams*)
            aPromise := thd.AsyncCall(task.taskFuncName, task.taskFuncParams*)
            aPromise.Then((result) => (task.FinishFunc.Call(thd, result), this.ExitThread(thd), this.StartThreads()))
            aPromise.Catch((exception) => (task.ErrorFunc.Call(thd, exception), this.ExitThread(thd), this.StartThreads()))
        }
    }

    ; 退出线程
    ExitThread(thd) {
        if thd.task.HasMethod("ExitFunc")
            thd.task.ExitFunc.Call(thd)
        this.runningThreads.Delete(thd.ThreadID)
        thd.ExitApp()
        thd := ""
    }

    ; 暂停线程池内的全部任务
    Pause() {
        for threadID, thd in this.runningThreads
            thd.Pause(true)
    }

    ; 恢复线程池内的全部任务
    Resume() {
        for threadID, thd in this.runningThreads
            thd.Pause(false)
    }

    ; 结束线程池内的全部任务
    StopAllTasks() {
        this.tasks.Length := 0
        for threadID, thd in this.runningThreads {
            thd.ExitApp()
            this.runningThreads.Delete(threadID)
        }
    }
}