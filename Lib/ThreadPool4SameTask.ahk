/************************************************************************
 * @description 重复任务的线程池类
 * @author nnrxin
 * @date 2024-09-06
 * @version 0.0.0
 * 无锁版本
 ***********************************************************************/
class ThreadPool4SameTask {
    /**
     * function 添加任务到线程池
     * @param maxThreads           线程池最大线程数
     * @param task.taskScript      任务脚本
     * @param task.taskFuncName    执行函数名称
     * @param task.taskFuncParams  执行函数参数
     * @param task.StartFunc       任务开始的回调函数 => StartFunc(thd, startFuncParams*)
     * @param task.startFuncParams 执行参数
     * @param task.FinishFunc      任务结束的回调函数 => FinishFunc(thd, result)
     * @param task.ErrorFunc       处理错误函数 => ErrorFunc(thd, exception)
     * @param task.ExitFunc        退出线程时的回调函数 => ExitFunc(thd)
     */
    __New(maxThreads, task) {
        this.maxThreads := maxThreads
        this.tasks := []
        this.runningThreads := Map()
        this.runningTasks := Map()
        this.inUnloading := false
        ;重复的任务
        this.taskScript      := task.taskScript
        this.taskFuncName    := task.taskFuncName
        this.taskFuncParams  := task.HasProp("taskFuncParams") ? task.taskFuncParams : ""
        this.StartFunc       := task.HasMethod("StartFunc") ? task.StartFunc : unset
        this.startFuncParams := task.HasProp("startFuncParams") ? task.startFuncParams : ""
        this.FinishFunc      := task.HasMethod("FinishFunc") ? task.FinishFunc : unset
        this.ErrorFunc       := task.HasMethod("ErrorFunc") ? task.ErrorFunc : unset
        this.ExitFunc        := task.HasMethod("ExitFunc") ? task.ExitFunc : unset
    }

    __Delete() {
        this.inUnloading := true
        this.Unload()
    }
    
    ; 加载n个线程用于预加载
    LoadThreads(n := 1) {
        loop Min(n, this.maxThreads - this.runningThreads.Count) ; 防止超出最大线程数
            this.AddThread()
    }

    ; 增加新线程
    AddThread() {
        thd := Worker("Persistent`n" this.taskScript)
        thd.task := ""
        return this.runningThreads[thd.ThreadID] := thd
    }

    /**
     * function AddTask 添加任务到线程池
     * @param task  任务
     */
    AddTask(task) {
        this.inUnloading := false
        this.tasks.Push({
            taskFuncName    : task.HasProp("taskFuncName") ? task.taskFuncName : this.taskFuncName,
            taskFuncParams  : task.HasProp("taskFuncParams") ? task.taskFuncParams : this.taskFuncParams,
            StartFunc       : task.HasMethod("StartFunc") ? task.StartFunc : this.StartFunc,
            startFuncParams : task.HasProp("startFuncParams") ? task.startFuncParams : this.startFuncParams,
            FinishFunc      : task.HasMethod("FinishFunc") ? task.FinishFunc : this.FinishFunc,
            ErrorFunc       : task.HasMethod("ErrorFunc") ? task.ErrorFunc : this.ErrorFunc,
            ExitFunc        : task.HasMethod("ExitFunc") ? task.ExitFunc : this.ExitFunc
        })
        this.RunTask()
    }

    ; 启动任务
    RunTask() {
        while (!this.inUnloading && this.runningTasks.Count < this.maxThreads && this.tasks.Length > 0) {
            ; 尝试获得一个没有任务的线程,不存在时新建一个线程
            thd := ""
            for threadID, runningThread in this.runningThreads {
                if runningThread.task = "" {
                    thd := runningThread
                    break
                }
            }
            thd := thd ? thd : this.AddThread()
            ; 从任务列表里获得最早的任务
            this.runningTasks[thd.ThreadID] := thd.task := task := this.tasks.RemoveAt(1)
            Loop {
            } until thd.Ready
            aPromise := thd.AsyncCall(task.taskFuncName, task.taskFuncParams*) ; 异步执行
            if this.HasMethod("StartFunc")
                this.StartFunc.Call(thd, task.startFuncParams*) ; 异步执行开始的回调
            aPromise.Then(PromiseThen) ; 异步执行完的回调
            aPromise.Catch(PromiseCatch) ; 异步执行报错的回调
        }
        PromiseThen(result) {
            if this.inUnloading
                return
            if task.HasMethod("FinishFunc")
                task.FinishFunc.Call(thd, result)
            this.runningTasks.Delete(thd.ThreadID)
            thd.task := ""
            this.RunTask()
        }
        PromiseCatch(exception) {
            if this.inUnloading
                return
            if task.HasMethod("ErrorFunc")
                task.ErrorFunc.Call(thd, exception)
            this.runningTasks.Delete(thd.ThreadID)
            thd.task := ""
            this.RunTask()
        }
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

    ; 退出线程
    ExitThread(threadID) {
        thd := Worker(threadID)
        ;先尝试执行任务的退出函数,再尝试执行线程的退出函数
        if this.runningThreads[threadID].task && this.runningThreads[threadID].task.HasMethod("ExitFunc")
            this.runningThreads[threadID].task.ExitFunc.Call(thd)
        else if this.HasMethod("ExitFunc")
            this.ExitFunc.Call(thd)

        if this.runningTasks.Has(threadID)
            this.runningTasks.Delete(ThreadID)
        this.runningThreads.Delete(threadID)
        thd.task := ""
        thd.ExitApp()
        thd.Wait()
        thd := ""
    }


    ; 清空队列中的任务,正在执行的任务强行退出线程
    StopAllTask() {
        this.tasks.Length := 0
        this.inUnloading := true
        for ThreadID in this.runningTasks
            this.ExitThread(threadID)
    }

    ; 卸载任务及停止线程
    Unload() {
        this.inUnloading := true
        thds := this.runningThreads.Clone()
        for ThreadID in this.runningThreads
            this.ExitThread(threadID)
        for ThreadID in Worker
            if thds.Has(ThreadID)
                MsgBox "线程未正常退出: " ThreadID
        this.tasks.Length := 0
        this.runningTasks.Clear()
    }
}