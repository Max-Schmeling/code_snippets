def GetFocusProcess(mode):
    """
    Returns the currently focused/active
    program in Windows only. Three modes:
    (1) mode = "pid" - Process ID
    (2) mode = "title" - Window title
    (3) mode = "process" - Process/Program name + start time

    by Max Schmeling
    """
    #Dependancies:
    #from win32process import GetWindowThreadProcessId
    #from win32gui import GetForegroundWindow
    #from psutil import Process
    foreground_win_id = GetForegroundWindow()
    foreground_pid = GetWindowThreadProcessId(foreground_win_id)[1]
    if mode == "pid":
        return foreground_pid
    elif mode == "title":
        return GetWindowText(foreground_win_id)
    elif mode == "process":
        foreground_proc_name = Process(foreground_pid).name()
        #foreground_proc_tstart = Process(foreground_pid).started()
        return (foreground_proc_name, foreground_pid)
        
    return False





def GetSysInfo(self):
    """
    Returns CPU, RAM, Netspeed

    by Max Schmeling
    """
    
    # Get and format percentual CPU usage (Psutil.cpu_percent)
    cpu_use = Psutil.cpu_percent(interval=1, percpu=False)
    cpu_str = "{}%".format(cpu_use)
    
    # Get and format RAM percentual usage (Psutil.virtual_memory)
    ram_use = Psutil.virtual_memory().percent
    ram_str = "{}%".format(ram_use)
    
    # Get and format upload/download rates (pyspeedtest)
    try:
        downrate = SpeedTest().download()
        downrate, downrate_unit = DataSpeedUnit(downrate)
        downrate_str = "{}{}".format(int(downrate), downrate_unit)
        
        uprate = SpeedTest().upload()
        uprate, uprate_unit = DataSpeedUnit(uprate)
        uprate_str = "{}{}".format(int(uprate), uprate_unit)
    except Exception as e:
        print(e)
        uprate_str = "-"
        downrate_str = "-"
        
    return [cpu_str, ram_str, downrate_str, uprate_str]
