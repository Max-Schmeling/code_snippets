import os
import inspect
import datetime

class Logger(object):
    TIME_FORMAT = "%Y-%m-%d - %H:%M:%S"
    FILE_EXT = "log"
    def __init__(self, directory, filebasename): #rollover=True
        self.directory = directory
        self.filebasename = filebasename
        self.index = 0
        
    def filepath(self):
        filename = self.filebasename + str(self.index) + "." + self.FILE_EXT
        return os.path.join(self.directory, filename)
        
    def log(self, level, msg, stdout=False):
        level = level.strip().upper()
        assert level in ("DEBUG", "INFO", "WARNING", "ERROR")
        time = datetime.datetime.now().strftime(self.TIME_FORMAT)
        caller = inspect.getframeinfo(inspect.stack()[1][0])
        filename = os.path.basename(caller.filename)
        text = "[{}][{}][{}:L{}][{}()]: {}".format(level, time, filename, caller.lineno,
                                                caller.function, msg)
        if stdout: print(text)
        if os.path.isfile(self.filepath()): 
            text = "\n" + text
        with open(self.filepath(), "a", encoding="utf-8") as logfile:
            logfile.write(text)
