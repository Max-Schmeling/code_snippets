def Timer(action):
    """
    Time code sections by measuring
    the time difference between the
    start time and the current time.
    Takes an action (start or stop).
    "Stop" returns the time passed.
    """
    if action == "start":
        global time_start
        time_start = time()
        return True
    
    elif action == "stop":
        unit = "seconds"
        time_passed = time() - time_start

        if time_passed >= 86400: #convert to days
            time_passed = time_passed / 86400
            unit = "days"
        elif time_passed >= 3600: #convert to hours
            time_passed = time_passed / 3600
            unit = "hours"
        elif time_passed >= 60: #convert to minutes
            time_passed = time_passed / 60
            unit = "minutes"
            
        return [round(time_passed, 1), unit]
