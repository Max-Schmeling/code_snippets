def ProgressBar(achieved, total, prefix="", suffix=""):
    """
    Creates a visual CLI progress bar by
    visualizing the ratio between the total
    and what has been achieved so far.

    Taken from SO and reworked by Max Schmeling
    """
    progr_norm = int(achieved/total*20)
    
    block = "█" #█ https://stackoverflow.com/questions/3173320/text-progress-bar-in-the-console
    blocks = progr_norm+1

    if blocks > 21:
        blocks = 21
        progr_perc = 100
        progr_norm = 20
        
    space = " "
    post_space = 20-progr_norm

    progr_graph = "{}{}".format(blocks*block, post_space*space)

    progress = int(achieved/total*100)
    print("{}|{}|{}%| ({} of {} {})".format(prefix, progr_graph, progress, achieved, total, suffix), end="\r")


def PrettyNumber(number, spacer=","):
    """
    Iterates through an Integer and sets dots
    as determined by dot_symb "." or "," every
    third digit for better readability.
    Example: NumberDotter(123456789)
    Output: "123.456.789"

    by Max Schmeling
    """
    
    string = str(number)
    spaced_list = []
    flag = 0
    
    for d in reversed(string):
        if flag == 3:#and string[string.index(d)-1] not in (".", ","):
            spaced_list.insert(0, spacer)
            flag = 0
        spaced_list.insert(0, d)
        flag += 1
        
    return "".join(spaced_list)
