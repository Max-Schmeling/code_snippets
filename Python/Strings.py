def Tagcleaner(text):
    """
    Remove "<>" tags from text
    and return cleaned text
    """
    # import re
    cleanr = re.compile('<.*?>')
    text_clean = re.sub(cleanr, '', text)
    return text_clean


def StringStatistics(string):
    """
    Gets formal analysis info
    from string. Amount of chars,
    words, lines and spaces and
    the charset of the string
    """
    charnum = 0 # Counts chars
    wordnum = 0 # Counts words
    linenum = 0 # Counts lines
    spacenum = 0 # Counts spaces
    charset = [] # holds chars in string (no repition)
    line = ""
    word = ""

    for c in string:
        # Character count
        charnum += 1

        # Word and Space count
        if c != " ":
            word += c
        else:
            wordnum += 1
            spacenum += 1
            word = ""

        # Line count
        if c != "\n":
            line += c
        else:
            linenum += 1
            wordnum += 1
            line = ""

        # Character set update
        if not c in charset:
            if c != "\n":
                if c != " " and not c.isspace():
                    charset.append(c)
        

    return [charnum, wordnum, linenum, spacenum, " ".join(charset)]
