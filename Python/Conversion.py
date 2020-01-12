def DataSpeedUnit(speed):
    """
    Convert data speed to the next
    best suited data size unit

    by Max Schmeling
    """
    units = ['bps', 'Kbps', 'Mbps', 'Gbps']
    unit = 0
    while speed >= 1024:
        speed /= 1024
        unit += 1
    return '%0.2f %s' % (speed, units[unit])


def FileSizeConverter(size, factor=1024):
    """
    Convert byte size to the next
    best suited data size unit

    by Max Schmeling
    """
    units = ["Bytes", "KB", "MB", "GB" ,"TB"]
    unit = 0

    while size >= factor:
        size /= factor
        unit += 1

    return [round(size, 2), units[unit]]


def DegreeToRadian(angle):
    """ Converts degrees to radian """
    return angle * math.pi / 180


def RgbToHex(r, g, b):
    return "#{0:02x}{1:02x}{2:02x}".format(r, g, b).upper()


def HexToRgb(hexcode):
    return tuple(int(hexcode[i:i+2], 16) for i in (0, 2 ,4))


def RandomHexColor():
    return "{:06x}".format(random.randint(0, 0xFFFFFF)).upper()


def HexColorShade(hexcode, factor=0.6):
    """ Creates a shade of the given color as determined by <factor> """
    r, g, b = tuple(int(hexcode.lstrip("#")[i:i+2], 16)*factor for i in (0, 2 ,4))
    newhex = "#{0:02x}{1:02x}{2:02x}".format(int(r), int(g), int(b)).upper()
    return newhex
