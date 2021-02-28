def mean(data):
    """ Return the sample arithmetic mean of data. """
    n = len(data)
    if n < 1:
        raise ValueError('mean requires at least one data point')
    return sum(data)/n

def median(data):
    """ Return the median of data """
    n = len(data)
    if n < 1:
        raise ValueError('median requires at least one data point')
    idx1 = int(n/2-1)
    idx2 = int(n/2)
    if n % 2 == 0:
        return mean(data[idx1:idx2])
    else:
        return data[idx1]

def _ss(data):
    """ Return sum of square deviations of sequence data. """
    c = mean(data)
    ss = sum((x-c)**2 for x in data)
    return ss

def stddev(data, ddof=0):
    """
    Calculates the population standard deviation
    by default; specify ddof=1 to compute the sample
    standard deviation. Requires function _ss().
    """
    n = len(data)
    if n < 2:
        raise ValueError('variance requires at least two data points')
    ss = _ss(data)
    pvar = ss/(n-ddof)
    return pvar**0.5
