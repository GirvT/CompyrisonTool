def compareAll(*args):
    # All input must be lists
    matchList = list()
    maxList = list()
    trackList = list()
    for arg in args:
        maxList.append(len(arg) - 1)
        trackList.append(0)
    while trackList[-1] != maxList[-1]:
        for i in range(0, len(maxList)):
            if trackList[i] > maxList[i]:
                try:
                    trackList[i + 1] += 1
                except:
                    pass
                else:
                    trackList[i] = 0
        if equalAll(list(args), trackList):
            matchList.append(trackList.copy())
        trackList[0] += 1
    return matchList

def equalAll(args, trackList):
    equalList = list()
    for i in range(0, len(trackList)):
        equalList.append(args[i][trackList[i]])
    if len(set(equalList)) == 1:
        return True
    else:
        return False
