import math


def coord_to_nesw(lon, lat):
    """Converts lon/lat to 1°2'3" E 1°2'3" N"""
    # Lon
    lon_dir = "E"
    if lon < 0:
        lon = abs(lon)
        lon_dir = "W"
    lon_degs = int(lon)
    rem = (lon - lon_degs) * 60
    lon_mins = int(rem)
    lon_secs = (rem - lon_mins) * 60
    full_lon = "{:0>2.0f}°{:0>2.0f}'{:0>5.2f}\"{}".format(lon_degs,
                                                          lon_mins,
                                                          lon_secs,
                                                          lon_dir,
                                                          )
    # Lat
    lat_dir = "N"
    if lat < 0:
        lat = abs(lat)
        lat_dir = "S"
    lat_degs = int(lat)
    rem = (lat - lat_degs) * 60
    lat_mins = int(rem)
    lat_secs = (rem - lat_mins) * 60
    full_lat = "{:0>2.0f}°{:0>2.0f}'{:0>5.2f}\"{}".format(lat_degs,
                                                          lat_mins,
                                                          lat_secs,
                                                          lat_dir,
                                                          )
    return "{} {}".format(full_lon, full_lat)


def nesw_to_coord(nesw):
    """Converts 1°2'3" E 1°2'3" N to lon/lat"""
    chars_to_remove = (" ", ",", ":", ";", "-", "&", "/", "\\")
    for c in chars_to_remove:
        nesw = nesw.replace(c, "")
    N = ""
    E = ""
    S = ""
    W = ""
    buffer = ""
    for c in nesw:
        if c == "N":
            N = buffer
            buffer = ""
        elif c == "E":
            E = buffer
            buffer = ""
        elif c == "S":
            S = buffer
            buffer = ""
        elif c == "W":
            W = buffer
            buffer = ""
        else:
            buffer += c
    assert ("" in (N, S)) and ("" in (E, W))

    if S == "":
        lat_dir = "N"
        lat_str = N
    else:
        lat_dir = "S"
        lat_str = S
    if W == "":
        lon_dir = "E"
        lon_str = E
    else:
        lon_dir = "W"
        lon_str = W

    lon_degs, lon_mins, lon_secs = multi_split(lon_str, ("°", "'", '"'))
    lon = round(sum([float(lon_degs), float(lon_mins) / 60, float(lon_secs) / 3600]), 10)
    lat_degs, lat_mins, lat_secs = multi_split(lat_str, ("°", "'", '"'))
    lat = round(sum([float(lat_degs), float(lat_mins) / 60, float(lat_secs) / 3600]), 10)
    assert (lon >= 0.0) and (lat >= 0.0)

    if lon_dir == "W":
        lon = lon.__neg__()
    else:
        pass
    if lat_dir == "S":
        lat = lat.__neg__()
    else:
        pass

    return lon, lat


def multi_count(targets, text):
    text = text.upper()
    return sum([text.count(c) for c in targets])


def normalise(text: str):
    """Detect if nesw or coord and produce coord"""
    chars_to_catch = ("N", "E", "S", "W", "°", "'", '"')
#    is_nesw = False
#    for c in chars_to_catch:
#        if c in text.upper():
#            is_nesw = True
#            break
#        else:
#            pass
    is_nesw = multi_count(chars_to_catch, text) >= 2

    if is_nesw:
        lon, lat = nesw_to_coord(text)
    else:
        parts = text.split(",")
        lon, lat = parts[0], parts[1]
        lon = float(lon.strip())
        lat = float(lat.strip())

    return lon, lat


def validate(text: str):
    try:
        lon, lat = normalise(text)
        return True
    except ValueError:
        print("ValueError in coord check.")
    except AssertionError:
        print("Assertion Error in coord check")
    except IndexError:
        print("Index Error in coord check")
    return False


def multi_split(text, seps=(",",)):
    output = [text]
    for s in seps:
        new_output = []
        for o in output:
            res = o.split(s)
            for r in res:
                new_output.append(r)
        output = new_output
    for i in range(output.count("")):
        output.remove("")
    return output


def distance(start, end):
    """Calc distance between start and end, start and end should be (lon, lat)"""
    earth_radius = 6364900  # Earth radius for central UK, metres
    lat1 = math.radians(start[1])
    lat2 = math.radians(end[1])
    lat_delta = math.radians(end[1] - start[1])
    lon_delta = math.radians(end[0] - start[0])
    a = math.sin(lat_delta/2) *\
        math.sin(lat_delta/2) +\
        math.cos(lat1) *\
        math.cos(lat2) *\
        math.sin(lon_delta/2) *\
        math.sin(lon_delta/2)
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    d = earth_radius * c
    return d
