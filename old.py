import re
import datetime
import easygui
import coord
import csv
import traceback


class Point():
    def __init__(self,
                 name="",
                 latitude=0.0,
                 longitude=0.0,
                 text="",
                 attributes=None
                 ):
        self._name = name
        self._longitude = longitude
        self._latitude = latitude
        self._text = text
        self._attributes = attributes


def explicit_strip(text: str, target: str):
    assert type(text) == str
    assert type(target) == str
    change = True
    while change:
        change = False
        if text.startswith(target):
            text = text[len(target):]
            change = True
        if text.endswith(target):
            text = text[:len(text) - len(target)]
    return text


def multi_strip(text: str,
                targets: list,
                ):
    assert type(text) == str
    assert type(targets) == list
    change = True
    while change:
        previous = text
        for t in targets:
            text = explicit_strip(text, t)
        if text == previous:
            change = False
    return text


def process_lines(data_lines):
    points = []
    current_point = {"name": "",
                     "description": "",
                     "longitude": 0.0,
                     "latitude": 0.0}
    in_point = False
    in_name = False
    in_description = False
    in_lat_lon = False

    errors = []

    for i, line in enumerate(data_lines):
        raw_line = line.strip()

        # Start new points and save finished points.
        if raw_line == "<Placemark>":
            current_point = {"name": "",
                             "description": "",
                             "longitude": 0.0,
                             "latitude": 0.0}
            in_point = True
            in_name = False
            in_description = False
            in_lat_lon = False
        elif raw_line == "</Placemark>":
            points.append(current_point)
            in_point = False
        else:
            pass

        # Track which part of the point we are reading.
        if in_point:
            if raw_line.startswith("<name>"):
                in_name = True
                raw_line = raw_line.replace("<name>", "")

            if raw_line.startswith("<description>"):
                in_description = True
                raw_line = raw_line.replace("<description>", "")

            if raw_line.startswith("<coordinates>"):
                in_lat_lon = True
                raw_line = raw_line.replace("<coordinates>", "")

            # Put the line's data in the right place depending on current area being read.
            if in_name:
                if raw_line.endswith("</name>"):
                    in_name = False
                    raw_line = raw_line.replace("</name>", "")
                current_point["name"] += "\n" + raw_line
                current_point["name"] = current_point["name"].strip()

            if in_description:
                if raw_line.endswith("</description>"):
                    in_description = False
                    raw_line = raw_line.replace("</description>", "")
                current_point["description"] += "\n" + raw_line
                current_point["description"] = current_point["description"].strip()
                current_point["description"] = multi_strip(current_point["description"],
                                                           ["]]>", "<![CDATA["])
            if in_lat_lon:
                if raw_line.endswith("</coordinates>"):
                    in_lat_lon = False
                    raw_line = raw_line.replace("</coordinates>", "")
                if raw_line.strip() == "":
                    pass
                else:
                    try:
                        current_point["longitude"], current_point["latitude"] = coord.normalise(raw_line)
                    except Exception as err:
                        errors.append({"text": "Failed to extract location",
                                       "id": current_point["name"],
                                       "line": i,
                                       "raw_line": '"' + raw_line + '"',
                                       "error": repr(err),
                                       "exception": traceback.format_exc()})
        else:
            pass
    return points, errors


def load_data(file_name=""):
    if file_name == "":
        file_name = easygui.fileopenbox()
    if file_name is None:
        return None

    with open(file_name) as fh:
        data_lines = fh.readlines()
    return data_lines


def spot_path(text):
    if 'src="' in text:
        start_pos = pos = re.search('src="', text).span()[1]
        stop = False
        while not stop:
            if text[pos] == '"':
                stop = True
            else:
                pos += 1
        path = text[start_pos: pos]
        path = multi_strip(path, ["file:",
                                  "/",
                                  ])
    elif text.lower().strip().startswith("found in"):
        path = multi_strip(text[len("found in"):], [" ", "\n", "\t", ":"])
    else:
        path = False

    return path


def find_paths(lines):
    new_lines = []
    paths = []
    for line in lines:
        path = spot_path(line)
        if path is False:
            if line.strip() == "":
                pass
            else:
                new_lines.append(line)
        else:
            paths.append(path)
    return new_lines, paths


def get_attributes(text):
    key_words = ["date",
                 "scale",
                 "library number",
                 "quality",
                 "run number",
                 "ref",
                 "other",
                 ]
    attributes = {}

    lines = [l.strip() for l in text.split("\n")]

    lines, paths = find_paths(lines)

    new_lines = []

    for line in lines:
        key_word_match = False
        for key_word in key_words:
            if line.lower().startswith(key_word.lower()):
                key_word_match = True
                key_valid = False
                n = 0
                new_key_word = key_word.lower()
                while not key_valid:
                    if new_key_word in attributes.keys():
                        n += 1
                        new_key_word = key_word.lower() + str(n)
                    else:
                        key_valid = True
                attributes[new_key_word] = multi_strip(line[len(key_word):], [":", " ", "\n", "\t"])
            else:
                pass
        if not key_word_match:
            new_lines.append(line)

    for i, path in enumerate(paths):
        attributes["path " + str(i + 1)] = path

    new_text = ""
    for line in new_lines:
        new_text += line + "\n"
    new_text = new_text.strip()

    return new_text, attributes


def convert_lines(data_lines):
    errors = []

    point_dicts, new_errors = process_lines(data_lines)

    errors = errors + new_errors

    points = []

    for pd in point_dicts:
        text, attributes = get_attributes(pd["description"])

        points.append(Point(pd["name"],
                            pd["latitude"],
                            pd["longitude"],
                            text,
                            attributes))

    return points, errors


def format_row(point):
    #Todo: RE-WRITE!
    return [point. longitude, point.latitude, point.name, point.description, point.date] + point.files


def export_points(points, file_name=""):
    #Todo: RE-WRITE!
    if file_name == "":
        file_name = easygui.filesavebox()
    if file_name is None:
        return None

    max_files = max([len(p.files) for p in points])

    headings = ["x", "y", "id", "description", "date"] + ["file " + str(i + 1) for i in range(max_files)]

    rows = [headings] + [format_row(p) for p in points]

    fh = open(file_name, "w")
    writer = csv.writer(fh, lineterminator='\n')
    writer.writerows(rows)
    fh.close()


def error_explorer(errors):
    rows = [e["text"] + " in " + e["id"] + " - caused: " + repr(e["error"]) for e in errors]
    choice = True
    while choice is not None:
        choice = easygui.choicebox(choices=rows)
        if choice is None:
            pass
        else:
            err = errors[rows.index(choice)]
            msg = ""
            for k in err:
                msg += "\n  " + k + ":\n" + str(err[k]) + "\n"
            easygui.msgbox(msg)


def get_fields(dicts):
    fields = []
    for d in dicts:
        for f in d:
            if f in fields:
                pass
            else:
                fields.append(f)
    return fields


def data_explorer(dicts):
    pass


def main():
    data_lines = load_data()
    msg = "{} lines of data loaded with an approximate {} points.\nContinue?".format(len(data_lines),
                                                                                     sum([int("</Placemark>" in line) for line in data_lines]))
    choice = easygui.buttonbox(msg, choices=["Yes", "No"])
    if choice == "No":
        return None
    dicts, errors = convert_lines(data_lines)
    fields = get_fields(dicts)
    msg = "{} points found.".format(len(dicts))
    msg += "\n\n{} errors.".format(len(errors))
    msg += "\n\nFields Found:\n"
    for f in fields:
        msg += f + "\n"

    while choice is not None:
        choices = ["View Data", "Export"]
        if len(errors) > 0:
            choices.insert(1, "View Errors")
        choice = easygui.buttonbox(msg, choices=choices)

        if choice == "View Data":
            data_explorer(dicts)
        elif choice == "View Errors":
            error_explorer(errors)
        elif choice is "Export":
            export_points(dicts)
        else:
            pass