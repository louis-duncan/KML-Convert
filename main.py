import csv
import coord
import traceback
import easygui
import re
import html2text


class Point:
    """An object to contain normalised point data."""

    def __init__(self,
                 name="",
                 latitude=0.0,
                 longitude=0.0,
                 text="",
                 attributes=None,
                 icon_type=None,
                 ):
        self._name = name
        self._longitude = longitude
        self._latitude = latitude
        self._text = text
        self._attributes = attributes
        self._icon_type = icon_type
        self.convert_text()

    def get_name(self):
        return self._name

    def get_lon_lat(self):
        return self._longitude, self._latitude

    def get_lon(self):
        return self._longitude

    def get_lat(self):
        return self._latitude

    def get_text(self):
        return self._text

    def get_attributes(self):
        return self._attributes

    def get_attribute(self, key, nul_resp=""):
        """Returns a specific attribute is it exists.
If not, returns empty string."""
        assert type(key) is str
        try:
            return self._attributes[key]
        except KeyError:
            return nul_resp

    def get_icon_type(self):
        return self._icon_type

    def set_text(self, text):
        self._text = text

    def set_attributes(self, attributes):
        self._attributes = attributes

    def convert_text(self):
        if self._text.startswith("<![CDATA[") and self._text.endswith("]]>"):
            self._text = multi_replace(self._text,
                                       ["<![CDATA[",
                                        "]]>",
                                        "<B>",
                                        "</B>",
                                        "<I>",
                                        "</I>",
                                        "<P>",
                                        "</P>,"
                                        "<H>",
                                        "</H"],
                                       "")
            self._text = self._text.replace("\n", "<br>")
            handler = html2text.HTML2Text()
            handler.body_width = 1000000
            self._text = handler.handle(self._text).strip()
            new_string = ""
            escaped = False
            for c in self._text:
                if escaped:
                    new_string += c
                    escaped = False
                else:
                    if c == "\\":
                        escaped = True
                    else:
                        new_string += c
                        escaped = False
            self._text = new_string

        self._text, self._attributes = get_attributes(self._text)


def explicit_strip(text: str, target: str):
    """Takes a string, and removes the target exactly from the beginning and end.
Is greedy.
RETURNS: Stripped string"""
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
            change = True
    return text


def multi_replace(text, targets, replacer):
    """Takes a string, and will replace all instances of all targets with the replacer"""
    for t in targets:
        text = text.replace(t, replacer)
    return text


def multi_strip(text: str,
                targets: list,
                ):
    """Takes a string, and list of strings.
Will remove all instances of the targets appearing at the start or end of the string.
Is greedy.
RETURNS: Stripped string"""
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
    place_ex = re.compile("<Placemark>.*?</Placemark>", re.DOTALL)
    name_ex = re.compile("<name>.*?</name>", re.DOTALL)
    description_ex = re.compile("<description>.*?</description>", re.DOTALL)
    coordinates_ex = re.compile("<coordinates>.*?</coordinates>", re.DOTALL)
    chunks = place_ex.findall(data_lines)
    points = []
    errors = []
    for chunk in chunks:
        try:
            new_name = multi_strip(name_ex.findall(chunk)[0], name_ex.pattern.split(".*?"))
            new_description = multi_strip(description_ex.findall(chunk)[0], description_ex.pattern.split(".*?"))
            new_lon, new_lat = coord.normalise(multi_strip(coordinates_ex.findall(chunk)[0],
                                                           coordinates_ex.pattern.split(".*?")))
            new_point = Point(new_name, new_lat, new_lon, new_description)
            points.append(new_point)
        except IndexError:
            errors.append({"text": "Missing attribute in:\n" + chunk})
    return points, errors


def load_data(file_name=None):
    """Optionally takes a file name.
RETURNS: File contents as string."""
    if file_name is None:
        file_name = easygui.fileopenbox(default="*.KML", filetypes=("*.KML",))
    if file_name is None:
        return None

    with open(file_name) as fh:
        data_lines = fh.read()  # fh.readlines()
    return data_lines


def spot_path(text):
    """Takes a string and looks for a file path with prefix 'src="'.
RETURNS: The path if a path is found, else: returns None"""
    test1 = re.search('src=".*">', text)
    test2 = re.search('!\[\]\(.*\)', text)
    if test1 is not None:
        start_pos, end_pos = test1.span()
        path = multi_strip(text[start_pos: end_pos], ['src="', ' ">'])
        path = multi_strip(path, ["file:", "/"])
    elif test2 is not None:
        start_pos, end_pos = test2.span()
        path = multi_strip(text[start_pos: end_pos], ['![](', ')'])
        path = multi_strip(path, ["file:", "/"])
    elif text.lower().strip().startswith("found in"):
        path = multi_strip(text[len("found in"):], [" ", "\n", "\t", ":"])
    else:
        path = None
    return path


def find_paths(lines):
    """Takes a list of strings, and searches each line for file paths using 'spot_path'.
RETURNS: The text minus any lines in which paths were found, and a list of paths."""
    new_lines = []
    paths = []
    for line in lines:
        path = spot_path(line)
        if path is None:
            if line.strip() == "":
                pass
            else:
                new_lines.append(line)
        else:
            paths.append(path)
    return new_lines, paths


def get_attributes(text,
                   key_words=("date",
                              "scale",
                              "library number",
                              "quality",
                              "run number",
                              "ref",
                              "other",
                              "alt name",
                              "opened",
                              "status dec 1944",
                              "current use",
                              "remarks",
                              ),
                   ):
    """Takes a long string, and searches for lines starting with given key words.
Default key words are as above, but a custom list of key words can be defined.
RETURNS: The text as given minus any lines starting with a key word,
and a dict containing values against keys which are the key word by which they were found."""
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
                        new_key_word = key_word.lower() + " " + str(n)
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


def error_explorer(errors):
    """Takes a list of error, and presents an easygui box to view them.
Will take a list of dict which have at least the keys 'text', 'id', and 'error',
other keys are displayed in the message box."""
    rows = [str([i + " : " + e[i] for i in e]) for e in errors]
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
    """Take a list of dicts.
RETURNS: A list containing all the unique/distinct keys used in the dicts given."""
    assert type(dicts) in (tuple, list)
    fields = []
    for d in dicts:
        if type(d) is dict:
            for f in d:
                if f in fields:
                    pass
                else:
                    fields.append(f)
        else:
            pass
    return sorted(fields)


def format_row(point, headings):
    """Takes a list of point objects and a list of heading strings.
RETURNS: A list of values as dictated by the heading values. Non existent values will be filled with an empty string"""
    row_values = []

    for field in headings:
        if field == "x":
            row_values.append(point.get_lon())
        elif field == "y":
            row_values.append(point.get_lat())
        elif field == "id":
            row_values.append(point.get_name())
        elif field == "text":
            row_values.append(point.get_text())
        else:
            row_values.append(point.get_attribute(field))

    return row_values


def export_points(points, file_name="", fields=None):
    """Takes a list of points and an option CSV file name.
Writes the points data to the CSV file."""
    if file_name == "":
        file_name = easygui.filesavebox()
    if file_name is None:
        return None

    if fields is None:
        fields = get_fields([p.get_attributes() for p in points])

    headings = ["x", "y", "id", "text"] + fields

    rows = [headings] + [format_row(p, headings) for p in points]

    fh = open(file_name, "w")
    writer = csv.writer(fh, lineterminator='\n')
    writer.writerows(rows)
    fh.close()


def main():
    data_lines = load_data()
    if data_lines is None:
        return None
    msg = "{} lines of data loaded with an approximate {} points.\nContinue?".format(data_lines.count("\n"),
                                                                                     data_lines.count("</Placemark>"))
    choice = easygui.buttonbox(msg, choices=["Yes", "No"])
    if choice == "No" or choice is None:
        return None
    points, errors = process_lines(data_lines)
    fields = get_fields([p.get_attributes() for p in points])
    msg = "{} points found.".format(len(points))
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
            pass
            # data_explorer(dicts)
        elif choice == "View Errors":
            error_explorer(errors)
        elif choice is "Export":
            export_points(points, fields=fields)
            easygui.msgbox("Done!\n\nThe program will now close.")
            choice = None
        else:
            pass
