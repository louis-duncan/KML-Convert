import csv
import os
import subprocess

import coord
import traceback
import easygui
import re
import html2text


class UserCancelError(Exception):
    pass


class Point:
    """An object to contain normalised point data."""

    def __init__(self,
                 name="",
                 latitude=0.0,
                 longitude=0.0,
                 text="",
                 attributes=None,
                 style=None,
                 ):
        self._name = name
        self._longitude = longitude
        self._latitude = latitude
        self._text = text
        self._attributes = attributes
        self._style = style
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

    def set_text(self, text):
        self._text = text

    def set_attributes(self, attributes):
        self._attributes = attributes

    def convert_text(self):
        if self._text.startswith("<![CDATA[") and self._text.endswith("]]>"):
            self._text = multi_replace(self._text,
                                       ["<![CDATA[",
                                        "]]>"],
                                       "")
            self._text = self._text.replace("\n", "<br>")
            handler = html2text.HTML2Text()
            handler.ignore_emphasis = True
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

    def one_line(self, width=100):
        text = self._name + ": " + self._text.replace("\n", " ")
        if len(text) > width:
            text = text[: width - 3] + "..."
        return text

    def easy_view(self):
        names = ["Name:",
                 "Location:",
                 "Text:",
                 ]
        values = [self._name,
                  coord.coord_to_nesw(self._longitude,
                                      self._latitude),
                  self._text,
                  ]
        for a in self._attributes:
            names.append(a.capitalize() + ":")
            values.append(self._attributes[a])

        text = ""
        for i in range(len(names)):
            text += names[i].ljust(max([len(n) for n in names]) + 1) + "\n"
            text += values[i] + "\n\n"

        easygui.msgbox(text, self._name)

    def get_style(self):
        return self._style


class Converter:
    def __init__(self, input_path):
        self._input_path = input_path
        self._output_path = ""
        self._data_lines = ""
        self._points = []
        self._errors = []
        self._styles = []
        self._style_maps = []

        self._placemark_expressions = ["<Placemark.*?>.*?</Placemark.*?>"]

        self.decide_output_path()
        self.load_data()

    def decide_output_path(self):
        path, filename = os.path.split(self._input_path)
        self._output_path = os.path.join(path, "output", str(filename.rsplit(".", 1)[0]) + ".csv")

    def load_data(self):
        self._data_lines, fn = load_data(self._input_path)

    def convert(self):
        self.decode_styles()
        self._points, self._errors = process_lines(self._data_lines, self._placemark_expressions)

    def get_errors(self):
        return self._errors

    def get_number_of_lines(self):
        return self._data_lines.count("\n") + 1

    def get_points_estimate(self):
        return sum([len(re.compile(e, re.DOTALL).findall(self._data_lines)) for e in self._placemark_expressions])

    def get_number_of_points(self):
        return len(self._points)

    def get_fields(self):
        """Take a list of dicts.
        RETURNS: A list containing all the unique/distinct keys used in the dicts given."""
        fields = []
        for p in self._points:
            d = p.get_attributes()
            if d is None:
                pass
            else:
                for f in d:
                    if f in fields:
                        pass
                    else:
                        fields.append(f)
        return sorted(fields)

    def get_pre_info(self):
        text = """{} lines loaded from:
{}
{} points expected.""".format(self.get_number_of_lines(),
                              self._input_path,
                              self.get_points_estimate())
        return text

    def get_post_info(self):
        formatted_fields = ""
        for f in self.get_fields():
            formatted_fields += f + "\n"
        formatted_fields = formatted_fields.strip()
        text = """Input File:
{}

Output File:
{}

{} points found with {} errors.

Fields:
{}""".format(self._input_path, self._output_path, len(self._points), len(self._errors), formatted_fields)
        return text

    def export(self):
        export_points(self._points, self._output_path)

    def get_input_path(self):
        return self._input_path

    def explore(self):
        formatted_fields = ""
        for f in self.get_fields():
            formatted_fields += "- " + f + "\n"
        formatted_fields = formatted_fields.strip()
        choices = ["View Points",
                   "View Errors",
                   "Change Output\nLocation",
                   "Back"]
        while True:
            text = """Input File:
{}
Output File:
{}

Points: {}
Errors: {}
Fields:
{}""".format(self._input_path,
             self._output_path,
             len(self._points),
             len(self._errors),
             formatted_fields)
            choice = easygui.buttonbox(text, "Data Explorer", choices)
            if choice == choices[0]:
                point_choices = [p.one_line() for p in self._points]
                point_choice = easygui.choicebox("Choose a point to view:", "Points", point_choices)
                if point_choice is not None:
                    self._points[point_choices.index(point_choice)].easy_view()
                else:
                    pass
            elif choice == choices[1]:
                error_explorer(self._errors)
            elif choice == choices[2]:
                new_path = easygui.filesavebox("Choose Save Location", default=self._output_path, filetypes=["*.csv"])
                if new_path is None:
                    pass
            else:
                raise UserCancelError()

    def edit_parameters(self):
        done = False
        while not done:
            para_choice = easygui.choicebox("Choose Parameter to Edit:",
                                            "Parameter Edit - {}".format(os.path.basename(self._input_path)),
                                            choices=["Show File Location..."] + self._placemark_expressions + [
                                                "Add New Expression..."])
            if para_choice == "Show File Location...":
                subprocess.Popen(r'explorer /select,"{}"'.format(self.get_input_path()))
            elif para_choice in self._placemark_expressions:
                pos = self._placemark_expressions.index(para_choice)
                new = easygui.enterbox("Edit expression, leave blank to remove:",
                                       default=para_choice)
                if new is None:
                    pass
                else:
                    if new == "":
                        self._placemark_expressions.remove(para_choice)
                    else:
                        self._placemark_expressions[pos] = new
            elif para_choice == "Add New Expression...":
                default = "<ABC>.*?</ABC>"
                new = easygui.enterbox("Type regular expression to find points:",
                                       default=default)
                if new is None:
                    pass
                else:
                    if new in self._placemark_expressions or new == default:
                        easygui.msgbox("Expression is duplicate or the same as the template.\nNot added.")
                    else:
                        self._placemark_expressions.append(new)
            else:
                done = True

    def decode_styles(self):
        styles_ex = re.compile('<Style id=".*?">.*?</Style>', re.DOTALL)
        maps_ex = re.compile('<StyleMap id=".*?">.*?</StyleMap>', re.DOTALL)
        self._styles = [Style(st) for st in styles_ex.findall(self._data_lines)]
        self._style_maps = [StyleMap(mt) for mt in maps_ex.findall(self._data_lines)]

    def get_style_icon(self, style_id):
        if style_id is None:
            return None
        else:
            result = self.search_styles(style_id)

            if result is None:
                result = self.search_style_maps(style_id)

            if type(result) == Style:
                return result.get_icon()
            elif type(result) == StyleMap:
                mapped_style = result.get_style()
                return self.search_styles(mapped_style).get_icon()
            else:
                return None

    def search_styles(self, test):
        result = None
        for s in self._styles:
            if s.check_id(test.strip("#")):
                result = s
                break
        return result

    def search_style_maps(self, test):
        result = None
        for s in self._style_maps:
            if s.check_id(test.strip("#")):
                result = s
                break
        return result

    def give_icons(self):
        # Todo: Write
        pass


class Style:
    def __init__(self, text=None):
        self._text = text
        self._id = ""
        self._icon = None
        self.decode()

    def decode(self):
        id_ex = re.compile('<Style id=".*?">')
        icon_ex = re.compile('<Icon>.*?</Icon>', re.DOTALL)
        href_ex = re.compile('<href>.*?</href>', re.DOTALL)
        self._id = multi_strip(id_ex.findall(self._text)[0],
                               ['<Style id="',
                                '">',
                                "#"]
                               )
        self._icon = multi_strip(href_ex.findall(icon_ex.findall(self._text)[0])[0],
                                 ["<href>",
                                  "</href>",
                                  ],
                                 )

    def get_id(self):
        return self._id

    def get_icon(self):
        return self._icon

    def check_id(self, test):
        return self._id == test


class StyleMap:
    """Defines a mapping from a style id to several other based on context."""

    def __init__(self, text):
        self._id = ""
        self._pairs = {}
        self._text = text
        self.decode()

    def decode(self):
        id_ex = re.compile('<StyleMap id=".*?">')
        map_id = multi_strip(id_ex.findall(self._text)[0], ['<StyleMap id="', '">', "#"])
        key_ex = re.compile('<key>.*?</key>', re.DOTALL)
        url_ex = re.compile('<styleUrl>.*?</styleUrl>', re.DOTALL)
        keys = [multi_strip(k, ["<key>", "</key>"]) for k in key_ex.findall(self._text)]
        urls = [multi_strip(s, ["<styleUrl>", "</styleUrl>"]) for s in url_ex.findall(self._text)]
        style_pairs = {}
        for i in range(len(keys)):
            style_pairs[keys[i]] = urls[i]
        self._id = map_id
        self._pairs = style_pairs

    def get_id(self):
        return self._id

    def check_id(self, test):
        return test == self._id

    def check_style(self, test):
        found = False
        k = None
        for k in self._pairs:
            if self._pairs[k] == test:
                found = True
                break
        if found:
            return k
        else:
            return False

    def get_style(self, key="normal"):
        try:
            return self._pairs[key]
        except ValueError:
            return

    def get_style_pairs(self):
        return self._pairs


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


def process_lines(data_lines, expressions=("<Placemark.*?>.*?</Placemark>",)):
    place_exes = [re.compile(e, re.DOTALL) for e in expressions]
    name_ex = re.compile("<name>.*?</name>", re.DOTALL)
    description_ex = re.compile("<description>.*?</description>", re.DOTALL)
    coordinates_ex = re.compile("<coordinates>.*?</coordinates>", re.DOTALL)
    style_ex = re.compile("<styleUrl>.*?</styleUrl>")
    chunks = []
    for ex in place_exes:
        chunks += ex.findall(data_lines)
    points = []
    errors = []
    c = 0
    for chunk in chunks:
        if (c - 1) % 1000 == 0:
            print("Processed {} points of {}.".format(c - 1, len(chunks)))
        try:
            new_name = multi_strip(name_ex.findall(chunk)[0], name_ex.pattern.split(".*?"))
            new_lon, new_lat = coord.normalise(multi_strip(coordinates_ex.findall(chunk)[0],
                                                           coordinates_ex.pattern.split(".*?")))
            new_description = description_ex.findall(chunk)
            if len(new_description) == 0:
                new_description = ""
            else:
                new_description = multi_strip(new_description[0],
                                              description_ex.pattern.split(".*?"))
            new_style = style_ex.findall(chunk)
            if len(new_style) == 0:
                new_style = None
            else:
                new_style = multi_strip(new_style[0],
                                        style_ex.pattern.split(".*?"))
            new_point = Point(new_name, new_lat, new_lon, new_description, style=new_style)
            points.append(new_point)
        except IndexError:
            err = {"text": "Missing attribute in:\n" + chunk,
                   "chunk number": c,
                   "exception": traceback.format_exc()}
            try:
                err["name"] = new_name
            except NameError:
                pass
            try:
                err["description"] = str(new_description)
            except NameError:
                pass
            try:
                err["lon/lat"] = str((new_lon, new_lat))
            except NameError:
                pass

            errors.append(err)
        c += 1
    return points, errors


def load_data(file_name=None):
    """Optionally takes a file name.
RETURNS: File contents as string."""
    if file_name is None:
        file_name = easygui.fileopenbox(default="*.KML", filetypes=("*.KML",))
    if file_name is None:
        raise UserCancelError()

    with open(file_name, errors="surrogateescape") as fh:
        data_lines = fh.read()  # fh.readlines()
    return data_lines, file_name


def spot_path(text):
    """Takes a string and looks for a file path with prefix 'src="'.
RETURNS: The path if a path is found, else: returns None"""
    test1 = re.search('src=".*">', text)
    test2 = re.search(r'!\[\]\(.*\)', text)
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


def insert_attribute(attributes, key, data):
    n = 2
    new_key = key
    while new_key in attributes:
        new_key = key + str(n)
        n += 1
    attributes[new_key] = data
    return attributes


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
                              "description",
                              "location",
                              "condition",
                              "long",
                              "lat",
                              "title",
                              "county",
                              "type",
                              "ngr",
                              "comments",
                              "site",
                              "date found",
                              "date fell",
                              "date made safe",
                              "date from",
                              ),
                   separators=(":", "=")
                   ):
    """Takes a long string, and searches for lines starting with given key words.
Default key words are as above, but a custom list of key words can be defined.
RETURNS: The text as given minus any lines starting with a key word,
and a dict containing values against keys which are the key word by which they were found."""
    attributes = {}

    lines = [l.strip() for l in text.split("\n")]

    new_text = ""

    for line in lines:
        found = False
        found_key = ""
        found_data = ""
        for s in separators:
            parts = line.split(s, 1)
            if len(parts) > 1:
                found_key = multi_strip(parts[0], [" ", "\t", "-", ":", "#", "\\", "="]).lower()
                if found_key in key_words:
                    found = True
                    found_data = multi_strip(parts[1], [" ", "\t", "\n", "-", ":", "#", "\\", "="])
                    break
            else:
                pass
        else:
            pass

        if not found:
            path = spot_path(line)
            if path is not None:
                found = True
                found_key = "path"
                found_data = path

        if found:
            attributes = insert_attribute(attributes, found_key, found_data)
        else:
            new_text += line + "\n"

    return new_text.strip(), attributes


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

    # Ensure the dir exists.
    path = os.path.dirname(file_name)
    if not os.path.exists(path):
        os.mkdir(path)

    if fields is None:
        fields = get_fields([p.get_attributes() for p in points])

    headings = ["x", "y", "id", "text"]
    if "icon" in fields:
        headings.append("icon")
        fields.remove("icon")
    headings = headings + fields

    rows = [headings] + [format_row(p, headings) for p in points]

    fh = open(file_name, "w", errors="surrogateescape")
    writer = csv.writer(fh, lineterminator='\n')
    writer.writerows(rows)
    fh.close()


def single_file():
    try:
        data_lines, input_path = load_data()
    except UserCancelError:
        return None
    msg = "{} lines of data loaded with an approximate {} points.\nContinue?".format(data_lines.count("\n"),
                                                                                     data_lines.count("</Placemark>"))
    choice = easygui.buttonbox(msg, choices=["Yes", "No"])
    if choice == "No" or choice is None:
        raise UserCancelError()
    points, errors = process_lines(data_lines)
    fields = get_fields([p.get_attributes() for p in points])
    output_path = input_path.rsplit(".", 1)[0] + ".csv"
    done = False
    while not done:
        msg = "{} points found.".format(len(points))
        msg += "\n\n{} errors.".format(len(errors))
        msg += "\n\nFields Found:\n"
        for f in fields:
            msg += f + "\n"
        msg += "\nOutput Path:\n" + output_path

        choices = ["View Data", "Export", "Change Output\nLocation"]
        if len(errors) > 0:
            choices.append("View Errors")
        choice = easygui.buttonbox(msg, choices=choices)

        if choice == choices[0]:
            pass
            # data_explorer(points)
        elif choice == choices[1]:
            export_points(points, output_path, fields=fields)
            easygui.msgbox("Done!\n\nThe program will now close.")
            done = True
        elif choice == choices[2]:
            new_path = easygui.filesavebox(default=output_path, filetypes=["*.csv"])
            if new_path is not None:
                output_path = new_path
        elif choice == "View Errors":
            error_explorer(errors)
        else:
            raise UserCancelError()


def data_explorer(converters):
    choices = [c.get_input_path() for c in converters]
    while True:
        choice = easygui.choicebox("Choose a file to review:", "Data Explorer", choices)
        if choice is None:
            raise UserCancelError()
        converters[choices.index(choice)].explore()


def multi_file():
    file_paths = easygui.fileopenbox(default="*.kml", multiple=True)
    if file_paths is None:
        raise UserCancelError()

    converters = [Converter(path) for path in file_paths]
    started = False
    choices = ["Continue", "Edit\nParameters"]
    while not started:
        pre_msg = ""
        for c in converters:
            pre_msg += "{}:\nExpecting {} points from {} lines.\n\n".format(os.path.basename(c.get_input_path()),
                                                                            c.get_points_estimate(),
                                                                            c.get_number_of_lines())
        choice = easygui.buttonbox(pre_msg, "Multi File Convert", choices)
        if choice == choices[0]:
            started = True
        elif choice == choices[1]:
            file_choices = [os.path.basename(c.get_input_path()) for c in converters]
            file_choice = easygui.choicebox("Select file:",
                                            "Parameter Edit",
                                            choices=file_choices)
            converters[file_choices.index(file_choice)].edit_parameters()
        else:
            raise UserCancelError()

    for c in range(len(converters)):
        converters[c].convert()

    post_msg = "{} points found with {} errors.".format(sum([c.get_number_of_points() for c in converters]),
                                                        sum([len(c.get_errors()) for c in converters]))
    choices = ["View Data\nby File", "Export", "Cancel"]
    done = False
    while not done:
        choice = easygui.buttonbox(post_msg, "Multi File Convert", choices)
        if choice == choices[0]:
            try:
                data_explorer(converters)
            except UserCancelError:
                pass
        elif choice == choices[1]:
            text = ""
            for c in converters:
                try:
                    print("Exporting points from {}.".format(os.path.basename(c.get_input_path())))
                    c.export()
                    text += "Successfully exported {}\n\n".format(c.get_input_path())
                except Exception as err:
                    text += "Error while exporting {}: {}\n\n".format(c.get_input_path(), err)
            easygui.msgbox(text, "Export Complete")
            done = True
        else:
            raise UserCancelError()


if __name__ == "__main__":
    # run = True
    # while run:
    #    try:
    #        main()
    #    except UserCancelError:
    #        run = False
    pass
