import xlwings as xw
import re

_alignment_map = {
    'hcenter': ['h', xw.constants.HAlign.xlHAlignCenter],
    'center_across_selection': ['h', xw.constants.HAlign.xlHAlignCenterAcrossSelection],
    'hdistributed': ['h', xw.constants.HAlign.xlHAlignDistributed],
    'fill': ['h', xw.constants.HAlign.xlHAlignFill],
    'general': ['h', xw.constants.HAlign.xlHAlignGeneral],
    'hjustify': ['h', xw.constants.HAlign.xlHAlignJustify],
    'left': ['h', xw.constants.HAlign.xlHAlignLeft],
    'right': ['h', xw.constants.HAlign.xlHAlignRight],
    'bottom': ['v', xw.constants.VAlign.xlVAlignBottom],
    'vcenter': ['v', xw.constants.VAlign.xlVAlignCenter],
    'vdistributed': ['v', xw.constants.VAlign.xlVAlignDistributed],
    'vjustify': ['v', xw.constants.VAlign.xlVAlignJustify],
    'top': ['v', xw.constants.VAlign.xlVAlignTop],
}

_fpattern_map = {
    'none': 0,  # xlNone
    'solid': 1,  # xlSolid
    'gray50': 2,  # xlGray50
    'gray75': 3,  # xlGray75
    'gray25': 4,  # xlGray25
    'horstripe': 5,  # xlHorizontalStripe
    'verstripe': 6,  # xlVerticalStripe
    'diagstripe': 8,  # xlDiagonalDown
    'revdiagstripe': 7,  # xlDiagonalUp
    'diagcrosshatch': 9,  # xlDiagonalCrosshatch
    'thinhorstripe': 11,  # xlThinHorizontalStripe
    'thinverstripe': 12,  # xlThinVerticalStripe
    'thindiagstripe': 14,  # xlThinDiagonalDown
    'thinrevdiagstripe': 13,  # xlThinDiagonalUp
    'thinhorcrosshatch': 15,  # xlThinHorizontalCrosshatch
    'thindiagcrosshatch': 16,  # xlThinDiagonalCrosshatch
    'thickdiagcrosshatch': 10,  # xlThickDiagonalCrosshatch
    'gray12p5': 17,  # xlGray12.5
    'gray6p25': 18  # xlGray6.25
}

_border_side_map = {
    'none': None,
    'inner': None,
    'outer': None,
    'all': None,
    'left': 7,
    'top': 8,
    'bottom': 9,
    'right': 10,
    'inner_vert': 11,
    'inner_hor': 12,
    'down_diagonal': 5,
    'up_diagonal': 6
}

_border_style_map = {
    'continue': 1,
    'dash': 2,
    'dot': 3,
    'dash_dot': 4,
    'dash_dot_dot': 5,
    'slant_dash': 6,
    'thick_dash': 8,
    'double': 9,
    'thick_dash_dot_dot': 11
}

_border_weight_map = {
    'thin': 2,
    'thick': 3,
    'thicker': 4
}

_border_custom = {
    'none': None,
    'all_thin': ['all', 'continue', 'thin'],
    'all_thick': ['all', 'continue', 'thick'],
    'inner_thin': ['inner', 'continue', 'thin'],
    'inner_thick': ['inner', 'continue', 'thick'],
    'outer_thin': ['outer', 'continue', 'thin'],
    'outer_thick': ['outer', 'continue', 'thick']
}

_cpdpuxl_color_map = {
    "darkred": "#C00000",
    "red": "#FF0000",
    "orange": "#FFC000",
    "yellow": "#FFFF00",
    "lightgreen": "#92D050",
    "green": "#00B050",
    "lightblue": "#00B0F0",
    "blue": "#0070C0",
    "darkblue": "#002060",
    "purple": "#7030A0",
    "grey": "#808080",
    "grey25": "#BFBFBF",
    "white": "#FFFFFF",
    "bluegray": "#44546A",
    "msblue": "#4472C4",
    "msorange": "#ED7D31",
    "msgray": "#A5A5A5",
    "msyellow": "#FFC000",
    "mslightblue": "#5B9BD5",
    "msgreen": "#70AD47",
    "bluegray80": "#D6DCE4",
    "msblue80": "#D9E1F2",
    "msorange80": "#FCE4D6",
    "msgray80": "#EDEDED",
    "msyellow80": "#FFF2CC",
    "mslightblue80": "#DDEBF7",
    "msgreen80": "#E2EFDA",
    "bluegray60": "#ACB9CA",
    "msblue60": "#B4C6E7",
    "msorange60": "#F8CBAD",
    "msgray60": "#DBDBDB",
    "msyellow60": "#FFE699",
    "mslightblue60": "#BDD7EE",
    "msgreen60": "#C6E0B4",
}


def print_cell_attributes(file, sheet_name, lcrange):
    lcsheet = xw.Book(file).sheets[sheet_name]
    color_range = lcsheet.range(lcrange)

    cell_colors = {}
    for cell in color_range:
        cell_address = cell.address
        rgb_int = int(cell.api.Font.Color)
        red = rgb_int % 256
        green = (rgb_int // 256) % 256
        blue = (rgb_int // 256 ** 2) % 256
        hex_color = f"#{red:02X}{green:02X}{blue:02X}"
        cell_colors[cell_address] = hex_color

    for address, color in cell_colors.items():
        print(f"Cell {address} has color {color}")


def _extract_tuple(s):
    pattern = r'\((\d+,\s*\d+,\s*\d+)\)'
    matches = list(re.finditer(pattern, s))
    if len(matches) == 0:
        return None, s.strip()
    elif len(matches) == 1:
        match = matches[0]
        tuple_str = match.group(1)
        color_tuple = tuple(map(int, tuple_str.split(',')))
        remaining_str = s[:match.start()] + s[match.end():]
        return color_tuple, remaining_str.strip()
    else:
        raise ValueError(f"Multiple tuples found in '{s}'")


def _is_number(s: str):
    pattern = re.compile(r'^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$')
    return bool(pattern.match(s))


def hex_to_int(hex: str):
    hex = hex.lstrip('#')
    red = int(hex[:2], 16)
    green = int(hex[2:4], 16)
    blue = int(hex[4:], 16)
    return red | (green << 8) | (blue << 16)


def is_valid_hex_color(s):
    pattern = r'^#[0-9A-F]{6}$'
    return bool(re.match(pattern, s, re.IGNORECASE))


def _is_valid_rgb(rgb):
    if not isinstance(rgb, (list, tuple)) or len(rgb) != 3:
        return False
    return all(isinstance(n, int) and 0 <= n <= 255 for n in rgb)


class RangeOperator:

    def __init__(self, xwrange: xw.Range) -> None:
        self.xwrange = xwrange

    def format(
            self,
            font: str | tuple | list = None,
            font_name: str = None,
            font_size: str = None,
            font_color: str | tuple = None,
            italic: bool = None,
            bold: bool = None,
            underline: bool = None,
            strikeout: bool = None,
            align: str | list = None,
            merge: bool = None,
            wrap: bool = None,
            width = None,
            height = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            appendix: bool = False
    ) -> None:

        if appendix:
            print('Please choose one value from the corresponding parameter: \n'
                  f'align: {list(_alignment_map.keys())}; \n'
                  f'fill_pattern: {list(_fpattern_map.keys())};\n'
                  f'border_custom: {list(_border_custom.keys())};\n')

        # Font Attributes
        ##################################
        if font:
            if isinstance(font, tuple):
                self.xwrange.font.color = font
            elif isinstance(font, (int, float)):
                self.xwrange.font.size = font
            elif isinstance(font, str):
                color, remaining = _extract_tuple(font)
                if color:
                    self.xwrange.font.color = color
                for item in remaining.split(','):
                    item = item.strip()
                    if _is_number(item):
                        self.xwrange.font.size = item
                    elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                        self.xwrange.font.color = item
                    elif item == 'bold':
                        self.xwrange.font.bold = True
                    elif item == 'italic':
                        self.xwrange.font.italic = True
                    elif item == 'underline':
                        self.xwrange.api.Font.Underline = True
                    elif item == 'strikeout':
                        self.xwrange.api.Font.Strikethrough = True
                    else:
                        self.xwrange.font.name = item
            elif isinstance(font, list):
                for item in font:
                    if isinstance(item, tuple):
                        self.xwrange.font.color = item
                    elif isinstance(item, (int, float)):
                        self.xwrange.font.size = item
                    elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                        self.xwrange.font.color = item
                    elif isinstance(item, str) and item == 'bold':
                        self.xwrange.font.bold = True
                    elif isinstance(item, str) and item == 'italic':
                        self.xwrange.font.italic = True
                    elif isinstance(item, str) and item == 'underline':
                        self.xwrange.api.Font.Underline = True
                    elif isinstance(item, str) and item == 'strikeout':
                        self.xwrange.api.Font.Strikethrough = True
                    else:
                        self.xwrange.font.name = item

        if font_name:
            self.xwrange.font.name = font_name

        if font_size is not None:
            self.xwrange.font.size = font_size

        if font_color:
            self.xwrange.font.color = font_color

        if italic is not None:
            self.xwrange.font.italic = italic

        if bold is not None:
            self.xwrange.font.bold = bold

        if underline is not None:
            self.xwrange.api.Font.Underline = underline

        if strikeout is not None:
            self.xwrange.api.Font.Strikethrough = strikeout

        # Align Attributes
        ##################################
        def _alignfunc(alignkey):
            if alignkey in ['center', 'justify', 'distributed']:
                self.xwrange.api.VerticalAlignment = _alignment_map['v' + alignkey][1]
                self.xwrange.api.HorizontalAlignment = _alignment_map['h' + alignkey][1]
            elif _alignment_map[alignkey][0] == 'v':
                self.xwrange.api.VerticalAlignment = _alignment_map[alignkey][1]
            elif _alignment_map[alignkey][0] == 'h':
                self.xwrange.api.HorizontalAlignment = _alignment_map[alignkey][1]
            elif align not in _alignment_map.keys():
                raise ValueError(f'Alignment {alignkey} is not supported')
            return

        if align:
            if isinstance(align, str):
                for item in align.split(','):
                    item = item.strip()
                    _alignfunc(item)
            elif isinstance(align, list):
                for item in align:
                    _alignfunc(item)

        # Merge and Wrap Attributes
        ##################################
        if merge:
            xw.apps.active.api.DisplayAlerts = False
            self.xwrange.api.MergeCells = merge
            _alignfunc('center')
            xw.apps.active.api.DisplayAlerts = True

        # noinspection PySimplifyBooleanCheck
        if merge == False:
            if self.xwrange.api.MergeCells:
                self.xwrange.unmerge()

        if wrap is not None:
            self.xwrange.api.WrapText = wrap

        # Width and Height Attributes
        ##################################
        '''
        default width and height for a excel cell is 8.54 (around ...) and 14.6
        '''
        if width:
            self.xwrange.api.EntireColumn.ColumnWidth = width

        if height:
            self.xwrange.api.RowHeight = height

        # Border Attributes
        ##################################
        if border:
            border_side = 'all'
            border_style = 'continue'
            weight = 1

            if isinstance(border, str) and border.strip() == 'none':
                for i in range(1, 12):
                    self.xwrange.api.Borders(i).LineStyle = 0
            else:
                if isinstance(border, str) and border.strip() in list(_border_custom.keys()):
                    border_para = _border_custom[border.strip()]
                elif isinstance(border, str):
                    border_para = [i.strip() for i in border.split(',')]
                elif isinstance(border, list):
                    border_para = [i.strip() for i in border]
                else:
                    raise ValueError(
                        'Invalid boarder specification, please use check_para=True to see the valid lists.')

                for item in border_para:
                    if isinstance(item, str) and item in list(_border_weight_map.keys()):
                        weight = _border_weight_map[item]
                    elif isinstance(item, str) and item in list(_border_side_map.keys()):
                        border_side = item
                    elif isinstance(item, str) and item in list(_border_style_map.keys()):
                        border_style = item
                    else:
                        raise ValueError(
                            'Invalid boarder specification, please use check_para=True to see the valid lists.')

                if border_side == 'none':
                    for i in range(1, 12):
                        self.xwrange.api.Borders(i).LineStyle = 0
                elif border_side == 'all':
                    self.xwrange.api.Borders.LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders.Weight = weight
                elif border_side == 'inner':
                    self.xwrange.api.Borders(11).LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders(11).Weight = weight
                    self.xwrange.api.Borders(12).LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders(12).Weight = weight
                elif border_side == 'outer':
                    for i in range(7, 11):
                        self.xwrange.api.Borders(i).LineStyle = _border_style_map[border_style]
                        self.xwrange.api.Borders(i).Weight = weight
                elif border_side in _border_side_map.keys():
                    self.xwrange.api.Borders(_border_side_map[border_side]).LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders(_border_side_map[border_side]).Weight = weight

        # Fill Attributes
        ##################################
        if fill:
            if isinstance(fill, list):
                patternlist, colorlist = [], []

                for item in fill:
                    if isinstance(item, (tuple, list, str)) and item in _fpattern_map.keys():
                        patternlist.append(item)
                    if isinstance(item, str) and is_valid_hex_color(item):
                        colorlist.append(item)
                    elif isinstance(item, (list, tuple)) and _is_valid_rgb(item):
                        colorlist.append(item)

                leftover = [item for item in fill if item not in patternlist + colorlist]
                if len(leftover) > 0:
                    raise ValueError('Invalid input. Please check if pattern or color are specified correctly.')
                else:
                    if (len(fill) == 1) or (len(fill) == 2 and 'solid' in fill):
                        for item in fill:
                            if isinstance(item, tuple):
                                self.xwrange.api.Interior.Color = xw.utils.rgb_to_int(item)
                            elif item in list(_fpattern_map.keys()):
                                self.xwrange.api.Interior.Pattern = _fpattern_map[item]
                            elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                                self.xwrange.api.Interior.Color = hex_to_int(item)
                    elif len(fill) == 2 and 'solid' not in fill:
                        for item in fill:
                            if isinstance(item, tuple):
                                self.xwrange.api.Interior.PatternColor = xw.utils.rgb_to_int(item)
                            elif item in list(_fpattern_map.keys()):
                                self.xwrange.api.Interior.Pattern = _fpattern_map[item]
                            elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                                self.xwrange.api.Interior.PatternColor = hex_to_int(item)
                    else:
                        raise ValueError(
                            "Can only accept 2 parameters (one for pattern and one for color) at most when passing a list object to 'fill'.")

            elif isinstance(fill, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill)
                self.xwrange.api.Interior.Color = foreground_color_int
            elif isinstance(fill, str):
                patternkeys = '(' + '|'.join(_fpattern_map.keys()) + ')'
                compiled_patternkeys = re.compile(patternkeys, re.IGNORECASE)
                patternlist = re.findall(compiled_patternkeys, fill)
                firstpattern = patternlist[0] if len(patternlist) >= 1 else None

                colorrule = r'#(?:[0-9a-fA-F]{3}){1,2}|\(\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*,\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*,\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*\)'
                colorlist = re.findall(colorrule, fill)

                leftover = fill
                for item in patternlist + colorlist:
                    leftover = leftover.replace(item, '').strip()

                if len(leftover.replace(',', '')) > 0:
                    raise ValueError('Incorrect pattern or color specified, please check')
                elif len(patternlist) > 1 or len(colorlist) > 1:
                    raise ValueError('Can not specify more than one color or more than one pattern, respectively.')
                else:
                    if len(patternlist) == 1:
                        self.xwrange.api.Interior.Pattern = _fpattern_map[patternlist[0]]
                    if len(colorlist) == 1:
                        if '#' in colorlist[0]:
                            color_int = hex_to_int(colorlist[0])
                        else:
                            colortuple = tuple(map(int, colorlist[0].replace('(', '').replace(')', '').split(',')))
                            color_int = xw.utils.rgb_to_int(colortuple)

                        if firstpattern is None or firstpattern == 'solid':
                            self.xwrange.api.Interior.Color = color_int
                        else:
                            self.xwrange.api.Interior.PatternColor = color_int

        if fill_pattern:
            self.xwrange.api.Interior.Pattern = _fpattern_map[fill_pattern]

        if isinstance(fill_fg, tuple):
            foreground_color_int = xw.utils.rgb_to_int(fill_fg)
            self.xwrange.api.Interior.PatternColor = foreground_color_int
        elif isinstance(fill_fg, str):
            self.xwrange.api.Interior.PatternColor = hex_to_int(fill_fg)

        if isinstance(fill_bg, tuple):
            background_color_int = xw.utils.rgb_to_int(fill_bg)
            self.xwrange.api.Interior.Color = background_color_int
        elif isinstance(fill_bg, str):
            self.xwrange.api.Interior.Color = hex_to_int(fill_bg)

        return

    def clear(self):
        self.xwrange.clear()


class cpdStyle:
    def __init__(self, **kwargs):
        self.format_dict = kwargs

    def __hash__(self):
        # Convert the format_dict to a tuple of items, which is hashable, and then hash it
        # Note: This assumes that all values in the dictionary are also hashable
        return hash(tuple(sorted(self.format_dict.items())))

    def __eq__(self, other):
        # Check if the other object is an instance of cpdStyle and if their format_dicts are equal
        return isinstance(other, cpdStyle) and self.format_dict == other.format_dict


def parse_format_rule(rule):
    if isinstance(rule, cpdStyle):
        return rule.format_dict

    elif not isinstance(rule, str):
        raise ValueError('format prompt key word must be str')

    promptlist = [prompt.strip() for prompt in rule.split(',')]
    return_dict = {}

    def _parse_str_format_key(prompt):
        result = {}
        keysmatch = {
            'italic': {'italic': True},
            'noitalic': {'italic': False},
            'bold': {'bold': True},
            'nobold': {'bold': False},
            'underline': {'underline': True},
            'nounderline': {'underline': False},
            'strikeout': {'strikeout': True},
            'nostrikeout': {'strikeout': False},
            'merge': {'merge': True},
            'nomerge': {'merge': False},
            'wrap': {'wrap': True},
            'nowrap': {'wrap': False},
        }
        patterns = {
            r'font_name=(.*)': ['font_name', lambda local_match: local_match.group(1)],
            r'font_size=(.*)': ['font_size', lambda local_match: float(local_match.group(1))],
            r'font_color=(.*)': ['font_color', lambda local_match: local_match.group(1)],
            r'align=(.*)': ['align', lambda local_match: local_match.group(1)],
            r'width=(.*)': ['width', lambda local_match: float(local_match.group(1))],
            r'height=(.*)': ['height', lambda local_match: float(local_match.group(1))],
            r'border=(.*)': ['border', lambda local_match: float(local_match.group(1))],
            r'(#[A-Z0-9]{6})': ['fill', lambda local_match: local_match.group(1)],
        }

        if prompt in keysmatch.keys():
            result.update(keysmatch['prompt'])

        if prompt in _cpdpuxl_color_map.keys():
            lc_hex = _cpdpuxl_color_map[prompt]
            result.update({'fill': lc_hex})

        for pattern, value in patterns.items():
            match = re.fullmatch(pattern, prompt)
            if match:
                append_dict = {value[0]: value[1](match)}
                result.update(append_dict)

        return result

    for term in promptlist:
        return_dict.update(_parse_str_format_key(term))

    return return_dict


if __name__ == '__main__':
    wb = xw.Book('sampledf.xlsx')
    sheet = wb.sheets[0]  # Reference to the first sheet

    # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
    # my_range = sheet.range("H2:I4")

    # Step 3: Create an object of the RangeOperator class with the specified range
    # a = RangeOperator(my_range)
    # a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border='outer, thicker')    # print(a.range)
    # a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border=['inner', 'thin'])    # print(a.range)
    # a.format(width=20, height=15)    # print(a.range)

    # my_range = sheet.range("A1:B12")
    # a = RangeOperator(my_range)
    # style = cpdStyle(font=['bold', 'strikeout', 12.5, (0,0,0)])
    # a.format(**style.format_dict)

    print_cell_attributes('sampledf.xlsx', 'Sheet3', 'A1:A34')
    print(parse_format_rule('red, font_size=12'))
