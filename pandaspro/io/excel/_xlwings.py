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


def extract_tuple(s):
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


def _is_valid_hex_color(s):
    pattern = r'^#[0-9A-F]{6}$'
    return bool(re.match(pattern, s, re.IGNORECASE))


def _is_valid_rgb(rgb):
    if not isinstance(rgb, (list, tuple)) or len(rgb) != 3:
        return False
    return all(isinstance(n, int) and 0 <= n <= 255 for n in rgb)


def color_to_int(color: str | tuple):
    if isinstance(color, str) and _is_valid_hex_color(color):        
        hex = color.lstrip('#')
        red = int(hex[:2], 16)
        green = int(hex[2:4], 16)
        blue = int(hex[4:], 16)
        return red | (green << 8) | (blue << 16)
    elif isinstance(color, tuple) and _is_valid_rgb(color):
        return xw.utils.rgb_to_int(color)
    else:
        return None


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
                color, remaining = extract_tuple(font)
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

        # Merge Attributes
        ##################################
        if merge:
            xw.apps.active.api.DisplayAlerts = False
            self.xwrange.api.MergeCells = merge
            _alignfunc('center')
            xw.apps.active.api.DisplayAlerts = True

        elif not merge:
            if self.xwrange.api.MergeCells:
                self.xwrange.unmerge()

        # Border Attributes
        ##################################
        if border:
            border_side = 'all'
            border_style = 'continue'
            weight = 1

            if isinstance(border, str) and border.strip() == 'none':
                for i in range(1, 12):
                    self.xwrange.api.Borders(i).LineStyle = 0

            if isinstance(border, str) and border.strip() in list(_border_custom.keys()):
                border_para = _border_custom[border.strip()]

            elif isinstance(border, str) and border.strip() != 'none':

                if extract_tuple(border)[0]:
                    border = border.replace(' ', '')
                    tuple_str = str(extract_tuple(border)[0]).replace(' ', '')
                    border_para = [i for i in [i.strip() for i in border.replace(tuple_str, '').split(',')] if i != '']
                    border_para.append(extract_tuple(border)[0])

                elif _is_valid_hex_color(border):
                    color_tpl = color_to_int(border)
                    border_color = xw.utils.rgb_to_int(color_tpl)

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
                for i in range(7,11):
                    self.xwrange.api.Borders(i).LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders(i).Weight = weight

            elif border_side in _border_side_map.keys():
                self.xwrange.api.Borders(_border_side_map[border_side]).LineStyle = _border_style_map[border_style]
                self.xwrange.api.Borders(_border_side_map[border_side]).Weight = weight

        # Fill Attributes
        ##################################
        if fill:
            def fill_with_mylist(fill_list):
                def find_pattern(mylist):
                    result = []
                    for local_item in mylist:
                        if isinstance(local_item, (tuple, list, str)) and local_item in _fpattern_map.keys():
                            result.append(local_item)
                    return result

                def find_colors(mylist):
                    result = []
                    for local_item in mylist:
                        if isinstance(local_item, str) and _is_valid_hex_color(local_item):
                            result.append(local_item)
                        elif isinstance(local_item, (list, tuple)) and _is_valid_rgb(local_item):
                            result.append(local_item)
                    return result

                # Parse the list and get the Pattern and Color Lists (should be only 1 or none)
                patternlist = find_pattern(fill_list)
                colorlist = find_colors(fill_list)
                leftover = [item for item in fill_list if item not in patternlist + colorlist]
                if len(leftover) > 0 or len(patternlist) > 1 or len(colorlist) > 1:
                    raise ValueError(
                        'Invalid input. Please check if pattern or color are specified correctly. At most 1 color and 1 pattern')

                # Create patter and color parameter
                pattern = patternlist[0] if len(patternlist) == 1 else None
                color = colorlist[0] if len(colorlist) == 1 else None

                if pattern:
                    self.xwrange.api.Interior.Pattern = _fpattern_map[pattern]

                if color:
                    if pattern == 'solid':
                        self.xwrange.api.Interior.Color = color_to_int(color)
                    else:
                        self.xwrange.api.Interior.PatternColor = color_to_int(color)

            if isinstance(fill, list):
                fill_with_mylist(fill)

            elif isinstance(fill, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill)
                self.xwrange.api.Interior.Color = foreground_color_int

            elif isinstance(fill, str):
                def parse_my_str():
                    pass
                fill_list_from_str = parse_my_str(fill)
                fill_with_mylist(fill_list_from_str)


                # patternkeys = '(' + '|'.join(_fpattern_map.keys()) + ')'
                # compiled_patternkeys = re.compile(patternkeys, re.IGNORECASE)
                # patternlist = re.findall(compiled_patternkeys, fill)
                # firstpattern = patternlist[0] if len(patternlist) >= 1 else None
                #
                # colorrule = r'#(?:[0-9a-fA-F]{3}){1,2}|\(\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*,\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*,\s*(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\s*\)'
                # colorlist = re.findall(colorrule, fill)
                #
                # leftover = fill
                # for item in patternlist + colorlist:
                #     leftover = leftover.replace(item, '').strip()
                #
                # if len(leftover.replace(',', '')) > 0:
                #     raise ValueError('Incorrect pattern or color specified, please check')
                # elif len(patternlist) > 1 or len(colorlist) > 1:
                #     raise ValueError('Can not specify more than one color or more than one pattern, respectively.')
                # else:
                #     if len(patternlist) == 1:
                #         self.xwrange.api.Interior.Pattern = _fpattern_map[patternlist[0]]
                #     if len(colorlist) == 1:
                #         if '#' in colorlist[0]:
                #             colorint = color_to_int(colorlist[0])
                #         else:
                #             colortuple = tuple(map(int, colorlist[0].replace('(', '').replace(')', '').split(',')))
                #             colorint = xw.utils.rgb_to_int(colortuple)
                #
                #         if firstpattern is None or firstpattern == 'solid':
                #             self.xwrange.api.Interior.Color = colorint
                #         else:
                #             self.xwrange.api.Interior.PatternColor = colorint

        if fill_pattern:
            self.xwrange.api.Interior.Pattern = _fpattern_map[fill_pattern]

        if fill_fg:
            if isinstance(fill_fg, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill_fg)
                self.xwrange.api.Interior.PatternColor = foreground_color_int
            elif isinstance(fill_fg, str):
                self.xwrange.api.Interior.PatternColor = color_to_int(fill_fg)

        if fill_bg:
            if isinstance(fill_bg, tuple):
                background_color_int = xw.utils.rgb_to_int(fill_bg)
                self.xwrange.api.Interior.Color = background_color_int
            elif isinstance(fill_bg, str):
                self.xwrange.api.Interior.Color = color_to_int(fill_bg)

        return

    def clear(self):
        self.xwrange.clear()


if __name__ == '__main__':
    wb = xw.Book('sampledf.xlsx')
    sheet = wb.sheets[0]  # Reference to the first sheet

    # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
    my_range = sheet.range("H2:I4")

    # Step 3: Create an object of the RangeOperator class with the specified range
    a = RangeOperator(my_range)
    a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border='outer, thicker')    # print(a.range)
    a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border=['inner', 'thin'])    # print(a.range)


