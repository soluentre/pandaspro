style_sheets = {
    # Text Styles
    'heading1': {
        'font_size=14': 'cell',
        'bold': 'cell',
        'font_color=#0070C0': 'cell'
    },
    'heading2': {
        'font_size=12': 'cell',
        'bold': 'cell',
        'font_color=#0070C0': 'cell'
    },
    'note1': {
        'font_size=10': 'cell',
        'font_color=#A6A6A6': 'cell',
        'italic': 'cell'
    },

    # DataFrame Styles
    'black': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'black; font_color=white': 'header_outer'
    },
    'blue': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'fill=#8ABDFF; font_color=black; wrap; bold': 'header_outer'
    },
    'bluelist': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'fill=#8ABDFF; font_color=black; wrap; bold; height=30': 'header_outer'
    },
    'darkblue': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': ['all', 'header_outer'],
        'blue80; font_color=white; wrap': 'header_outer'
    },
    'green': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'green80; font_color=black; wrap': 'header_outer'
    },
    'index_merge': {
        'merge': 'index_merge_inputs(level=__index__, columns=__columns__)',
        'border=outer_thick': [
            'index_levels',
            'index_hsections(level=__index__)'
        ]
    },
    'total': {
        'border=outer_thick': ['columns(c=Total, header=True)', 'index_outer'],
    },
    'pivot_gray_lastrow': {
        'border=outer_thick; fill=#EDEDED': 'bottom1'
    }
}