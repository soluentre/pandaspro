import sys
import warnings

if sys.version_info >= (3, 12):
    warnings.warn(
        "\n\n⚠️ pandaspro is currently not supporting Python 3.12 or above, please use 3.8 - 3.11 versions\n",
        UserWarning,
        stacklevel=2
    )

from pandaspro.email.api import (
    emailfetcher,
    create_mail_class
)

# noinspection PyProtectedMember
from pandaspro.core.api import (
    bdate,
    dfilter,
    FramePro,
    lowervarlist,
    tab,
    str2list,
    wildcardread,
    parse_method,
    parse_wild,
    df_with_index_for_mask,
    create_column_color_dict,
    csort,
    cpdBaseFrameMapper,
    cpdBaseFrameList,
    consecgrouper,
    replace_left_with_right
)

from pandaspro.cpdbase.api import (
    cpdBaseFrame,
    FilesVersionParser
)

from pandaspro.date.api import (
    DatePro
)

from pandaspro.io.api import (
    CellPro,
    index_cell,
    cell_index,
    resize,
    offset,
    getrange,
    PutxlSet,
    pwread,
    WorkbookExportSimplifier,
    fw
)

from pandaspro.sampledf.api import (
    sysuse_countries,
    sysuse_auto,
    wbuse_pivot
)

from pandaspro.pdfs.api import (
    merge_pdfs
)

excel_d = WorkbookExportSimplifier().declare_workbook
