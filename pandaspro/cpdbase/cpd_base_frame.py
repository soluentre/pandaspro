import pandas as pd
import pandaspro
import pandaspro as cpd
from abc import ABC
import inspect
from pandaspro import FramePro, cpdBaseFrameMapper
from pandaspro.io.excel.putexcel import PutxlSet
from pandaspro.cpdbase.design import cpdBaseFrameDesign
from pandaspro.cpdbase.files_version_parser import FilesVersionParser
import textwrap


# from pandaspro.utils.cpd_logger import cpdLogger


def extract_params(func):
    signature = inspect.signature(func)
    pos_params = [
        name for name, param in signature.parameters.items()
        if param.default == inspect.Parameter.empty
    ]
    kw_params_with_defaults = {
        name: param.default for name, param in signature.parameters.items()
        if param.default != inspect.Parameter.empty
    }
    return pos_params, kw_params_with_defaults


def cpdBaseFrame(
        path: str = None,
        file_type: str = 'csv',
        load = None,
        default_version: str = 'latest',
        prefix: str = None,
        dateid: str = '%Y%m%d',
        sheet_name: str | int = 0,
        cellrange: str = 'A1',
        fiscal_year_end: str = '06-30',
        uid: str = None,
        rename_status: str = 'Process',
        imr: dict = None,
        exr: dict = None,
        dvl: str | list = None,
        export_file: str = None,
        export_status: bool = False,
        ps: pandaspro.io.excel.putexcel.PutxlSet = None,
        **custom_attrs
):
    def decorator(myclass):
        class CombinedClass(myclass, cpdBaseFrameDesign, FramePro, ABC):
            @classmethod
            def get_file_versions_parser(cls):
                this_prefix = myclass.__name__ if prefix is None else prefix
                fvp = FilesVersionParser(
                    path=cls.get_path(),
                    class_prefix=this_prefix,
                    dateid_expression=dateid,
                    file_type=file_type,
                    fiscal_year_end=fiscal_year_end
                )
                # print(fvp.list_all_files())
                return fvp

            @classmethod
            def get_path(cls):
                if path and (hasattr(myclass, 'get_path') or hasattr(myclass, 'path')):
                    raise AttributeError(
                        'Declaring both path argument (1) in @decorator and (2) in-class path/get_path at the same time is not allowed, please only declare one.')
                elif path and not (hasattr(myclass, 'get_path') or hasattr(myclass, 'path')):
                    return path
                else:
                    if hasattr(myclass, 'get_path') and hasattr(myclass, 'path'):
                        raise AttributeError('Declaring in-class path and get_path is not allowed, only declare one.')
                    elif hasattr(myclass, 'path'):
                        return myclass.path
                    elif hasattr(myclass, 'get_path'):
                        return myclass.get_path()
                    else:
                        raise TypeError(
                            "Can't instantiate abstract class MyConcreteClass with abstract method get_path.")

            @classmethod
            def read_table(cls, version):
                this_fvp = cls.get_file_versions_parser()
                filename = this_fvp.get_file(version)
                file_fullpath = cls.get_path() + f'/{filename}'

                if file_type == 'csv':
                    return cpd.pwread(file_fullpath, cellrange=cellrange, low_memory=False)
                elif file_type == 'xlsx':
                    try:
                        input_data = cpd.pwread(file_fullpath, sheet_name=sheet_name, cellrange=cellrange)
                    except Exception as e:
                        ps = PutxlSet(file_fullpath)
                        ps.close()
                        input_data = cpd.pwread(file_fullpath, sheet_name=sheet_name, cellrange=cellrange)
                        PutxlSet(file_fullpath)
                    return input_data
                else:
                    raise ValueError('Invalid file type, can only read .csv/.xlsx format.')

            @staticmethod
            def load(data, **kwargs):
                return data

            @classmethod
            def get_process_method(cls):
                if load and hasattr(myclass, 'load'):
                    raise AttributeError(
                        'Declaring both load argument (1) in @decorator and (2) in-class load attribute/method at the same time is not allowed, please only declare one.')
                elif hasattr(myclass, 'load'):
                    return myclass.load
                elif load:
                    return load
                else:
                    return CombinedClass.load

            def __init__(self, *args, **kwargs):
                cpd_kwargs = extract_params(CombinedClass.get_process_method())[1]
                uid_kwarg = {'uid': kwargs.pop('uid', uid)}
                fvp_kwarg = {'fvp': kwargs.pop('fvp', CombinedClass.get_file_versions_parser())}
                version_kwarg = {'version': kwargs.pop('version', default_version)}
                rename_status_kwarg = {'rename_status': kwargs.pop('rename_status', rename_status)}
                import_rename_kwarg = {'import_rename': kwargs.pop('import_rename', imr)}
                export_rename_kwarg = {'export_rename': kwargs.pop('export_rename', exr)}

                # Report exporting ...
                export_path_kwarg = {'export_file': kwargs.pop('export_file', export_file)}
                export_status_kwarg = {'export_status': kwargs.pop('export_status', export_status)}
                export_ps_kwarg = {'ps': kwargs.pop('ps', ps)}

                # Custom attrs
                custom_attrs_saver = {}
                for attr_name, attr_value in custom_attrs.items():
                    custom_attrs_saver[attr_name] = kwargs.pop(attr_name, attr_value)
                other_kwargs = {key: kwargs.pop(key, value) for key, value in cpd_kwargs.items()}

                # self.debug.info(f'[cpd_kwargs]: {cpd_kwargs}')
                # self.debug.info(f'[version_kwarg]: {version_kwarg}')
                # self.debug.info(f'[other_kwargs]: {other_kwargs}')
                # self.debug.info(f'[kwargs]: {kwargs}')
                # self.debug.info(f'[args]: {args}')

                if args or kwargs:
                    name_map = {}
                    # self.debug_info_lv1('Inside __init__')
                    # self.logger.info(f'Entered Above Part of init: args: **{type(args)}**, kwargs: **{type(kwargs)}**')
                    # self.logger.debug(f'Seeing values -> args: **{args}**, kwargs: **{kwargs}**')
                    try:
                        super(CombinedClass, self).__init__(*args, uid=uid, rename_status=rename_status, **kwargs)
                    except ValueError as e:
                        raise ValueError(textwrap.dedent(f'''
                            --------------------------------------
                            {e}
                            Please only pass key-word arguments when you want to create new instances of the <{myclass.__name__}>. 
                            You are passing {args} as positional arguments, which can not be parsed by the pandas DataFrame constructor
                        '''))
                    except TypeError as e:
                        raise TypeError(textwrap.dedent(f'''
                            --------------------------------------
                            {e}
                            Please check that you don't pass any extra key-word arguments besides those you declared in load defined in class <{myclass.__name__}>.
                            For the load method defined, the class constructor can only take the following kwargs: {list(other_kwargs.keys())}  
                        '''))
                else:
                    # self.logger.info('Entered Below Part of init: no args or kwargs detected')
                    raw_frame, name_map = CombinedClass.read_table(**version_kwarg)
                    # 先应用 import_rename，再传给 load 方法处理
                    if import_rename_kwarg['import_rename'] is not None:
                        raw_frame = raw_frame.rename(columns=import_rename_kwarg['import_rename'])
                    processed_frame = CombinedClass.get_process_method()(raw_frame, **other_kwargs)
                    super(CombinedClass, self).__init__(processed_frame, uid=uid, rename_status=rename_status)  # Ensure DataFrame initialization

                self.fvp = fvp_kwarg['fvp']
                self.version = version_kwarg['version']
                self.uid = uid_kwarg['uid']
                self.rename_status = rename_status_kwarg['rename_status']
                self.import_mapper = cpdBaseFrameMapper(import_rename_kwarg['import_rename'])
                
                # 优化 export_mapper: 如果 import_rename 已经应用，需要更新 export_mapper 的键
                # name_map 结构: {当前小写列名: Excel原始列名}
                # 目标: export_mapper = {最终当前列名: Excel原始列名}
                # noinspection PyUnboundLocalVariable
                final_export_map = name_map.copy()
                
                # 如果 import_rename 存在，需要将 export_mapper 的键更新为 import_rename 后的列名
                if import_rename_kwarg['import_rename'] is not None:
                    updated_export_map = {}
                    for current_col, excel_col in final_export_map.items():
                        # 如果当前列名在 import_rename 中被进一步重命名
                        if current_col in import_rename_kwarg['import_rename']:
                            # 使用 import_rename 后的最终列名作为 key
                            final_col = import_rename_kwarg['import_rename'][current_col]
                            updated_export_map[final_col] = excel_col
                        else:
                            # 如果没有被重命名，保持原样
                            updated_export_map[current_col] = excel_col
                    final_export_map = updated_export_map
                
                # 应用额外的 export_rename 配置
                if export_rename_kwarg['export_rename'] is not None:
                    final_export_map.update(export_rename_kwarg['export_rename'])
                
                self.export_mapper = cpdBaseFrameMapper(final_export_map)

                self.get_filename = self.fvp.get_file(self.version)
                self.get_filename_full = self.get_path() + '/' + self.get_filename
                self.get_version = self.fvp.get_file_version_str(self.version)
                self.get_vo = self.fvp.get_file_version_dt(self.version)
                self.vo = self.get_vo
                self.get_more_info = self.fvp.get_suffix(self.version)

                # Report Exporting
                self.export_file = export_path_kwarg['export_file']
                self.export_status = export_status_kwarg['export_status']
                self.ps = export_ps_kwarg['ps']

                # Custom Attributes
                for attr_name, attr_value in custom_attrs.items():
                    # setattr(self, attr_name, attr_value)
                    self.__dict__[attr_name] = attr_value
                self.custom_attrs_saver = cpdBaseFrameMapper(custom_attrs_saver)

            @property
            def _constructor(self):
                def _c(*args, **kwargs):
                    custom_kwargs = {key: getattr(self, key) for key in self.custom_attrs_saver.dict}
                    kwargs.update(custom_kwargs)
                    return CombinedClass(
                        *args,
                        version=self.version,
                        uid=self.uid,
                        rename_status=self.rename_status,
                        import_rename=self.import_mapper.dict,
                        export_rename=self.export_mapper.dict,
                        export_file=self.export_file,
                        export_status=self.export_status,
                        ps=self.ps,
                        **kwargs
                    )
                return _c

            def __getattr__(self, item):
                override_list = []

                if hasattr(super(self.__class__, self), item) and not item.startswith(tuple(override_list)):
                    return getattr(super(self.__class__, self), item)

                # elif item.startswith('cpdpvt_'):
                #     pivot_info = item[4:]
                #     variables = pivot_info.split('__')
                #
                #     if len(variables) == 2:
                #         pivot_index, pivot_columns = variables
                #
                #         if self.uid is None:
                #             idvar = self.columns[self.notnull().all()].tolist()
                #         else:
                #             idvar = self.uid
                #
                #         if self.rename_status == 'Export':
                #             pivot_index = self.export_mapper.dict[pivot_index]
                #             pivot_columns = self.export_mapper.dict[pivot_columns]
                #             idvar = self.export_mapper.dict[idvar]
                #
                #         return FramePro(
                #             self.pivot_table(
                #                 index=pivot_index,
                #                 columns=pivot_columns,
                #                 values=idvar,
                #                 aggfunc='count',
                #                 margins=True,
                #                 margins_name='Total'
                #             )
                #         )
                #     else:
                #         raise ValueError('pvt_ for sob class must have 2 vars seperated by double underline mark __')
                # elif item.startswith('quickview_'):

                else:
                    return super().__getattr__(item)

            @property
            def er(self):
                self.rename_status = 'Export'
                return self.rename(columns=self.export_mapper.dict)

            def export_build(self, update_export_path: str = None):
                if update_export_path is not None and isinstance(update_export_path, str):
                    self.export_file = update_export_path

                if isinstance(self.export_file, str):
                    self.ps = cpd.PutxlSet(self.export_file)
                    self.export_status = True
                else:
                    raise ValueError('export attribute for the object is not str and can not be parsed by PutxlSet')

            @property
            def _parse_default_view_list(self):
                if dvl is None:
                    return self.columns.to_list()
                else:
                    if isinstance(dvl, str):
                        return self.cvar(dvl)
                    elif isinstance(dvl, list):
                        return dvl
                    else:
                        raise TypeError('Invalid object type for default_view_list parameter')

            @property
            def dvl(self):
                viewlist = self._parse_default_view_list
                return self[viewlist]

            def _parse_adjust_column(self, varlist):
                # Decide the adjust_column (list of vars)
                if isinstance(varlist, str):
                    adjust_column = cpd.parse_wild(varlist, self.columns)
                elif isinstance(varlist, list):
                    adjust_column = varlist
                else:
                    raise ValueError('input varlist must be either str or list')
                pd.set_option('Display.Max_columns', None)
                return adjust_column

            def dvlmore(self, varlist):
                if dvl is None:
                    raise ValueError(
                        "CANNOT CALL: [bmore] method cannot be called when default_list_view is at its default value None, which indicates method [b] will return all columns and no need to browse more columns")

                adjust_column = self._parse_adjust_column(varlist)
                return self[self._parse_default_view_list + adjust_column]

            def dvlless(self, varlist):
                adjust_column = self._parse_adjust_column(varlist)
                original_list = self._parse_default_view_list
                new_list = [item for item in original_list if item not in adjust_column]
                return self[new_list]

        CombinedClass.__name__ = myclass.__name__

        return CombinedClass

    return decorator


if __name__ == '__main__':
    @cpdBaseFrame(
        path=r'C:\Users\wb539289\OneDrive - WBG\K - Knowledge Management\Databases\Staff on Board Database\csv',
        export_file=r'C:\Users\wb539289\OneDrive - WBG\Desktop\temp_data.xlsx',
        default_version='latest_month',
        uid='upi',
        msg=123
    )
    class SOB(pd.DataFrame):

        @staticmethod
        def load(data, region=None):
            print(region)
            return data

        # noinspection PyAttributeOutsideInit
        def update_msg(self):
            self.msg = 1

    df1 = SOB(region='balabala')
    df1.export_build()
    # print(df1.vo)
    # v = df1.vo
    # df2 = df1.inlist('upi', 83315)
    # print(df2.cpdtabd_gender)
