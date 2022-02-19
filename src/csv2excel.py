import os
from pathlib import Path
from typing import List, Tuple
import shutil
from copy import copy
from multiprocessing import Pool
from functools import partial

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles.protection import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from tqdm import tqdm


class CSV2ExcelProcessor:

    @staticmethod
    def _get_files(dir_path: str, files_type: List[str] = 'csv') -> List[Path]:
        files_type = [files_type] if files_type is str else files_type
        files = os.listdir(dir_path)
        files = [Path(dir_path) / file for file in files if file.split('.')[-1] in files_type]
        return files

    @staticmethod
    def get_max_csv_cordinality(dir_path: str):
        files = CSV2ExcelProcessor._get_files(dir_path)
        max_rows = 0
        for file in tqdm(files):
            df = pd.read_csv(file, header=None)
            if max_rows < df.shape[0]:
                max_rows = df.shape[0]
        return max_rows

    @staticmethod
    def _replace_escapes(text: str) -> str:
        escapes = ''.join([chr(char) for char in range(1, 32)])
        translator = str.maketrans('', '', escapes)
        return text.translate(translator)

    @staticmethod
    def _filter_ma_indicators(df: pd.DataFrame) -> pd.DataFrame:
        filter_values = [509, 510, 511, 512]
        return df[~df[6].isin(filter_values)]

    @staticmethod
    def _load_csv(csv_path: str, columns_to_int: List[int]) -> pd.DataFrame:
        csv_data = pd.read_csv(csv_path, header=None)
        csv_data = csv_data.loc[1:, :]
        csv_data.fillna('', inplace=True)
        csv_data = csv_data.applymap(CSV2ExcelProcessor._replace_escapes)
        for idx in columns_to_int:
            csv_data[idx] = pd.to_numeric(csv_data[idx].apply(lambda x: x.split('.')[0]))
        csv_data = CSV2ExcelProcessor._filter_ma_indicators(csv_data)
        return csv_data

    # @staticmethod
    # def _load_excel(excel_path: str, sheet_list: List[str], engine: str = 'pyxlsb') -> Dict[str, pd.DataFrame]:
    #     xlsb_data = pd.read_excel(excel_path, sheet_name=sheet_list, engine=engine)
    #     return xlsb_data

    @staticmethod
    def _restore_formatting(excel_writer: pd.ExcelWriter,
                            sheet_name: str,
                            unprotected_col_ids: List[int]
                            ):

        sheet = excel_writer.book[sheet_name]
        styles = []

        for i, row in enumerate(sheet.rows):
            if i < 2:
                continue
            elif i > 300:
                break
            for j, cell in enumerate(row):
                if i == 2:
                    if j in unprotected_col_ids:
                        protection = Protection(locked=False)
                        cell.protection = protection
                    else:
                        protection = cell.protection
                    style = {'color': cell.fill, 'protection': protection}
                    styles.append(style)
                    continue
                cell.fill = copy(styles[j]['color'])
                cell.protection = copy(styles[j]['protection'])

        CSV2ExcelProcessor._restore_all_dropdowns(excel_writer.book)

    @staticmethod
    def _restore_dropdowns(wb: Workbook, sheet_name: str, apply_range: str, formula: str) -> None:
        data_val = DataValidation(type="list", formula1=formula)
        wb[sheet_name].add_data_validation(data_val)
        data_val.add(apply_range)

    @staticmethod
    def _restore_all_dropdowns(wb: Workbook):
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Template', 'AQ3:AQ501', '=dropdowns!$M$2')
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Add New Records', 'E2:E45', '=dropdowns!$H$2:$H$3')
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Add New Records', 'F2:F45', '=dropdowns!$I$2:$I$6')
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Add New Records', 'G2:G45', '=dropdowns!$A$2:$A$24')
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Add New Records', 'H2:H45', '=dropdowns!$F$2:$F$1002')
        CSV2ExcelProcessor._restore_dropdowns(wb, 'Add New Records', 'L2:L45', '=dropdowns!$K$2:$K$3')

    @staticmethod
    def _prepare_template_copy(csv_file: str,
                               template_path: str,
                               output_dir: str,
                               skip_existing: bool = False) -> Tuple[bool, str]:
        csv_stem = Path(csv_file).stem
        output_filename = csv_stem + '_' + Path(template_path).name
        output_file_path = Path(output_dir) / output_filename
        if skip_existing and output_file_path.exists():
            skip = True
        else:
            shutil.copy(template_path, output_file_path)
            skip = False
        return skip, str(output_file_path)

    @staticmethod
    def sheets2csv(excel_dir: str,
                   sheet_name: str,
                   header_row_n: int,
                   col_range: str,
                   output_csv_path: str) -> None:

        excel_files = CSV2ExcelProcessor._get_files(excel_dir, ['xlsx', 'xlsb'])
        if len(excel_files) == 0:
            print(f'{excel_dir} has no excel files')
            return

        merged_df = pd.read_excel(excel_files[0], sheet_name=sheet_name, usecols=col_range, dtype=str)
        merged_df['file_name'] = Path(excel_files[0]).name
        col_names = list(merged_df.columns)
        if len(excel_files) > 1:
            for file in tqdm(excel_files[1:]):
                df = pd.read_excel(file, sheet_name=sheet_name,
                                   usecols=col_range,
                                   dtype=str)
                df['file_name'] = Path(file).name
                merged_df = pd.concat([merged_df, df])
        merged_df[col_names].to_csv(output_csv_path)

    @staticmethod
    def csv2template(csv_dir: str,
                     template_path: str,
                     output_dir: str,
                     unprotected_col_ids: List[int],
                     columns_to_int: List[int],
                     sheet_name: str,
                     skip_existing: bool = False):

        template_type = template_path.split('.')[-1]
        if template_type != 'xlsx':
            print(f'Template type must be in \"xlsl\" format, use Save As in Excel to convert it.')
            return

        csv_files = CSV2ExcelProcessor._get_files(csv_dir)
        if len(csv_files) == 0:
            print(f'No csv files in {csv_dir}. Please check your input dir.')
            return

        failed_files = []
        files_tuple = []
        for csv_file in csv_files:

            skip, output_file_path = CSV2ExcelProcessor._prepare_template_copy(str(csv_file),
                                                                               template_path,
                                                                               output_dir,
                                                                               skip_existing)
            if skip:
                print(f'{output_file_path} already exists. Skipping...')
                continue

            files_tuple.append((csv_file, output_file_path))
        f = partial(CSV2ExcelProcessor._process_one_file,
                    unprotected_col_ids=unprotected_col_ids,
                    columns_to_int=columns_to_int,
                    sheet_name=sheet_name)

        with Pool(8) as p:
            p.map(f, files_tuple)

        if len(failed_files) > 0:
            print('Processing done. Failed to process these files:')
            [print(file) for file in failed_files]

    @staticmethod
    def _process_one_file(input_output_files: Tuple[str, str],
                          unprotected_col_ids: List[int],
                          columns_to_int: List[int],
                          sheet_name: str):
        try:
            csv_file, output_file_path = input_output_files
            print(f'{csv_file} --> {output_file_path}')
            csv_df = CSV2ExcelProcessor._load_csv(csv_file, columns_to_int)
            with pd.ExcelWriter(output_file_path, mode='a', if_sheet_exists='overlay') as excel_writer:
                csv_df.to_excel(excel_writer,
                                sheet_name=sheet_name,
                                index=False,
                                startcol=4,
                                startrow=2,
                                header=None)
                CSV2ExcelProcessor._restore_formatting(excel_writer, sheet_name, unprotected_col_ids)

        except Exception as e:
            print(f"Couldn't process file: {csv_file}")
            print(f'Exception details: {str(e)}')
            if output_file_path is not None:
                Path(output_file_path).unlink(missing_ok=True)
            raise EnvironmentError(f"Can't process {csv_file}")