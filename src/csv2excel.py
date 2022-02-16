import os
from pathlib import Path
from typing import List, Tuple
import shutil
from copy import copy
from multiprocessing import Pool
from functools import partial

import pandas as pd
from openpyxl.styles.protection import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from tqdm import tqdm


class CSV2ExcelProcessor:

    @staticmethod
    def _get_files(dir_path: str, files_type: str = 'csv') -> List[Path]:
        files = os.listdir(dir_path)
        files = [Path(dir_path) / file for file in files if file.split('.')[-1] == files_type]
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
        df = df[~df[6].isin(filter_values)]
        return df

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

        data_val = DataValidation(type="list", formula1='=dropdowns!$M$2')
        sheet.add_data_validation(data_val)
        data_val.add('AQ3:AQ501')

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

            skip, output_file_path = CSV2ExcelProcessor._prepare_template_copy(csv_file,
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