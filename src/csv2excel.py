import os
from pathlib import Path
from typing import List, Tuple, Dict
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
    def _delete_escapes(df: pd.DataFrame):
        for col_name in df.columns:
            if df[col_name].dtype.name == 'object':
                df[col_name] = df[col_name].astype(str).apply(CSV2ExcelProcessor._replace_escapes)
        return df

    @staticmethod
    def _filter_ma_indicators(df: pd.DataFrame) -> pd.DataFrame:
        filter_values = [509, 510, 511, 512]
        return df[~df[6].isin(filter_values)]

    @staticmethod
    def _load_csv(csv_path: str, columns_to_int: List[int], filter_ma: bool) -> pd.DataFrame:
        csv_data = pd.read_csv(csv_path, header=None)
        csv_data = csv_data.loc[1:, :]
        csv_data.fillna('', inplace=True)
        csv_data = csv_data.applymap(CSV2ExcelProcessor._replace_escapes)
        for idx in columns_to_int:
            csv_data[idx] = pd.to_numeric(csv_data[idx].apply(lambda x: x.split('.')[0]))
        if filter_ma:
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
    def _prepare_template_copy(input_file: str,
                               template_path: str,
                               output_dir: str,
                               skip_existing: bool = False) -> Tuple[bool, str]:
        input_file_stem = Path(input_file).stem
        # bvd_num = input_file_stem.split('_')[0]
        output_filename = input_file_stem + '_' + Path(template_path).name
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
        excel_files.sort(reverse=True)
        if len(excel_files) == 0:
            print(f'{excel_dir} has no excel files')
            return

        merged_df = pd.read_excel(excel_files[0], sheet_name=sheet_name, dtype=str)
        # col_names = list(merged_df.columns)
        merged_df['file_name'] = Path(excel_files[0]).name
        if len(excel_files) > 1:
            for file in tqdm(excel_files[1:]):
                df = pd.read_excel(file, sheet_name=sheet_name,
                                   dtype=str)
                df['file_name'] = Path(file).name
                merged_df = pd.concat([merged_df, df])
        new_cols = [col for col in merged_df.columns if col != 'file_name']
        new_cols.append('file_name')
        merged_df[new_cols].to_csv(output_csv_path)

    @staticmethod
    def csv2template(input_dir: str,
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

        input_files = CSV2ExcelProcessor._get_files(input_dir, files_type=['csv', 'xlsx', 'xlsb'])
        if len(input_files) == 0:
            print(f'No input files in {input_dir}. Please check your input dir.')
            return

        failed_files = []
        files_tuple = []
        for data_file in input_files:

            skip, output_file_path = CSV2ExcelProcessor._prepare_template_copy(str(data_file),
                                                                               template_path,
                                                                               output_dir,
                                                                               skip_existing)
            if skip:
                print(f'{output_file_path} already exists. Skipping...')
                continue

            files_tuple.append((data_file, output_file_path))
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
        input_file, output_file = input_output_files
        file_type = str(input_file).split('.')[-1]
        try:
            if file_type == 'csv':
                CSV2ExcelProcessor._process_input_csv(input_file,
                                                      sheet_name,
                                                      columns_to_int,
                                                      unprotected_col_ids,
                                                      output_file)
            elif file_type in ['xlsx', 'xlsb']:
                CSV2ExcelProcessor._process_input_excel(input_file,
                                                        columns_to_int,
                                                        unprotected_col_ids,
                                                        output_file)
        except Exception as e:
            print(f"Couldn't process file: {input_file}")
            print(f'Exception details: {str(e)}')
            if output_file is not None:
                Path(output_file).unlink(missing_ok=True)
            raise EnvironmentError(f"Can't process {input_file}")

    @staticmethod
    def _process_input_csv(csv_file: str,
                           sheet_name: str,
                           columns_to_int: List[int],
                           unprotected_col_ids: List[int],
                           output_file: str) -> None:

        csv_df = CSV2ExcelProcessor._load_csv(csv_file, columns_to_int, filter_ma=True)
        with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as excel_writer:
            csv_df.to_excel(excel_writer,
                            sheet_name=sheet_name,
                            index=False,
                            startcol=4,
                            startrow=2,
                            header=None)
            CSV2ExcelProcessor._restore_formatting(excel_writer, sheet_name, unprotected_col_ids)

    @staticmethod
    def _process_input_excel(excel_file: str,
                             columns_to_int: List[int],
                             unprotected_col_ids: List[int],
                             output_file: str) -> None:
        sheets = ['Template', 'Add New Records']
        df_dict = pd.read_excel(excel_file, sheet_name=sheets, header=None)
        template_df = df_dict['Template']
        template_df.fillna('', inplace=True)
        template_df = template_df.astype('str')
        template_df = template_df.applymap(CSV2ExcelProcessor._replace_escapes)

        col_names = template_df.iloc[1, :].tolist()
        uuid_pos = col_names.index('uuid')
        template_new_ver = 'Unit' in col_names
        part1_end_col = 30 if template_new_ver else uuid_pos + 23
        part3_col = col_names.index('Internal Comment')
        temp_part1 = template_df.iloc[2:501, uuid_pos: part1_end_col]
        temp_part2 = template_df.iloc[2:501, uuid_pos + 23:uuid_pos + 25]
        temp_part3 = template_df.iloc[2:501, part3_col]
        temp_part1 = CSV2ExcelProcessor._to_numeric(temp_part1, columns_to_int)

        add_record_df = df_dict['Add New Records']
        add_record_df.fillna('', inplace=True)
        add_record_df = add_record_df.astype('str')
        add_record_df = add_record_df.applymap(CSV2ExcelProcessor._replace_escapes)
        # add_rec_int_cols = [0, 1, 5, 6]
        add_record_df = add_record_df.iloc[1:, :]
        if template_new_ver:
            add_record_df = CSV2ExcelProcessor._to_numeric(add_record_df, [0, 1, 6, 7])
        else:
            add_record_df = CSV2ExcelProcessor._to_numeric(add_record_df, [0, 1, 5, 6])
        part1_end_col = 15 if template_new_ver else 3
        addrec_part1 = add_record_df.iloc[:44, 0:part1_end_col]
        addrec_part2 = add_record_df.iloc[:44, 3:9]
        addrec_part3 = add_record_df.iloc[:44, 9:13]

        with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='overlay') as excel_writer:
            temp_part1.to_excel(excel_writer, sheet_name='Template', index=False, startcol=4, startrow=2, header=None)
            if not template_new_ver:
                temp_part2.to_excel(excel_writer, sheet_name='Template', index=False, startcol=28, startrow=2, header=None)
            temp_part3.to_excel(excel_writer, sheet_name='Template', index=False, startcol=42, startrow=2, header=None)
            addrec_part1.to_excel(excel_writer, sheet_name='Add New Records', index=False, startcol=0, startrow=1, header=None)
            if not template_new_ver:
                addrec_part2.to_excel(excel_writer, sheet_name='Add New Records', index=False, startcol=4, startrow=1, header=None)
                addrec_part3.to_excel(excel_writer, sheet_name='Add New Records', index=False, startcol=11, startrow=1, header=None)

            CSV2ExcelProcessor._restore_formatting(excel_writer, 'Template', unprotected_col_ids)

    @staticmethod
    def _to_numeric(df: pd.DataFrame, columns_to_int: List[int]) -> pd.DataFrame:
        for idx in columns_to_int:
            df.iloc[:, idx] = pd.to_numeric(df.iloc[:, idx], errors='ignore')
        return df
