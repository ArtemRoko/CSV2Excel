import os
from pathlib import Path
from typing import List, Tuple
import shutil
from copy import copy

import pandas as pd
from openpyxl.styles.protection import Protection
from tqdm import tqdm


class CSV2ExcelProcessor:

    @staticmethod
    def get_files(dir_path: str, files_type: str = 'csv') -> List[Path]:
        files = os.listdir(dir_path)
        files = [Path(dir_path) / file for file in files if file.split('.')[-1] == files_type]
        return files

    @staticmethod
    def replace_escapes(text: str) -> str:
        escapes = ''.join([chr(char) for char in range(1, 32)])
        translator = str.maketrans('', '', escapes)
        return text.translate(translator)

    @staticmethod
    def load_csv(csv_path: str, columns_to_int: List[int]) -> pd.DataFrame:
        csv_data = pd.read_csv(csv_path, header=None)
        csv_data = csv_data.loc[1:, :]
        csv_data.fillna('', inplace=True)
        csv_data = csv_data.applymap(CSV2ExcelProcessor.replace_escapes)
        for idx in columns_to_int:
            csv_data[idx] = csv_data[idx].apply(lambda x: x.split('.')[0])
        return csv_data

    @staticmethod
    def restore_formatting(excel_writer: pd.ExcelWriter,
                           sheet_name: str,
                           unprotected_col_ids: List[int]
                           ):

        sheet = excel_writer.book[sheet_name]
        styles = []

        for i, row in enumerate(sheet.rows):
            if i < 2:
                continue
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

    @staticmethod
    def prepare_template_copy(csv_file: str,
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

        csv_files = CSV2ExcelProcessor.get_files(csv_dir)
        if len(csv_files) == 0:
            print(f'No csv files in {csv_dir}. Please check your input dir.')
            return

        failed_files = []

        for csv_file in tqdm(csv_files):
            output_file_path = None
            try:
                skip, output_file_path = CSV2ExcelProcessor.prepare_template_copy(csv_file,
                                                                                  template_path,
                                                                                  output_dir,
                                                                                  skip_existing)
                if skip:
                    print(f'{output_file_path} already exists. Skipping...')
                    continue

                csv_df = CSV2ExcelProcessor.load_csv(csv_file, columns_to_int)
                with pd.ExcelWriter(output_file_path, mode='a', if_sheet_exists='overlay') as excel_writer:
                    csv_df.to_excel(excel_writer,
                                    sheet_name=sheet_name,
                                    index=False,
                                    startcol=4,
                                    startrow=2,
                                    header=None)
                    CSV2ExcelProcessor.restore_formatting(excel_writer, sheet_name, unprotected_col_ids)

            except Exception as e:
                print(f"Couldn't process file: {csv_file}")
                print(f'Exception details: {str(e)}')
                if output_file_path is not None:
                    Path(output_file_path).unlink(missing_ok=True)
                    failed_files.append(csv_file)

        if len(failed_files) > 0:
            print('Processing done. Failed to process these files:')
            [print(file) for file in failed_files]
