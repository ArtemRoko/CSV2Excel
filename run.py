from src.csv2excel import CSV2ExcelProcessor
from pathlib import Path

if __name__ == '__main__':

    # Set params
    csv_dir = './input_csv'  # Program will process every csv file in this dir
    template_file = './template/CGV_enrichment_template_v6.0.xlsx'  # Should be in xlsx format
    sheet_name = 'Template'  # Where to copy data from CSVs
    unprotected_columns_ids = [22, 23, 24, 26, 27, 28, 29, 42]  # Columns that should be editable
    columns_to_int = [2, 6, 17, 18, 19]  # IDs of columns in csv file that will be cast to int
    output_dir = './output_excel'  # output file name = [csv_name]_[template_name].xlsx
    skip_existing = True  # skip processing if an output file already exists

    excel_dir = './input_excel'
    output_csv = str(Path(output_dir) / 'add_new_record.csv')
    add_new_record_header = 1
    add_new_record_colrange = 'A:M'

    # CSV2ExcelProcessor.csv2template(csv_dir=csv_dir,
    #                                 template_path=template_file,
    #                                 output_dir=output_dir,
    #                                 unprotected_col_ids=unprotected_columns_ids,
    #                                 columns_to_int=columns_to_int,
    #                                 sheet_name=sheet_name,
    #                                 skip_existing=skip_existing)

    CSV2ExcelProcessor.sheets2csv(excel_dir,
                                  'Add New Records',
                                  add_new_record_header,
                                  add_new_record_colrange,
                                  output_csv)


