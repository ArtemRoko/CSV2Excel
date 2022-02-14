from src.csv2excel import CSV2ExcelProcessor
import time

if __name__ == '__main__':

    # Set params
    csv_dir = './input_csv'  # Program will process every csv file in this dir
    template_file = './template/CGV_enrichment_template_v4.0.xlsx'  # Should be in xlsx format
    sheet_name = 'Template'  # Where to copy data from CSVs
    unprotected_columns_ids = [22, 23, 24, 26, 27, 28, 29, 42]  # Columns that should be editable
    columns_to_int = [2, 6, 17, 18, 19]  # IDs of columns in csv file that will be cast to int
    output_dir = './output_excel'  # output file name = [csv_name]_[template_name].xlsx
    skip_existing = True  # skip processing if an output file already exists

    t1 = time.monotonic()
    CSV2ExcelProcessor.csv2template(csv_dir=csv_dir,
                                    template_path=template_file,
                                    output_dir=output_dir,
                                    unprotected_col_ids=unprotected_columns_ids,
                                    columns_to_int=columns_to_int,
                                    sheet_name=sheet_name,
                                    skip_existing=skip_existing)
    t2 = time.monotonic()
    print(f'One file per {(t2-t1)/25:.2f} s')

