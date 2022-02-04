from src.csv2excel import CSV2ExcelProcessor

if __name__ == '__main__':

    # Set params
    csv_dir = './input_csv'
    template_file = './template/CGV_enrichment_template_v2.xlsx'  # Should be converted in xlsx format
    sheet_name = 'Template'  # Were to copy data from csv
    unprotected_columns_ids = [22, 23, 24, 26, 27, 28, 31]  # Columns that should be editable
    output_dir = './output_excel'

    CSV2ExcelProcessor.csv2template(csv_dir=csv_dir,
                                    template_path=template_file,
                                    output_dir=output_dir,
                                    unprotected_col_ids=unprotected_columns_ids,
                                    sheet_name=sheet_name)
    