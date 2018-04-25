from shutil import copyfile
import openpyxl

def write_to_template(all_transactions, file_destination, sheet_name):
    active_file = openpyxl.load_workbook(file_destination, keep_vba=True)
    active_sheet = active_file.get_sheet_by_name(sheet_name)
    row = 2
    for transaction in all_transactions:
        for key in transaction.template_col_dict.keys():
            cell = key + str(row)
            to_print = transaction.template_col_dict[key]
            if to_print:
                active_sheet[cell] = to_print
        row += 1
    active_file.save(file_destination)

def publish_to_template(all_transactions,file_destination,template,sheet_name):
    copyfile(template, file_destination)
    write_to_template(all_transactions,file_destination,sheet_name)