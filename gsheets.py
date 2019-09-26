"""to download a google sheet and write to output file"""
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def oauth_scope(secrets_loc='./secrets', secrets_file='client_secret.json'):
    """
    :param secrets_loc: location of secrets file
    :param secrets_file: name of secreats file
    :return: gsspread authorised object that can read from google drive api
    """
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(f'{secrets_loc}/{secrets_file}',
                                                             scope)
    return gspread.authorize(creds)

def read_gsheet(gc, wkbook, wksheet, outloc, outfile, outfiletype):
    """
    uses pandas to read each values and used the first row as column names
    :param gc: authorised instance of gspread
    :param wkbook: name of google sheet workbook
    :param wksheet: name of google sheet worksheet
    :return: pandas DataFrame of data and write file to specifiec location in csv or pandas
    """
    worksheet = gc.open(wkbook).worksheet(wksheet)
    values = worksheet.get_all_values()
    df = pd.DataFrame(values[1:], columns=values[0])
    if outfiletype.lower() == 'csv':
        df.to_csv(f'{outloc}/{outfile}.csv', encoding='utf-8')
    elif outfiletype.lower() == 'excel':
        df.to_excel(f'{outloc}/{outfile}.xlsx', encoding='utf-8')

    return df

def main(gc, list_to_download: list):
    """
    download contents of specified gheets to location, outputfile will be
    the same as the workbook concatened with worksheet
    :param gc: uthorised instance of gspread
    :param list_to_download: list of workbooks, worksheet and output files types
    """
    for dl in list_to_download:
        # read_gsheet(gc, dl[0], dl[1], dl[2], f'{dl[0]}_{dl[1]}', dl[3])
        try:
            read_gsheet(gc, dl[0], dl[1], dl[2], f'{dl[0]}_{dl[1]}', dl[3])
            print(f'saved data from {dl[0]} - {dl[1]} to {dl[0]}_{dl[1]} as {dl[3]}')
        except:
            print(f'Error retreiving data from {dl[0]} - {dl[1]}')


if __name__ == '__main__':

    # authorise gspread
    gc = oauth_scope(secrets_loc='./secrets',
                     secrets_file='client_secret.json')

    # specify files to download
    save_loc = './data'
    file_type = 'excel'  #['excel', 'csv']
    downloads = [['OKR2 Key Activity Tracker', 'Data', save_loc, file_type],
                 ['OKR3 Key Activity Tracker', 'Data', save_loc, file_type],
                 ['OKR4 Key Activity Tracker', 'Data', save_loc, file_type],
                 ['OKR5 Key Activity Tracker', 'Data', save_loc, file_type]]

    # call main download function
    main(gc, downloads)


