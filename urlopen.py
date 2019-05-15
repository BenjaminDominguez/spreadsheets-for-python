import pandas as pd
import webbrowser
import sys

def open_spreadsheet(spreadsheet_name):
    """
    Sheet should have a Company Name and URL column
    """

    df = pd.read_excel('spreadsheets/{0}'.format(spreadsheet_name))

    return df

def url_open_generator(spreadsheet_name):

    df = open_spreadsheet(spreadsheet_name)

    for index, row in df.iterrows():
        yield (row['URL'], row['Company Name'])

def open_urls():
    """
    Example argv set: python urlopen.py realestate --save
    """

    save = False
    
    try:
        save = True if sys.argv[2] == '--save' else False
    except IndexError:
        # no argument given
        pass

    print(save)

    spreadsheet_name = sys.argv[1] + '.xlsx'

    gen = url_open_generator(spreadsheet_name)
    df = open_spreadsheet(spreadsheet_name)

    generator_on = True

    print('Press CTRL-C and then enter to quit program..')
    print('--'*40)
    print('Press enter to open a new url')
    print('Type \'s\' after a url to save to a new dataframe')
    
    #Init new rows list to be concatenated into a single dataframe at the end
    new_rows = []
    while generator_on:
        try:
            url, company_name = next(gen)
            print('Company name: ', company_name)
            webbrowser.open(url)

            x = input()
            if x == 's':
                new_row = df[df['URL'] == url]
                new_rows.append(new_row)

            else:
                continue

        except KeyboardInterrupt:
            generator_on = False

            if save:
                df = pd.concat(new_rows, ignore_index=True)
                print("""
                CSV file generated from saved results:
                \n
                {0}
                """.format(df.head()))

                df.to_excel('spreadsheets/outputs/{0}'.format(sys.argv[1] + '_output.xlsx'))

open_urls()