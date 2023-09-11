import tableauserverclient as TSC
import pandas as pd

tableau_auth = TSC.TableauAuth('username', 'Password')
server = TSC.Server('https://tableau.website.net/',use_server_version=True)

with server.auth.sign_in(tableau_auth):
    # get view's id
    req_option = TSC.RequestOptions()
    # create filter on all views by the view name
    req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name,
                                     TSC.RequestOptions.Operator.Equals,
                                             'tableau workbook sheet name'))
    # get all views meeting requirement
    matching_views, pagination_item = server.views.get(req_option)
    if len(matching_views) == 1:
        # when we only get one view, then we'll continue to get csv data of the view
        VIEW_ID = matching_views[0].id
        print("The view id is: " + VIEW_ID)
        default_view = matching_views[0]
        # create filters on csv data of view
        csv_req_option = TSC.CSVRequestOptions()
        csv_req_option.vf('ColumnNametoFilter', 'ColumnValuetoFilter')
        # Populate and save the CSV data as 'view_csv.csv'
        server.views.populate_csv(default_view, csv_req_option)
        with open(f'view.csv', 'wb') as f:
            # Perform byte join on the CSV data
            f.write(b''.join(default_view.csv))


import pandas as pd
df = pd.read_csv('view.csv')



#rearranging order of column names
df = df[['x1', 'x2', 'x3']]



# also Order the list by x1, x2, x3
df = df.sort_values(by=['x1', 'x2', 'x3'],ascending=True)



# resetting index
df.reset_index(drop=True, inplace=True)



import win32com.client

if not df.empty:
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "emailname1; emailname2"
    mail.Subject = 'ableau workbook'
    mail.HTMLBody = '''The following is the requested tableau workbook.\n\n
               {}'''.format(df.to_html())

    mail.CC = 'emailname'
    mail.send

