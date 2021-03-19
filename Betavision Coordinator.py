import tkinter as tk
import smartsheet
import getpass
import win32com.client as client
import webbrowser



# Smartsheet client acess token
smartsheet_client = smartsheet.Smartsheet(smartsheet_token)

# Get order submission sheet by ID
os_sheet=smartsheet_client.Sheets.get_sheet('7385615210702724')

# Get order form folder
order_form_folder=smartsheet_client.Folders.get_folder('6576000201975684')

# Get order form template sheet by ID
order_form_sheet=smartsheet_client.Sheets.get_sheet('4332222128908164')

# Get PM sheet by ID
pm_sheet=smartsheet_client.Sheets.get_sheet('4115641859893124')

user=getpass.getuser()

if user=='yuero':
    name='Roy Yue'
    title='Technical Coordinator | Custom Product'
    email_cc="Andrew.Cotter@arcteryx.com"
elif user=='cottera':
    name='Andrew Cotter'
    title='Technical Lead | Custom Product'
    email_cc="roy.yue@arcteryx.com"
elif user=='serbanc':
    name='Cristian Serban'
    recipients = ["royyue87@gmail.com", "royyue87@hotmail.com"]
    title='Decoration Services Coordinator | Custom Product' 
    email_cc="; ".join(recipients)



squad={'roy.yue@arcteryx.com':'Roy Yue',
'andrew.cotter@arcteryx.com':'Andrew Cotter',
'blair.gorrell@arcteryx.com':'Blair Gorrell',
'cristian.serban@arcteryx.com':'Cristian Serban',
'kaley.robinson@arcteryx.com':'Kaley Robinson',
'adam.hoogeveen@arcteryx.com':'Adam Hoogeveen',
'daniel.vanderhouwen@arcteryx.com':'Daniel Van Der Houwen',
'duncan.railton@arcteryx.com':'Duncan Railton',
'gustavo.ayresnetto@arcteryx.com':'Gustavo Ayres Netto',
'josh.banks@arcteryx.com':'Josh Banks',
'kat.redmond@arcteryx.com':'Katherine Redmond',
'matthew.reale@arcteryx.com':'Matthew Reale',
'scott.fierbach@arcteryx.com':'Scott Fierbach',
'kyle.goertzen@arcteryx.com':'Kyle Goertzen',
'gina.wong@arcteryx.com': 'Gina Wong',
'gustavo.ayresnetto@amersports.com': 'Gus'
}

# Put OS column titles and id in a dictionary
os_sheet_columns = smartsheet_client.Sheets.get_columns(
  7385615210702724,       
  include_all=True)
os_dict={}
column_title_list=[]
column_id_list=[]
for column in os_sheet_columns.data:
    column_title_list.append(column.title)
    column_id_list.append(column.id)
for y in range(0,24):
    os_dict[column_title_list[y]]=column_id_list[y]


# Put order form titles and id in a dictionary
order_form_dict={}
order_form_sheet_columns=smartsheet_client.Sheets.get_columns(
  4332222128908164,       
  include_all=True)
of_column_title_list=[]
of_column_id_list=[]
for column in order_form_sheet_columns.data:
    of_column_title_list.append(column.title)
    of_column_id_list.append(column.id)
for y in range(0,18):
    order_form_dict[of_column_title_list[y]]=of_column_id_list[y]


height=220
width=520

root =tk.Tk()
root.title("Betavision Coordinator 1.0.0")
canvas=tk.Canvas(root,height=height, width=width, bg='#cccccc')


frame=tk.Frame(root, bg='#cccccc',height=height, width=width)
frame.place(rely=0, relx=0, relwidth=1, relheight=1)

info=tk.StringVar()
form_link=tk.StringVar()
system_information=tk.StringVar()
error_system_information=tk.StringVar()
message_global_link=tk.StringVar()
message_system_info=tk.StringVar()
global_link_message=tk.StringVar()

project_id_list=[]
for row in os_sheet.rows:
    project_id_list.append(row.get_column(os_dict['Project ID']).value)


def read(entry):
    info.set("")
    form_link.set("")
    system_information.set("")
    global_link_message.set("")
    for row in os_sheet.rows:
        if row.get_column(os_dict['Project ID']).value==entry:
            accountName=row.get_column(os_dict['Account Name']).value
            orderNum=row.get_column(os_dict['Order Number']).value
            orderType=row.get_column(os_dict['Order Type']).value
            orderForm=row.get_column(os_dict['Order Form Link']).value
            sales=row.get_column(os_dict['Sales']).value
            region=row.get_column(os_dict['Production Facility']).value

            if orderNum==str:
                order_info=f'{accountName} | {orderNum} | {orderType} | {squad[sales]} | {region}'
            else:
                order_info=f'{accountName} | {str(int(orderNum))} | {orderType} | {squad[sales]} | {region}'

            
            if orderForm==None:
                global_link_message.set('')
            else:
                # Get Global Link
                for sheet in order_form_folder.sheets:
                    if sheet.permalink==orderForm:
                        response=smartsheet_client.Sheets.get_publish_status(sheet.id)    
                        if response.read_only_full_url==None:
                            sheetToPublish = smartsheet_client.Sheets.set_publish_status(sheet.id,smartsheet.models.SheetPublish({'readOnlyFullEnabled': True}))
                            sheetToPublishData=sheetToPublish.data
                            global_link=sheetToPublishData.read_only_full_url
                            global_link_message.set(global_link)
                        else:
                            global_link_message.set(response.read_only_full_url)




            info.set(order_info)
            form_link.set(orderForm)

            message_info=tk.Entry(frame,bg='#bbbbbb',textvariable=info, bd=0,state='readonly', relief='flat')
            message_info.place(height=30, width=390, x=100, y=50)
            message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
            message_form_link.place(height=30, width=390, x=100, y=90)
            global_link_field=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=global_link_message,bd=0,state='readonly', relief='flat')
            global_link_field.place(height=30, width=390, x=100, y=130)

    if entry in project_id_list:
        system_text='Successful!'
        fg_color='#48ff00'
    else:
        system_text='Error, typo on your entry!'
        fg_color='#ff0000'
    system_information.set(system_text)
    system_field=tk.Entry(frame, bg='#bbbbbb',font=("Arial", 8),textvariable=system_information,bd=0,state='readonly', relief='flat',fg=fg_color)
    system_field.place(height=30, width=390, x=100, y=170)

    
def create(entry):
    info.set("")
    form_link.set("")
    system_information.set("")
    global_link_message.set("")
    if entry in project_id_list:
        for row in os_sheet.rows:
            if row.get_column(os_dict['Project ID']).value==entry:
                row_id=row.id
                accountName=row.get_column(os_dict['Account Name']).value
                orderNum=row.get_column(os_dict['Order Number']).value
                orderForm=row.get_column(os_dict['Order Form Link']).value
                orderType=row.get_column(os_dict['Order Type']).value
                sales=row.get_column(os_dict['Sales']).value
                region=row.get_column(os_dict['Production Facility']).value

                if orderNum==str:
                    order_info=f'{accountName} | {orderNum} | {orderType} | {squad[sales]} | {region}'
                else:
                    order_info=f'{accountName} | {str(int(orderNum))} | {orderType} | {squad[sales]} | {region}'
            
                info.set(order_info)

                message_info=tk.Entry(frame,bg='#bbbbbb',textvariable=info, bd=0,state='readonly', relief='flat')
                message_info.place(height=30, width=390, x=100, y=50)
                message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
                message_form_link.place(height=30, width=390, x=100, y=90)

                

                # Put order form titles and id in a dictionary
                order_form_dict={}
                order_form_sheet_columns=smartsheet_client.Sheets.get_columns(
                4332222128908164,       
                include_all=True)
                of_column_title_list=[]
                of_column_id_list=[]
                for column in order_form_sheet_columns.data:
                    of_column_title_list.append(column.title)
                    of_column_id_list.append(column.id)
                for y in range(0,18):
                    order_form_dict[of_column_title_list[y]]=of_column_id_list[y]


                if orderForm ==None:
                    system_text='Successful!'
                    fg_color='#48ff00'
                    # Create order form and link
                    def updateOrderInfo(rowId, columnName, orderInfo, sheetId):
                        new_cell = smartsheet.models.Cell()
                        new_cell.column_id = order_form_dict[columnName]
                        new_cell.value = orderInfo
                        new_cell.strict = False
                        new_row = smartsheet.models.Row()
                        new_row.id = rowId
                        new_row.cells.append(new_cell)
                        updated_row = smartsheet_client.Sheets.update_rows(
                            sheetId,      # sheet_id
                            [new_row])
                    
                    
                    updateOrderInfo(3771591160817540, 'Project ID', "", 4332222128908164)
                    updateOrderInfo(3771591160817540, 'Project ID', entry, 4332222128908164)
                    # updateOrderInfo(8275190788188036, 'Project ID', 'Customer Info', 4332222128908164)
                    # updateOrderInfo(8275190788188036, 'Project ID', 'Customer Infomation', 4332222128908164)
                    
                    # Copy template and save in order form folder
                    if type(orderNum)==str:
                        newOrderSheet = smartsheet_client.Sheets.copy_sheet(4332222128908164,
                        smartsheet.models.ContainerDestination(
                            {'destination_type': 'folder',
                            'destination_id': 6576000201975684,
                            'new_name': orderNum+"-"+accountName+"-"+orderType}), 
                            include='cellLinks')
                    else:
                        newOrderSheet = smartsheet_client.Sheets.copy_sheet(4332222128908164,
                        smartsheet.models.ContainerDestination(
                            {'destination_type': 'folder',
                            'destination_id': 6576000201975684,
                            'new_name': str(int(orderNum))+"-"+accountName+"-"+orderType}), 
                            include='all')


                    newOrderSheet_data=newOrderSheet.result
                    form_link.set(newOrderSheet_data.permalink)

                    # Post form link to OS sheet
                    new_cell = smartsheet.models.Cell()
                    new_cell.column_id = os_dict['Order Form Link']
                    new_cell.value = newOrderSheet_data.permalink
                    new_cell.strict = False

                    # Build the row to update
                    new_row = smartsheet.models.Row()
                    new_row.id = row_id
                    new_row.cells.append(new_cell)

                    # Update rows
                    updated_row = smartsheet_client.Sheets.update_rows(
                    7385615210702724,      # sheet_id
                    [new_row])
                    

                    # Publish the order form sheet
                    sheetToPublish = smartsheet_client.Sheets.set_publish_status(
                    newOrderSheet_data.id,       # sheet_id
                    smartsheet.models.SheetPublish({
                    'readOnlyFullEnabled': True
                    })
                    )

                    sheetToPublishData=sheetToPublish.data
                    global_link_message.set(sheetToPublishData.read_only_full_url)
            
                else:
                    system_text='You already created order form'
                    fg_color='#FFD300'
                    form_link.set(orderForm)
                    for sheet in order_form_folder.sheets:
                        if sheet.permalink==orderForm:
                            response=smartsheet_client.Sheets.get_publish_status(sheet.id)    
                            if response.read_only_full_url==None:
                                sheetToPublish = smartsheet_client.Sheets.set_publish_status(sheet.id,smartsheet.models.SheetPublish({'readOnlyFullEnabled': True}))
                                sheetToPublishData=sheetToPublish.data
                                global_link=sheetToPublishData.read_only_full_url
                                global_link_message.set(global_link)
                            else:
                                global_link_message.set(response.read_only_full_url)



            message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
            message_form_link.place(height=30, width=390, x=100, y=90)
            global_link_field=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=global_link_message,bd=0,state='readonly', relief='flat')
            global_link_field.place(height=30, width=390, x=100, y=130)




    else:
        system_text='Error, typo on your entry!'
        fg_color='#ff0000'

    system_information.set(system_text)
    system_field=tk.Entry(frame, bg='#bbbbbb',font=("Arial", 8),textvariable=system_information,bd=0,state='readonly', relief='flat',fg=fg_color)
    system_field.place(height=30, width=390, x=100, y=170)

def open(entry):
    info.set("")
    form_link.set("")
    system_information.set("")
    global_link_message.set("")

    if entry not in project_id_list:
        system_text='Error, typo on your entry!'
        fg_color='#ff0000'

    else:
        for row in os_sheet.rows:
            if row.get_column(os_dict['Project ID']).value==entry:
                accountName=row.get_column(os_dict['Account Name']).value
                orderNum=row.get_column(os_dict['Order Number']).value
                orderType=row.get_column(os_dict['Order Type']).value
                orderForm=row.get_column(os_dict['Order Form Link']).value
                sales=row.get_column(os_dict['Sales']).value
                region=row.get_column(os_dict['Production Facility']).value

                if orderNum==str:
                    order_info=f'{accountName} | {orderNum} | {orderType} | {squad[sales]} | {region}'
                else:
                    order_info=f'{accountName} | {str(int(orderNum))} | {orderType} | {squad[sales]} | {region}'

                
                if orderForm==None:
                    global_link_message.set('')


                else:
                    # Get Global Link
                    for sheet in order_form_folder.sheets:
                        if sheet.permalink==orderForm:
                            response=smartsheet_client.Sheets.get_publish_status(sheet.id)    
                            if response.read_only_full_url==None:
                                sheetToPublish = smartsheet_client.Sheets.set_publish_status(sheet.id,smartsheet.models.SheetPublish({'readOnlyFullEnabled': True}))
                                sheetToPublishData=sheetToPublish.data
                                global_link=sheetToPublishData.read_only_full_url
                                global_link_message.set(global_link)
                            else:
                                global_link_message.set(response.read_only_full_url)
                



                if orderForm!= None:
                    webbrowser.open(orderForm)  


                info.set(order_info)
                form_link.set(orderForm)

                message_info=tk.Entry(frame,bg='#bbbbbb',textvariable=info, bd=0,state='readonly', relief='flat')
                message_info.place(height=30, width=390, x=100, y=50)
                message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
                message_form_link.place(height=30, width=390, x=100, y=90)
                global_link_field=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=global_link_message,bd=0,state='readonly', relief='flat')
                global_link_field.place(height=30, width=390, x=100, y=130)
            
                if orderForm ==None:
                    system_text='No order form created!'
                    fg_color='#ff0000'
                else:
                    system_text='Successful!'
                    fg_color='#48ff00'


                    
    system_information.set(system_text)
    system_field=tk.Entry(frame, bg='#bbbbbb',font=("Arial", 8),textvariable=system_information,bd=0,state='readonly', relief='flat',fg=fg_color)
    system_field.place(height=30, width=390, x=100, y=170)

def send(entry):
    info.set("")
    form_link.set("")
    system_information.set("")
    global_link_message.set("")

    globalLink=''
    project_id_pm=[]
    for row in pm_sheet.rows:
        project_id_pm.append(row.get_column(8779048005461892).value)
    
    
    if entry in project_id_pm:
        info.set("")
        form_link.set("")
        system_information.set("")
        global_link_message.set("")


        # Put PM column titles and id in a dictionary
        pm_sheet_columns = smartsheet_client.Sheets.get_columns(
        4115641859893124,       
        include_all=True)
        pm_dict={}
        column_title_list=[]
        column_id_list=[]
        for column in pm_sheet_columns.data:
            column_title_list.append(column.title)
            column_id_list.append(column.id)
        for y in range(0,64):
            pm_dict[column_title_list[y]]=column_id_list[y]

        for row in pm_sheet.rows:
            if row.get_column(pm_dict['Project ID']).value==entry:
                st_num=row.get_column(pm_dict['ST Number']).value
                decorations=row.get_column(pm_dict['Decoration Details']).value
                technical=row.get_column(pm_dict['Technical Rep']).value
                coordinator=row.get_column(pm_dict['Project Mgmt Rep']).value




        
        for row in os_sheet.rows:
            if row.get_column(os_dict['Project ID']).value==entry:
                row_id=row.id
                accountName=row.get_column(os_dict['Account Name']).value
                orderNum=row.get_column(os_dict['Order Number']).value
                orderForm=row.get_column(os_dict['Order Form Link']).value
                orderType=row.get_column(os_dict['Order Type']).value
                sales=row.get_column(os_dict['Sales']).value
                region=row.get_column(os_dict['Production Facility']).value

                if orderNum==str:
                    order_info=f'{accountName} | {orderNum} | {orderType} | {squad[sales]} | {region}'
                else:
                    order_info=f'{accountName} | {str(int(orderNum))} | {orderType} | {squad[sales]} | {region}'
            
                info.set(order_info)

                message_info=tk.Entry(frame,bg='#bbbbbb',textvariable=info, bd=0,state='readonly', relief='flat')
                message_info.place(height=30, width=390, x=100, y=50)
                message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
                message_form_link.place(height=30, width=390, x=100, y=90)

                

                # Put order form titles and id in a dictionary
                order_form_dict={}
                order_form_sheet_columns=smartsheet_client.Sheets.get_columns(
                4332222128908164,       
                include_all=True)
                of_column_title_list=[]
                of_column_id_list=[]
                for column in order_form_sheet_columns.data:
                    of_column_title_list.append(column.title)
                    of_column_id_list.append(column.id)
                for y in range(0,18):
                    order_form_dict[of_column_title_list[y]]=of_column_id_list[y]


                if orderForm ==None:
                    # Create order form and link
                    def updateOrderInfo(rowId, columnName, orderInfo, sheetId):
                        new_cell = smartsheet.models.Cell()
                        new_cell.column_id = order_form_dict[columnName]
                        new_cell.value = orderInfo
                        new_cell.strict = False
                        new_row = smartsheet.models.Row()
                        new_row.id = rowId
                        new_row.cells.append(new_cell)
                        updated_row = smartsheet_client.Sheets.update_rows(
                            sheetId,      # sheet_id
                            [new_row])
                    
                    
                    updateOrderInfo(3771591160817540, 'Project ID', "", 4332222128908164)
                    updateOrderInfo(3771591160817540, 'Project ID', entry, 4332222128908164)
                    
                    # Copy template and save in order form folder
                    if type(orderNum)==str:
                        newOrderSheet = smartsheet_client.Sheets.copy_sheet(4332222128908164,
                        smartsheet.models.ContainerDestination(
                            {'destination_type': 'folder',
                            'destination_id': 6576000201975684,
                            'new_name': orderNum+"-"+accountName+"-"+orderType}), 
                            include='cellLinks')
                    else:
                        newOrderSheet = smartsheet_client.Sheets.copy_sheet(4332222128908164,
                        smartsheet.models.ContainerDestination(
                            {'destination_type': 'folder',
                            'destination_id': 6576000201975684,
                            'new_name': str(int(orderNum))+"-"+accountName+"-"+orderType}), 
                            include='all')


                    newOrderSheet_data=newOrderSheet.result
                    form_link.set(newOrderSheet_data.permalink)

                    # Post form link to OS sheet
                    new_cell = smartsheet.models.Cell()
                    new_cell.column_id = os_dict['Order Form Link']
                    new_cell.value = newOrderSheet_data.permalink
                    new_cell.strict = False

                    # Build the row to update
                    new_row = smartsheet.models.Row()
                    new_row.id = row_id
                    new_row.cells.append(new_cell)

                    # Update rows
                    updated_row = smartsheet_client.Sheets.update_rows(
                    7385615210702724,      # sheet_id
                    [new_row])
                    

                    # Publish the order form sheet
                    sheetToPublish = smartsheet_client.Sheets.set_publish_status(
                    newOrderSheet_data.id,       # sheet_id
                    smartsheet.models.SheetPublish({
                    'readOnlyFullEnabled': True
                    })
                    )

                    sheetToPublishData=sheetToPublish.data
                    global_link_message.set(sheetToPublishData.read_only_full_url)
            
                else:
                    form_link.set(orderForm)
                    for sheet in order_form_folder.sheets:
                        if sheet.permalink==orderForm:
                            response=smartsheet_client.Sheets.get_publish_status(sheet.id)    
                            if response.read_only_full_url==None:
                                sheetToPublish = smartsheet_client.Sheets.set_publish_status(sheet.id,smartsheet.models.SheetPublish({'readOnlyFullEnabled': True}))
                                sheetToPublishData=sheetToPublish.data
                                global_link=sheetToPublishData.read_only_full_url
                                global_link_message.set(global_link)
                                global_link=global_link
                            else:
                                global_link_message.set(response.read_only_full_url)
                                global_link=response.read_only_full_url


                outlook=client.Dispatch("Outlook.Application")
                message=outlook.CreateItem(0)
                message.Display()
                message.To="production@corporateimageoutfitters.com"
                message.CC=email_cc
                message.Subject="Arc'teryx Order Submission | {} | {}".format(accountName, str(int(orderNum)))
                message.HTMLBody=f'''

                <p>Hello, </p>

                <p>New order coming your way</p>

                <p><strong>ST #:</strong>  {int(st_num)}</p>
                <p><strong>Decoration Types:</strong>  {decorations}</p>
                <p><strong>Technical Rep:</strong>  {technical}</p>
                <p><strong>Project Coordinator:</strong> {coordinator}</p>
                <p><strong>Order Form Link:</strong> {global_link}</p>

                <p>Thanks!</p><br>

                <p style='font-weight:bold; font-size:14'>
                {name}<br>
                <span style='font-weight:normal'>{title}</span><br>
                ARC’TERYX Equipment | A Division of Amer Sports Canada Inc.</p>'''

                system_text='Successful!'
                fg_color='#48ff00'

    else:
        system_text='Error, order is not in PM sheet!'
        fg_color='#ff0000'





    system_information.set(system_text)
    message_form_link=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=form_link,bd=0,state='readonly', relief='flat')
    message_form_link.place(height=30, width=390, x=100, y=90)
    global_link_field=tk.Entry(frame, bg='#bbbbbb', font=("Arial", 8),textvariable=global_link_message,bd=0,state='readonly', relief='flat')
    global_link_field.place(height=30, width=390, x=100, y=130) 
    system_field=tk.Entry(frame, bg='#bbbbbb',font=("Arial", 8),textvariable=system_information,bd=0,state='readonly', relief='flat',fg=fg_color)
    system_field.place(height=30, width=390, x=100, y=170)

















        
button_read=tk.Button(frame, text='READ',command=lambda: read(entry.get()))
button_create=tk.Button(frame, text='CREATE',command=lambda: create(entry.get()))
button_open=tk.Button(frame, text='OPEN',command=lambda: open(entry.get()))
button_send=tk.Button(frame, text='SEND',command=lambda: send(entry.get()))
label_project_id = tk.Label(frame,text='Project ID:', bg='#cccccc')
label_order_link=tk.Label(frame,text='Order Form:', bg='#cccccc')
label_global=tk.Label(frame,text='Global Link:', bg='#cccccc')
label_order_info=tk.Label(frame,text='Order Info:', bg='#cccccc')
label_log=tk.Label(frame,text='System Info:', bg='#cccccc')
entry=tk.Entry(frame,bg='#ffffff')

message_form_link=tk.Message(frame, bg='#eeeeee', textvariable=form_link)
message_global=tk.Message(frame, bg='#eeeeee', textvariable=message_global_link)
message_system=tk.Message(frame, bg='#eeeeee', textvariable=message_system_info)
message=tk.Message(frame, bg='#eeeeee')



label_project_id.place(height=30, width=100, x=0, y=10)
entry.place(height=30, width=150, x=100, y=10)
button_read.place(height=30, width=50, x=260, y=10)
button_create.place(height=30, width=50, x=320, y=10)
button_open.place(height=30, width=50, x=380, y=10)
button_send.place(height=30, width=50, x=440, y=10)
label_order_info.place(height=30, width=100, x=0, y=50)
message.place(height=30, width=390, x=100, y=50)
label_order_link.place(height=30, width=100, x=0, y=90)
message_form_link.place(height=30, width=390, x=100, y=90)
label_global.place(height=30, width=100, x=0, y=130)
message_global.place(height=30, width=390, x=100, y=130)
label_log.place(height=30, width=100, x=0, y=170)
message_system.place(height=30, width=390, x=100, y=170)



canvas.pack()
root.mainloop()
