import pandas as pd
import win32com.client as win32                                                                                         # pip install pywin32
import re
import os
import shutil
import warnings
from datetime import datetime
from datetime import date

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

original_df = r'\\ant.amazon.com\dept-as\HND10\JP_F3\3.F3_Retail\2.F3_Instock\Vendor Operation Calendar\★vendor連絡先リスト_Master.xlsx'
vendors = pd.read_excel(original_df, sheet_name="Master_List", usecols="B,D:N,R,T,U", header=1)
vm = pd.read_excel(original_df, sheet_name="VM名称")
df = pd.merge(vendors, vm, on='VM_name', how='left')                                                                    # left join two tables

parent_vendor_code = list(df.iloc[:, 1])
child_vendor_code = df.iloc[:, 2:11]                                                                                    # related vendor codes

tidy_df = pd.DataFrame()

child_list = []
parent_list = []
vendor_name_list = []
vendor_email_list = []
person_in_charge_list = []
vm_login_list = []
vm_name_list = []
vm_email_list = []

for i in range(len(parent_vendor_code)):                                                                                # merge parent vendor codes into tidy dataset
    child_list.append(parent_vendor_code[i])
    parent_list.append(parent_vendor_code[i])
    vendor_name_list.append(df.iloc[i, 0])
    vendor_email_list.append(df.iloc[i, 13])
    person_in_charge_list.append(df.iloc[i, 14])
    vm_login_list.append(df.iloc[i, 12])
    vm_name_list.append(df.iloc[i, 15])
    vm_email_list.append(df.iloc[i, 16])

for i in range(len(child_vendor_code)):                                                                                 # merge child vendor codes into tidy dataset
    for j in range(len(child_vendor_code.columns)):
        if len(str(child_vendor_code.iloc[i, j]).replace(" ", "")) == 5:                                                # make sure the cell contains a vendor code
            child_list.append(child_vendor_code.iloc[i, j])
            parent_list.append(df.iloc[i, 1])
            vendor_name_list.append(df.iloc[i, 0])
            vendor_email_list.append(df.iloc[i, 13])
            person_in_charge_list.append(df.iloc[i, 14])
            vm_login_list.append(df.iloc[i, 12])
            vm_name_list.append(df.iloc[i, 15])
            vm_email_list.append(df.iloc[i, 16])

for i in range(len(vendor_name_list)):
    nb_rep = 1
    while (nb_rep):                                                                                                     # remove parenthesis
        (vendor_name_list[i], nb_rep) = re.subn(r'\([^()]*\)', '', vendor_name_list[i])

    vendor_name_list[i] = re.sub("[\[].*?[\]]", "", vendor_name_list[i])
    vendor_name_list[i] = re.sub("[\(\（].*?[\)\）]", "", vendor_name_list[i])

    vendor_name_list[i] = vendor_name_list[i].strip()                                                                   # remove spaces

    if vendor_name_list[i].startswith("株式会社"):
        vendor_name_list[i] = vendor_name_list[i][4:]

    if not vendor_name_list[i].endswith("株式会社"):
        vendor_name_list[i] += "株式会社"

for i in range(len(person_in_charge_list)):                                                                             # replace commas with japanese commas
    if len(str(person_in_charge_list[i])) > 2:
        person_in_charge_list[i] = str(person_in_charge_list[i]).replace(", ", "、")

tidy_df['child_vendor_code'] = pd.Series(child_list)
tidy_df['parent_vendor_code'] = pd.Series(parent_list)
tidy_df['vendor_name'] = pd.Series(vendor_name_list)
tidy_df['vendor_email'] = pd.Series(vendor_email_list)
tidy_df['person_in_charge'] = pd.Series(person_in_charge_list)
tidy_df['vm_login'] = pd.Series(vm_login_list)
tidy_df['vm_name'] = pd.Series(vm_name_list)
tidy_df['vm_email'] = pd.Series(vm_email_list)

tidy_df.to_excel('auto_generated_vm_list.xlsx', index=False)

with open('vendorlist.txt', 'w+') as f_vendorlist:
    f_vendorlist.write("A tidy dataset is exported to root directory.")

"""
Read txt files and write dataframe into template file.
"""

date_yymmdd = str(datetime.today().strftime('%Y-%m-%d')).replace("-", "")[2:]
current_month = str(date.today().month)

folder_path = r'\\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\vendor_code'
template_file = r'\\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\Vendor送付ファイル\Template\F3_NewProduct_Request_Vendor名_alias_date_vssc_factoryid_ver2.xlsx'
new_path = r'\\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\Vendor送付ファイル\\'+current_month+"月送付分"

if not os.path.exists(new_path):
    os.makedirs(new_path)

involved_vendors = []
empty_vendors = []
invalid_codes = []

for file in os.listdir(folder_path):
    file_name = str(os.path.splitext(file)[0])
    file_extension = str(os.path.splitext(file)[1])

    if file_extension == ".txt":
        if len(file_name) == 5:
            with open(folder_path+r'\\'+str(file)) as txt_file:
                file_content = str(txt_file.read())

                if len(file_content) > 79:
                    involved_vendors.append(file_name)

                    file_path = new_path+r"\F3_NewProduct_Request_"+str(file_name)+"_alias_"+date_yymmdd+"_vssc_factoryid.xlsx"
                    shutil.copyfile(template_file, file_path)

                    file_df = pd.read_csv(folder_path+r'\\'+str(file), sep="\t", encoding='ANSI')

                    file_df['sku'] = file_df['sku'].str.replace('?', 'ー', regex=True)
                    file_df['category'] = file_df['category'].str.replace('?', 'ー', regex=True)
                    file_df['sub_category'] = file_df['sub_category'].str.replace('?', 'ー', regex=True)
                    file_df['maker'] = file_df['maker'].str.replace('?', 'ー', regex=True)
                    file_df['brand'] = file_df['brand'].str.replace('?', 'ー', regex=True)

                    writer = pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay')
                    file_df.to_excel(writer, sheet_name='データ貼り付け', index=False)

                    file_df['maker'].to_excel(writer, sheet_name='ベンダー様記入シート', index=False, header=None, startcol=3, startrow=6)
                    file_df['sku'].to_excel(writer, sheet_name='ベンダー様記入シート', index=False, header=None, startcol=5, startrow=6)
                    file_df['jan'].to_excel(writer, sheet_name='ベンダー様記入シート', index=False, header=None, startcol=6, startrow=6)

                    writer.save()

                else:
                    empty_vendors.append(file_name)

        else:
            invalid_codes.append(file_name)

involved_vendors = list(set(involved_vendors))

with open('attachment.txt', 'w+') as f_attachment:
    f_attachment.write(str(len(involved_vendors)) + " vendor codes should be reported to the vendors. Please find them in vendor_code.xlsx.\n")
    f_attachment.write(str(len(empty_vendors)) + " vendor codes are empty. Please find them in empty_vendor_code.xlsx and notify the vendor manager.")

    if len(invalid_codes):
        f_attachment.write("\n\n")

        for invalid_code in invalid_codes:
            f_attachment.write(invalid_code + "\n")

        f_attachment.write("These vendor codes are not valid. Please double-check.")

pd.DataFrame(involved_vendors).to_excel('vendor_code.xlsx', header=False, index=False)
pd.DataFrame(empty_vendors).to_excel('empty_vendor_code.xlsx', header=False, index=False)

"""
Insert attachment(s) into emails.
Generate emails in Outlook.
"""

outlook = win32.Dispatch('outlook.application')

df = pd.read_excel("auto_generated_vm_list.xlsx")
sub_df = df[['parent_vendor_code', 'child_vendor_code']]
input_df = pd.read_excel("vendor_code.xlsx", header=None)

if not input_df.empty:
    involved_vendors = list(pd.read_excel("vendor_code.xlsx", header=None).iloc[:, 0])
    involved_parent_vendors = []
    codes_not_found = []
    invalid_codes = []

    child_vendor_code = df.iloc[:, 0]
    parent_vendor_code = df.iloc[:, 1]

    for i in range(len(involved_vendors)):
        if len(str(involved_vendors[i])) == 5:                                                                          # check if the vendor code is valid
            try:
                involved_parent_vendors.append(str(df[df.eq(involved_vendors[i]).any(1)].iloc[0, 1]))                   # find the parent of vendor code

            except:
                codes_not_found.append(involved_vendors[i])

        else:
            invalid_codes.append(involved_vendors[i])

    involved_parent_vendors = list(set(involved_parent_vendors))                                                        # remove duplicates

    for i in range(len(involved_parent_vendors)):
        current_parent_vendor = involved_parent_vendors[i]
        current_child_vendors = []
        attachments = []

        vendor_email = str(df[df.eq(current_parent_vendor).any(1)].iloc[0, 3])
        vm_email = str(df[df.eq(current_parent_vendor).any(1)].iloc[0, 7])
        vendor_name = str(df[df.eq(current_parent_vendor).any(1)].iloc[0, 2])
        person_in_charge = str(df[df.eq(current_parent_vendor).any(1)].iloc[0, 4])
        vm_name = str(df[df.eq(current_parent_vendor).any(1)].iloc[0, 6])

        child_index = sub_df.index[sub_df['parent_vendor_code'] == current_parent_vendor].tolist()

        for index in child_index:
            current_child_vendors.append(sub_df.iloc[index, 1])

        for file in os.listdir(new_path):
            file_name = str(os.path.splitext(file))[24:29]
            file_path = new_path + r'\F3_NewProduct_Request_' + str(file_name) + "_alias_" + date_yymmdd + "_vssc_factoryid.xlsx"

            if file_name in current_child_vendors:
                attachments.append(file_path)

        attachments = list(set(attachments))                                                                            # remove duplicates

        mail = outlook.CreateItem(0)
        mail.To = vendor_email
        mail.Cc = vm_email+"; rbs-jp-vulcan-operation@amazon.com"
        mail.Subject = "新規お取り扱いの候補リストについてご確認をお願い致します("+current_month+"月分)"
        mail.HtmlBody = (
          """
          <html>
              <body>
                  <p>
                      <span style="font-family:&quot;Meiryo UI&quot;; font-size:10pt">
          """ + vendor_name +
          """
                      </span>
                  </p>
                  <p>
                      <span style="font-family:&quot;Meiryo UI&quot;; font-size:10pt">
          """ + person_in_charge +
          """
                      </span>
                  </p>
                  <p>
                      <span style="font-family:&quot;Meiryo UI&quot;; font-size:10pt">
          """ + vm_name +
          """
                      </span>
                   </p>
               </body>
          </html>
          """)

        for attachment in attachments:
            mail.Attachments.Add(attachment)

        mail.Display(False)                                                                                             # save the email as a draft

    if (len(codes_not_found)) == 0:
        with open('email.txt', 'w+') as f_email:
            f_email.write(str(len(involved_parent_vendors)) + " email(s) generated. All succeeded.")

    else:
        with open('email.txt', 'w+') as f_email:
            f_email.write(str(len(involved_parent_vendors)) + " email(s) generated.\n")

        for code in codes_not_found:
            with open('email.txt', 'w+') as f_email:
                f_email.write(code + "\n")

        with open('email.txt', 'w+') as f_email:
            f_email.write("Vendor code(s) above cannot be found in ★vendor連絡先リスト_Master.xlsx. Please double-check with the vendor manager.")

else:
    with open('email.txt', 'w+') as f_email:
        f_email.write("Error: No vendor codes are given.")
