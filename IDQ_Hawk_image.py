import pandas as pd
import win32com.client as win32                                                                                         # pip install pywin32
import os
import re
import warnings
from datetime import datetime

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

current_date = str(datetime.today().strftime('%Y-%m-%d'))
username = os.getenv('username')

with open('name.txt', encoding='utf-8') as f:
    signature = str(f.read())
    signature = signature.replace('\n', '<br/>')

file_path = "results.txt"
idq_df = pd.read_csv(file_path, sep="\t")

idq_df.drop(idq_df.index[idq_df['merchant_name'] != 'Retail Merchant'], inplace=True)
idq_df.drop(idq_df.index[idq_df['country'] != 'JP'], inplace=True)
idq_df.drop(idq_df.index[idq_df['gl_product_group_desc'] == 'gl_fresh_produce'], inplace=True)
idq_df.drop(idq_df.index[~idq_df['total_image_count'].isin([0, 1, 2])], inplace=True)

original_df = r'\\ant.amazon.com\dept-as\HND10\JP_F3\3.F3_Retail\2.F3_Instock\Vendor Operation Calendar\★vendor連絡先リスト_Master.xlsx'
vendors = pd.read_excel(io=original_df, sheet_name="Master_List", usecols="B,D:N,R,T,U", header=1)
vm = pd.read_excel(io=original_df, sheet_name="VM名称", usecols="A:C")
merged_vendor_df = pd.merge(vendors, vm, on='VM_name', how='left')

parent_vendor_code = list(merged_vendor_df.iloc[:, 1])
child_vendor_code = merged_vendor_df.iloc[:, 2:11]

tidy_df = pd.DataFrame()

child_list = []
parent_list = []
vendor_name_list = []
vendor_email_list = []
person_in_charge_list = []
vm_login_list = []
vm_name_list = []
vm_email_list = []

for i in range(len(parent_vendor_code)):
    child_list.append(parent_vendor_code[i])
    parent_list.append(parent_vendor_code[i])
    vendor_name_list.append(merged_vendor_df.iloc[i, 0])
    vendor_email_list.append(merged_vendor_df.iloc[i, 13])
    person_in_charge_list.append(merged_vendor_df.iloc[i, 14])
    vm_login_list.append(merged_vendor_df.iloc[i, 12])
    vm_name_list.append(merged_vendor_df.iloc[i, 15])
    vm_email_list.append(merged_vendor_df.iloc[i, 16])

for i in range(len(child_vendor_code)):
    for j in range(len(child_vendor_code.columns)):
        if len(str(child_vendor_code.iloc[i, j]).replace(" ", "")) == 5:
            child_list.append(child_vendor_code.iloc[i, j])
            parent_list.append(merged_vendor_df.iloc[i, 1])
            vendor_name_list.append(merged_vendor_df.iloc[i, 0])
            vendor_email_list.append(merged_vendor_df.iloc[i, 13])
            person_in_charge_list.append(merged_vendor_df.iloc[i, 14])
            vm_login_list.append(merged_vendor_df.iloc[i, 12])
            vm_name_list.append(merged_vendor_df.iloc[i, 15])
            vm_email_list.append(merged_vendor_df.iloc[i, 16])

for i in range(len(vendor_name_list)):
    nb_rep = 1
    while (nb_rep):
        (vendor_name_list[i], nb_rep) = re.subn(r'\([^()]*\)', '', vendor_name_list[i])

    vendor_name_list[i] = re.sub("[\[].*?[\]]", "", vendor_name_list[i])
    vendor_name_list[i] = re.sub("[\(\（].*?[\)\）]", "", vendor_name_list[i])

    vendor_name_list[i] = vendor_name_list[i].strip()

    if vendor_name_list[i].startswith("株式会社"):
        vendor_name_list[i] = vendor_name_list[i][4:]

    if not vendor_name_list[i].endswith("株式会社"):
        vendor_name_list[i] += "株式会社"

for i in range(len(person_in_charge_list)):
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

image_df = pd.DataFrame()
image_df['ASIN'] = idq_df['asin']
image_df['商品名'] = ""
image_df['gl_product_group_desc'] = idq_df['gl_product_group_desc']
image_df['ベンダーコード'] = ""
image_df['画像'] = idq_df['total_image_count']
image_df['画像の追加数'] = ""
image_df['ベンダーコード2'] = ""
image_df['連絡先'] = ""
image_df['宛名'] = ""
image_df['対応（送信）'] = ""
image_df['arias'] = ""
image_df['date'] = ""
image_df['memo'] = ""

image_df.to_excel(current_date + '_IDQ Hawk_Image.xlsx', index=False)

image_lst = image_df['画像'].tolist()
image2_lst = []

for i in image_lst:
    j = "画像を" + str(3-int(i)) + "枚以上追加ください。登録画像に商品裏面の画像がない場合は、商品裏面の画像を追加ください。"
    image2_lst.append(j)

image2_df = pd.DataFrame(image2_lst)
writer = pd.ExcelWriter(current_date + '_IDQ Hawk_Image.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
image2_df.to_excel(writer, index=False, header=None, startcol=5, startrow=1)
writer.save()

image_df = pd.read_excel(current_date + '_IDQ Hawk_Image.xlsx')

for i in range(len(image_df)):
    if len(image_df.iloc[i, 5]) < 48:
        number = image_df.iloc[i, 5][3]
        image_df.loc[i, '画像の追加数'] = "画像を" + number + "枚以上追加ください。登録画像に商品裏面の画像がない場合は、商品裏面の画像を追加ください。"

image_df.to_excel(current_date + '_IDQ Hawk_Image.xlsx', index=False)

if os.path.exists('f3ast.xlsx'):
    f3ast_df = pd.read_excel('f3ast.xlsx', sheet_name="export", usecols="G,K,N")
    merged_df = pd.merge(image_df, f3ast_df, left_on='ASIN', right_on='asin', how='left')

    writer = pd.ExcelWriter(current_date + '_IDQ Hawk_Image.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    merged_df['product_title'].to_excel(writer, index=False, header=None, startcol=1, startrow=1)
    merged_df['vendor_code'].to_excel(writer, index=False, header=None, startcol=3, startrow=1)
    writer.save()

    image_df = pd.read_excel(current_date + '_IDQ Hawk_Image.xlsx')
    image_df = image_df.astype(str)

    merged_vendor_image_df = pd.merge(image_df, tidy_df, left_on='ベンダーコード', right_on='child_vendor_code', how='left')

    writer = pd.ExcelWriter(current_date + '_IDQ Hawk_Image.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    merged_vendor_image_df['parent_vendor_code'].to_excel(writer, index=False, header=None, startcol=6, startrow=1)
    merged_vendor_image_df['vendor_email'].to_excel(writer, index=False, header=None, startcol=7, startrow=1)
    merged_vendor_image_df['person_in_charge'].to_excel(writer, index=False, header=None, startcol=8, startrow=1)
    writer.save()

    image_df = pd.read_excel(current_date + '_IDQ Hawk_Image.xlsx')
    image_df = image_df.astype(str)

    exceptions_df = pd.read_excel('exceptions.xlsx', header=None)
    exceptions_lst = exceptions_df[0].tolist()

    for i in range(len(image_df)):
        if image_df.iloc[i, 3] == "nan":
            image_df.loc[i, '対応（送信）'] = "不要"
            image_df.loc[i, 'memo'] = "vendor codeなし"
            image_df.loc[i, 'arias'] = username
            image_df.loc[i, 'date'] = current_date

        else:
            if image_df.iloc[i, 3] in exceptions_lst:
                image_df.loc[i, '対応（送信）'] = "不要"
                image_df.loc[i, 'memo'] = "vendor code除外"
                image_df.loc[i, 'arias'] = username
                image_df.loc[i, 'date'] = current_date

            else:
                if image_df.iloc[i, 6] == "nan":
                    image_df.loc[i, '対応（送信）'] = "不要"
                    image_df.loc[i, 'memo'] = "連絡先なし"
                    image_df.loc[i, 'arias'] = username
                    image_df.loc[i, 'date'] = current_date

                else:
                    if image_df.iloc[i, 7] == "nan" or image_df.iloc[i, 7] == "(Core BS経由のコミュニケーション)":
                        image_df.loc[i, '対応（送信）'] = "不要"
                        image_df.loc[i, 'memo'] = "連絡先なし"
                        image_df.loc[i, 'arias'] = username
                        image_df.loc[i, 'date'] = current_date

                    else:
                        if image_df.iloc[i, 8] == "nan":
                            image_df.loc[i, '対応（送信）'] = "不要"
                            image_df.loc[i, 'memo'] = "宛名なし"
                            image_df.loc[i, 'arias'] = username
                            image_df.loc[i, 'date'] = current_date

                        else:
                            image_df.loc[i, '対応（送信）'] = "要"
                            image_df.loc[i, 'memo'] = ""
                            image_df.loc[i, 'arias'] = username
                            image_df.loc[i, 'date'] = current_date

        if image_df.iloc[i, 1] == "nan":
            image_df.loc[i, '商品名'] = ""

        if image_df.iloc[i, 3] == "nan":
            image_df.loc[i, 'ベンダーコード'] = ""

        if image_df.iloc[i, 6] == "nan":
            image_df.loc[i, 'ベンダーコード2'] = ""

        if image_df.iloc[i, 7] == "nan":
            image_df.loc[i, '連絡先'] = ""

        if image_df.iloc[i, 8] == "nan":
            image_df.loc[i, '宛名'] = ""

    image_df.to_excel(current_date + '_IDQ Hawk_Image.xlsx', index=False)

    involved_df = image_df[image_df['対応（送信）'] == "要"]

    parent_lst = list(involved_df.iloc[:, 6])
    parent_lst = list(set(parent_lst))
    current_parent_df = pd.DataFrame()
    sent_parent_lst = []

    attachments_path = 'attachments_' + current_date

    if not os.path.exists(attachments_path):
        os.makedirs(attachments_path)

    outlook = win32.Dispatch('outlook.application')

    for current_parent_code in parent_lst:
        current_parent_df = involved_df
        current_parent_df = current_parent_df[(current_parent_df.ベンダーコード2 == current_parent_code)]
        current_parent_df = current_parent_df.iloc[:, 0:6]
        current_parent_df.sort_values(by=['ベンダーコード', 'gl_product_group_desc'], ascending=[True, True])
        current_parent_df.to_excel(attachments_path + r'\画像登録依頼シート' + current_parent_code + '.xlsx', index=False)

        vendor_email = str(involved_df[involved_df.eq(current_parent_code).any(1)].iloc[0, 7])
        person_in_charge = str(involved_df[involved_df.eq(current_parent_code).any(1)].iloc[0, 8])

        attachment = os.path.abspath(attachments_path + r'\画像登録依頼シート' + current_parent_code + '.xlsx')

        mail = outlook.CreateItem(0)
        mail.To = vendor_email
        mail.Subject = "[Amazonフレッシュ][画像追加登録のお願い]"
        mail.HtmlBody = (
            """
            <html>
                <body>
                    <p>
                        <span style="font-family:&quot;Meiryo UI&quot;; font-size:10pt">
            """ + person_in_charge +
            """
                        </span>
                    </p>
                    <p>
                        <span style="font-family:&quot;Meiryo UI&quot;; font-size:10pt">
            """ + signature +
            """
                        </span>
                    </p>
                </body>
            </html>
            """)

        mail.Attachments.Add(attachment)
        mail.Display(False)

        sent_parent_lst.append(current_parent_code)

    pd.DataFrame(sent_parent_lst).to_excel('sent_vendor_codes.xlsx', header=False, index=False)
