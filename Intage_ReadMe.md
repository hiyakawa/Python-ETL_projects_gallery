# Intage Automation Tool

#### Developed by Muyuan Zhang @RBS FLEX Team (zhangmuy@amazon.com)

## Overview

此python应用程序是为日本亚马逊fresh Intage Automation任务而设计。请依照以下步骤运行程序。在运行过程中，用户不需要输入任何值。如需停止运行，请随时关闭黑色窗口。

双击intage_auto_tool\intage.exe，程序将在黑色窗口中运行。

运行两分钟后，intage_auto_tool文件夹中会生成vendorlist.txt和auto_generated_vm_list.xlsx。文件内容示例：

vendorlist.txt
A tidy dataset is exported to root directory.

auto_generated_vm_list.xlsx
child_vendor_code|parent_vendor_code|vendor_name|vendor_email|person_in_charge|vm_login|vm_name|vm_email

此excel文件是\\ant.amazon.com\dept-as\HND10\JP_F3\3.F3_Retail\2.F3_Instock\Vendor Operation Calendar\★vendor連絡先リスト_Master.xlsx的清洁版本。

在程序运行过程中，请不要更改或删除这两个文件。

继续运行三十分钟左右，所有邮件附件将生成在\\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\Vendor送付ファイル\*月送付分。一个名为'*月送付分'的文件夹将被自动生成，其中*为当前月份。用户不需要自行创建文件夹或excel模板。附件命名方式为F3_NewProduct_Request_*vendor code*_alias_*YYMMDD*_vssc_factoryid.xlsx，其中*YYMMDD*为当前日期。

intage_auto_tool文件夹中会生成vendor_code.xlsx、empty_vendor_code.xlsx和attachment.txt。

\\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\vendor_code的非空txt文件涉及的所有vendor code存储在vendor_code.xlsx中。空txt文件涉及的所有vendor code存储在empty_vendor_code.xlsx中。

用户可以在attachment.txt找到需要对应的vendor code个数和txt内容为空的vendor code个数。文件内还列出了所有未在★vendor連絡先リスト_Master.xlsx中找到的vendor code (如NO_VENDOR_CODE) 。

接下来，邮件草稿会在Outlook中自动生成，附件也将自动插入。初次使用本程序时，请检查邮件主题、公司名称、收件人、附件和署名并确认无误。

所有邮件生成后，intage_auto_tool文件夹中会生成email.txt。文件内容示例：

email.txt
\# email(s) generated. All succeeded.

\# 为生成的邮件个数，小于等于需要对应的vendor code总数。email.txt生成后，程序即结束，黑色窗口自行关闭。

如有任何问题或后续需求，请联系开发者。

Copyright(c)	07/01/2022 Muyuan Zhang


*** English version ***

# Intage Automation Tool

#### Developed by Muyuan Zhang @RBS FLEX Team (zhangmuy@amazon.com)

## Overview

This is a python application developed for Amazon JP fresh Intage Automation task. Please follow the steps below to run the application. Users are not required to make any inputs during the whole process. Please close the console window anytime if you need to stop the application.

Please double-click intage_auto_tool\intage.exe and run the application in the console window.

vendorlist.txt and auto_generated_vm_list.xlsx will be generated in the intage_auto_tool folder after running for 2 minutes.

Examples:
vendorlist.txt
A tidy dataset is exported to root directory.

auto_generated_vm_list.xlsx
child_vendor_code|parent_vendor_code|vendor_name|vendor_email|person_in_charge|vm_login|vm_name|vm_email

This excel file is a tidy version of \\ant.amazon.com\dept-as\HND10\JP_F3\3.F3_Retail\2.F3_Instock\Vendor Operation Calendar\★vendor連絡先リスト_Master.xlsx.

Please do not modify or delete the files above while the application is running.

All the attachments will be generated in \\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\Vendor送付ファイル\*月送付分 after another 30 minutes. A new folder named '*月送付分' will be automatically created, where * is current month. Users are not required to create any folders or excel templates. Attachments are named as F3_NewProduct_Request_*vendor code*_alias_*YYMMDD*_vssc_factoryid.xlsx, where *YYMMDD* is the date today.

vendor_code.xlsx, empty_vendor_code.xlsx and attachment.txt will be generated in the intage_auto_tool folder.

vendor_code.xlsx contains involved vendor codes from txt files in \\ant\dept-eu\EUF3\F3-Reporting\Country_WBRs\JP\Selection\MSC_JAN_list\vendor_code. empty_vendor_code.xlsx contains vendor codes whose txt files are empty.

Users can find the number of involved vendor codes and empty vendor codes in attachment.txt. Any vendor codes that cannot be found in ★vendor連絡先リスト_Master.xlsx (NO_VENDOR_CODE, etc.) are also listed.

Email drafts will be generated in Outlook with automatically inserted attachments. Please double-check the subject, company name, title, attachments and sign for the first time you run this application.

email.txt will be generated in the intage_auto_tool folder after all the email drafts are displayed.

Example:
email.txt
\# email(s) generated. All succeeded.

\# is the total number of emails, which is equal to or smaller than the number of involved vendor codes. After email.txt is generated, the application will finish and the console window will close.

Please contact the developer if there are any issues or further requirements.

Copyright(c)	07/01/2022 Muyuan Zhang
