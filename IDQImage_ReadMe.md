# IDQ Hawk Image Automation Tool

#### Developed by Muyuan Zhang @RBS FLEX Team (zhangmuy@amazon.com)

## Overview

此python应用程序是为日本亚马逊fresh IDQ Hawk Image自动化任务而设计。请依照以下步骤运行程序。如需停止运行，请随时关闭黑色窗口。

请注意两个压缩包对应不同功能，idqhawk_image_display.zip的功能是将生成的邮件存为草稿，而idqhawk_image_send.zip则会直接发送邮件。除此之外，这两个应用程序的使用步骤和输出文件完全相同。以下使用步骤均以idqhawk_image_display.zip为例。

!在运行idqhawk_image_send.zip内的应用程序前，请务必检查源数据、签名和例外vendor codes等对应的输入文档!

首先，请按照SOP给出的地址，点击Job #22268993 Revision #2: Combined ASIN Level Data with Merchant Name旁的时钟图标，进入Status为Success且Dataset Date距今最近的一条结果，在网页右侧选中UTF-8与Text (请注意此处与SOP要求不同)，下载results.txt。如果下载的文件没有拓展名，请手动添加.txt为文件后缀。

解压idqhawk_image_display.zip后，将results.txt移入idqhawk_image_display文件夹。请检查并修改此文件夹内的exceptions.xlsx (如有无需作业的vendor code更新) 和name.txt (此文件内的文字将会出现在所有邮件签名处)。如需在签名内换行，请按enter键正常换行。

检查上述三个输入文件 (idqhawk_image_display\results.txt, idqhawk_image_display\name.txt, idqhawk_image_display\exceptions.xlsx) 无误后，双击idqhawk_image_display\idqhawk_image_display.exe，程序将在黑色窗口中运行。在运行过程中，请不要打开、更改或删除这三个文件。

运行约半分钟后，idqhawk_image_display文件夹中会生成名为当日日期+_IDQ Hawk_Image.xlsx的excel文件，同时，黑色窗口自动关闭。文件内容示例：

*2022-07-13_IDQ Hawk_Image.xlsx*

ASIN|商品名|gl_product_group_desc|ベンダーコード|画像|画像の追加数|ベンダーコード2|連絡先|宛名|対応（送信）|arias|date|memo

此excel包含所有需要作业的ASIN。请复制ASIN列的所有行，进入F3AST系统，按照SOP步骤导出结果，将导出文件另存为idqhawk_image_display\f3ast.xlsx作为第四个输入文件。同样，在运行过程中，请不要打开、更改或删除这个文件。

再次双击idqhawk_image_display\idqhawk_image_display.exe，程序将在黑色窗口中运行。

idqhawk_image_display\当日日期+_IDQ Hawk_Image.xlsx会被写入最终结果和完成日期。

所有邮件附件将生成在idqhawk_image_display\attachments_+当日日期文件夹。附件命名方式为"画像登録依頼シート+vendor code.xlsx"。

接下来，邮件草稿会在Outlook中自动生成，附件也将自动插入。初次使用本程序时，请检查邮件主题、公司名称、收件人、附件和署名并确认无误。

所有邮件生成后，idqhawk_image_display文件夹中会生成sent_vendor_codes.xlsx，包含生成的所有邮件涉及的parent vendor code，总行数等于生成的邮件个数。sent_vendor_codes.xlsx生成后，程序即结束，黑色窗口自行关闭。

如有任何问题或后续需求，请联系开发者。

Copyright(c)	07/01/2022 Muyuan Zhang
