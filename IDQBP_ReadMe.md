IDQ Hawk Bullet Point Automation Tool

Developed by Muyuan Zhang @RBS FLEX Team (zhangmuy@amazon.com)

此python应用程序是为日本亚马逊fresh IDQ Hawk Bullet Point自动化任务而设计。请依照以下步骤运行程序。如需停止运行，请随时关闭黑色窗口。

#### 准备输入文件
首先，请按照SOP给出的地址https://datacentral.a2z.com/dw-platform/servlet/dwp/template/EtlViewExtractJobs.vm/job_profile_id/10358527，点击Job #22268993 Revision #2: Combined ASIN Level Data with Merchant Name旁的时钟图标，进入Status为Success且Dataset Date距今最近的一条结果，在网页右侧选中UTF-8与Text (请注意此处与SOP要求不同)，下载results.txt。如果下载的文件没有拓展名，请手动添加.txt为文件后缀。

#### 第一次运行
解压idqhawk_bulletpoint.zip后，将results.txt移入idqhawk_bulletpoint文件夹。在运行过程中，请不要打开、更改或删除这个输入文件。双击idqhawk_bulletpoint\idqhawk_bulletpoint.exe，程序将在黑色窗口中运行。

运行约一分钟后，idqhawk_bulletpoint文件夹中会生成名为当日日期+_IDQ Hawk_BulletPoint.xlsx的excel文件，同时，黑色窗口自动关闭。文件内容示例：

2022-07-20_IDQ Hawk_BulletPoint.xlsx
ASIN|商品名|gl_product_group_desc|原産国&原産地?|Bullet Pointの数|Bullet Pointの追加数（数式あり）|SC対応|DP反映|Comment|Arias|Date

此excel包含所有需要作业的ASIN。请复制ASIN列的所有行，进入F3AST系统，在https://www.amznselection.com/sourcing解除筛选，并选取Region为UFF-Kanagawa/UJF1/J00S，结果约有75,000条。
如果ASIN数超过8000，请选择Select All Rows，点击网页右上的导出按钮，导出所有结果，将导出的excel文件另存为f3ast.xlsx。此过程约耗时30min。
如果ASIN数少于8000，可以仅查询涉及的ASIN，导出所有结果，将导出的excel文件另存为f3ast.xlsx。

Select Central Next的具体查找和导出流程请见idqhawk_bulletpoint\sc_next.PNG的截图示例。请务必注意，搜索字段宁多勿少，且必填字段必须准确。请将导出文件另存为idqhawk_bulletpoint\sc_next.xlsx。

#### 第二次运行
再次双击idqhawk_bulletpoint\idqhawk_bulletpoint.exe，程序将在黑色窗口中运行。在运行过程中，请不要打开、更改或删除三个输入文件。idqhawk_bulletpoint\当日日期+_IDQ Hawk_BulletPoint.xlsx会被写入商品名。SC対応、DP反映和完成日期等需要作业者后续手动填写。

参考结果和供上传的结果将分别生成在idqhawk_bulletpoint\当日日期+_raw_result.xlsx和idqhawk_bulletpoint\当日日期+_upload_result.xlsx。当日日期+_upload_result.xlsx生成后，程序即结束，黑色窗口自行关闭。文件内容示例：

2022-07-20_raw_result.xlsx
asin|商品名|gl_product_group|bp_count|重量-内容量|bullet_point.value|bullet_point#2.value|bullet_point#3.value|bullet_point#4.value|bullet_point#5.value|商品サイズ|ブラント名|メーカー名|原産国名|原産地|商品の重量|内容量|原材料

这些字段分别代表：
ASIN;
商品名;
产品线分类;
原有bullet point个数;
商品重量减内容量的值;
原有bullet point的前五个;
商品サイズ;
ブラント名;
メーカー名;
原産国名;
原産地;
商品の重量;
内容量;
原材料。

2022-07-20_upload_result.xlsx
asin|sc_vendor_name|商品名|gl_product_group|bp_count|重量-内容量|重量あり|bullet_point.value|bullet_point#2.value|bullet_point#3.value|bullet_point#4.value|bullet_point#5.value|6|7|8|9|10|11|12|13

这些字段分别代表：
ASIN;
vendor (均为internal_JP_high);
商品名;
产品线分类;
原有bullet point个数;
商品重量减内容量的值;
前五个bullet point内是否含有商品重量;
已填充的五个bullet point;
供候补的 (最多) 八个bullet point。

#### 后续作业
请注意，当日日期+_upload_result.xlsx内"重量-内容量"字段非空的行对应的ASIN前五个bullet point同时含有重量及内容量。请在DP核对并寻找更合适的属性以填充bullet point。

请筛选当日日期+_upload_result.xlsx内"bullet_point#5.value"字段为空的行，这些ASIN的所有信息不足以填充5个bullet point。请在DP核对并寻找更合适的属性以填充bullet point。

如有任何问题或后续需求，请联系开发者。

Copyright(c)	07/20/2022 Muyuan Zhang
