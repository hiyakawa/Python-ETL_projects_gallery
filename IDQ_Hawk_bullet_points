import pandas as pd
import numpy as np
import os
import warnings
from datetime import datetime

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

current_date = str(datetime.today().strftime('%Y-%m-%d'))

file_path = "results.txt"
idq_df = pd.read_csv(file_path, sep="\t")

idq_df.drop(idq_df.index[idq_df['merchant_name'] != 'Retail Merchant'], inplace=True)
idq_df.drop(idq_df.index[idq_df['country'] != 'JP'], inplace=True)
idq_df.drop(idq_df.index[~idq_df['bullet_point_count'].isin([0, 1, 2, 3, 4])], inplace=True)

idq_df.to_excel(current_date + '_IDQ Hawk_BulletPoint.xlsx', index=False)

bp_df = pd.DataFrame()
bp_df['ASIN'] = idq_df['asin']
bp_df['商品名'] = ""
bp_df['gl_product_group_desc'] = idq_df['gl_product_group_desc']
bp_df['原産国&原産地?'] = ""
bp_df['Bullet Pointの数'] = idq_df['bullet_point_count']
bp_df['Bullet Pointの追加数（数式あり）'] = ""
bp_df['SC対応'] = ""
bp_df['DP反映'] = ""
bp_df['Comment'] = ""
bp_df['Arias'] = ""
bp_df['Date'] = ""

bp_df.to_excel(current_date + '_IDQ Hawk_BulletPoint.xlsx', index=False)

bp_lst = bp_df['Bullet Pointの数'].tolist()
bp2_lst = []
bp_gl_lst = bp_df['gl_product_group_desc'].tolist()
bp2_gl_lst = []

for i in bp_lst:
    j = "Bullet Pointを" + str(5-int(i)) + "つ以上追加ください。"
    bp2_lst.append(j)

for i in bp_gl_lst:
    if i == "gl_fresh_produce":
        bp2_gl_lst.append("N")

    else:
        bp2_gl_lst.append("")

bp2_df = pd.DataFrame(bp2_lst)
bp2_gl_df = pd.DataFrame(bp2_gl_lst)
writer = pd.ExcelWriter(current_date + '_IDQ Hawk_BulletPoint.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
index_country_bp = bp_df.columns.get_loc('原産国&原産地?')
index_required_bp = bp_df.columns.get_loc('Bullet Pointの追加数（数式あり）')
bp2_gl_df.to_excel(writer, index=False, header=None, startcol=index_country_bp, startrow=1)
bp2_df.to_excel(writer, index=False, header=None, startcol=index_required_bp, startrow=1)
writer.save()

bp_df = pd.read_excel(current_date + '_IDQ Hawk_BulletPoint.xlsx')

if os.path.exists('f3ast.xlsx'):
    f3ast_df = pd.read_excel('f3ast.xlsx', sheet_name="export", usecols="G,K,N")
    merged_df = pd.merge(bp_df, f3ast_df, left_on='ASIN', right_on='asin', how='left')

    writer = pd.ExcelWriter(current_date + '_IDQ Hawk_BulletPoint.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')
    index_name = bp_df.columns.get_loc('商品名')
    merged_df['product_title'].to_excel(writer, index=False, header=None, startcol=index_name, startrow=1)
    writer.save()

    bp_df = pd.read_excel(current_date + '_IDQ Hawk_BulletPoint.xlsx')

    if os.path.exists('sc_next.xlsx'):
        scn_df = pd.read_excel('sc_next.xlsx')
        scn_df = scn_df.astype(str)

        scn_df['重量-内容量'] = ""
        scn_df['商品サイズ'] = ""
        scn_df['ブラント名'] = ""
        scn_df['メーカー名'] = ""
        scn_df['原産地'] = ""
        scn_df['商品の重量'] = ""
        scn_df['内容量'] = ""
        scn_df['原材料'] = ""

        index_brand = scn_df.columns.get_loc('brand.value')
        index_manufacturer = scn_df.columns.get_loc('manufacturer.value')
        index_region = scn_df.columns.get_loc('subregion_of_origin.value')

        index_grossweight_value = scn_df.columns.get_loc('item_package_weight.normalized_value.value')
        index_grossweight_unit = scn_df.columns.get_loc('item_package_weight.normalized_value.unit')
        index_grossweight_value_override = scn_df.columns.get_loc('item_package_weight.value')
        index_grossweight_unit_override = scn_df.columns.get_loc('item_package_weight.unit')

        index_netweight_value = scn_df.columns.get_loc('item_weight.normalized_value.value')
        index_netweight_unit = scn_df.columns.get_loc('item_weight.normalized_value.unit')
        index_netweight_value_override = scn_df.columns.get_loc('item_weight.value')
        index_netweight_unit_override = scn_df.columns.get_loc('item_weight.unit')

        index_len = scn_df.columns.get_loc('item_dimensions.length.normalized_value.value')
        index_wid = scn_df.columns.get_loc('item_dimensions.width.normalized_value.value')
        index_height = scn_df.columns.get_loc('item_dimensions.height.normalized_value.value')

        index_ingredients = scn_df.columns.get_loc('ingredients.value')

        newindex_gross_minus_net = scn_df.columns.get_loc('重量-内容量')
        newindex_size = scn_df.columns.get_loc('商品サイズ')
        newindex_brand = scn_df.columns.get_loc('ブラント名')
        newindex_maker = scn_df.columns.get_loc('メーカー名')
        newindex_region = scn_df.columns.get_loc('原産地')
        newindex_grossweight = scn_df.columns.get_loc('商品の重量')
        newindex_netweight = scn_df.columns.get_loc('内容量')
        newindex_ingredients = scn_df.columns.get_loc('原材料')

        pound_to_gram = 453.592
        inch_to_mm = 25.4

        for i in range(len(scn_df)):
            if scn_df.iloc[i, index_brand] != "nan" and scn_df.iloc[i, index_brand] != "不明" and scn_df.iloc[i, index_brand] != "XXXXXbrand":
                scn_df.loc[i, 'ブラント名'] = "ブラント名: " + str(scn_df.iloc[i, index_brand])

            if scn_df.iloc[i, index_manufacturer] != "nan" and scn_df.iloc[i, index_manufacturer] != "XXXXX":
                scn_df.loc[i, 'メーカー名'] = "メーカー名: " + str(scn_df.iloc[i, index_manufacturer])

            if scn_df.iloc[i, index_region] != "nan":
                if scn_df.iloc[i, index_region].startswith('原産'):
                    scn_df.loc[i, '原産地'] = str(scn_df.iloc[i, index_region])
                else:
                    scn_df.loc[i, '原産地'] = "原産地: " + str(scn_df.iloc[i, index_region])

            if scn_df.iloc[i, index_grossweight_value] != "nan":
                if scn_df.iloc[i, index_grossweight_unit] == "pounds":
                    if float(scn_df.iloc[i, index_grossweight_value])*pound_to_gram < 1:
                        scn_df.loc[i, '商品の重量'] = "商品の重量: " + str(round(float(scn_df.iloc[i, index_grossweight_value])*pound_to_gram, 2)) + " g"

                    else:
                        scn_df.loc[i, '商品の重量'] = "商品の重量: " + str(round(float(scn_df.iloc[i, index_grossweight_value])*pound_to_gram)) + " g"

                else:
                    scn_df.loc[i, '商品の重量'] = "商品の重量: " + str(scn_df.iloc[i, index_grossweight_value_override]) + " " + str(scn_df.iloc[i, index_grossweight_unit_override])

                if scn_df.iloc[i, newindex_grossweight] == "商品の重量: 0.00 g":
                    scn_df.loc[i, '商品の重量'] = ""

                if scn_df.iloc[i, newindex_grossweight] == "商品の重量: 1.0 g":
                    scn_df.loc[i, '商品の重量'] = "商品の重量: 1 g"

            if scn_df.iloc[i, index_netweight_value] != "nan":
                if scn_df.iloc[i, index_netweight_unit] == "pounds":
                    if float(scn_df.iloc[i, index_netweight_value])*pound_to_gram < 1:
                        scn_df.loc[i, '内容量'] = "内容量: " + str(round(float(scn_df.iloc[i, index_netweight_value])*pound_to_gram, 2)) + " g"

                    else:
                        scn_df.loc[i, '内容量'] = "内容量: " + str(round(float(scn_df.iloc[i, index_netweight_value])*pound_to_gram)) + " g"

                else:
                    scn_df.loc[i, '内容量'] = "内容量: " + str(scn_df.iloc[i, index_netweight_value_override]) + " " + str(scn_df.iloc[i, index_netweight_unit_override])

                if scn_df.iloc[i, newindex_netweight] == "内容量: 0.0 g":
                    scn_df.loc[i, '内容量'] = ""

                if scn_df.iloc[i, newindex_netweight] == "内容量: 1.0 g":
                    scn_df.loc[i, '内容量'] = "内容量: 1 g"

            if scn_df.iloc[i, index_height] != "nan" and scn_df.iloc[i, index_height] != "0.0":
                scn_df.loc[i, '商品サイズ'] = "商品サイズ (高さ×奥行×幅): " + str(round(float(scn_df.iloc[i, index_len])*inch_to_mm)) + "mm×" + str(round(float(scn_df.iloc[i, index_wid])*inch_to_mm)) + "mm×" + str(round(float(scn_df.iloc[i, index_height])*inch_to_mm)) + "mm"

                if scn_df.iloc[i, newindex_size].startswith("商品サイズ (高さ×奥行×幅): 0mm"):
                    scn_df.loc[i, '商品サイズ'] = ""

            if scn_df.iloc[i, index_ingredients] != "nan" and scn_df.iloc[i, index_ingredients] != "―":
                if "原材料" in scn_df.iloc[i, index_ingredients]:
                    scn_df.loc[i, '原材料'] = str(scn_df.iloc[i, index_ingredients])

                else:
                    scn_df.loc[i, '原材料'] = "原材料: " + str(scn_df.iloc[i, index_ingredients])

        scn_tidy_df = pd.DataFrame()
        scn_tidy_df['ASIN'] = scn_df['asin']
        scn_tidy_df['重量-内容量'] = scn_df['重量-内容量']
        scn_tidy_df['bp_1'] = scn_df['bullet_point.value']
        scn_tidy_df['bp_2'] = scn_df['bullet_point#2.value']
        scn_tidy_df['bp_3'] = scn_df['bullet_point#3.value']
        scn_tidy_df['bp_4'] = scn_df['bullet_point#4.value']
        scn_tidy_df['bp_5'] = scn_df['bullet_point#5.value']
        scn_tidy_df['商品サイズ'] = scn_df['商品サイズ']
        scn_tidy_df['ブラント名'] = scn_df['ブラント名']
        scn_tidy_df['メーカー名'] = scn_df['メーカー名']
        scn_tidy_df['原産地'] = scn_df['原産地']
        scn_tidy_df['商品の重量'] = scn_df['商品の重量']
        scn_tidy_df['内容量'] = scn_df['内容量']
        scn_tidy_df['原材料'] = scn_df['原材料']

        if os.path.exists('sc_legacy.xlsx'):
            scl_df = pd.read_excel('sc_legacy.xlsx', header=1)
            scl_df = scl_df.astype(str)

            scl_df['原産国名'] = ""

            index_country = scl_df.columns.get_loc('country_of_origin.value')
            newindex_country = scl_df.columns.get_loc('原産国名')

            if os.path.exists('countrycode.xlsx'):
                countrycode_df = pd.read_excel('countrycode.xlsx')
                countrycode_df = countrycode_df.astype(str)
                index_country_name = countrycode_df.columns.get_loc('日本語名')

                for i in range(len(scl_df)):
                    if scl_df.iloc[i, index_country] != "" and scl_df.iloc[i, index_country] != "nan" and countrycode_df['国名コード'].str.contains(scl_df.iloc[i, index_country]).any():
                        index_cur_country = int(np.where(countrycode_df['国名コード'] == scl_df.iloc[i, index_country])[0])
                        scl_df.loc[i, '原産国名'] = "原産国名: " + str(countrycode_df.iloc[index_cur_country, index_country_name])

                    if scl_df.iloc[i, newindex_country] == "原産国名: ナミビア共和国":
                        scl_df.loc[i, '原産国名'] = ""

            scl_tidy_df = pd.DataFrame()
            scl_tidy_df['ASIN'] = scl_df['asin']
            scl_tidy_df['原産国名'] = scl_df['原産国名']

            merged_scn_scl_df = pd.merge(scn_tidy_df, scl_tidy_df, on='ASIN', how='left')
            merged_bp_sc_df = pd.merge(bp_df, merged_scn_scl_df, on='ASIN', how='left')
            merged_bp_sc_df = merged_bp_sc_df.astype(str)

            merged_bp_sc_df['bp_count'] = ""

            index_gl = merged_bp_sc_df.columns.get_loc('原産国&原産地?')
            index_first_bp = merged_bp_sc_df.columns.get_loc('bp_1')

            for i in range(len(merged_bp_sc_df)):
                if merged_bp_sc_df.iloc[i, index_gl] == "N":
                    merged_bp_sc_df.loc[i, '原産国名'] = ""
                    merged_bp_sc_df.loc[i, '原産地'] = ""

                bp_counter = 0

                for j in range(index_first_bp, (int(index_first_bp) + 5)):
                    if merged_bp_sc_df.iloc[i, j] != "nan":
                        bp_counter += 1

                    if "ブラント" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, 'ブラント名'] = ""

                    if "メーカー" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, 'メーカー名'] = ""

                    if "原産国" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '原産国名'] = ""

                    if "原産地" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '原産地'] = ""

                    if "重量" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '商品の重量'] = ""

                    if "内容量" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '内容量'] = ""

                    if "サイズ" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '商品サイズ'] = ""

                    if "原材料" in merged_bp_sc_df.iloc[i, j]:
                        merged_bp_sc_df.loc[i, '原材料'] = ""

                merged_bp_sc_df.loc[i, 'bp_count'] = str(bp_counter)

            result_df = pd.DataFrame()
            result_df['asin'] = merged_bp_sc_df['ASIN']
            result_df['商品名'] = merged_bp_sc_df['商品名']
            result_df['gl_product_group'] = merged_bp_sc_df['gl_product_group_desc']
            result_df['bp_count'] = merged_bp_sc_df['bp_count']
            result_df['重量-内容量'] = merged_bp_sc_df['重量-内容量']
            result_df['bullet_point.value'] = merged_bp_sc_df['bp_1']
            result_df['bullet_point#2.value'] = merged_bp_sc_df['bp_2']
            result_df['bullet_point#3.value'] = merged_bp_sc_df['bp_3']
            result_df['bullet_point#4.value'] = merged_bp_sc_df['bp_4']
            result_df['bullet_point#5.value'] = merged_bp_sc_df['bp_5']
            result_df['商品サイズ'] = merged_bp_sc_df['商品サイズ']
            result_df['ブラント名'] = merged_bp_sc_df['ブラント名']
            result_df['メーカー名'] = merged_bp_sc_df['メーカー名']
            result_df['原産国名'] = merged_bp_sc_df['原産国名']
            result_df['原産地'] = merged_bp_sc_df['原産地']
            result_df['商品の重量'] = merged_bp_sc_df['商品の重量']
            result_df['内容量'] = merged_bp_sc_df['内容量']
            result_df['原材料'] = merged_bp_sc_df['原材料']

            index_result_gross_weight = result_df.columns.get_loc('商品の重量')
            index_result_net_weight = result_df.columns.get_loc('内容量')

            result_df = result_df.replace(np.nan, "", regex=True)
            result_df = result_df.replace("nan", "", regex=True)

            for i in range(len(result_df)):
                if len(result_df.iloc[i, index_result_gross_weight]) > 5:
                    if len(result_df.iloc[i, index_result_net_weight]) > 3:
                        result_df.loc[i, '重量-内容量'] = float(result_df.iloc[i, index_result_gross_weight][7:-2]) - float(result_df.iloc[i, index_result_net_weight][5:-2])

            result_df.to_excel(current_date + '_raw_result.xlsx', index=False)

            first_result_df = pd.DataFrame()
            first_result_df['asin'] = result_df['asin']
            first_result_df['sc_vendor_name'] = "internal_JP_high"
            first_result_df['商品名'] = result_df['商品名']
            first_result_df['gl_product_group'] = result_df['gl_product_group']
            first_result_df['bp_count'] = result_df['bp_count']
            first_result_df['重量-内容量'] = result_df['重量-内容量']
            first_result_df['重量あり'] = ""

            second_result_df = pd.DataFrame()
            second_result_df['asin'] = result_df['asin']
            second_result_df['bullet_point.value'] = result_df['bullet_point.value']
            second_result_df['bullet_point#2.value'] = result_df['bullet_point#2.value']
            second_result_df['bullet_point#3.value'] = result_df['bullet_point#3.value']
            second_result_df['bullet_point#4.value'] = result_df['bullet_point#4.value']
            second_result_df['bullet_point#5.value'] = result_df['bullet_point#5.value']
            second_result_df['6'] = result_df['商品サイズ']
            second_result_df['7'] = result_df['ブラント名']
            second_result_df['8'] = result_df['メーカー名']
            second_result_df['9'] = result_df['原産国名']
            second_result_df['10'] = result_df['原産地']
            second_result_df['11'] = result_df['商品の重量']
            second_result_df['12'] = result_df['内容量']
            second_result_df['13'] = result_df['原材料']

            temp_second_result_df = second_result_df.replace("", np.nan, regex=True)

            def squeeze_nan(x):
                original_columns = x.index.tolist()

                squeezed = x.dropna()
                squeezed.index = [original_columns[n] for n in range(squeezed.count())]

                return squeezed.reindex(original_columns, fill_value=np.nan)

            temp_second_result_df = temp_second_result_df.apply(squeeze_nan, axis=1)

            merged_result_df = pd.merge(first_result_df, temp_second_result_df, on='asin', how='left')

            index_result_first_bp = merged_result_df.columns.get_loc('bullet_point.value')
            index_compare = merged_result_df.columns.get_loc('重量-内容量')

            upload_result = merged_result_df.replace(np.nan, "", regex=True)

            for i in range(len(upload_result)):
                has_gross_weight = False
                has_net_weight = False

                for j in range(index_result_first_bp, (int(index_result_first_bp) + 5)):
                    if "商品の重量" in upload_result.iloc[i, j]:
                        has_gross_weight = True
                    elif "内容量" in upload_result.iloc[i, j]:
                        has_net_weight = True

                if (not has_gross_weight) or (not has_net_weight):
                    upload_result.loc[i, '重量-内容量'] = ""

                if has_gross_weight:
                    upload_result.loc[i, '重量あり'] = "Y"

            upload_result = upload_result.replace(np.nan, "", regex=True)
            upload_result.to_excel(current_date + '_upload_result.xlsx', index=False)
