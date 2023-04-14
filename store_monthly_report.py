import streamlit as st
import pandas as pd
import openpyxl as xl
from glob import glob
import datetime as dt
import time
import numpy as np
import os
import jaconv
import csv
from datetime import datetime as dt
import datetime
import mojimoji
import glob
import chardet
from datetime import date



st.write('<h3>承天貿易有限会社月次ストア毎集計システム</h3>', unsafe_allow_html=True)
st.write('注意：ゆうパック代金は明細が出ないため中央値での概算となります。', unsafe_allow_html=True)

if st.checkbox('使用方法を表示'):
    st.write('①左のサイドバーより各種データを読み込みます<br>②卸金額不明分の内容を確認し、NEWNEW_DBへ商品登録を依頼します。<br>\
             ③自動計算されないファイルをダウンロードし手動で運賃の計算を行います。<br>※リロードすると各種ファイルの登録が消える場合がありますので、その場合は再度読み込みを行なってください。', unsafe_allow_html=True)
if st.checkbox('各ファイルのダウンロードについてを表示（編集中）'):
    st.write('<h5>● JPON出力データ</h5>', unsafe_allow_html=True)
    st.write('オーダー検索を開き、左上の検索パターンから”データ分析用出力”発送日に集計期間を入力し、検索。データを全選択して”CSV出力ボタン”を押してCSVファイル出力ウインドウを開く、パターン名称より”データ抽出(池田 月次集計用 tanaorosi”\
             を選んで、CSV出力ボタンを押す。(出力先フォルダは任意に変更してください。)')
    st.write('<h5>● NEWNEW_DB.xlsx</h5>', unsafe_allow_html=True)
    st.write('共有フォルダより最新のNEWNEW_DB.xlsxを読み込んでください。')

st.sidebar.write('<h4>必要なファイルを読み込んでいきます</h4>', unsafe_allow_html=True)

st.sidebar.write('①　各種ファイルを読み込んでいきます')

st.sidebar.write('■　JPON出力データの読み込み')
uploaded_file = st.sidebar.file_uploader("JPON出力集計用データの読み込みファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    main_data_df = pd.read_csv(uploaded_file, encoding="cp932")

num_rows, num_cols = main_data_df.shape
st.write('入力集計データ行数（総数）:', num_rows)

main_data_df['配送伝票番号'] = main_data_df['配送伝票番号'].astype(float).astype(str)
main_data_df['配送伝票番号'] = main_data_df['配送伝票番号'].str.slice(0, -2)


#メインデータの処理ここから

main_data_df = main_data_df.replace(' ', '').replace('　', '')#空白の削除
main_data_df = main_data_df.fillna(0)


main_data_df = main_data_df.replace('通常', '', regex=True)
main_data_df = main_data_df.replace('予約', '', regex=True)
main_data_df = main_data_df.replace('\(携帯\)', '', regex=True)


#メインデータのデータ型変更と半角変換


main_data_df['変換商品名'] = main_data_df['変換商品名'].astype(str)
main_data_df['変換商品名'] = main_data_df['変換商品名'].apply(mojimoji.zen_to_han)
main_data_df['変換商品名'] = main_data_df['変換商品名'].str.strip()
#/メインデータのデータ型変更と半角変換、空白の除去

main_data_df['配送伝票番号'] = main_data_df['配送伝票番号'].astype(str)
main_data_df['配送伝票番号'].replace('.0', '')


main_data_df['受注番号'] = main_data_df['受注番号'].astype(str)

main_data_df = main_data_df[['受注番号', '顧客分類', '発送日', '配送伝票番号', '配送業者', '変換商品名', '購入品数量' ]]



#送料自動計算できない分だけ抽出
toll = main_data_df.loc[main_data_df['配送業者'] == '佐川急便[チャーター便]']
nituu_tp = main_data_df.loc[main_data_df['配送業者'] == '日通トランスポート']
kyushu_k = main_data_df.loc[main_data_df['配送業者'] == '九州航空']




#メインデータの分岐、セット商品と単品
main_data_df_set_item = main_data_df[main_data_df['変換商品名'].str.contains('&', '*')]
main_data_df_single_item = main_data_df[~main_data_df['変換商品名'].str.contains('&', '*')]
#/メインデータの分岐、セット商品と単品





#分岐データをさらに分岐、
main_data_df_set_item_ship = main_data_df_set_item.drop_duplicates(subset='配送伝票番号')
main_data_df_single_item_ship = main_data_df_single_item.drop_duplicates(subset='配送伝票番号')






#disp
if st.checkbox('分岐データを表示'):
    st.write('main_data_df_single_item')
    st.write(main_data_df_single_item)
    st.write('main_data_df_set_item')
    st.write(main_data_df_set_item)
    st.write('main_data_df_single_item_ship')
    st.write(main_data_df_single_item_ship)
    st.write('main_data_df_set_item_ship')
    st.write(main_data_df_set_item_ship)
    num_rows, num_cols = main_data_df_single_item.shape
    st.write('シングル商品行数:', num_rows)
    num_rows, num_cols = main_data_df_set_item.shape
    st.write('セット商品行数:', num_rows)







#単品の処理開始
#変換商品名を※で区切り、商品名と数量に分割
df_ = main_data_df_single_item['変換商品名'].str.split('*', expand=True)
#分割した項目にカラム名を設定
df_.columns = ['item_name', 'quantity',]
#欠損値に「１」を代入
df_ = df_.fillna('1')      

#インデックスをキーにしてマージ
main_data_df_single_item = pd.merge(main_data_df_single_item, df_, left_index=True, right_index=True)
#quantityを整数型に変換
main_data_df_single_item['quantity'] = main_data_df_single_item['quantity'].astype('int64')
#販売商品数量を計算し、カラムを追加。
main_data_df_single_item['販売数量'] = main_data_df_single_item['購入品数量'] * main_data_df_single_item['quantity']






st.sidebar.write('■　NEWNEW_DB')
uploaded_file = st.sidebar.file_uploader("NEWNEW_DB.xlsxファイル最新版を読み込んでください", type=["xlsx"])
if uploaded_file is not None:
    df_newnew_db = pd.read_excel(uploaded_file)

df_newnew_db.replace(' ', '').replace('　', '')#空白の削除
#変換商品名を半角変換
df_newnew_db['変換商品名'] = df_newnew_db['変換商品名'].astype(str)
df_newnew_db['変換商品名'] = df_newnew_db['変換商品名'].apply(mojimoji.zen_to_han)
df_newnew_db['変換商品名'] = df_newnew_db['変換商品名'].str.strip()

df_newnew_db['卸価格'] = df_newnew_db['卸価格'].fillna(0).astype(int)





st.sidebar.write('■　佐川急便')
uploaded_file = st.sidebar.file_uploader("佐川急便（承天貿易）運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    sagawa_syouten = pd.read_csv(uploaded_file, encoding="cp932")
sagawa_syouten['お問合せNO'] = sagawa_syouten['お問合せNO'].astype(str)
sagawa_syouten = sagawa_syouten[['顧客管理番号', 'お問合せNO', '運賃合計金額']]
sagawa_syouten = sagawa_syouten.rename(columns={'運賃合計金額':'佐川急便運賃' })



uploaded_file = st.sidebar.file_uploader("佐川急便（GOODLIFE）運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    sagawa_goodlife = pd.read_csv(uploaded_file, encoding="cp932")
sagawa_goodlife['お問合せNO'] = sagawa_goodlife['お問合せNO'].astype(str)
sagawa_goodlife = sagawa_goodlife[['顧客管理番号', 'お問合せNO', '運賃合計金額']]
sagawa_goodlife = sagawa_goodlife.rename(columns={'運賃合計金額':'佐川急便運賃' })


uploaded_file = st.sidebar.file_uploader("佐川急便（昌隆）運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    sagawa_masataka = pd.read_csv(uploaded_file, encoding="cp932")
sagawa_masataka['お問合せNO'] = sagawa_masataka['お問合せNO'].astype(str)
sagawa_masataka = sagawa_masataka[['顧客管理番号', 'お問合せNO', '運賃合計金額']]
sagawa_masataka = sagawa_masataka.rename(columns={'運賃合計金額':'佐川急便運賃' })


sagawa = pd.concat([sagawa_syouten, sagawa_goodlife, sagawa_masataka])


st.sidebar.write('■　西濃運輸')
uploaded_file = st.sidebar.file_uploader("西濃運輸 運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    seinou = pd.read_csv(uploaded_file, encoding="cp932")
seinou['原票No.'] = seinou['原票No.'].astype(str)
seinou = seinou[['原票No.', '合計']]
seinou = seinou.rename(columns={'合計':'西濃運輸運賃' })

st.sidebar.write('■　セイノーSSX')
uploaded_file = st.sidebar.file_uploader("セイノーSSX 運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    ssx = pd.read_csv(uploaded_file, encoding="cp932")
    
ssx['伝票番号'] = ssx['伝票番号'].astype(str)
ssx['伝票番号'] = ssx['伝票番号'].str.lstrip('0')
ssx['合計'] = ssx['合計'].str.lstrip('0')


ssx = ssx[['伝票番号', '合計']]
ssx = ssx.rename(columns={'合計':'ssx運賃' })

#伝票番号に伝票番号が含まれる行を削除
ssx = ssx[~(ssx['伝票番号']=='伝票番号')]

ssx['ssx運賃'] = ssx['ssx運賃'].astype(int)






st.sidebar.write('■　福山通運')
uploaded_file = st.sidebar.file_uploader("福山通運 運賃ファイルを読み込んでください", type="csv")
if uploaded_file is not None:
    fukuyama = pd.read_csv(uploaded_file, encoding="cp932")
    
fukuyama['原票番号'] = fukuyama['原票番号'].astype(str)

fukuyama = fukuyama[['原票番号', '運賃']]
fukuyama = fukuyama.rename(columns={'運賃':'福山通運運賃' })





if st.checkbox('読み込んだ各運送会社の送料表を表示する'):
    

    st.write('佐川急便　承天貿易')
    st.write(sagawa_syouten)
    st.write('佐川急便　GOODLIFE')
    st.write(sagawa_goodlife)
    st.write('佐川急便　昌隆')
    st.write(sagawa_masataka)
    st.write('西濃運輸')
    st.write(seinou)
    st.write('SSX')
    st.write(ssx)
    st.write('福山通運')
    st.write(fukuyama)



shipper_list = main_data_df['配送業者'].unique()


if st.checkbox('今月の使用運送会社を表示する'):
    st.table(shipper_list)


#読み込みここまで









#ここまででマージの準備完了


#商品原価集計のためNEWMEW_DBとマージ
main_data_df_single_item = pd.merge(main_data_df_single_item, df_newnew_db, left_on='item_name', right_on='変換商品名', how='left')

if st.checkbox('single_item,newnew_dbのマージ結果を表示'):
    st.write(main_data_df_single_item)


#卸金額が入っていないレコードを抽出
main_data_df_single_item_null = main_data_df_single_item[main_data_df_single_item['卸価格'].isnull()]
st.write('◆ single_item_卸金額不明分')
st.write(main_data_df_single_item_null)

main_data_df_single_item = main_data_df_single_item[['受注番号', '顧客分類', '発送日', '配送伝票番号', '配送業者', '卸価格', '変換商品名_y', '販売数量']]


#金額計算
main_data_df_single_item['商品代金'] = main_data_df_single_item['卸価格'] * main_data_df_single_item['販売数量']





#single運賃情報をマージ
main_data_df_single_item_ship = pd.merge(main_data_df_single_item_ship, sagawa, left_on='配送伝票番号', right_on='お問合せNO', how='left')



main_data_df_single_item_ship = pd.merge(main_data_df_single_item_ship, seinou, left_on='配送伝票番号', right_on='原票No.', how='left')

main_data_df_single_item_ship = pd.merge(main_data_df_single_item_ship, ssx, left_on='配送伝票番号', right_on='伝票番号', how='left')

main_data_df_single_item_ship = pd.merge(main_data_df_single_item_ship, fukuyama, left_on='配送伝票番号', right_on='原票番号', how='left')



#ゆうぱっく運賃の中央値624円を日本郵便に入力
main_data_df_single_item_ship.loc[main_data_df_single_item_ship['配送業者'] == '日本郵便', 'ゆうパック運賃'] = 624

#クリックポスト運賃の税抜き169円をクリックポストに入力
main_data_df_single_item_ship.loc[main_data_df_single_item_ship['配送業者'] == 'クリックポスト', 'クリックポスト運賃'] = 169



#定形外郵便のみ抽出
teikeigai = main_data_df_single_item.loc[main_data_df_single_item['配送業者'] == '定形外郵便']

teikeigai = pd.merge(teikeigai, df_newnew_db, left_on='変換商品名_y', right_on='変換商品名', how='left')
teikeigai = teikeigai[['受注番号', '顧客分類', '変換商品名', '定形外送料', ]]


main_data_df_single_item_ship = pd.merge(main_data_df_single_item_ship, teikeigai, on='受注番号', how='outer')

main_data_df_single_item_ship = main_data_df_single_item_ship \
[['受注番号', '顧客分類_x', '佐川急便運賃', '西濃運輸運賃', 'ssx運賃', '福山通運運賃', 'ゆうパック運賃','クリックポスト運賃','定形外送料']]


main_data_df_single_item_ship_total = main_data_df_single_item_ship.groupby('顧客分類_x').sum()

main_data_df_single_item_ship_total.rename(columns={'顧客分類_x':'顧客分類'}, inplace=True)





#商品原価集計のためNEWMEW_DBとマージ
main_data_df_set_item = pd.merge(main_data_df_set_item, df_newnew_db, left_on='変換商品名', right_on='変換商品名', how='left')

if st.checkbox('set_item,newnew_dbのマージ結果を表示'):
    st.write(main_data_df_set_item)


#卸金額が入っていないレコードを抽出
main_data_df_set_item_null = main_data_df_set_item[main_data_df_set_item['卸価格'].isnull()]
st.write('◆ set_item_卸金額不明分')
st.write(main_data_df_set_item_null)



main_data_df_set_item = main_data_df_set_item[['受注番号', '顧客分類', '発送日', '配送伝票番号', '配送業者', '卸価格', '変換商品名', '購入品数量']]


#金額計算
main_data_df_set_item['商品代金'] = main_data_df_set_item['卸価格'] * main_data_df_set_item['購入品数量']





#set運賃情報をマージ
main_data_df_set_item_ship = pd.merge(main_data_df_set_item_ship, sagawa, left_on='配送伝票番号', right_on='お問合せNO', how='left')

main_data_df_set_item_ship = pd.merge(main_data_df_set_item_ship, seinou, left_on='配送伝票番号', right_on='原票No.', how='left')

main_data_df_set_item_ship = pd.merge(main_data_df_set_item_ship, ssx, left_on='配送伝票番号', right_on='伝票番号', how='left')

main_data_df_set_item_ship = pd.merge(main_data_df_set_item_ship, fukuyama, left_on='配送伝票番号', right_on='原票番号', how='left')



#ゆうぱっく運賃の中央値624円を日本郵便に入力
main_data_df_set_item_ship.loc[main_data_df_set_item_ship['配送業者'] == '日本郵便', 'ゆうパック運賃'] = 624

#クリックポスト運賃の税抜き169円をクリックポストに入力
main_data_df_set_item_ship.loc[main_data_df_set_item_ship['配送業者'] == 'クリックポスト', 'クリックポスト運賃'] = 169



#定形外郵便のみ抽出
teikeigai = main_data_df_set_item.loc[main_data_df_set_item['配送業者'] == '定形外郵便']

teikeigai = pd.merge(teikeigai, df_newnew_db, left_on='変換商品名', right_on='変換商品名', how='left')
teikeigai = teikeigai[['受注番号', '顧客分類', '変換商品名', '定形外送料', ]]


main_data_df_set_item_ship = pd.merge(main_data_df_set_item_ship, teikeigai, on='受注番号', how='outer')

main_data_df_set_item_ship = main_data_df_set_item_ship[['受注番号', '顧客分類_x', '佐川急便運賃', '西濃運輸運賃', 'ssx運賃',\
                                                          '福山通運運賃', 'ゆうパック運賃','クリックポスト運賃','定形外送料']]




main_data_df_set_item_ship_total = main_data_df_set_item_ship.groupby('顧客分類_x').sum()

main_data_df_set_item_ship_total.rename(columns={'顧客分類_x':'顧客分類'}, inplace=True)







main_data_df_set_item = main_data_df_set_item.groupby('顧客分類').sum()



main_data_df_set_item = main_data_df_set_item[['購入品数量', '商品代金']]

main_data_df_set_item.rename(columns={'購入品数量':'販売数量'}, inplace=True)





#商品代金の総合集計
item_price_total = pd.concat([main_data_df_single_item, main_data_df_set_item])

item_price_total =  item_price_total.groupby('顧客分類').sum()

item_price_total = item_price_total[['販売数量', '商品代金']]




#運賃の総合集計
item_price_total_ship = pd.concat([main_data_df_single_item_ship_total, main_data_df_set_item_ship_total])

item_price_total_ship =  item_price_total_ship.groupby('顧客分類_x').sum()



item_price_total_ship['運賃'] = item_price_total_ship['佐川急便運賃'] + item_price_total_ship['西濃運輸運賃'] + item_price_total_ship['ssx運賃']\
      + item_price_total_ship['福山通運運賃'] + item_price_total_ship['ゆうパック運賃'] + item_price_total_ship['クリックポスト運賃'] + item_price_total_ship['定形外送料']

item_price_total_ship.rename(columns={'顧客分類_x':'顧客分類'}, inplace=True)

item_price_total_ship = item_price_total_ship.reset_index(drop=False)
item_price_total = item_price_total.reset_index(drop=False)




result = pd.merge(item_price_total_ship, item_price_total, left_on='顧客分類_x', right_on='顧客分類', )


result = result[['顧客分類_x', '販売数量', '商品代金', '運賃' ]]

result.rename(columns={'顧客分類_x':'顧客分類'}, inplace=True)



st.write('<h4>出力結果</h4>※一部運賃を含まないので注意すること<br>', unsafe_allow_html=True)

st.dataframe(result)
csv = result.to_csv().encode('cp932')#データフレームをCSVにして、
st.download_button(label='出力結果ダウンロード', data=csv, file_name='result.csv', mime='text/csv')#そのCSVをダウンロード
st.write('※佐川急便(チャーター便)、日通トランスポート、九州航空は含まないので別途計算します。')




st.write('<h4>自動計算されない配送方法ファイルダウンロード</h4>', unsafe_allow_html=True)
st.write('佐川急便(チャーター)')
st.write(toll)
csv = toll.to_csv().encode('cp932')#データフレームをCSVにして、
st.download_button(label='↑佐川急便(チャーター便)データダウンロード', data=csv, file_name='toll.csv', mime='text/csv')#そのCSVをダウンロード



st.write(kyushu_k)
csv = kyushu_k.to_csv().encode('cp932')#データフレームをCSVにして、
st.download_button(label='↑九州航空データダウンロード', data=csv, file_name='kyushu_k.csv', mime='text/csv')#そのCSVをダウンロード



st.write(nituu_tp)
csv = nituu_tp.to_csv().encode('cp932')#データフレームをCSVにして、
st.download_button(label='↑日通トランスポートデータダウンロード', data=csv, file_name='nituu_tp.csv', mime='text/csv')#そのCSVをダウンロード


