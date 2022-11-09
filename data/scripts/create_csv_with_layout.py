import csv

with open('contract.csv', mode='w', encoding='utf-8-sig') as csvfile:
    writer = csv.writer(csvfile, delimiter=',')

    row = ['','', '売買契約 兼 請求書']
    writer.writerow(row)

    row = ['No.161761376809', '', '', '', '', '', '契約日', '2021/04/05']
    writer.writerow(row)
    row = ['', '', '', '', '', '', '更新日', '2021/04/05']
    writer.writerow(row)

    row = ['会社名', '北日本・遊機株式会社ＡＣディビジョン', '', 'フリガナ', 'キタニホン・ユウキＡＣディビジョン']
    writer.writerow(row)
    row = ['郵便番号', '007-0827']
    writer.writerow(row)
    row = ['住所', '北海道札幌市東区東雁来七条２丁目８－３', '', '', '', '', 'P-SENSOR 会員番号', '8240-2413-3628']
    writer.writerow(row)
    row = ['TEL', '011-299-8900', '', 'FAX', '011-299-8900', '', '担当名', '田中']
    writer.writerow(row)
    
    row = []
    writer.writerow(row)

    row = ['商品名']
    writer.writerow(row)
    row = ['機種名', '中分類', '数量', '単価', '金額']
    writer.writerow(row)
    row = ['Ｓアメリカン番長　ＨＥＹ！鏡Ｂ２', '本体', '2', '121000', '121000']
    writer.writerow(row)
    row = ['Ｓアメリカン番長　ＨＥＹ！鏡Ｂ２', '本体', '2', '121000', '121000']
    writer.writerow(row)

    row = []
    writer.writerow(row)

    row = ['商品名そのほか']
    writer.writerow(row)
    row = ['書類', '数量', '単価', '金額']
    writer.writerow(row)
    row = ['出庫手数料', '2', '12000', '24000']
    writer.writerow(row)
    row = ['出庫手数料', '2', '12000', '24000']
    writer.writerow(row)

    row = []
    writer.writerow(row)

    row = ['機械発送日', '2021/04/05', '', '小計', '820000']
    writer.writerow(row)
    row = ['備考', '', '', '消費税 (10%)', '1200']
    writer.writerow(row)
    row = ['', '', '', '保険代 (非課税)', '1200']
    writer.writerow(row)
    row = ['', '', '', '合計', '13200']
    writer.writerow(row)

    row = ['運送方法', '発送', '', '御請求金額', '13200']
    writer.writerow(row)
    row = ['お支払方法', '振込', '', '支払期限', '2021/04/05']
    writer.writerow(row)

    row = []
    writer.writerow(row)

    row = ['商品発送先', '', '', '書類発送先']
    writer.writerow(row)
    row = ['会社名', 'さくら企画', '', '会社名', 'アイエス販売']
    writer.writerow(row)
    row = ['住所', '大阪市中央区上本町西１－４－２３', '', '住所', '埼玉県鴻巣市大間３－６－８']
    writer.writerow(row)
    row = ['TEL', '090-2359-4826', '', 'TEL', '048-580-3000']
    writer.writerow(row)
    row = ['FAX', '06-6561-6556', '', 'FAX', '048-580-3001']
    writer.writerow(row)
    row = ['到着予定日', '4/5/2021', '', '到着予定日', '4/5/2021']
    writer.writerow(row)
