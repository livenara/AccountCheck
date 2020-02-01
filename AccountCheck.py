"""

銀行サイトからCSVでエクスポートした仮想口座と通常口座の入金データが混ざったものを
それぞれ事業所の経理で基幹システムに取込処理できるよう
事業所別に基幹取り込み形式のテキストを出力する。

"""

# python3.6
import pandas as pd
import xlwt, openpyxl
import os, glob, shutil, csv
import datetime

# 処理時注意リスト 転記用エクセルに追加記載
def KaisyuRan(KouzaBango):

    CoutionList = [
        ['39*****','0','得意先番号注意','**商店'],
        ['39*****','0','リベート注意','**フーズ'],
        ['39*****','0','別に本部あり','****'],
        ['39*****','0','**食品と一緒1837368','***ｷｶｸ'],
        ['39*****','0','****と混同注意','***の駅'],
        ['39*****','0','****と混同注意','****'],
        ['39*****','0','別に本部あり','***'],
        ['39*****','0','880+販促100','*****'],
        ['18*****','0','別本部有466001と467001','****フーズ']
        ]

    for Coution in CoutionList:
        if KouzaBango == Coution[0]:
            Comment = Coution[2]
            break

    return Comment

def AccountCheck():

    # ファイル用日付
    FileDate = datetime.datetime.today().strftime("%Y%m%d")

    KouzaList         = [] # 口座番号管理リスト.csv
    ImportDataList    = [] # 取込データList

    JigyoSyoList = ["大阪事業本部","東京大田事業所","東京世田谷事業所"] # 口座番号内-事業所データ
    DelList = ['18**100','18**101','18**200'] # 仮想口座番号でない口座番号

    # 入金データ取込用フォルダ名
    FileName = glob.glob('取込用\\*')[0].replace('取込用\\', '')
    print(FileName)
    ImportFileDir = "G:\\マイドライブ\\PyAccountCheck\\取込用\\"
    ImportFile = ImportFileDir + FileName

    # 仮想口座管理CSVのデータで突き合わせているので最新のものを読み込む
    DataDir = "G:\\マイドライブ\\999_*****データ\\"
    KouzaFileName = DataDir + "仮想口座.csv" # グーグルドライブに仮想口座割当からエクスポートしたCSV


    # 管理エクセル転記用エクセル　出力ファイル名
    ExportDir = "G:\\マイドライブ\\PyAccountCheck\\出力\\"
    ExportFileName = str(FileDate) + "_大阪事業所用エクセルデータ.xls"
    ExportFile = ExportDir + ExportFileName


    # 最初のデータ読み込み (口座番号管理.csv)
    KouzaList = [] # 口座番号管理リスト.csv

    def FirstDataImport(KouzaFileName):

        with open(KouzaFileName,"r",encoding="utf-8") as f:
            data = csv.reader(f)
            for x in data:
                KouzaList.append(x)

        return KouzaList


    # 取り込みCSVファイル　＞差分チェック機能を入れたい
    # 手作業で入れてもらっているのでデータ内容が安定していない。
    ImportDataList = [] # 取込データList
    def ImportCsvFile(ImportFile):
        with open(ImportFile, "r", encoding="sjis")as f:
            data = csv.reader(f)
            for x in data:
                ImportDataList.append(x)
        return ImportDataList


    # 口座番号管理を回して取込データの口座番号が合えば(得意先コード,口座番号,事業所)を吐き出す。
    # FOR文で遅くなってきたらpandasに切り替えていく
    def KouzaCheck(KouzaBango):

        Dic = {"TokuiCode" : "","KouzaBango": "","Jigyosyo": ""} # 辞書を先に設定しておく

        for x in KouzaList:

            if KouzaBango == x[11]: # 個店番号が空白でなければ処理

                zero = str(x[1]).strip().zfill(8) # 得意先番号8桁ゼロ埋め
                Dic = {"TokuiCode" : zero, "KouzaBango":x[10], "Jigyosyo":x[9]} # 得意先コード、口座番号、事業所番号

            elif KouzaBango == x[10]: # 個店番号が空白の時処理

                zero = str(x[1]).strip().zfill(8) # 得意先番号8桁ゼロ埋め
                Dic = {"TokuiCode" : zero, "KouzaBango":x[10], "Jigyosyo":x[9]} # 得意先コード、口座番号、事業所番号

        return Dic


    # コピー用エクセルデータ作成
    NyukinCsvCopyData = []
    CopyDataList = []
    def ExcelCopyData(ImportDataList, JigyoSyoList, DelList):

        for i in ImportDataList:

            #日付
            #Date = datetime.datetime.today().strftime("%Y/%m/%d")
            Date = str(datetime.datetime.today().strftime("%Y")) + "/" + i[7][-4:-2] + "/" + i[7][-2:]

            #口座番号
            KouzaBango = i[14][-7:] #下7桁が口座番号 "804*******" -> "*******"

            #口座番号から{得意先コード,口座番号,事業所}を吐き出す。
            Dic = KouzaCheck(KouzaBango)

            #入金額
            if type(i[10]) is int:
                Nyukingaku = int(i[10])
            else:
                Nyukingaku = i[10]

            #手数料
            if type(i[11]) is int:
                Tesuryou = int(i[11])
            else:
                Tesuryou = i[11]

            #回収額
            Kaisyu = KaisyuRan(KouzaBango)

            #得意先コード
            TokuiCode = Dic["TokuiCode"]

            #振込人カナ
            KouzaKana = i[15].strip()

            #口座仕向銀行
            KouzaShimukeGinko = i[16].strip()

            #口座仕向支店
            KouzaShimukeMise = i[17].strip()

            #事業所
            Jigyousyo = Dic["Jigyosyo"]

            #まとめリスト化
            CopyDataList = [Date,FileName,Nyukingaku,Tesuryou,Kaisyu,TokuiCode,KouzaKana,KouzaShimukeGinko,KouzaShimukeMise,KouzaBango,Jigyousyo]
            
            #まとめのまとめ
            NyukinCsvCopyData.append(CopyDataList)

            #まとめリスト初期化
            CopyDataList = []

        # タイトルの一行目だけ消す
        NyukinCsvCopyData.pop(0)

        # 仮想口座番号で無い ['18**100','18**101','18**200'] の口座番号が入った配列を除いて行く
        OsakaCopyExcelList = []
        for x in NyukinCsvCopyData:
            if str(x[9]) == "0" or str(x[9]) == DelList[0] or str(x[9]) == DelList[1] or str(x[9]) == DelList[2]:
                pass
            elif x[10] == JigyoSyoList[0]: # JigyoSyoList[0] 大阪事業本部
                OsakaCopyExcelList.append(x)


        # pandasでエクセルを作成する。

        df = pd.DataFrame(OsakaCopyExcelList)
        df.to_excel(ExportFile, index=False, header=False) # インデックスとヘッダーを消す



    #入金チェックファイル作成
    KouzaList = FirstDataImport(KouzaFileName) # 最初のデータ読み込み (口座番号管理.csv)
    ImportDataList = ImportCsvFile(ImportFile) # 取り込みCSVファイル
    ExcelCopyData(ImportDataList, JigyoSyoList, DelList) # 入金用チェックファイルコピーデータ作成

    #▲入金チェックファイル作成まで


    #--------------------------------------------------------------------------------------------------

    #取り込み用データを作る　タブ形式テキストファイル

    OsakaList = []
    OotaList = []
    SetagayaList = []

    i = 0

    MainList = []
    for x in ImportDataList:

        KouzaBango = x[14]

        if KouzaBango == DelList[0]:
            pass
        elif KouzaBango == DelList[1]:
            pass
        elif KouzaBango == DelList[2]:
            pass
        else:
            MainList.append(x)

    #print(MainList)

    for x in MainList:

        KouzaBango = x[14][-7:]


        #数値をint値に直すと取り込みが出来る。
        try:
            x[0] = int(x[0])
            x[5] = int(x[5])
            x[7] = int(x[7])
            x[10] = int(x[10])
            x[11] = int(x[11])
            x[13] = int(x[13])
        except Exception as e:
            pass

        for y in KouzaList:

            if KouzaBango == y[11]:

                if y[9] == "大阪事業本部":
                    if not str(KouzaBango[-7:-5]) == "18":
                        OsakaList.append(x)
                    continue
                elif y[9] == "東京大田事業所":
                    OotaList.append(x)
                    continue
                elif y[9] == "東京世田谷事業所":
                    SetagayaList.append(x)
                    continue

            elif KouzaBango == y[10]:

                if y[9] == "大阪事業本部":
                    # 大阪事業本部分は18～の口座番号を除く
                    if not str(KouzaBango[-7:-5]) == "18":
                        OsakaList.append(x)
                    continue
                elif y[9] == "東京大田事業所":
                    OotaList.append(x)
                    continue
                elif y[9] == "東京世田谷事業所":
                    SetagayaList.append(x)
                    continue

    #print(list(map(list, set(map(tuple, OsakaList)))))
    #print(list(map(list, set(map(tuple, SetagayaList)))))
    #print(list(map(list, set(map(tuple, OotaList)))))

    OsakaList = list(map(list, set(map(tuple, OsakaList)))) #リストの中のリスト重複を削除する。

    if len(OsakaList):

        df = pd.DataFrame(OsakaList)
        OsakaCsvFileName = "出力\\" + str(FileDate) + "_Osaka.txt"
        df.to_csv(OsakaCsvFileName , sep="\t", encoding="shift_jis", index=False, header=False) #大阪用ファイル出力



    SetagayaList = list(map(list, set(map(tuple, SetagayaList)))) #リストの中のリスト重複を削除する。

    if len(SetagayaList):

        df = pd.DataFrame(SetagayaList)
        SetagayaCsvFileName = "出力\\" + str(FileDate) + "_Setagaya.txt"
        df.to_csv(SetagayaCsvFileName , sep="\t", encoding="shift_jis", index=False, header=False) #世田谷用ファイル出力

    OotaList = list(map(list, set(map(tuple, OotaList)))) #リストの中のリスト重複を削除する。

    if len(OotaList):

        df = pd.DataFrame(OotaList)
        OotaCsvFileName = "出力\\" + str(FileDate) + "_Oota.txt"
        df.to_csv(OotaCsvFileName , sep="\t", encoding="shift_jis", index=False, header=False) #大田用ファイル出力

    #print(len(OotaList))
    #print(len(SetagayaList))
    #print(len(OsakaList))

#--------------------------------------------------------------------------------------------------

AccountCheck() #入金データ入CSVを各事業所別に基幹システム取込形式へ

