import xlwings as xw
import urllib.parse
import base64
import logging.config
import requests
import json

# ログ設定ファイルからログ設定を読み込み
logging.config.fileConfig('C:\Work\logging.conf')

logger = logging.getLogger()

logger.info('---- Rakuten searchOrder ----')

def searchOrderMain():
    sh = xw.sheets.active
    # シートからアクセストークン取得
    SECRET = xw.Range("Secret").value
    LICKEY = xw.Range("licenseKey").value
    AuthStr = base64.b64encode((SECRET + ":" + LICKEY).encode())

    #取得処理
    REQ_URL = "https://api.rms.rakuten.co.jp/es/2.0/order/searchOrder/"
    AuthData = "ESA " + AuthStr.decode()
    headers = {'Content-Type':'application/json; charset=utf-8','Authorization':AuthData}
    post_data = {}  #辞書として定義

    #orderProgressList
    orderProgressListArr = []
    for i in range(9):
        if xw.Range("orderProgressList_" + str(i+1)).value == "○":
            orderProgressListArr.append(str((i+1)*100))
    orderProgressList = '[' +','.join(orderProgressListArr) + ']'

#    post_data["orderProgressList"] = orderProgressList  #辞書要素追加

    #dateType
    for i in range(6):
        if xw.Range("dateType_" + str(i+1)).value == "○":
            dateTypeNum = i + 1

    post_data["dateType"] = dateTypeNum  #辞書要素追加

    #startDatetime
    startDate8= str(int(xw.Range("startDatetime").value))
    startDatetimeStr = str(startDate8[0:4]) + "-" + str(startDate8[4:6]) + "-" + str(startDate8[6:8]) + "T00:00:00+0900"

    post_data["startDatetime"] = startDatetimeStr  #辞書要素追加

    #endDatetime
    endDate8= str(int(xw.Range("endDatetime").value))
    endDatetimeStr = str(endDate8[0:4]) + "-" + str(endDate8[4:6]) + "-" + str(endDate8[6:8]) + "T00:00:00+0900"

    post_data["endDatetime"] = endDatetimeStr  #辞書要素追加

    #settlementMethod
    if xw.Range("settlementMethod_1").value == "○":
        post_data["settlementMethod"] = 2  #辞書要素追加
    elif xw.Range("settlementMethod_2").value == "○":
        post_data["settlementMethod"] = 9  #辞書要素追加

    #searchKeywordType
    keywordflg = 0
    for i in range(6):
        if xw.Range("searchKeywordType_" + str(i+1)).value == "○":
            post_data["searchKeywordType"] = i  # 辞書要素追加
            keywordflg = 1

    #searchKeyword
    if keywordflg:
        post_data["searchKeyword"] = xw.Range("searchKeyword").value  # 辞書要素追加

    #asurakuFlag
    if xw.Range("asurakuFlag_1").value == "○":
        post_data["asurakuFlag"] = 1  #辞書要素追加

    # Level 2: PaginationRequestModel
    PaginationRequestModelDic = {} #辞書として定義
    PaginationRequestModelDic["requestRecordsAmount"] = 1000
    PaginationRequestModelDic["requestPage"] = 1
    post_data["PaginationRequestModel"] = PaginationRequestModelDic  # 辞書要素追加

    #Level 3: SortModel
    SortModelDic = {} #辞書として定義
    SortModelDic["sortColumn"] = 1  #注文日時順
    SortModelDic["sortDirection"] = 1   #昇順
    post_data["SortModelList"] = SortModelDic  # 辞書要素追加

    proxies = {
        'http': 'http://127.0.0.1:8888',
        'https': 'http://127.0.0.1:8888',
    }

    #res = requests.post(REQ_URL, json=post_data, headers=headers,proxies=proxies,verify=False)
    res = requests.post(REQ_URL, json=post_data, headers=headers,verify=False)

    rescorde = str(res.status_code)
    logger.info('HTTP Status:' + rescorde)
    logger.info(res)
    resdata = res.json()
    if rescorde == "200" and resdata['MessageModelList'][0]['messageCode'] == "ORDER_EXT_API_SEARCH_ORDER_INFO_101":
        #受注番号出力
        orderNumberList = resdata['orderNumberList']
        orderNumberListCol = convert_1d_to_2d(orderNumberList,1) #2次元配列
        # xw.Range('A1').value = "受注番号"
        # xw.Range('A2').value = orderNumberListCol

        #受注番号ごとに詳細取得
        ShowOrderDetails(orderNumberList,AuthData)

    if rescorde == "200" and resdata['MessageModelList'][0]['messageCode'] == "ORDER_EXT_API_SEARCH_ORDER_INFO_102":
        xw.Range('A1').value = "検索結果は０件です。検索条件を確かめてください"
    else:
        logger.error('API was failed.Rescode:' + rescorde)
        logger.error(resdata['MessageModelList'][0]['messageCode'])
        logger.error(resdata['MessageModelList'][0]['message'])
        xw.Range('C75').value = "エラーが発生しました。ログを確認してください。"
    print("test")

#1次元配列を2次元にする関数（拾い物）
def convert_1d_to_2d(l, cols):
    return [l[i:i + cols] for i in range(0, len(l), cols)]

#引数の受注番号リストそれぞれについて詳細を取得し表示
def ShowOrderDetails(OrderNumList,AuthData):
    sh = xw.sheets.active
    # 出力ヘッダー配列
    OutHeaders = [\
        #Level 2: OrderModel
        "受注番号","ステータス","注文日時", "注文確認日時", \
        "発送完了報告日時", "お届け日指定", "お届け時間帯", \
        "コメント", "ギフト希望","複数送付先", "離島", \
        "利用端末", "楽天確認中", "商品金額", "送料", "代引料", \
        "決済手数料", "合計金額", "請求金額", "店舗クーポン", "モールクーポン", \
        "あす楽","ひとことメモ", \
        #Level 3: OrdererModel
        "発Z1", "発Z2", "発県", "発市", "発住所", "発姓", "発名", "発TEL1", "発TEL2", "発TEL3", "発mail", \
        # Level 3: SettlementModel
        "支払い方法", \
        #Level 3: PointModel
        "ポイント利用額", \
        #Level 3: WrappingModel
        "包装名", "包装料金", \
        #Level 3: PackageModel
        "個別送料", "個別代引料", "個別商品金額", "個別合計金額", "個別のし", \
        #Level 4: SenderModel
        "送Z1","送Z2", "送県", "送市", "送住所", "送姓", "送名",  "送TEL1", "送TEL2", "送TEL3", "送mail", "送離島", \
        #Level 4: ItemModel
        "商品明細ID", "商品名", "商品ID", "単価", "数量", "商品毎価格", "単品配送", \
        #Level 4: ShippingModel
        "配送伝票No", "配送会社", "発送日"
            ]
    # LabelsOrderModel = {\
    #              "orderNumber":"受注番号","orderProgress":"ステータス","orderDatetime":"注文日時","shopOrderCfmDatetime":"注文確認日時",\
    #              "shippingCmplRptDatetime":"発送完了報告日時","deliveryDate":"お届け日指定","shippingTerm":"お届け時間帯", \
    #              "remarks":"コメント","giftCheckFlag":"ギフト希望","severalSenderFlag":"複数送付先","isolatedIslandFlag":"離島",\
    #              "carrierCode":"利用端末","rakutenConfirmFlag":"楽天確認中","goodsPrice":"商品金額","postagePrice":"送料","deliveryPrice":"代引料",\
    #              "paymentCharge":"決済手数料","totalPrice":"合計金額","requestPrice":"請求金額","couponShopPrice":"店舗クーポン","couponOtherPrice":"モールクーポン",\
    #              "asurakuFlag":"あす楽","memo":"ひとことメモ"\
    #     }
    LabelsOrderModel = {\
        "受注番号":"orderNumber","ステータス":"orderProgress","注文日時":"orderDatetime","注文確認日時":"shopOrderCfmDatetime",\
        "発送完了報告日時":"shippingCmplRptDatetime","お届け日指定":"deliveryDate","お届け時間帯":"shippingTerm", \
        "コメント":"remarks","ギフト希望":"giftCheckFlag","複数送付先":"severalSenderFlag","離島":"isolatedIslandFlag",\
        "利用端末":"carrierCode","楽天確認中":"rakutenConfirmFlag","商品金額":"goodsPrice","送料":"postagePrice","代引料":"deliveryPrice",\
        "決済手数料":"paymentCharge","合計金額":"totalPrice","請求金額":"requestPrice","店舗クーポン":"couponShopPrice","モールクーポン":"couponOtherPrice",\
        "あす楽":"asurakuFlag","ひとことメモ":"memo", \
        "発Z1":"zipCode1", "発Z2": "zipCode2", "発県": "prefecture", "発市": "city", "発住所": "subAddress", \
        "発姓": "familyName", "発名": "firstName", \
        "発TEL1": "phoneNumber1", "発TEL2": "phoneNumber2", "発TEL3": "phoneNumber3", "発mail": "emailAddress", \
        "個別送料": "postagePrice", "個別代引料": "deliveryPrice", "個別商品金額": "goodsPrice", "個別合計金額": "totalPrice", \
        "個別のし": "noshi" , \
        "送Z1": "zipCode1", "送Z2": "zipCode2", "送県": "prefecture", "送市": "city", "送住所": "subAddress", \
        "送姓": "familyName", "送名": "firstName", \
        "送TEL1": "phoneNumber1", "送TEL2": "phoneNumber2", "送TEL3": "phoneNumber3", "送mail": "emailAddress", \
        "送離島": "isolatedIslandFlag", \
        "商品明細ID": "itemDetailId", "商品名": "itemName", "商品ID": "itemId", "単価": "price", "数量": "units", \
        "商品毎価格": "priceTaxIncl", "単品配送": "isSingleItemShipping" \
        }
    LabelsOrdererModel ={\
                            "発Z1":"zipCode1", "発Z2":"zipCode2", "発県":"prefecture", "発市":"city", "発住所":"subAddress",\
                            "発姓":"familyName", "発名":"firstName", \
                            "発TEL1":"phoneNumber1", "発TEL2":"phoneNumber2", "発TEL3":"phoneNumber3", "発mail":"emailAddress"\
        }
    # LabelsOrdererModel ={\
    #                         "zipCode1":"発Z1", "zipCode2":"発Z2", "prefecture":"発県", "city":"発市", "subAddress":"発住所",\
    #                         "familyName":"発姓", "firstName":"発名", \
    #                         "phoneNumber1":"発TEL1", "phoneNumber2":"発TEL2", "phoneNumber3":"発TEL3", "emailAddress":"発mail"\
    #     }
    LabelsSettlementModel={"支払い方法":"settlementMethod"}
    LabelsPointModel ={"ポイント利用額":"usedPoint"}
    LabelsWrappingModel ={"包装名":"name","包装料金":"price"}
    # LabelsSettlementModel={"settlementMethod":"支払い方法"}
    # LabelsPointModel ={"usedPoint":"ポイント利用額"}
    # LabelsWrappingModel ={"name":"包装名","price":"包装料金"}
    # LabelsPackageModel = {\
    #     "postagePrice":"個別送料","deliveryPrice":"個別代引料","goodsPrice":"個別商品金額","totalPrice":"個別合計金額",\
    #     "noshi":"個別のし"\
    #     }
    LabelsPackageModel = {\
        "個別送料":"postagePrice","個別代引料":"deliveryPrice","個別商品金額":"goodsPrice","個別合計金額":"totalPrice",\
        "個別のし":"noshi"\
        }
    LabelsSenderModel = { \
        "送Z1":"zipCode1", "送Z2":"zipCode2", "送県":"prefecture", "送市":"city", "送住所":"subAddress", \
        "送姓":"familyName", "送名":"firstName", \
        "送TEL1":"phoneNumber1", "送TEL2":"phoneNumber2", "送TEL3":"phoneNumber3", "送mail":"emailAddress", \
        "送離島":"isolatedIslandFlag"\
        }
    LabelsItemModel ={\
        "商品明細ID":"itemDetailId","商品名":"itemName","商品ID":"itemId","単価":"price","数量":"units",\
        "商品毎価格":"priceTaxIncl","単品配送":"isSingleItemShipping" \
        }

    # LabelsSenderModel = { \
    #     "zipCode1":"送Z1", "zipCode2": "送Z2", "prefecture": "送県", "city": "送市", "subAddress": "送住所", \
    #     "familyName": "送姓", "firstName": "送名", \
    #     "phoneNumber1": "送TEL1", "phoneNumber2": "送TEL2", "phoneNumber3": "送TEL3", "emailAddress": "送mail", \
    #     "isolatedIslandFlag":"送離島"\
    #     }
    # LabelsItemModel ={\
    #     "itemDetailId":"商品明細ID","itemName":"商品名","itemId":"商品ID","price":"単価","units":"数量",\
    #     "priceTaxIncl":"商品毎価格","isSingleItemShipping":"単品配送"\
    #     }
    LabelsShippingModel ={"shippingNumber":"配送伝票No","deliveryCompanyName":"配送会社","shippingDate":"発送日"}

    # 取得データパラメータの配列
    L1_Base = ["MessageModelList","OrderModelList"]
    L2_MessageModel=["messageType","messageCode","message","orderNumber"]
    L2_OrderModel=[\
        "orderNumber","orderProgress","orderDatetime","shopOrderCfmDatetime","orderFixDatetime",\
        "shippingInstDatetime","shippingCmplRptDatetime","cancelDueDate","deliveryDate","shippingTerm", \
        "remarks","giftCheckFlag","severalSenderFlag","equalSenderFlag","isolatedIslandFlag","rakutenMemberFlag", \
        "carrierCode","orderType","cautionDisplayType","rakutenConfirmFlag","goodsPrice","postagePrice","deliveryPrice",\
        "paymentCharge","paymentChargeTaxRate","totalPrice","requestPrice","couponAllTotalPrice","couponShopPrice",\
        "couponOtherPrice","additionalFeeOccurAmountToUser","additionalFeeOccurAmountToShop","asurakuFlag","memo",\
        "OrdererModel","SettlementModel","DeliveryModel","PointModel","WrappingModel1","WrappingModel2",\
        "PackageModelList","CouponModelList","ChangeReasonModelList","TaxSummaryModelList"\
        ]
    L3_OrdererModel=[\
        "zipCode1","zipCode2","prefecture","city","subAddress","familyName","firstName","familyNameKana","firstNameKana",\
        "phoneNumber1","phoneNumber2","phoneNumber3","emailAddress",\
        ]
    L3_SettlementModel =["settlementMethodCode","settlementMethod","rpaySettlementFlag"]
    L3_DeliveryModel =["deliveryName","deliveryClass"]
    L3_PointModel =["usedPoint"]
    L3_WrappingModel =["title","name","price","includeTaxflag"]
    L3_PackageModel = [\
        "basketId","postagePrice","postageTaxRate","deliveryPrice","deliveryTaxRate","goodsTax","goodsPrice","totalPrice",\
        "noshi","packageDeleteFlag","SenderModel","ItemModelList","ShippingModelList","DeliveryCvsModel","defaultDeliveryCompanyCode"\
        ]
    L4_SenderModel = [ \
        "zipCode1","zipCode2","prefecture","city","subAddress","familyName","firstName","familyNameKana",\
        "firstNameKana","phoneNumber1","phoneNumber2","phoneNumber3","isolatedIslandFlag"\
        ]
    L4_ItemModel =[ \
        "itemDetailId","itemName","itemId","itemNumber","manageNumber","price","units",\
        "includePostageFlag","includeTaxFlag","includeCashOnDeliveryPostageFlag","selectedChoice",\
        "pointRate","pointType","inventoryType","delvdateInfo","restoreInventoryFlag","dealFlag","drugFlag",\
        "deleteItemFlag","taxRate","priceTaxIncl","isSingleItemShipping"\
        ]
    L4_ShippingModel =["shippingDetailId","shippingNumber","deliveryCompany","deliveryCompanyName","shippingDate"]
    L3_CouponModel =[ \
        "couponCode","itemId","couponName","couponSummary","couponCapital","couponCapitalCode",\
        "expiryDate","couponPrice","couponUnit","couponTotalPrice"\
        ]
    L3_TaxSummaryModel = [ \
        "taxRate","reqPrice","reqPriceTax","totalPrice","paymentCharge","couponPrice","point"\
        ]

    #取得処理
    REQ_URL = "https://api.rms.rakuten.co.jp/es/2.0/order/getOrder/"
    headers = {'Content-Type':'application/json; charset=utf-8','Authorization':AuthData}
    post_data = {}  #辞書として定義

    #受注番号リストを１００ずつ区切って二次元配列化
    SlicedList = [OrderNumList[i:i + 100] for i in range(0, len(OrderNumList), 100)]

    for i in range(len(SlicedList)):
        post_data["orderNumberList"] = SlicedList[i]
        post_data["version"] = 1    #固定値

        res = requests.post(REQ_URL, json=post_data, headers=headers,verify=False)

        rescorde = str(res.status_code)
        logger.info('getOreder:' + str(i+1) +'回目:' + rescorde)
        logger.info('HTTP Status:' + rescorde)
        logger.info(res)
        resdata = res.json()
        if rescorde == "200" and resdata['MessageModelList'][0]['messageCode'] == "ORDER_EXT_API_GET_ORDER_INFO_101":
            #Excelに出力
            #ヘッダ出力
            xw.Range('A1').value = OutHeaders

            outModel = ""
            #Level 2:OrderModelListに対し要求した注文番号の数分だけ処理する
            for i in range(len(OrderNumList)):
                j = 2   #出力行No.
                # 出力したい要素を取り出す
                for label in OutHeaders:
                    # 対象要素が含まれるLevel3以降のModelを記録しておく
                    if label == "発Z1":
                        outModel = "OrdererModel"
                    elif label == "支払い方法":
                        outModel = "SettlementModel"
                    elif label == "ポイント利用額":
                        outModel = "PointModel"
                    elif label == "包装名":
                        outModel = "WrappingModel1"
                    elif label == "個別送料":
                        outModel = "PackageModelList"
                    # 以降はLevel4
                    elif label == "送Z1":
                        outModel = "SenderModel"
                    elif label == "商品明細ID":
                        outModel = "ItemModelList"
                    elif label == "配送伝票No":
                        outModel = "ShippingModelList"

                    if outModel == "":
                        #Level2
                        xw.Range((j,OutHeaders.index(label)+1)).value = resdata['OrderModelList'][i+1][LabelsOrderModel[label]]
                    elif outModel == "OrdererModel" or outModel == "SettlementModel" or outModel == "PointModel" or outModel == "WrappingModel1":
                        #Level3 1行
                        xw.Range((j,OutHeaders.index(label)+1)).value = resdata['OrderModelList'][i+1][outModel][LabelsOrderModel[label]]
                    elif outModel == "PackageModelList":
                        #PakageModel 複数行あり
                        PackageListLen = len(resdata['OrderModelList'][i+1][outModel]["PackageModelList"])
                        for k in range(PackageListLen):
                            xw.Range((j,OutHeaders.index(label)+1)).value = \
                                resdata['OrderModelList'][i+1][outModel]["PackageModelList"][k+1][LabelsOrderModel[label]]




            # orderNumberList = resdata['orderNumberList']
            # orderNumberListCol = convert_1d_to_2d(orderNumberList,1) #2次元配列
            # xw.Range('A1').value = "受注番号"
            # xw.Range('A2').value = orderNumberListCol

    return True


if __name__ == '__main__':
    searchOrderMain()
