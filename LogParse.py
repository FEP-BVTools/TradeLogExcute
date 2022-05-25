# -*- coding: utf-8 -*-
"""
Created on Mon Mar 21 08:28:14 2022

@author: TomTsai


變更目的:
1.移除EZC加值補送Log&Log3用的檔案
2.批量解譯
3.將ICash解譯Log檔名結合測項
4.能自動將Log解譯結果自動歸回各自資料夾

參數:
測試大項
測試細項
待回歸檔案名

回歸流程:
1.確認檔案有含'pending'
2.取得檔案回歸位置
3.移動完成後,將檔名設為正常檔名
4.


確認是否有業者代碼資料夾

確認是否為檔案

取得檔案列表

移除非資料夾項目於列表

進入後再次確認檔案列表

進行解析





"""
import os
from  ParseFuc import TradeLogParse

CompanyList=["EZC","KRTC","ICASH"]

if __name__ == '__main__':
    while(1):
        Pause=input('執行解析')
        print('離開請自行關閉~~')
        if  os.path.exists("Datas"):
            TradeFolderList=os.listdir("./Datas")
    
            for TradeFolderName in TradeFolderList:
    
                if TradeFolderName=="EZC":
                    print("執行{}分析".format(TradeFolderName))
                    TradeLogParse.EZCParse()

                elif TradeFolderName=="KRTC":
                    print("執行{}分析".format(TradeFolderName))
                    TradeLogParse.KRTCExcelParse()
                    
                elif TradeFolderName=="ICASH":
                    print("執行{}分析".format(TradeFolderName))
                    TradeLogParse.ICashParse()

            print('解析程序結束')
            print('--------------------------------------------------------------')
        else:
            print("無Datas資料夾")

