# -*- coding: utf-8 -*-
"""
Created on Mon Mar 21 08:28:14 2022

@author: ed_liu
"""
import os
from  ParseFuc import TradeLogParse

#確認該資料夾以下所有層是否含有KRTC

if __name__ == '__main__':
    while(1):
        a=input('執行解析')
        print('離開請自行關閉~~')
        if  os.path.exists("Datas"):
            TradeFolderList=os.listdir("./Datas")
    
            for x in TradeFolderList:
    
                if x=="EZC":
                    print("執行{}分析".format(x))
                    TradeLogParse.EZCParse()

                elif x=="KRTC":
                    print("執行{}分析".format(x))
                    TradeLogParse.KRTCExcelParse()
                    
                elif x=="ICASH":
                    print("執行{}分析".format(x))
                    TradeLogParse.ICashParse()

            print('解析程序結束')
            print('--------------------------------------------------------------')
        else:
            print("無Datas資料夾")

