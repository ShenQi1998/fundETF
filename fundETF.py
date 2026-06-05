# V1.0 2026-06-05 获取上交所/深交所全量ETF数据    Author shenqi  


import sys
import requests
import time
import json
from datetime import datetime, timedelta
from openpyxl import Workbook


#自定义参数 可修改  Begin
#深证ETF日期，必须是交易日，格式 "2026-06-04" ，默认是当前日期前一天 ，
STAT_DATE = ""
#标记上证基金代码
SHMark = ['510330' , '510050']    
#标记深证基金 
SZMark = ['159001' , '159003']  
#自定义参数 可修改  End


headers = {
        "Accept": "*/*",
        "Referer": "https://www.sse.com.cn/",
    }
fundNameMap = {}


# 打印日志方法
debugFlag = '0'
def printff ( obj , level ):
    if( debugFlag == '0' ):
        if( level == 'info' ):
            print( obj )
    else:
        print( obj )


def getAllFoudName( responseJson ):
    fundNameList= []
    for index,row in enumerate(responseJson['result']) :
        fundNameList.append(row['SEC_CODE'])
        if( (index+1) % 300 == 0 ):
            fundNameStr = ','.join(fundNameList)
            fundNameList= []
            getFoudName(fundNameStr)

    if( len(fundNameList) != 0):
        fundNameStr = ','.join(fundNameList)
        getFoudName(fundNameStr)
  
    printff( fundNameMap , "debug" )


def getFoudName( fundNameStr ):
    printff( fundNameStr , "debug" )
    url2 = "https://query.sse.com.cn/security/stock/queryExpandName.do"
    params2 = {
        "jsonCallBack": "jsonpCallback00000002",
        "secCodes": fundNameStr,
        "_": int(time.time() * 1000)
    }
    response2 = requests.get(url2, headers=headers , params = params2)
    responseJson2 = json.loads(response2.text[22:-1])
    printff( responseJson2['result'] , "debug" )
    for row in responseJson2['result'] :
        fundNameMap[row[0]] = row[1]



#入口
if __name__ == "__main__":

    ## step1 初始化日期
    if( STAT_DATE == ""):
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        STAT_DATE = yesterday.strftime("%Y-%m-%d")
    printff( "获取日期" + STAT_DATE , "info" )


    ## step2 获取上交易所ETF数据
    printff( "获取上交所数据Begin>>>>>>>" , "info" )

    url = "https://query.sse.com.cn/commonQuery.do"

    params = {
        "jsonCallBack": "jsonpCallback00000001",
        "isPagination": "false",
        "sqlId":"COMMON_SSE_ZQPZ_ETFZL_XXPL_ETFGM_SEARCH_L",
        "STAT_DATE":STAT_DATE,
        "_": int(time.time() * 1000)
    }

    response = requests.get(url, headers=headers , params = params)
    responseJson = json.loads(response.text[22:-1])
    printff( "responseJson['result']" , "debug" )


    ## step3 获取上交所基金中文名
    getAllFoudName(responseJson)
    printff( "获取上交所数据End>>>>>>>" , "info" )


    ## step4 获取深交易所ETF数据
    printff( "获取深交所数据Begin>>>>>>>" , "info" )

    urlSZ = "https://www.szse.cn/api/report/ShowReport/data"
    paramsSZ = {
        "SHOWTYPE":"JSON",
        "CATALOGID":"scsj_fund_jjgm",
        "TABKEY":"tab1",
        "PAGENO":"1",
        "txtStart": STAT_DATE,
        "txtEnd": STAT_DATE ,
        "jjlb":"ETF",
        "random":"0.6720094301496038",
    }
    nowPage = 1 
    lastPage = False
    responseJsonSZData = []

    while not lastPage: 
        paramsSZ["PAGENO"] = nowPage
        responseSZ = requests.get(urlSZ,  params = paramsSZ)
        responseJsonSZ = json.loads(responseSZ.text)
        pagecount = responseJsonSZ[0]['metadata']['pagecount']
        if( pagecount == 0 ):
            printff( "执行失败：获取深交所数据失败,请检查日期" , "info" )
            sys.exit(1)
        if( nowPage == pagecount  ):
            lastPage = True 
        nowPage = nowPage + 1
  
        printff( responseJsonSZ[0]['data'] , "debug" )
        printff( "已获取" + str(nowPage - 1) + "/" + str(pagecount) , "info" )

        responseJsonSZData = responseJsonSZData +  responseJsonSZ[0]['data'] 

    printff( responseJsonSZData , "debug" )
    printff( "获取深交所数据End>>>>>>>" , "info" )


    # ## step5 写入EXCEL
    printff( "开始写入数据>>>>>>>>>>>" , "info" )

    wb = Workbook()
    nowTime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")  # 2026-06-05-143045
    file_path = 'C://' + nowTime + '.xlsx'


    sheet1 = wb.active
    sheet1.title = '上交所数据' + STAT_DATE
    sheet1.append(['日期', '基金代码', '基金名称' , '份额' , '标记'])
    for row in responseJson['result'] :
        mark = "1" if row['SEC_CODE'] in SHMark else ""
        sheet1.append( [ row['STAT_DATE'] ,row['SEC_CODE'] ,  fundNameMap[row['SEC_CODE']] , row['TOT_VOL'] , mark ] )


    sheet2 = wb.create_sheet('深交所数据' + STAT_DATE )
    sheet2.append(['日期', '基金代码', '基金简称' , '基金规模(万份)' , '标记'])
    for row in responseJsonSZData :
        mark = "1" if row['fund_code'] in SZMark else ""
        sheet2.append( [ row['size_date'] , row['fund_code'] ,  row['security_short_name'] , row['current_size'] , mark ] )


    wb.save( file_path )
    printff( "开始写入数据完成>>>>>>>>>>>" , "info" )
    printff( "执行成功，EXCEL已创建至C盘根目录，文件名《" + nowTime + ".xlsx》" , "info" )



