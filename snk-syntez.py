    # -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
import shutil
import openpyxl                       # Для .xlsx
#import xlrd                          # для .xls
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, openX, sheetByName
import csv
import requests, lxml.html



def getXlsxString(sh, i, in_columns_j):
    impValues = {}
    for item in in_columns_j.keys() :
        j = in_columns_j[item]
        if item in ('закупка','продажа','цена') :
            if getCellXlsx(row=i, col=j, isDigit='N', sheet=sh).find('Звоните') >=0 :
                impValues[item] = '0.1'
            else :
                impValues[item] = getCellXlsx(row=i, col=j, isDigit='Y', sheet=sh)
            #print(sh, i, sh.cell( row=i, column=j).value, sh.cell(row=i, column=j).number_format, currencyType(sh, i, j))
        elif item == 'валюта_по_формату':
            impValues[item] = currencyType(row=i, col=j, sheet=sh)
        else:
            impValues[item] = getCellXlsx(row=i, col=j, isDigit='N', sheet=sh)
    return impValues



def convert2csv( cfg ):
    csvFName  = cfg.get('basic','filename_out')
    priceFName= cfg.get('basic','filename_in')
    sheetName = cfg.get('basic','sheetname')
    
    book = openpyxl.load_workbook(filename = priceFName, read_only=False, keep_vba=False, data_only=False)  # xlsx
    sheet = book.worksheets[0]                                                                              # xlsx
    log.info('-------------------  '+sheet.title +'  ----------')                                           # xlsx
#   sheetNames = book.get_sheet_names()                                                                     # xlsx

#   book = xlrd.open_workbook( priceFName.encode('cp1251'), formatting_info=True)                       # xls
#   sheet = book.sheets()[0]                                                                            # xls
#   log.info('-------------------  '+sheet.name +'  ----------')                                        # xls

    out_cols = cfg.options("cols_out")
    out_template = {}
    for vName in out_cols :
         out_template[vName] = cfg.get("cols_out", vName)
    
    in_cols = cfg.options("cols_in")
    in_cols_j = {}
    for vName in in_cols :
         in_cols_j[vName] = cfg.getint("cols_in",  vName)

    brands = cfg.options('discount')
    discount = {}
    for vName in brands :
        discount[vName] = (100 - int(cfg.get('discount',vName)))/100
        print(vName, discount[vName])

    outFile = open( csvFName, 'w', newline='', encoding='CP1251', errors='replace')
    csvWriter = csv.DictWriter(outFile, fieldnames=out_cols )
    csvWriter.writeheader()

    '''                                            # Блок проверки свойств для распознавания групп      XLSX                                  
    for i in range(2393, 2397):                                                         
        i_last = i
        ccc = sheet.cell( row=i, column=in_cols_j['группа'] )
        print(i, ccc.value)
        print(ccc.font.name, ccc.font.sz, ccc.font.b, ccc.font.i, ccc.font.color.rgb, '------', ccc.fill.fgColor.rgb)
        print('------')
    '''
    ssss    = []
    brand   = ''
    grp     = ''
    subgrp1 = ''
    subgrp2 = ''
    brand_koeft = 1
    recOut  ={}

#   for i in range(1, sheet.nrows) :                                    # xls
    for i in range(1, sheet.max_row +1) :                               # xlsx
        i_last = i
        try:
            ccc = sheet.cell( row=i, column=in_cols_j['группа'] )
            ccc2= sheet.cell( row=i, column=in_cols_j['цена'] )
            if   ccc.value == None:                                     # Пустая строка
                 pass
            elif ccc.font.color.rgb == 'FFFFFFFF':                      # Бренд
                brand=ccc.value
                grp = ''
                subgrp1 = ''
                subgrp2 = ''
                try:
                    print(brand)
                    brand_koeft = discount[brand.lower()]
                except Exception as e:
                    log.error('Exception: <' + str(e) + '> Ошибка назначения скидки в файле конфигурации' )
                    brand_koeft = 1
            elif ccc.fill.fgColor.rgb == 'FF746FEE':                    # Группа
                grp = ccc.value
                try:
                    num = float(grp[ :grp.find(' ')])
                except Exception as e:
                    grp = ccc.value
                else:
                    grp = grp[ grp.find(' ')+1:]
                subgrp1 = ''
                subgrp2 = ''
            elif ccc.fill.fgColor.rgb == 'FFE6E6E6':                    # Подгруппа-1
                subgrp1 = ccc.value
                try:
                    num = float(subgrp1[ :subgrp1.find(' ')])
                except Exception as e:
                    subgrp1 = ccc.value
                else:
                    subgrp1 = subgrp1[ subgrp1.find(' ')+1:]
                subgrp2 = ''
            elif ccc.fill.fgColor.rgb == 'FFFAFAFA':                    # Подгруппа-2
                subgrp2 = ccc.value
            elif ccc2.value == None:                                    # Пустая строка
                pass
                #print( 'Пустая строка. i=', i )
            elif ccc.font.b == False:                                   # Обычная строка
                impValues = getXlsxString(sheet, i, in_cols_j)
                impValues['бренд'] = brand
                impValues['группа'] = grp
                impValues['подгруппа'] = subgrp1+' '+subgrp2

                for outColName in out_template.keys() :
                    shablon = out_template[outColName]
                    for key in impValues.keys():
                        if shablon.find(key) >= 0 :
                            shablon = shablon.replace(key, impValues[key])
                    if (outColName == 'закупка') and ('*' in shablon) :
                        vvvv = float( shablon[ :shablon.find('*')     ] )
                        #print(vvvv)
                        shablon = str( float(vvvv) * brand_koeft )
                    recOut[outColName] = shablon

#                recOut['бренд'] = brand
#                recOut['группа'] = grp
#                recOut['подгруппа'] = subgrp1+' '+subgrp2
                csvWriter.writerow(recOut)

            else :                                                      # нераспознана строка
                log.info('Не распознана строка ' + str(i) + '<' + ccc.value + '>' )

        except Exception as e:
            if str(e) == "'NoneType' object has no attribute 'rgb'":
                pass
            else:
                log.debug('Exception: <' + str(e) + '> при обработке строки ' + str(i) +'.' )

    log.info('Обработано ' +str(i_last)+ ' строк.')
    outFile.close()



def download( cfg ):
    retCode     = False
    filename_new= cfg.get('download','filename_new')
    filename_old= cfg.get('download','filename_old')
    if cfg.has_option('download','login'):    login       = cfg.get('download','login'    )
    if cfg.has_option('download','password'): password    = cfg.get('download','password' )
    url_download_page= cfg.get('download','url_download_page'   )
    url_base         = cfg.get('download','url_base' )
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.0; rv:14.0) Gecko/20100101 Firefox/14.0.1',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Language':'ru-ru,ru;q=0.8,en-us;q=0.5,en;q=0.3',
               'Accept-Encoding':'gzip, deflate',
               'Connection':'keep-alive',
               'DNT':'1'
              }
    try:
        s = requests.Session()
        r = s.get(url_download_page,  headers = headers)
        page = lxml.html.fromstring(r.text)
        for item in page.xpath('//a'):
            if item.text == u"Полный прайс-лист компании СНК-СИНТЕЗ":
               #print(item.attrib)
               url_file = item.get('href')
        r = s.get(url_base + url_file)
        log.debug('Загрузка файла %16d bytes   --- code=%d', len(r.content), r.status_code)
        retCode = True
    except Exception as e:
        log.debug('Exception: <' + str(e) + '>')

    if os.path.exists( filename_new) and os.path.exists( filename_old): 
        os.remove( filename_old)
        os.rename( filename_new, filename_old)
    if os.path.exists( filename_new) :
        os.rename( filename_new, filename_old)
    f2 = open(filename_new, 'wb')                                  # Теперь записываем файл
    f2.write(r.content)
    f2.close()
    if filename_new[-4:] == '.zip':                                # Архив. Обработка не завершена
        log.debug( 'Zip-архив. Разархивируем '+ filename_new)
        dir_befo_download = set(os.listdir(os.getcwd()))
        os.system('unzip -oj ' + filename_new)
        dir_afte_download = set(os.listdir(os.getcwd()))
        new_files  = list( dir_afte_download.difference(dir_befo_download))
        filename_in= cfg.get('basic','filename_in')
        if os.path.exists(filename_in): os.remove( filename_in)
        os.rename( new_files[0], filename_in)
    return True



def is_file_fresh(fileName, qty_days):
    qty_seconds = qty_days *24*60*60 
    if os.path.exists( fileName):
        price_datetime = os.path.getmtime(fileName)
    else:
        log.error('Не найден файл  '+ fileName)
        return False

    if price_datetime+qty_seconds < time.time() :
        file_age = round((time.time()-price_datetime)/24/60/60)
        log.error('Файл "'+fileName+'" устарел!  Допустимый период '+ str(qty_days)+' дней, а ему ' + str(file_age) )
        return False
    else:
        return True



def config_read( cfgFName ):
    cfg = configparser.ConfigParser(inline_comment_prefixes=('#'))
    if  os.path.exists('private.cfg'):     
        cfg.read('private.cfg', encoding='utf-8')
    if  os.path.exists(cfgFName):     
        cfg.read( cfgFName, encoding='utf-8')
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    return cfg



def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def processing(cfgFName):
    log.info('----------------------- Processing '+cfgFName )
    cfg = config_read(cfgFName)
    filename_out  = cfg.get('basic','filename_out')
    filename_in= cfg.get('basic','filename_in')
    
    if cfg.has_section('download'):
        result = download(cfg)
    if is_file_fresh( filename_in, int(cfg.get('basic','срок годности'))):
        #os.system( dealerName + '_converter_xlsx.xlsm')
        convert2csv(cfg)
    folderName = os.path.basename(os.getcwd())
    if os.path.exists( filename_out): shutil.copy2( filename_out, 'c://AV_PROM/prices/' +folderName+'/'+filename_out)
    if os.path.exists( 'python.log'): shutil.copy2( 'python.log', 'c://AV_PROM/prices/' +folderName+'/python.log')
    if os.path.exists( 'python.1'  ): shutil.copy2( 'python.log', 'c://AV_PROM/prices/' +folderName+'/python.1'  )
    


def main( dealerName):
    """ Обработка прайсов выполняется согласно файлов конфигурации.
    Для этого в текущей папке должны быть файлы конфигурации, описывающие
    свойства файла и правила обработки. По одному конфигу на каждый 
    прайс или раздел прайса со своими правилами обработки
    """
    make_loger()
    log.info('          '+dealerName )
    for cfgFName in os.listdir("."):
        if cfgFName.startswith("cfg") and cfgFName.endswith(".cfg"):
            processing(cfgFName)


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( myName)
