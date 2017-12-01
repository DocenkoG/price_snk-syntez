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
from   price_tools import getCellXlsx, getCell, quoted, dump_cell, currencyType, subInParentheses
import csv



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



def convert2csv( dealerName, csvFName ):
    cfgFName   = ('cfg_'+dealerName+'.cfg').lower()
    fileNameIn = ('new_'+dealerName+'.xlsx').lower()
    basicNamelist, basic = config_read( cfgFName, 'basic' )

    book = openpyxl.load_workbook(filename = fileNameIn, read_only=False, keep_vba=False, data_only=False)  # xlsx
    sheet = book.worksheets[0]                                                                              # xlsx
    log.info('-------------------  '+sheet.title +'  ----------')                                           # xlsx
#   sheetNames = book.get_sheet_names()                                                                     # xlsx

#   book = xlrd.open_workbook( fileNameIn.encode('cp1251'), formatting_info=True)                       # xls
#   sheet = book.sheets()[0]                                                                            # xls
#   log.info('-------------------  '+sheet.name +'  ----------')                                        # xls

    out_cols, out_template = config_read(cfgFName, 'cols_out')
    in_cols,  in_cols_j    = config_read(cfgFName, 'cols_in')
    brands,   discount     = config_read(cfgFName, 'discount')
    for k in in_cols_j.keys():
        in_cols_j[k] = int(in_cols_j[k])
    for k in discount.keys():
        discount[k] = (100 - int(discount[k]))/100

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

                recOut['бренд'] = brand
                recOut['группа'] = grp
                recOut['подгруппа'] = subgrp1+' '+subgrp2
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



def config_read( cfgFName, partName ):
    log.debug('Reading config ' + cfgFName )
    config = configparser.ConfigParser()
    keyList = []
    keyDict = {}
    if  os.path.exists(cfgFName):     
        config.read( cfgFName, encoding='utf-8')
        keyList = config.options(partName)
        for vName in keyList :
            if ('' != config.get(partName, vName)) :
                keyDict[vName] = config.get(partName, vName)
    else: 
        log.debug('Нет файла конфигурации '+cfgFName)
    
    return keyList, keyDict



def download( dealerName ):
    pathDwnld = './tmp'
    pathPython2 = 'c:/Python27/python.exe'
    retCode = False
    fUnitName = os.path.join( dealerName +'_unittest.py')
    if  not os.path.exists(fUnitName):
        log.debug( 'Отсутствует юниттест для загрузки прайса ' + fUnitName)
    else:
        dir_befo_download = set(os.listdir(pathDwnld))
        os.system( fUnitName)                                                           # Вызов unittest'a
        dir_afte_download = set(os.listdir(pathDwnld))
        new_files = list( dir_afte_download.difference(dir_befo_download))
        if len(new_files) == 1 :   
            new_file = new_files[0]                                                     # загружен ровно один файл. 
            new_ext  = os.path.splitext(new_file)[-1]
            DnewFile = os.path.join( pathDwnld,new_file)
            new_file_date = os.path.getmtime(DnewFile)
            log.info( 'Скачанный файл ' +DnewFile + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) ) )
            if new_ext == '.zip':                                                       # Архив. Обработка не завершена
                log.debug( 'Zip-архив. Разархивируем.')
                work_dir = os.getcwd()                                                  
                os.chdir( os.path.join( pathDwnld ))
                dir_befo_download = set(os.listdir(os.getcwd()))
                os.system('unzip -oj ' + new_file)
                os.remove(new_file)   
                dir_afte_download = set(os.listdir(os.getcwd()))
                new_files = list( dir_afte_download.difference(dir_befo_download))
                if len(new_files) == 1 :   
                    new_file = new_files[0]                                             # разархивирован ровно один файл. 
                    new_ext  = os.path.splitext(new_file)[-1]
                    DnewFile = os.path.join( os.getcwd(),new_file)
                    new_file_date = os.path.getmtime(DnewFile)
                    log.debug( 'Файл из архива ' +DnewFile + ' имеет дату ' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(new_file_date) )     )
                    DnewPrice = DnewFile
                elif len(new_files) >1 :
                    log.debug( 'В архиве не единственный файл. Надо разбираться.')
                    DnewPrice = "dummy"
                else:
                    log.debug( 'Нет новых файлов после разархивации. Загляни в папку юниттеста поставщика.')
                    DnewPrice = "dummy"
                os.chdir(work_dir)
            elif new_ext in ( '.csv', '.htm', '.xls', '.xlsx'):
                DnewPrice = DnewFile                                             # Имя скачанного прайса
            if DnewPrice != "dummy" :
                FoldName = 'old_' + dealerName + new_ext                         # Старая копия прайса, для сравнения даты
                FnewName = 'new_' + dealerName + new_ext                         # Предыдущий прайс, с которым работает макрос
                if  (not os.path.exists( FnewName)) or new_file_date>os.path.getmtime(FnewName) : 
                    log.debug( 'Предыдущего прайса нет или он устарел. Копируем новый.' )
                    if os.path.exists( FoldName): os.remove( FoldName)
                    if os.path.exists( FnewName): os.rename( FnewName, FoldName)
                    shutil.copy2(DnewPrice, FnewName)
                    retCode = True
                else:
                    log.debug( 'Предыдущий прайс не старый, копироавать не надо.' )
                # Убрать скачанные файлы
                if  os.path.exists(DnewPrice):  os.remove(DnewPrice)   
            
        elif len(new_files) == 0 :        
            log.debug( 'Не удалось скачать файл прайса ')
        else:
            log.debug( 'Скачалось несколько файлов. Надо разбираться ...')

    return retCode




def make_loger():
    global log
    logging.config.fileConfig('logging.cfg')
    log = logging.getLogger('logFile')



def main( dealerName):
    make_loger()
    log.info('         '+dealerName )
    csvFName   = ('csv_'+dealerName+'.csv').lower()

    #s  download( dealerName)
    convert2csv( dealerName, csvFName)
    if os.path.exists( csvFName    ) : shutil.copy2( csvFName ,    'c://AV_PROM/prices/' + dealerName +'/'+csvFName )
    if os.path.exists( 'python.log') : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.log')
    if os.path.exists( 'python.1'  ) : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.1'  )


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( 'snk-syntez')
