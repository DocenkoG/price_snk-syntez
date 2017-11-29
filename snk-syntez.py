# -*- coding: UTF-8 -*-
import os
import os.path
import logging
import logging.config
import sys
import configparser
import time
# import elittech_downloader
import shutil
import csv



def convert2csv( dealerName ):
    inFile  = open( 'new_'+dealerName +'.csv', 'r', newline='')
    outFile = open(        dealerName +'.csv', 'w', newline='')
    csvReader = csv.DictReader(inFile)
    csvWriter = csv.DictWriter(outFile, fieldnames=[
        'бренд',
        'группа',
        'подгруппа',
        'код',
        'код производителя',
        'наименование',
        'описание',
        'закупка',
        'продажа',
        'валюта',
        'наличие',
        '?'])

    for k in range (0, len(csvReader.fieldnames)):
        csvReader.fieldnames[k] = csvReader.fieldnames[k].lower()
    print(csvReader.fieldnames)
    csvWriter.writeheader()
    recOut = {}
    for recIn in csvReader:
        recOut['бренд']        = recIn['бренд']
        recOut['группа']       = recIn['группа']
        recOut['подгруппа']    = recIn['подгруппа']
        recOut['код']          = recIn['наименование']
        recOut['код производителя'] = recIn['код производителя']
        recOut['наименование'] = recIn['бренд']+' '+recIn['наименование']+' '+recIn['описание']
        recOut['описание']     = recIn['бренд']+' '+recIn['наименование']+' '+recIn['описание']+' Код продавца: '+recIn['артикул']+' код производителя: '+recIn['код производителя']
        recOut['продажа']      = recIn['розничная']
        try:
            recOut['закупка']  = float(recIn['розничная']) * 0.7
        except:
            recOut['закупка']  = 0.1
        #
        recOut['валюта']       = recIn['валюта']
        recOut['наличие']      = recIn['наличие']
        recOut['?']            = '?'
        #print(recOut)
        csvWriter.writerow(recOut)
    log.info('Обработано '+ str(csvReader.line_num) +'строк.')
    inFile.close()
    outFile.close()



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
                print(DnewFile)
                print(new_file)
                os.system('unzip -oj ' + new_file)
                os.remove(new_file)   
                dir_afte_download = set(os.listdir(os.getcwd()))
                new_files = list( dir_afte_download.difference(dir_befo_download))
                print(new_files)
                if len(new_files) == 1 :   
                    new_file = new_files[0]                                             # разархивирован ровно один файл. 
                    new_ext  = os.path.splitext(new_file)[-1]
                    DnewFile = os.path.join( os.getcwd(),new_file)
                    new_file_date = os.path.getmtime(DnewFile)
                    print(DnewFile)
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
    if  download( dealerName):
        pass
        #convert2csv( dealerName )
    priceName = dealerName+'.csv'
    if os.path.exists( priceName   ) : shutil.copy2( priceName,    'c://AV_PROM/prices/' + dealerName +'/'+priceName)
    if os.path.exists( 'python.log') : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.log')
    if os.path.exists( 'python.1'  ) : shutil.copy2( 'python.log', 'c://AV_PROM/prices/' + dealerName +'/python.1'  )


if __name__ == '__main__':
    myName = os.path.basename(os.path.splitext(sys.argv[0])[0])
    mydir    = os.path.dirname (sys.argv[0])
    print(mydir, myName)
    main( 'snk-syntez')
