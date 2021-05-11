import xlrd
import re
import time
import ftplib
import shutil
import datetime
import os
import ftplib
from translit_and_dir_create import *


# Копируем и передаем на сервер фалй прайса
# Создаем новое имя для прайса, который скопируем на сервер. Оно включает дату
now = str(datetime.datetime.now())
den=now[8:10]
mes=now[5:7]
god=now[:4]

#new_name='Price Energopress TNPA '+den+'-'+mes+'-'+god+'.xls'
#new_nameJ='Price Energopress Journals '+den+'-'+mes+'-'+god+'.xls'
#path=r'\\Buh\подписка\Прайсы'
path=os.getcwd()
price='Price Energopress Journals 20-11-2020.xlsx'

#Создаем HTML страницу из файла прайса

#fp = open(os.path.join(path,'Прайс для сайта.html'), 'w')

fp = open('index.html', 'w', encoding="utf-8")

nazvanie=' '
postanovlenie=' '
cena="0.00"
kol='  '
idList=''
themeList=''
if  os.path.isfile(os.path.join(path, 'miss_cover.txt')):
    os.remove(os.path.join(path, 'miss_cover.txt'))

host = "vh116.hoster.by"
ftp_user = "energeti"
ftp_password = "Eev0quai"

while True:
    copy_to_server=input('Копировать файлы на сервер? (y/n) ')
    if copy_to_server=='y' or copy_to_server=='n':
        break



#Открываем FTP
with ftplib.FTP(host, ftp_user, ftp_password) as con:
    #con = ftplib.FTP(host, ftp_user, ftp_password)

    con.__class__.encoding = "utf-8"



    # Записываем дату обновления

    #Создаем заголовок таблицы


    with open (r'part1.html', 'r', encoding="utf-8") as fl:
        part1=fl.read()
        part1=part1.replace('{{PriceFile}}',price)


    fp.write(part1)





    #Открываем файл Прайс.xls
    rbook = xlrd.open_workbook(os.path.join(path,price))
    # Получаем доступ к листу
    rsheet=rbook.sheet_by_index(0)
    i=0
    theme=''
    for str_num in range(12,rsheet.nrows):
        i+=1

    #Создаем тело таблицы
        razdel=str(rsheet.cell(str_num,0).value)

        psevdo=str(rsheet.cell(str_num,1).value).strip()
        #**********************************************************

        theme=str(rsheet.cell(str_num,0).value).strip() if len(str(rsheet.cell(str_num,0).value).strip())>0 else theme
        npp=str(rsheet.cell(str_num,1).value).strip()
        nazvanie=str(rsheet.cell(str_num,2).value)
        hpl=rsheet.hyperlink_map.get((str_num,3))
        url_cell=str(rsheet.cell(str_num,3).value)
        url = '(No URL)' if hpl is None else hpl.url_or_path
        # Заменяем разрывы строк внутри ячеек на <br>
        #nazvanie=re.sub('[\r\n]','<br>', nazvanie)
        print(i,'   ',nazvanie)
        # После предлогов ставим неразрывный пробел
        #nazvanie=re.sub('(\s[а-я]{1,3})\s',r'\1&nbsp;' , nazvanie)
        # После предлогов, стоящих перед кавычками ставим неразрывный пробел
        #nazvanie=re.sub('(["|«|(][а-я,А-Я]{1,3})\s',r'\1&nbsp;' , nazvanie)
        nazvanie=re.sub('^Сборник','<b>СБОРНИК</b>', nazvanie)
        nazvanie=re.sub('^СБОРНИК','<b>СБОРНИК</b>', nazvanie)
        postanovlenie=str(rsheet.cell(str_num,4).value)
        postanovlenie=postanovlenie.replace('°',' ')
        # Заменяем разрывы строк внутри ячеек на <br>
        postanovlenie=re.sub('[\r\n]','<br>', postanovlenie)
        # После предлогов ставим неразрывный пробел
        #postanovlenie=re.sub('(\s[а-я]{1,3})\s',r'\1&nbsp;' , postanovlenie)
        # Если внутри есть 2018 или 2019 окрашиваем его жирным черным
        #postanovlenie=postanovlenie.replace('2018','<span style="color: black; font-size:120%"><b>2018</b></span> ')
        #postanovlenie=postanovlenie.replace('2019','<span style="color: black; font-size:120%"><b>2019</b></span> ')

        cena=str(rsheet.cell(str_num,3).value).strip()
        cena=cena.replace(',','.')
        cena=re.sub('[\r\n]', '', cena)

        # Формируем цену с двумя знаками после запятой: чтобы не округляло 14.80 в 14.8
        if len(cena)>0:
            if cena[0].isdigit():
                #преобразуем в плавающее
                #cena =  float(rsheet.cell(str_num,4).value)
                cena =  float(cena)
                #преобразуем в форма два знака до, два после запятой
                cena="%3.2f" % (cena)

            cena=cena.replace(',','.')

        #schet='<a href="_files/'+psevdo+'_schet.pdf" download>Скачать счет</a>'
        count='<input id="count_'+psevdo+'" type="text" value="" \
                title="Введите необходимое Вам количество экз." style="width:75%; text-align:center ">\n'\
                # onchange="sum_calculate(this.value)


        if nazvanie=='Название':
            continue
        if not razdel.strip() and not nazvanie.strip():
            continue
        if razdel.strip():
            if razdel.find('НОВЫЕ')>-1:
                fp.write('<tr class="razdel"><th colspan="7" align="center" style="color:red;">'+razdel+'</th></tr>\n')
            else:
                fp.write('<tr class="razdel"  class="'+theme.replace(' ','')+'"><th colspan="7" align="center"><a name="'+razdel+'"></a>'+razdel+'</th></tr>\n')
                # Добавляем название раздела в список разделов (тем)
                themeList=themeList+"'"+theme.replace('  ',' ')+"'," if len(theme.strip())>0 else themeList
            continue
        idList = idList+"'"+psevdo+"'," if len(psevdo.strip())>0 else idList

        fp.write('<tr class="'+theme.replace(' ','')+'" id="'+psevdo+'">\n')

        # fp.write('<td  class="col5p">'+npp+'<a name="'+psevdo+'"></a></td>\n') строка с якорем
        fp.write('<td  class="col5p">'+npp+'</td>\n')

        if url=='(No URL)':
            #fp.write('<td class="col45p"  id="'+psevdo+'_name" >'+nazvanie +'</td>'+'\n')
            url=translit(psevdo,nazvanie,cena,'',theme, con, copy_to_server)
        nazvanie=re.sub('форма 2020 г.','<b>форма 2020 г.</b>', nazvanie)
        fp.write('<td class="col45p"  id="'+psevdo+'_name" ><a href="'+url+'" target="blank">'+nazvanie+'</a></td>'+'\n')

        fp.write('<td class="col10p"  id="'+psevdo+'_price">'+cena+'</td>\n')
        fp.write('<td class="col10p" ><input id="'+psevdo+'_input_count"')
        fp.write('class="book_count" type="number" value="" size="2"')
        fp.write('title="Введите необходимое Вам количество экз."')
        fp.write('onchange="sum_calculate('+"'"+psevdo+"',"+cena+',this.value)"></td>\n')

        fp.write('<td class="col10p" >Без НДС</th>'+'\n')
        fp.write('<td class="col10p"  id="'+psevdo+'_sum_disc">&mdash;</td>'+'\n')
        fp.write('<td class="col10p" ><span id="'+psevdo+'_sum_k_opl">&mdash;</span></td>\n')

        fp.write('</tr>\n')


    with open (r'part3.html', 'r', encoding="utf-8") as fl:
        part3=fl.read()

    fp.write(part3)

    with open (r'partFooter.html', 'r', encoding="utf-8") as fl:
        partFooter=fl.read()


    fp.write (partFooter)

    fp.write("\n</body>\n<script>\n")
    fp.write("var idList=["+idList+"];\n")
    fp.write("var themeList=["+themeList+"];\n")

    fp.write("insertThemes ();\n")
    fp.write("WindowOnfocus ();\n")


    #fp.write("gotoAncor();\n")



    fp.write("</script>\n</html>")
    fp.close()


if copy_to_server=='y':
    with ftplib.FTP(host, ftp_user, ftp_password) as con:
        con.__class__.encoding = "utf-8"
        CopyFileToFTP(con, '', 'index.html')
        CopyFileToFTP(con, '', price)
