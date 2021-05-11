import os, xlrd



path = r"D:\cov\low"


os.chdir(path)




rbook = xlrd.open_workbook(r"C:\Users\User\Desktop\enp.by\ForCovers.xls")
# Получаем доступ к листу
rsheet=rbook.sheet_by_index(0)
i=0
for str_num in range(1,rsheet.nrows):
    fil_new_name=str(rsheet.cell(str_num,0).value).strip()+".jpg"
    #fil_new_name=str(rsheet.cell(str_num,0).value).strip()+" "+str(rsheet.cell(str_num,1).value).strip()+".jpg"
    fil_old_name="Обложки слияние и проверка-"+str(str_num).strip()+".jpg"
    os.rename(fil_old_name, fil_new_name)

