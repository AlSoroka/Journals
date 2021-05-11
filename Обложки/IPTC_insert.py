# Вставка IPTC информации в jpg
import datetime
from iptcinfo3 import IPTCInfo
import os

year
base = r'D:\enp.by\Develop\Журналы\_cover\\'
fileExt = '.JPG'
filelist = os.listdir(base)
for fl in filelist:
    fname=fl[:len(fl)-4].strip().split(' ',1)[1].strip()
    print(fname)
    info = IPTCInfo(base+fl)
    #info = IPTCInfo(base+fl, force=True)
    info['object name']=fname.encode(encoding = 'cp1251')
    info['headline'] =['Журналы по охране труда, промышленной безопасности, электробезопасности, газоснабжению, котельному оборудованию, производства работ '.encode(encoding = 'cp1251',errors = 'strict')]
    info['subject reference'] =['Журналы по охране труда, промышленной безопасности, электробезопасности, газоснабжению, котельному оборудованию, производства работ '.encode(encoding = 'cp1251',errors = 'strict')]
    info['caption/abstract'] =fname.encode(encoding = 'cp1251')
    info['keywords']=[fname.encode(encoding = 'cp1251')]
    info['source'] ='Энергопресс, BY'.encode(encoding = 'cp1251')
    info['writer/editor'] ='Энергопресс, BY'.encode(encoding = 'cp1251')
    info['contact'] =['enp.by', '+375-29-385-96-66']
    info['copyright notice']='© Энергопресс, 2020 г.'.encode(encoding = 'cp1251')
    info.save()










#info['contacts']=['www.enp.by', '+375-29-385-96-66']
#info.save()
#Print list of keywords, supplemental categories, contacts
"""
for a in info['caption/abstract']:
    print('Заголовок ', a.decode(encoding = 'UTF-8', errors = 'strict'))
for b in info['keywords']:
    print('Ключевые слова ', b.decode(encoding = 'cp1251', errors = 'strict'))
for c in info['supplementalCategories']:
    print('дополнительные категории ', c.decode(encoding = 'cp1251', errors = 'strict'))
for d in info['contacts']:
    print('Контакты ', d.decode(encoding = 'cp1251', errors = 'strict'))
"""

'''
#Get specific attributes…
caption = info['caption/abstract']

#Create object for file that may not have IPTC data
info = IPTCInfo('such_iptc.jpg', force=True)

#Add/change an attribute
info['caption/abstract'] = 'Witty caption here'
info['supplemental category'] = ['portrait']

#Save new info to file
info.save()
info.save_as('very_meta.jpg')
'''
