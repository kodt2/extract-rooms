from rooms import work_rooms, work_pdf
from names import work_names

import sys

def check_module(module_name, install_command):
    try:
        __import__(module_name)
    except ImportError:
        print(f"Module '{module_name}' not found. Please install it using: '{install_command}'")
        sys.exit(1)

check_module('openpyxl', 'pip install openpyxl')
check_module('datetime', 'pip install datetime')
check_module('fpdf', 'pip install fpdf')


with open('needed rooms.txt', encoding='utf-8') as f:
    arr = f.read()
    arr = arr.split('\n')
    for i in range(len(arr)):
        arr[i] = arr[i].split(' ')

with open("names.txt", encoding='utf-8') as f:
    arr_of_names = f.read()
    arr_of_names = arr_of_names.split('\n')
predsed_name = ''
predsed_phone = ''

K_surname = ''
K_intials = ''
B_name = ''
Z_name = ''
VAE = ""

KNV = K_surname+" "+K_intials
KaNV = K_surname+"a "+K_intials

async def schedule_picker(char):
    req = char.split('\n')
    for i in range(len(req)):
        req[i] = req[i].split(' ')
    print('0',req)
    arr = work_rooms(req, KaNV, KNV, Z_name, VAE, B_name, predsed_name, predsed_phone)
    s=''
    print(arr)
    for i in arr:
        s+=i[0]+' '+i[1]+' '+i[2]+' '+i[4]+' '+i[5]+' '+ i[6]+ ' '+i[7]+' '+ str(i[8])+'\n'
    #print(s)
    return s
async def pdf_picker(char):
    req = char.split('\n')
    for i in range(len(req)):
        req[i] = req[i].split(' ')
    arr,name = work_pdf(req, KaNV, KNV, Z_name, VAE, B_name, predsed_name, predsed_phone)
    arr.sort(key=lambda x: x[0])
    s=''
    print(arr)
    for i in arr:
        s += i[0] + ' ' + i[1] + ' ' + i[2] + ' ' + i[5] + '-' + i[6] + ' ' + i[7] + ' ауд. ' + str(i[8]) + ' корп.\n'
    print('dfb',s)
    return s, name

work_names(arr_of_names,KaNV, KNV, Z_name, VAE, B_name, predsed_name, predsed_phone)