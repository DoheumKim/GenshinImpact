
'''
# 저자 : DoheumKim
# 창작)신규 선령펫 온필드 캐릭따라 색상변경 <- 여기서 캐릭터 ib값 등 자동추가 툴
'''
# 참고: 이 스크립트는 3dmigtoGIMI 플러그인을 사용하여 생성된 모드에서만 작동합니다
# 원글: https://arca.live/b/genshinskinmode/96039430
# 엑셀파일 출처: https://docs.google.com/spreadsheets/d/1IXj6C-ZX6p7TxcUt6mR0zI8fRTJRZf4VAL6yjMDwITs/edit?usp=sharing
import os
from openpyxl import load_workbook

ini_file = ''
load_wb = ''
load_ws = ''

# --------------------------------------------------------------------------------------------------------------------------- #
def main():
    global load_wb,load_ws
    load_wb = load_workbook("4.3 hash hunting.xlsx", data_only=True)
    load_ws = load_wb['0']
    char = {}
    MainChar = False
    MainChar_txt = ''
    start_ini = 0
    start_txt = ''

    a,i = 0,3   #i == 셀에서 캐릭터이름이 시작되는 값
    result = ''
    while True:
        if load_ws[f'A{str(i)}'].value == None:break    #캐릭터 없으면 break
        if load_ws[f'E{str(i)}'].value == None:break    #원소 없으면 break
        if load_ws[f'F{str(i)}'].value == None:break    #무기 없으면 break
        else:
            char[load_ws[f'A{str(i)}'].value] = [load_ws[f'C{str(i)}'].value,load_ws[f'E{str(i)}'].value,load_ws[f'F{str(i)}'].value]
            a += 1   #char == {'캐릭터':[ib,원소,무기],'캐릭터2':[ib2,원소2,무기2]}
            i += 1
            if 'Traveler' in str(load_ws[f'A{str(i)}'].value):
                MainChar = True
    
    # ini파일
    fileEx = r'.ini'
    ini_file = [os.path.join(file) for file in os.listdir() if file.endswith(fileEx)]
    print('\n\n\nini파일목록: \n')
    for i in range(len(ini_file)):print(f'{str(i).zfill(2)}: {ini_file[i]}')
    print()
    if len(ini_file) > 1:ini_file = ini_file[int(input('바꿀 파일 선택(0부터 시작, 숫자만)ex: 0\n: '))]
    else:ini_file = ini_file[0]

    with open(ini_file, "r", encoding="utf-8") as f:
        if MainChar == True:
            MainChar_txt = f'\
[Constants]\n\
global persist $swapvar = 0\n\
global persist $MainChar = 1\n\
global persist $swap_WP = 0\n\
\n\
[KeySwap]\n\
key = ctrl + .\n\
type = cycle\n\
$swapvar = 1,2,3,4,5,6,7\n\
\n\
[KeySwapMainChar]\n\
key = 0\n\
type = cycle\n\
$MainChar = 1,2,3,5,7'
        lines = f.readlines()
        
        for i in range(len(lines)):
            result += lines[i]
            if ';' in lines[i] and start_ini < 3:
                start_ini += 1
                if start_ini == 3:
                    start_txt = result
                    result = ''
                    print(start_txt,'res',result)
                    start_ini = 4
                    break
        c = '\n이미 있는 캐릭: \n'

        for i,l in enumerate(lines):
            if '[KeySwapMainChar]' in lines[i]:MainChar_txt = f'\
[Constants]\n\
global persist $swapvar = 0\n\
global persist $MainChar = 1\n\
global persist $swap_WP = 0\n\
\n\
[KeySwap]\n\
key = ctrl + .\n\
type = cycle\n\
$swapvar = 1,2,3,4,5,6,7'
            if '[TextureOverride' in lines[i]:
                if lines[i][16:-10] in char:
                    del char[lines[i][16:-10]]
                    c += f'{lines[i][16:-10]} / '
        print(c[:-3])
    
    #print(char)
    if len(list(char.keys())) != 0:result += '\n\n; ------------------- [ADDED Overrides] -------------------'
    for i in range(len(list(char.keys()))):
        if list(char.values())[i][2] == 4:
            #result += '\n'
            result += f'\n[TextureOverride{list(char.keys())[i]}Position]\n'
            result += f'hash = {list(char.values())[i][0]}\n'
            result += 'match_priority = 1'
            result += '\n'
            if 'Traveler' in list(char.keys())[i]:result += '$element = $MainChar\n'
            else:result += f'$element = {list(char.values())[i][1]}\n'
        #result += f'$swap_WP = {list(char.values())[i][2]}\n'
    with open(ini_file, "w", encoding="utf-8") as f:
            start_txt += result
            f.write('')
            f.write(start_txt)
            #f.write(result)
#             if MainChar:
#                 f.write(MainChar_txt)
#                 f.write(result[132:])
#                 print('\n여행자용 스왑키 추가완료,키는 0')
#             else:
#                 f.write(
# '[Constants]\n\
# global persist $swapvar = 0\n\
# global persist $MainChar = 1\n\
# global persist $swap_WP = 0\n\
# \n\
# [KeySwap]\n\
# key = ctrl + .\n\
# type = cycle\n\
# $swapvar = 1,2,3,4,5,6,7')
#                 f.write(result[132:])
            if len(list(char.keys())) == 0:print('\n추가된 값 없음, 이미 추가 가능한 모든 값이 있습니다')
            else:
                print('\n추가된 값:')
                a = ''
                for i in range(len(list(char.keys()))):a += f'{str(i).zfill(2)}: {list(char.keys())[i]} / '
                print(a[:-3])

if __name__ == "__main__":
    #main(0,'')
    main()