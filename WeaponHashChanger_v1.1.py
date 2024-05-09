
'''
# 저자 : DoheumKim
'''
# 참고: 이 스크립트는 3dmigtoGIMI 플러그인을 사용하여 생성된 모드에서만 작동합니다
# v1.1 원신 4.5버전 대응(치오리 무기 추가)
# 엑셀파일:
# https://docs.google.com/spreadsheets/d/1F4w7Dn4foqaJUrmG52QQk3p2rDnvCTDd/edit?usp=sharing&ouid=102220522664008852910&rtpof=true&sd=true
import os
import re
import argparse
import hashlib
from openpyxl import load_workbook

ini_file,new_hash = '',''
load_wb = ''
load_ws = ''

# --------------------------------------------------------------------------------------------------------------------------- #
def load_excel(select_list):
    global load_wb,load_ws,new_hash
    load_wb = load_workbook("WP_list.xlsx", data_only=True)
    load_ws = load_wb[select_list[0]]
    comments_list = []
     
    if   select_list[0] == '한손검':        #무기종류별 무기개수, 무기 셀위치 설정
        if   select_list[1] == 5:WP_num,WP_start = 11,2
        elif select_list[1] == 4:WP_num,WP_start = 23,14
        elif select_list[1] == 3:WP_num,WP_start = 6,38
    elif select_list[0] == '장병기':
        if   select_list[1] == 5:WP_num,WP_start = 7,2
        elif select_list[1] == 4:WP_num,WP_start = 17,10
        elif select_list[1] == 3:WP_num,WP_start = 3,28     
    elif select_list[0] == '활':
        if   select_list[1] == 5:WP_num,WP_start = 8,2
        elif select_list[1] == 4:WP_num,WP_start = 22,11
        elif select_list[1] == 3:WP_num,WP_start = 5,34
    elif select_list[0] == '대검':
        if   select_list[1] == 5:WP_num,WP_start = 7,2
        elif select_list[1] == 4:WP_num,WP_start = 21,10
        elif select_list[1] == 3:WP_num,WP_start = 5,32
    elif select_list[0] == '법구':
        if   select_list[1] == 5:WP_num,WP_start = 11,2
        elif select_list[1] == 4:WP_num,WP_start = 19,14
        elif select_list[1] == 3:WP_num,WP_start = 5,34


    for i in range(1,WP_num+1):     #무기리스트 출력
        print(f'\n{str(i).zfill(2)}: {load_ws[f"B{str(WP_start+i-1)}"].value}',end='\t')
        coomments = load_ws[f"H{str(WP_start+i-1)}"].value
        exist = load_ws[f"C{str(WP_start+i-1)}"].value  #무기 해쉬값 없으면 없다고 표시
         
        HeadNor = load_ws[f"M{str(WP_start+i-1)}"].value
        BodyDif = load_ws[f"N{str(WP_start+i-1)}"].value
        BodyDifGuide = load_ws[f"O{str(WP_start+i-1)}"].value
        BodyLight = load_ws[f"P{str(WP_start+i-1)}"].value
        BodyMetal = load_ws[f"Q{str(WP_start+i-1)}"].value

        if coomments == '상시' or coomments == '한정' or coomments == '단조' or coomments == '낚시' or coomments == '이벤트':coomments += ' 무기'
        elif coomments[-2:] == '기행':coomments = coomments+' 무기'
        elif coomments[-2:] == '.2':coomments = '이나즈마 '+coomments[:-2]+' 무기'
        elif coomments[-2:] == '.3':coomments = '수메르 '+coomments[:-2]+' 무기'
        elif coomments[-2:] == '.4':coomments = '폰타인 '+coomments[:-2]+' 무기'
        elif coomments[-3:] == '.1a':coomments = '몬드 보물'+coomments[:-3]
        elif coomments[-3:] == '.1b':coomments = '리월 보물'+coomments[:-3]
        elif coomments == '퀘스트' or coomments == '플스전용' or coomments == '스타라이트' or coomments == '대화':pass
        else:coomments += ' 전무'
        if exist =='':print('(없음)',end='')
        print(f'\t\t({coomments})',end='')
        comments_list.append(coomments)

    select_WP = int(input('\n원하는 무기를 골라주세요(왼쪽의 번호로)\n: '))
    select_WP += WP_start-1

    print(f"\n{load_ws[f'B{select_WP}'].value}({comments_list[select_WP-WP_start]}) 선택됨\n")
    if input('계속하려면 ENTER를 누르세요\n') == '':main(1,select_WP)
    else:print('프로그램을 종료합니다')


def WP_change(value):
    select_list = list()
    for _ in range(4):
        if 0 <= value <= 4:
            continue
        else:
            value = int(input('0~4사이를 입력해주세요\n: '))

    rank = int(input('\n무기등급입력(5성이면 5입력)\n: '))
    if value == 0:select_list = ['한손검',rank]
    elif value == 1:select_list = ['장병기',rank]
    elif value == 2:select_list = ['활',rank]
    elif value == 3:select_list = ['대검',rank]
    elif value == 4:select_list = ['법구',rank]
    print(f'\n{select_list[1]}성 {select_list[0]} 선택됨\n----------')
    load_excel(select_list)
          
    
    
        

def main(t,select_WP):
    global ini_file,load_ws,new_hash
    if t == 0:
        fileEx = r'.ini'
        ini_file = [os.path.join(file) for file in os.listdir() if file.endswith(fileEx)]

        print('\n\n\nini파일목록: \n')
        for i in range(len(ini_file)):print(f'{i}: {ini_file[i]}')
        print()
        if len(ini_file) > 1:ini_file = ini_file[int(input('바꿀 파일 선택(0부터 시작, 숫자만)ex: 0\n: '))]
        else:ini_file = ini_file[0]

        print(f'\n\n무기를 선택해주세요\n0: 한손검    1: 장병기    2: 활    3: 대검    4: 법구\n숫자만 입력하세요,ex)0 -> 한손검\n: ',end='')
        WP_change(int(input()))
    else:
        char = ''
        a,b,num = 3,2,1
        with open(ini_file, "r", encoding="utf-8") as f:
            backup = open(f'백업_{ini_file[:-4]}.txt','w')
            lines = f.readlines()
            result = ''
            for i in range(len(lines)):result += lines[i]
            backup.write(result)
            backup.close()

            mod_data =[0,{'type':'','hash':''}]
            for i,l in enumerate(lines):
                if '[TextureOverride' in lines[i] and not '$active = 1' in lines[i+2]:
                    if mod_data[0] == 0:
                        if 'IB' in lines[i] or 'Head' in lines[i]:mod_data[1]['type'],char = 'ib','D'
                        elif 'Position' in lines[i]:mod_data[1]['type'],char = 'position_vb','E'
                        elif 'Blend' in lines[i]:mod_data[1]['type'],char = 'blend_vb','F'
                        elif 'Texcoord' in lines[i]:mod_data[1]['type'],char = 'texcoord_vb','G'
                        else:mod_data[1]['type'],char = 'draw_vb','C'
                        mod_data[1]['hash'] = lines[i+1]
                        mod_data[0] += 1
                        mod_data.append(2)
                        mod_data.append({'type':'','hash':''})
                    else:
                        num = a
                        if 'IB' in lines[i] or 'Head' in lines[i]:mod_data[num]['type'],char = 'ib','D'
                        elif 'Position' in lines[i]:mod_data[num]['type'],char = 'position_vb','E'
                        elif 'Blend' in lines[i]:mod_data[num]['type'],char = 'blend_vb','F'
                        elif 'Texcoord' in lines[i]:mod_data[num]['type'],char = 'texcoord_vb','G'
                        else:mod_data[num]['type'],char = 'draw_vb','C'

                        mod_data[num]['hash'] = lines[i+1]
                        a += 2
                        b += 1  
                        mod_data.append(b)
                        mod_data.append({'type':'','hash':''})
                    new_hash = load_ws[f'{char}{select_WP}'].value

                    if 'VertexLimitRaise' in lines[i]:
                            new_hash = mod_data[num]['hash'][7:]
                    #아직 엑셀파일에 없는 무기로 바꾸려할때 if문, 아니면 else문으로
                    if new_hash == '':result = result.replace(f"{mod_data[num]['hash']}",f";해쉬값 존재하지 않음\nhash = {mod_data[num]['hash'][7:]}\n")
                    else:result = result.replace(f"{mod_data[num]['hash']}",f";이전해쉬: {mod_data[num]['hash'][7:]}hash = {new_hash}\n")
                    mod_data[num]['hash'] = ''
                    
        with open(ini_file, "w", encoding="utf-8") as f:
            f.write('')
            f.write(result)
            print(f"{load_ws[f'B{select_WP}'].value}의 해쉬값으로 교체 완료")

if __name__ == "__main__":
    main(0,'')