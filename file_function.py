#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import shutil
import time
import pandas as pd
from tqdm.notebook import tqdm
import re
from pathlib import Path
import traceback

#########################################
# 채무자조회.xlsx -> dict
#########################################
def debtorInfoDict(path: str):
    """
    채무자조회.xlsx파일을 읽어 채무자키를 key로 하고 
    나머지컬럼을 value로 하는 dict반환
    dict["키"].컬럼명 으로 읽으면 편함
    """
    df_c = pd.read_excel(path)
    dict = {}
    for key, row in df_c.iterrows():
        dict[str(row.채무자키)] = row[1:]
    # row는 시리즈. 채무자키는 dict의 키로 넣었으니 row[1:]을 value로 넣자.
    # 시리지는 .컬럼명으로 읽으면 되니까 최종적으로
    # dict["20495151"].성명 이렇게 읽으면 된다.

    # 2차원 딕셔너리는 row[1:] 대신 아래를 사용
    # {'매각사구분':row.매각사구분, '성명' : row.성명, \
        # '주민번호인':row.주민번호인, '관리자기타':str(row.관리자기타), '보증인성명':row.보증인성명}
    return dict


#########################################
# rename
#########################################
def rename(src:str, dst:str)->None :
    """
    전체경로(파일명포함) 두개를 받아서 파일 이름 바꾸는 함수
    동일파일이 있는 경우, 넘버링
    """
    # dst dir과 file 분리하기
    # file 에서 일련번호 분리하기
    # while - 일련번호 다시 붙이기

    dir = os.path.split(dst)[0]
    f_name = os.path.split(dst)[1]
    stem = os.path.splitext(f_name)[0]
    ext = os.path.splitext(f_name)[1]
    
    temp = re.sub("[^가-힣]+$", "", stem)
    new_name = temp + ext
    
    i = 1
    while os.path.exists(dir+"/"+new_name):  # 작업디렉토리가 아니므로 풀경로
        new_name = temp + "_"+"("+str(i)+")"+ext
        i += 1

    dst_final = dir + "/" + new_name
    os.rename(src, dst_final)

    print(f, new_name)




#########################################
# 생성일, 수정일 보기 + time을 보기 좋게
#########################################


def c_time(path):
    a = time.ctime(os.path.getctime(path))
    b = time.strptime(a)
    c = time.strftime('%Y.%m.%d - %H:%M:%S', b)
    return c


def m_time(path):
    a = time.ctime(os.path.getmtime(path))
    b = time.strptime(a)
    c = time.strftime('%Y.%m.%d - %H:%M:%S', b)
    return c

#########################################
# 모든 파일의 정보를 추출하여 csv로 저장, df 반환
#########################################


def all_files(path, local):
    os.chdir(path)

    z = "/"
    f_name_dict = {}  # 중복파일명 숫자를 카운트 할 딕셔너리
    f_dir = []
    f_name = []
    name_count = []
    f_size = []
    create_time = []
    modify_time = []
    extension = []

    for root, dirs, files in tqdm(os.walk(path)):

        for f in files:
            path = root + z + f

            if f not in f_name:
                f_name_dict[f] = 1
            else:
                f_name_dict[f] = f_name_dict[f]+1

            f_dir.append(root)
            f_name.append(f)
            name_count.append(f_name_dict[f])
            f_size.append(os.path.getsize(path))
            create_time.append(c_time(path))
            modify_time.append(m_time(path))
            extension.append(os.path.splitext(f)[1])

    # 매각사 칼럼 추가하기
    import re
    df_matching = pd.read_excel(
        r'D:\전산\workspace\python_work\파일\매각사 이름매칭.xlsx')
    sell = []

    for i in f_dir:
        for index, row in df_matching.iterrows():
            if re.search(row[0], i):
                sell.append(row[1])
                continue

    df = pd.DataFrame({'경로': f_dir, '매각사': sell, '파일명': f_name, '중복수': name_count,
                      '크기': f_size, '생성일': create_time, '수정일': modify_time, '확장자': extension})
    df.to_csv(r'C:\Users\SL\Desktop/'+local +
              ' 모든 파일 정보.csv', encoding='utf-8-sig')

    return df


#########################################
# 공백을 _로 바꾸고 현재경로에 그대로 저장
#########################################
def change_spaceTo_(s):
    os.chdir(s)

    for root, __dirs__, files in tqdm(os.walk(s)):

        for f in files:
            f_s = root + "\\" + f
            f_d = root + "\\" + f.strip().replace(" ", "_")
            os.rename(f_s, f_d)


#########################################
# 보증인 끝으로 이동
#########################################
def change_guarantee(f_s):
    os.chdir(f_s)
    word = '보증인'
    index = 0
    for root, __dirs__, files in os.walk(f_s):
        for f in files:
            if re.search(word, f):
                f_d = re.sub(word, "", f)
                f_d = os.path.splitext(
                    f_d)[0]+"_"+word+os.path.splitext(f_d)[1]
                # f_d = re.sub('_{2,}', '_', f_d) 언더바 제거 따로 일괄작업

                # walk를 쓸 때는 언제나 root와 합쳐줘야. cwd와 다르니까!
                f_s = root + "\\" + f
                f_d = root + "\\" + f_d
                os.rename(f_s, f_d)

                index += 1
    print(str(index)+"건의 보증인 파일 이름 수정 완료")


#########################################
# 연속된 _를 제거
#########################################
def change__(s):
    os.chdir(s)
    import re

    for root, __dirs__, files in tqdm(os.walk(s)):

        for f in files:
            f_a = re.sub('_{2,}', '_', f)

            f_s = root + "\\" + f
            f_d = root + "\\" + f_a
            os.rename(f_s, f_d)


#########################################
# pdf류 아닌 파일 모두 이동시키기
#########################################

def not_pdf(file):  # 문자변환 여부 주의
    df_s = pd.read_csv(file)

    # pdf류 확장자 리스트
    extension = ['.jpeg', '.jpg', '.bmp', '.gif', '.pdf', '.png', '.tif']

    # 모든 파일 정보에서 확장자 대소문자 구분없이 비교하여 해당사항이 없으면 이동시키기
    # 폴더트리까지 그대로 이동하므로 파일명 겹칠 걱정은 하지 않아도 됨

    for index, row in df_s.iterrows():
        ext_check = False  # 기본 옮길 대상인 pdf류가 아니라고 설정
        for ext in extension:
            if re.match(ext, row.확장자, re.I):
                ext_check = True  # pdf류이면 기본값 변경
                continue

        if ext_check == False:  # 기본값에 변경이 없다면 파일 이동
            path_s = row.경로+"/"+row.파일명
            ######### 다운로드 폴더가 아닌 곳에 있다면 여기 수정해줘야함 ##########
            ######### C:\Users\SL\Downloads\ ##########
            path_d_d = "d:/기타확장자/"+row.경로[22:]
            path_d_f = path_d_d + "/" + row.파일명

            try:
                shutil.move(path_s, path_d_f)
            except:
                if not os.path.exists(path_d_d):
                    os.makedirs(path_d_d)
                    shutil.move(path_s, path_d_f)


#########################################
# 사이즈 같은 파일들만 딕셔너리로
#########################################
def same_size(path):
    import copy
    os.chdir(path)
    file_list = os.listdir(path)
    dict_size = {}

    for f in file_list:
        size = os.path.getsize(f)
        if size not in dict_size:
            dict_size[size] = [f]
        else:
            dict_size[size].append(f)  # 일련번호를 key로 하는 딕셔너리 만들기

    dict_size_2 = copy.deepcopy(dict_size)

    for key, value in dict_size.items():
        if len(value) == 1:
            del dict_size_2[key]
        else:
            continue

    return dict_size_2


#########################################
# 다른 폴더로 이동
#########################################
def move(path, f_d):
    os.chdir(path)
    file_list = os.listdir(path)
    index = 0
    p1 = re.compile(r'[^가-힣]+$')

    for f in tqdm(file_list):
        if os.path.isfile(f) & (f != "Thumbs.db"):
            stem = os.path.splitext(f)[0]
            ext = os.path.splitext(f)[1]

            temp = p1.sub("", stem)
            new_name = temp + ext  # 넘버링 제외된 파일명 + 확장자 붙여서 비교

            i = 1
            while os.path.exists(f_d+"/"+new_name):  # 작업디렉토리가 아니므로 풀경로
                new_name = temp + "_"+"("+str(i)+")"+ext
                i += 1

            f_d_final = f_d + "/" + new_name
            shutil.move(f, f_d_final)
            index += 1

    print(index, "개 파일 이동 완료")
    os.chdir('c:/')


#########################################
# 최종확인1 - 자기 폴더에서 공백과 언더바 점검후 넘버링도 새롭게(마지막 숫자가 +될수도 있음)
#########################################

def final_rename(path):
    f_d = path
    os.chdir(path)
    file_list = os.listdir(path)
    index = 0
    changed = []
    error = []
    p0 = re.compile(r'\s')
    p1 = re.compile('_{2,}')
    p2 = re.compile('복사본')
    p3 = re.compile(r'[^가-힣]+$')

    docu_kind = '원인서류|채권양도통지서|판결문|지급명령|이행권고|화해권고|타채|결정문|등본|초본|등,초본|등초본|외국인|개회|신복|파산'
    etc_kind = '보증인|재도|1차|2차|3차|4차'
    p_key = re.compile("[0-9]{8}")
    p_docu = re.compile(docu_kind)
    p_etc = re.compile(etc_kind)

    try:
        for f in tqdm(file_list):

            if os.path.isfile(f) & (f != "Thumbs.db"):

                fullname = Path(Path.cwd() / f)
                temp = fullname.stem
                new_name = ""

                # 키 뒤에, 문서종류, 기타정보 앞에 _ 없는 경우 수정
                res_d = p_docu.search(f)
                res_e = p_etc.search(f)

                if p_key.match(temp):
                    if temp[8] != "_":  # key 뒤에 언더바 없는 경우
                        temp = temp[:8] + "_" + temp[8:]
                else:
                    error.append(f)

                if res_d != None:
                    if temp[res_d.start()-1] == "_":
                        pass
                    else:  # 문서종류 앞이 _가 아닌경우
                        temp = temp[:res_d.start()] + "_" + \
                            temp[res_d.start():]

                    # 키와 문서 종류 사이에 언더바가 여러개인경우
                    name_before = temp[9:res_d.start()-1]  # _namebefore_
                    name_after = re.sub("_", "", name_before)
                    if name_after != name_before:
                        temp = temp[:9] + name_after + temp[res_d.start()-1:]

                else:
                    error.append(f)

                if res_e != None:
                    if temp[res_e.start()-1] == "_":
                        pass
                    else:  # 기타키워드 앞이 _가 아닌경우
                        temp = temp[:res_e.start()] + "_" + \
                            temp[res_e.start():]

                # 공백
                temp = temp.strip()
                temp = p0.sub("_", temp)

                # 연속 언더바
                temp = p1.sub('_', temp)

                # 끝이 한글이 아니거나 '복사본'인 경우(즉 모든 넘버링 및 기호 지우기)
                temp = p2.sub("", temp)
                temp = p3.sub("", temp)

                new_name = temp + fullname.suffix  # 넘버링 제외된 파일명 + 확장자 붙여서 비교

                if new_name == f:  # 달라진게 없다면
                    pass
                else:

                    i = 1
                    while os.path.exists(f_d+"/"+new_name):  # 작업디렉토리가 아니므로 풀경로
                        new_name = temp + "_"+"("+str(i)+")"+fullname.suffix
                        i += 1

                    f_d_final = f_d + "/" + new_name
                    shutil.move(f, f_d_final)
                    changed.append([f, new_name])
                index += 1

            else:
                pass

    except:
        print(traceback.format_exc())
        print(index, "번째 파일까지 처리하고 에러")

    print(index, "개의 파일 이름 변경 완료")
    print("error : ", *error, sep="\n")
    print("파일명 변경 목록 : ", *changed, sep="\n")
    os.chdir('c:/')

#########################################
# 최종확인2 - 마지막으로 3개 양식과 언버가 개수 확인
#########################################

# 등,초본이 초본을 포함해버린다. 이거 수정해야


def final_check(path):
    os.chdir(path)
    file_list = os.listdir(path)
    lista = []  # 메인 3개 항목 중 문제 발생
    listb = []  # 언더바 개수에 이상
    index = 0
    docu_kind = '원인서류|채권양도통지서|판결문|지급명령|이행권고|화해권고|타채|결정문|등본|초본|개회|신복|파산|외국인증명'
    p0 = re.compile(
        r'[\d]{8}_[a-zA-Z가-힣,()]+_('+docu_kind+')')
    p_ = re.compile('_')

    # try:
    for f in tqdm(file_list):

        if os.path.isfile(f) & (f != "Thumbs.db"):

            fullname = Path(Path.cwd() / f)
            temp = fullname.stem

            cond1 = (p0.match(temp) == None)
            num_ = len(p_.findall(temp))
            cond2 = (num_ < 2) | (num_ > 4)

            if cond1:
                lista.append(f)
                index += 1
            else:
                pass

            if cond2:
                listb.append(f)
                index += 1
            else:
                pass

    print(index, "개의 이상 탐색")
    print("메인3항목중 이상 발견", *lista, sep="\n")
    print("언더바 개수 이상 발견", *listb, sep="\n")
    os.chdir('c:/')

#########################################
# 파일명 구분기호별 분류하기
#########################################
# def filename_split(path) :
#   all = os.listdir(path)
#   files = [Path(path +'/'+x).stem for x in all if os.path.isfile(path +'/'+x)]
#           ########################확장자 포함하려면 이걸 그냥 x로
#   #files = [Path(path +'/'+x).stem for x in all if os.path.isfile(path +'/'+x)]
#   f_split_list = [f.split('_') for f in files]
#   return f_split_list
