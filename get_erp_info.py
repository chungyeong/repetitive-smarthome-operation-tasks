import pandas as pd

"""
Date : 2021/09/13 
Author : 이충영
Description :  erp 양식에서 필요한 정보(동, 호, 성명, 전화번호) 정리 후 excel파일로 추출 
"""
#설정 값 : 실행 전 변경 필수!
excel_file = '22.02.25.xls'  # erp excel 파일 위치 -> 수동으로 입력 필요
startPoint = 0 #숫자만 입력, 따움표 사용 X
rawDongName='동'
rawHoName='호'
rawName='성명'
rawPhoneNumPrimaryName='휴대폰'
rawPhoneNumSubName='집전화'
rawHostRel='소유주'
exportExcelName="export.xlsx"

# 이하 프로그램 코드 변경 X

erp_excel = pd.read_excel(excel_file, header=startPoint)  # 동, 호, 성명 등 시작점
rawColumns = [rawDongName, rawHoName, rawName, rawPhoneNumPrimaryName, rawPhoneNumSubName, rawHostRel]
print(erp_excel.columns.tolist())

raw_data = erp_excel.loc[:,rawColumns]  # 필요한 Column 선택
raw_data.rename(columns={rawDongName:'동',rawHoName:"호",rawName:"성명",rawPhoneNumPrimaryName:"휴대폰",rawPhoneNumSubName:"집전화",rawHostRel:"세대주관계"},inplace=True)
print(raw_data)

# erp 파일에 있는 동, 호, 성명, 휴대폰의 명칭 열 선택 및 삭제
category_index = raw_data[
    (raw_data['동'] == '동')
    & (raw_data['호'] == '호')
].index

category_del = raw_data.drop(category_index)

# erp 파일에서 비어 있는 열 삭제
blank_index = raw_data[
    # (raw_data['동'].isnull())
    # &(raw_data['호'].isnull())
    (raw_data['성명'].isnull())
    & (raw_data['휴대폰'].isnull())
].index
blank_category_del = category_del.drop(blank_index)

# index reset
reset_erp = blank_category_del.reset_index(drop=True)

# 병합 Cell 해제 후, 빈 값 보충
rows = len(reset_erp)
old_dong_value = reset_erp.loc[0, "동"]
old_ho_value = reset_erp.loc[0, "호"]
drop_index_list = []
for row in range(rows):
    # 동, 호수 빈칸 체우기
    new_dong_value = reset_erp.loc[row, "동"]

    if pd.isnull(new_dong_value):
        reset_erp.loc[row, "동"] = old_dong_value
    else:
        old_dong_value = new_dong_value

    new_ho_value = reset_erp.loc[row, "호"]

    if pd.isnull(new_ho_value):
        reset_erp.loc[row, "호"] = old_ho_value
    else:
        if type(new_ho_value) is str:
            if new_ho_value.count("-"):

                new_ho_value = new_ho_value.replace("-", "0")
                reset_erp.loc[row, "호"] = new_ho_value

                drop_index_list.append(row)

        old_ho_value = new_ho_value

    if reset_erp.loc[row, "동"] in [9999, 999, 900]:  # 9999동 등 Test 동 삭제
        drop_index_list.append(row)

    if pd.isnull(reset_erp.loc[row, "세대주관계"]):  # 세대주 관계 비어 있으면 정보없음으로 표기
        reset_erp.loc[row, "세대주관계"] = "정보없음"

    # 세대주관계 특수문자 있는 열 제거
    if any(sym in str(reset_erp.at[row, '세대주관계']) for sym in '.!@#$%^&*():'):
        drop_index_list.append(row)
    try:
    # ERP에 휴대폰 번호가 없으면 집전화에 있는 정보로 대체
        if pd.isnull(reset_erp.at[row, "휴대폰"]):  
            reset_erp.at[row, "휴대폰"] = reset_erp.at[row, "집전화"]
    except KeyError:
        continue

    if pd.isnull(reset_erp.at[row, "성명"]):  # 이름 빈 값 삭제
        drop_index_list.append(row)

    # 이름에 특수문자가 있으면 삭제
    if any(sym in str(reset_erp.at[row, '성명']) for sym in '.!@#$%^&*():-'):
        drop_index_list.append(row)

    if len(str(reset_erp.at[row, "성명"])) >= 20:
        reset_erp.at[row, "성명"] = reset_erp.at[row,
                                               "성명"][0:20]  # 이름 20자 이내로 수정

    phone_num = reset_erp.at[row, '휴대폰']
    str_phone_num = str(phone_num)

    if str_phone_num.isnumeric() == False:

        if str_phone_num.find('-') != -1:
            reset_erp.at[row, '휴대폰'] = str_phone_num.replace('-', '')  # 하이픈 삭제
            str_phone_num = reset_erp.at[row, '휴대폰']

            if str_phone_num.isnumeric() == False:  # 문자 포함 시, 핸드폰 번호 삭제
                drop_index_list.append(row)

    if str_phone_num != "nan":
        if len(str_phone_num) != 11:  # 핸드폰 번호 길이 오류, 공백으로 변경
            reset_erp.at[row, '휴대폰'] = ""

    if str_phone_num[0:3] != "010":  # 핸드폰 번호가 아닌 다른 번호 일 경우 공백으로 변경
        reset_erp.at[row, '휴대폰'] = ""

reset_erp = reset_erp.drop(drop_index_list)
reset_erp = reset_erp.loc[:, ['동', '호', '성명', '세대주관계', '휴대폰']]
print(reset_erp)
writer = pd.ExcelWriter(
    exportExcelName)  # pylint: disable=abstract-class-instantiated
reset_erp.to_excel(writer, index=False)
writer.close()
print("끄으으읕")
