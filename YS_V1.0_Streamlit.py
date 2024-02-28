import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from io import BytesIO

###################################################################################################################
st.title('양곡관리 입금내역 매칭 서비스 V1.0')
st.markdown(":green[*** 농협 인터넷뱅킹 입출금 엑셀파일과 비교하여 '본인부담금 입금여부'를 확인할 수 있습니다. ***]")

current_date = datetime.now() # 현재 날짜와 월을 표현합니다.
current_month = current_date.strftime('%Y년 %m월')  # 예: '2024년 02월'

###################################################################################################################

# 양곡 입력 데이터프레임 생성 및 엑셀 다운로드 
st.markdown("ㅇ 아래 '서식'을 다운로드 받아 '양곡배부 대상자'를 작성하시기 바랍니다.")
data = {
    '연번': [1,2,3],
    '구분': ['기초수급','생계급여','주거급여'],
    '성명': ['제이홉', '진', '슈가',],
    '시군구': ['양산시', '양산시', '양산시',],
    '행정동': ['물금읍', '동면','원동면'],
    '주소': ['경상남도 양산시 물금읍', '경상남도 양산시 동면','경상남도 양산시 원동면'],
    '세부주소': ['', '',''],
    '휴대전화번호': ['010-2233-4433', '010-2390-1234','010-3222-3333'],
    '자택전화번호': ['055-392-2222', '055-392-2224','055-392-2221'],
    '양곡수량': [1, 2, 3],
    '생년월일': ['1981-09-15', '1948-05-21', '1951-01-13'],
    '문자수신여부': ['Y', 'N', 'Y'],
    '가구원수(명)': [1,2,3],
    '본인부담금액(원)': [10000, 8000, 2000],}


dataframe = pd.DataFrame(data)

# 데이터프레임을 스트림릿 앱에 표시 (st.write 사용)

if 'dataframe' in locals():
    st.write(dataframe)
else:
    st.error("dataframe is not defined.") 

#데이터프레임을 엑셀 파일로 변환하는 함수 
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', mode='wb') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    processed_data = output.getvalue()
    return processed_data
#st.caption('*** 다운로드 버튼을 눌러 서식을 다운로드 받아, 등록 서식을 준수하여 작성하시기 바랍니다. ***')

# 다운로드 버튼 생성
excel_data = convert_df_to_excel(dataframe)  # 데이터프레임을 엑셀 데이터로 변환
st.download_button(label="📥 서식 다운로드(Excel)",
                   data=excel_data,
                   file_name=f'{current_month} 양곡관리 입금대상.xlsx',
                   mime='Excel')
st.caption('* 등록되는 개인정보는 수집되지 않으며, 개인정보는 해당 PC에 등록되며 프로그램 종료시 삭제됩니다. ')
st.markdown('<br>', unsafe_allow_html=True) #화면 공백 띄우기
###################################################################################################################

# 양곡파일 대상자 명단 업로드

st.markdown("ㅇ 파일 등록을 통해 양곡관리 입금여부를 확인하실 수 있습니다.")

file_uploader_label1 = '가. 양곡배부 대상자 명단을 등록하세요'
file1 = st.file_uploader(file_uploader_label1, type=['xls', 'xlsx'], key="file1")

if file1:
    대상자 = pd.read_excel(file1, sheet_name=0, dtype={
                                                        '생년월일': str,
                                                        '휴대전화번호': str,
                                                        '자택전화번호': str,                              
                                                      }) 
    if '문자수신여부' in 대상자.columns:
        대상자['문자수신여부'] = 대상자['문자수신여부'].str.upper()

    st.dataframe(대상자, use_container_width=True)

###################################################################################################################
# 농협 입출금내역 업로드
file_uploader_label2 = '나. 농협 거래내역 파일을 등록하세요(Excel)'
file2 = st.file_uploader(file_uploader_label2, type=['xls', 'xlsx'], key="file2")


# 입금자 대상자 데이터 처리
if file2 is not None:
    try:
        농협입금데이터 = pd.read_excel(file2, skiprows=[0,1,2,3,4,5,6,7,8])
        
        # 한글만 추출하는 함수
        def extract_korean(text): 
          return re.sub('[^가-힣]', '', text)

        # '거래기록사항' 컬럼에서 한글만 추출하여 새로운 '성명' 컬럼에 저장
        농협입금데이터['성명'] = 농협입금데이터['거래기록사항'].apply(extract_korean)
        
        농협추출데이터 = 농협입금데이터[['거래일자', '입금금액(원)', '성명','거래기록사항', '거래점']]
        #st.dataframe(농협추출데이터, use_container_width=True)
        # '성명' 컬럼이 대상자 데이터프레임에 있는지 확인
        print("대상자 컬럼:", 대상자.columns)

        # '성명' 컬럼이 농협추출데이터 데이터프레임에 있는지 확인
        print("농협추출데이터 컬럼:", 농협추출데이터.columns)
    except Exception as e:
        st.error(f"서식 또는 농협 거래내역 파일을 등록하시기 바랍니다.")

###################################################################################################################
# 파일 머지 및 검증
if '농협추출데이터' in locals():

   # 파일 머지 및 검증
    if file1 is not None and file2 is not None:
        결과 = pd.merge(대상자, 농협추출데이터, on='성명', how='outer')

        def compare_amounts(row):
            if pd.isna(row['본인부담금액(원)']) or pd.isna(row['입금금액(원)']):
                return '입금요청 및 신규 검토 대상'
            elif row['본인부담금액(원)'] == row['입금금액(원)']:
               return '정상'
            else:
                return '금액 확인필요'

        결과['검증결과'] = 결과.apply(compare_amounts, axis=1)
        
        정렬된_데이터프레임 = 결과[[
            '연번','검증결과', '구분','성명','시군구','행정동','주소','세부주소','휴대전화번호','자택전화번호','양곡수량','생년월일','문자수신여부','가구원수(명)','본인부담금액(원)',
            '입금금액(원)','거래점','거래일자','거래기록사항'
            ]]

###################################################################################################################
# 순서정렬
             
        검증결과_순서 = ['금액 확인필요','정상','입금요청 및 신규 검토 대상',] #사용자 지정 순서 정의

        # '검증결과' 열을 Categorical 타입으로 변환하고 사용자 지정 순서 적용
        정렬된_데이터프레임['검증결과'] = pd.Categorical(정렬된_데이터프레임['검증결과'], categories=검증결과_순서, ordered=True)

        # '검증결과'를 기준으로 데이터프레임 정렬
        정렬된_데이터프레임 = 정렬된_데이터프레임.sort_values(by=['검증결과','연번'])

        # 정렬된 데이터프레임을 Streamlit 앱에 표시
        st.dataframe(정렬된_데이터프레임, use_container_width=True)
        
###################################################################################################################
 
# 정상 납부자 현황       
        정상_결과 = 정렬된_데이터프레임[정렬된_데이터프레임['검증결과'] == '정상'] # '정상'으로 판별된 데이터만 필터링
        정상_납부자수 = 정상_결과['성명'].count() # '정상' 납부자 수 합계 계산
        정상_입금금액_합계 = 정상_결과['입금금액(원)'].sum() # '입금금액(원)' 합계 계산
        # 계산된 합계를 한 줄로 표현
        st.write(f" 검증결과 '정상' 납부자 수: {정상_납부자수}명, 입금금액 합계: {정상_입금금액_합계:,}원")
 
# 확인필요 납부자 현황        
        확인필요_결과 = 정렬된_데이터프레임[정렬된_데이터프레임['검증결과'] == '금액 확인필요'] # '정상'으로 판별된 데이터만 필터링
        확인필요_납부자수 = 확인필요_결과['성명'].count() # '정상' 납부자 수 합계 계산
        확인필요_입금금액_합계 = 확인필요_결과['입금금액(원)'].sum() # '입금금액(원)' 합계 계산
        # 계산된 합계를 한 줄로 표현
        st.write(f" 검증결과 '금액 확인필요' 납부자 수: {확인필요_납부자수}명, 확인금액 합계: {확인필요_입금금액_합계:,}원")        

###################################################################################################################
# 엑셀파일
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            정렬된_데이터프레임.to_excel(writer, index=False, sheet_name='Sheet1')
            
        excel_data = output.getvalue()

# 현재 날짜와 시간을 파일 이름에 포함
        from datetime import datetime
        current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"검증결과_{current_datetime}.xlsx"

# 다운로드 버튼 생성
        st.download_button(
            label="📥 검증결과 다운로드(Excel)",
            data=excel_data,
            file_name=file_name,
            mime='Excel'
        )