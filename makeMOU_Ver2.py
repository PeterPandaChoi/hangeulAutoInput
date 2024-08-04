#사전 준비사항, 아래 모듈 설치 및 한글세큐리티 모듈 등록
#pip install pywin32
#pip install pandas
#pip install openpyxl
#세큐리티 모듈 등록 필요

#모듈 임포트
import win32com.client as win32
import pandas as pd
import os

#파일 이름과 디렉토리 변수저장
BASE_DIR= os.path.dirname(__file__)
thisProgramName=os.path.basename(__file__)

#엑셀 불러오기
excel = pd.read_excel('MOU_form_content.xlsx')

#한글 열기 + 모듈 연결
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule('FilePathCheckDLL','SecurityModule')


while True:
  form = input("form의 양식을 선택해주세요. (1 또는 2), 종료(q)  :  ")
  
  if form=='1':
    hwp.Open(BASE_DIR+'\\MOU_form1.hwp',Format="HWP",arg="")
    break
  elif form=='2':
    hwp.Open(BASE_DIR+'\\MOU_form2.hwp',Format="HWP",arg="")
    break
  elif form=='3':
    hwp.Open(BASE_DIR+'\\MOU_form3.hwp',Format="HWP",arg="")
    break
  elif form=='q':
    exit()
  else:
    print("그런 양식은 없습니다. 다시 입력하세yo")

    
  

#한글에서 쓰이는 필드리스트 받아오기
hwp.GetFieldList()
hwp.GetFieldList(1)
hwp.GetFieldList(2)
field_list = [i for i in hwp.GetFieldList().split('\x02')]

#한글에서 매크로 실행 - 전체 선택 - 복사 - 맨 뒤로 커서 이동
hwp.Run('SelectAll')
hwp.Run('Copy')
hwp.MovePos(3)

#엑셀의 행 만큼 반복 - 붙여넣기 - 맨 뒤로 커서 이동
for i in range(len(excel)-1):
  hwp.Run('Paste')
  hwp.MovePos(3)

#각 페이지의 필드 값 바꿔주기 
for page in range(len(excel)):
  for field in field_list:
    hwp.PutFieldText(f'{field}{{{{{page}}}}}',excel[field].iloc[page])




#다른이름으로 저장!
hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = BASE_DIR+'\\MOU_form_revised'+str(form)+'.hwp'
hwp.HParameterSet.HFileOpenSave.Format = "HWP"
hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.Quit()

#PDF로 저장하기는 도장을 넣어야 되어서 비활성화함, 작동안됨
'''
hwp.HAction.GetDefault('FileSaveAsPdf',hwp.HParameterSet.HFileOpenSave.HSet)
hwp.HParameterSet.HFileOpenSave.filename = BASE_DIR+'MOU_form_revised.pdf'
hwp.HParameterSet.HFileOpenSave.Format = 'PDF'
hwp.HAction.Execute("FileSaveAsPdf",hwp.HParameterSet.HFileOpenSave.HSet)
'''