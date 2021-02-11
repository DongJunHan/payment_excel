from openpyxl import load_workbook


#시트 로드. data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
load_wb = load_workbook("./한동준_국민체크8906.xlsx",data_only=True)

#시트 이름으로 불러오기
load_ws = load_wb['Sheet1']

#A2 cell의 값을 가져오기.
print(load_ws['A2'].value)
