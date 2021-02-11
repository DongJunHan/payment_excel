from openpyxl import load_workbook
# from openpyxl import Workbook

path = "./한동준_국민체크8906.xlsx"
write_path = "./OOO_법인카드지출내역서1.xlsx"
def get_excel_data(path):
    #시트 로드. data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
    try:
        load_wb = load_workbook(path,data_only=True)
    except FileNotFoundError:
        print("get_excel_data There is no file")
        return

    #시트 이름으로 불러오기
    load_ws = load_wb['Sheet1']

    #cell의 데이터 가져오기.
    payment_list =[]
    for row in load_ws.rows:
        row_list = []
        for cell in row:
            row_list.append(cell.value)
        payment_list.append(row_list)
    return payment_list

def parsing_excel_data(data_list):
    cell_list = []
    for data in data_list:
        date = data[0]
        if date != "승인일":
            cell = []
            cell.append(data[0]) # date
            cell.append(data[1]) # time 내용을 점심식사, 저녁식사구분하기 위함.
            cell.append(data[6]) # shop name
            cell.append(data[7]) # shop 식대인지, 택시인지 구분하기 위함. TODO. 현재는 식대랑, 택시만 사용할거같지만 추후에 어떤것이 더 생길지 모름
            cell.append(data[10]) # price
            cell_list.append(cell)
    return cell_list

def write_excel(path,data_list):
    try:
        load_wb = load_workbook(path,data_only=True)
    except FileNotFoundError:
        print("write_excel There is no file")
        return
    sheet1 = load_wb.active

    # write_wb = Workbook()

data_list = get_excel_data(path)
parse_data = parsing_excel_data(data_list)
print(parse_data)
write_excel(write_path,parse_data)