import openpyxl

row_length = 474 #주소록의 행 방향 길이
name_column = 'B' #주소록의 이름 열
number_column = 'A' #주소록의 학번 열
Phone_column = 'C' #주소록의 전화번호 열
mail_column = 'G' #주소록의 이메일 열
write_column = 'M' # 참석 유무 기록 열

Survey_row_length = 53 #설문 조사의 행 방향 길이
Survey_name_column = 'D' #총회 참석 이름 열
Survey_number_column = 'E' #총회 참석 학번 열
Survey_column = 'B' #총회 참석 유무 열

# 엑셀 파일 불러오기
address_Workbook = openpyxl.load_workbook("C:/Users/--/Downloads/MR 주소록.xlsx")
Survey_results = openpyxl.load_workbook("C:/Users/--/Downloads/2023 MR 총회 참석 여부 조사 설문(응답).xlsx")
#시트 불러오기
address_sheet = address_Workbook['시트1']
Survey_results_sheet = Survey_results['시트1']

class Hash_table:
    def __init__(self, length = 5):
        self.max_len = length
        self.table = [[] for _ in range(self.max_len)]

    def hash(self, key):
        res = sum([ord(s) for s in key])
        return res % self.max_len

    def set(self, key, value): #해시 테이블에 key와 value를 넣는다.
        index = self.hash(key)
        self.table[index].append((key, value))

    def get(self, key):      #해시 테이블에서 key의 value를 찾는다.
        index = self.hash(key)
        value = self.table[index]
        if not value:         #찾는 키가 없으면 None을 반환
            return None
        for v in value:       #리스트에서 값을 찾아 반환 
            if v[0] == key:
                return v[1]
        return None
    
    def update(self, key, new_value):
        index = self.hash(key)
        values = self.table[index]
        if not values:
            return None
        for i, v in enumerate(values):
            if v[0] == key:
                self.table[index][i] = (key, new_value)
                return
        return None

hash_table = Hash_table(row_length)

#----------------------------------------------------------------------------------------#

#해시 테이블 입력
for row in range(2,row_length+1):
    key = address_sheet[name_column+str(row)].value + str(int(address_sheet[number_column+str(row)].value)) #이름+학번이 키 값(ex:임준범22)
    hash_table.set(key, row)
    if address_sheet[Phone_column+str(row)].value == None and address_sheet[mail_column+str(row)].value == None: #연락 못할 때
        address_sheet[write_column+str(row)] = '?'

#설문 조사 결과 엑셀 파일에 입력
address_sheet[write_column+str(1)] = "참여 여부"
for i in range(2, Survey_row_length+1):
    key = Survey_results_sheet[Survey_name_column+str(i)].value + str(int(Survey_results_sheet[Survey_number_column+str(i)].value))
    row = hash_table.get(key)
    if (row != None):
        if Survey_results_sheet[Survey_column+str(i)].value == '예':
            address_sheet[write_column+str(row)] = 'O'
        elif Survey_results_sheet[Survey_column+str(i)].value == '아니요':
            address_sheet[write_column+str(row)] = 'X'
    else:
        print(Survey_results_sheet[Survey_name_column+str(i)].value)
address_Workbook.save("C:/Users/--/Downloads/MR 주소록.xlsx")
