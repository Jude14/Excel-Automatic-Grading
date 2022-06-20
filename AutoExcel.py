from openpyxl import Workbook, workbook,load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font 

data = {
	"Joe": {
		"math": 65,
		"science": 78,
		"english": 98,
		"gym": 89
	},
	"Bill": {
		"math": 55,
		"science": 72,
		"english": 87,
		"gym": 95
	},
	"Tim": {
		"math": 100,
		"science": 45,
		"english": 75,
		"gym": 92
	},
	"Sally": {
		"math": 30,
		"science": 25,
		"english": 45,
		"gym": 100
	},
	"Jane": {
		"math": 100,
		"science": 100,
		"english": 100,
		"gym": 60
	}
}

wb = Workbook()
ws = wb.active
ws.title="Grades"

headings =  ['Name']+ list(data['Joe'].keys())+['std grd']
ws.append(headings)

print()
for person in data:
    grades = list(data[person].values())
    ws.append([person]+ grades)
ws.append(['subject avg'])

for col in range(2,len(data['Joe'])+2):
    char = get_column_letter(col)
    ws[char+str(len(data)+2)]= f"=SUM({char+'2'}:{char+str(len(data['Joe'])+2)})/{len(data)}"

for row in range(2,len(data)+2):
    ws[get_column_letter(len(data['Joe'])+2)+str(row)]=f"=SUM({'B'+str(row)}:{'E'+str(row)})/{len(data['Joe'])}"

for col in range(1, len(data)+2):
	ws[get_column_letter(col) + '1'].font = Font(bold=True, color="0099CCFF")

wb.save("NewGrades.xlsx") 










