import re
import xlsxwriter

LINE_COUNT = 33		# 33 lines, 34+ including top
ALPHABETIZE_STUDENTS = True
as_sheets = True

class_list = []

# read classes
with open("classes.txt", 'r') as f:
	curr = []
	for line in f:
		if line.find("PAGE") != -1:
			# if copy-pasted output, skip these lines
			continue
		if line == "\n":
			if len(curr) > 0:
				if curr[0].find('T ') != -1:
					class_list[-1] += curr
				else:
					class_list.append(curr)
			curr = []
		else:
			if line[-1] != '\n':
				line = line + '\n'
			curr.append(line)
if len(curr) > 0:
	class_list.append(curr)

if ALPHABETIZE_STUDENTS:
	for i in range(len(class_list)):
		# ignore first two and all teachers
		students_end = len(class_list[i])-1
		while class_list[i][students_end][0:2] == 'T ':		# teachers should go at the end no matter what
			students_end -= 1
		
		students = class_list[i][2:students_end+1]
		students.sort()
		class_list[i][2:students_end+1] = students
		class_list[i].insert(students_end+1, "\n")			# add an extra line to account for new students

DAYS_ORDER = {"mon": -3, "wed" : -2, "fri" : -1, "sat": 0, "sun": 1}

filtered_classes = []
for c in class_list:
	first_line = c[1].strip() if len(c) > 1 else ""
	for day, rank in DAYS_ORDER.items():
		if day in first_line.lower():
			filtered_classes.append(c)
			break

def get_day_and_time_order(c):
	first_line = c[1].strip() if len(c) > 1 else ""
	
	day_rank = 99
	for day, rank in DAYS_ORDER.items():
		if day in first_line.lower():
			day_rank = rank
			break
			
	time_match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*([aApP][mM])', first_line)
	if time_match:
		hours = int(time_match.group(1))
		minutes = int(time_match.group(2)) if time_match.group(2) else 0
		period = time_match.group(3)
		
		if period.lower() == 'pm' and hours != 12:
			hours += 12
		elif period.lower() == 'am' and hours == 12:
			hours = 0
				
		time_rank = hours * 60 + minutes
	else:
		time_rank = 9999
		
	return (day_rank, time_rank)

filtered_classes.sort(key=get_day_and_time_order)
print(filtered_classes)
pages = []
current_page = []
current_page_lines = 0

for c in filtered_classes:
	class_len = len(c)
	spacing = 1 if current_page_lines > 0 else 0
	total_needed = class_len + spacing
	
	if current_page_lines + total_needed <= LINE_COUNT:
		if spacing > 0:
			current_page.append(["\n"])
		current_page.append(c)
		current_page_lines += total_needed
	else:
		if current_page:
			pages.append(current_page)
		current_page = [c]
		current_page_lines = class_len

if current_page:
	pages.append(current_page)

if not as_sheets:
	with open("output.txt", 'w') as f:
		for page_num, page in enumerate(pages, 1):
			f.write(f"--- PAGE {page_num} ---\n")
			for block in page:
				for line in block:
					f.write(line)
			f.write("\n\n")
	print("Wrote to", "\"output.txt\"")
else:
	# format as xlsx
	workbook = xlsxwriter.Workbook("attendance log.xlsx")
	worksheet = workbook.add_worksheet()

	for page_num, page in enumerate(pages, 1):
		worksheet.set_column_pixels((page_num-1)*2, (page_num-1)*2, 150)
		worksheet.set_column_pixels((page_num-1)*2+1, (page_num-1)*2+1, 50)		# attempt to reset xlsx to default column size
		row = 0
		for block in page:
			for line in block:
				worksheet.write(row, (page_num-1)*2, line)
				worksheet.set_row_pixels(row=row, height=21)										# attempt to reset xlsx to default row size
				row += 1

	cell_format = workbook.add_format()
	cell_format.set_bottom(2)
	worksheet.set_row_pixels(row=32, height=21, cell_format=cell_format)

	workbook.close()
	print("Wrote to", "\"attendance log.xlsx\"")
