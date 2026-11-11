def has_numbers(inputString):
	return any(char.isdigit() for char in inputString)

with open("classes.txt", 'r') as f:
	curr = []
	for line in f:
		if line != '\n' and not has_numbers(line):
			if line.find('T ') != -1:
				line = line[0:4] + '\n'
			else:
				line = line[0:2] + '\n'
		curr.append(line)
	f.close()
with open("o.txt", 'w') as f:
	f.writelines(curr)