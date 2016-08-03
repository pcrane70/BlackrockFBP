# Check the lengths of both files

def check_file(f_name):

	with open(f_name, 'r') as f:
		amount = str(f.read()).count('first_name')

	return amount

file1 = 'C:/Users/muizyusuff/Desktop/dev/blackrock/BlackrockFBP/blackrock/compiled-data.txt'

file2 = 'C:/Users/muizyusuff/Desktop/dev/blackrock/BlackrockFBP/blackrock/final-object.txt'

print('File 1: ' + str(check_file(file1)))

print('File 2: ' + str(check_file(file2)))

# Both have exactly the same amount of people
