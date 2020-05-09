import xlsxwriter # using to create the spreadsheet
from PIL import Image
from operator import itemgetter #extracts elements of tuplules from a list of tuplules

#get image from the user
# input_img = input('Enter image name: ')
opened_image = Image.open('logo.png')

#store the size of the image as a tuplules
size = opened_image.size

#store the first element of given tupule in row and the second in col

row, col = size[0],size[1]

#set pixel limit of the picture, if it is bigger then it gets resized

max_rows, max_cols = 100, 100

#check to see if the image is resized

resized = False

if row > max_rows or col > max_cols:
    opened_image = opened_image.resize((max_rows,max_cols))
    #Set rezied to True
    print('Image has been resized from {}px by {}px to {}px by {}px'.format(size[0],size[1],max_rows,max_cols))
    resized = True

if not resized: # checks if image has been resized
    row = row*3 # 3 vertical cells make up ONE pixel in a spreadsheet, rows have to be multiplied by 3
else:
    row = max_rows*3
    col = max_cols

#return a list of tupules where each tupule contains the RGB value of one pixel.
# NOTE: Pixels are read from Top Left to Right

data = list(opened_image.getdata())
# print(data)

# Split each value of the list into their own individual colors. DO so by taking each elemenet of tupules

red_list = list(map(itemgetter(0), data)) # will return a list of R values of each pixel... RED Pixels stored at index 0
green_list = list(map(itemgetter(1),data)) # return green pixels at index 1
blue_list = list(map(itemgetter(2),data)) # return blue pixels at index 2


#create a new excel file to begin writing pixel values to

workbook = xlsxwriter.Workbook('excelImageView.xlsx')

worksheet = workbook.add_worksheet() # created a new worksheet within the new workbook

#track row numbers in order to fill/format with the right colors

redx, greenx,bluex = 0,1,2

# formatting 3 rows at a time, so we take the original pixel count or original rows as range for iteration

iteration_number = row //3

for _ in range(iteration_number):
    #format the red rows. For each row, formatting starts at right and ends at the Nth cell where N is the number of columns
    worksheet.conditional_format(redx,0,redx,col, {
    'type': '2_color_scale',
    'min_type': 'num',
    'max_type': 'num',
    'min_value': 0,
    'max_value': 255,
    'min_color': '#000000',
    'max_color': '#FF0000',
    })


    #format green rows
    worksheet.conditional_format(greenx,0,greenx,col, {
    'type': '2_color_scale',
    'min_type': 'num',
    'max_type': 'num',
    'min_value': 0,
    'max_value': 255,
    'min_color': '#000000',
    'max_color': '#00FF00',
    })



    #format blue rows
    worksheet.conditional_format(bluex,0,bluex,col, {
    'type': '2_color_scale',
    'min_type': 'num',
    'max_type': 'num',
    'min_value': 0,
    'max_value': 255,
    'min_color': '#000000',
    'max_color': '#0000FF',
    })

    #after each iteration increase respective rows by 3 to get the next set of 3

    redx +=3
    bluex+=3
    greenx+=3


#track the cell we are at any given time. cellx corresponds to the row number and celly to the column of a particular cells

cellx, celly = 0,0
count = 0 # track row type we are in. 0 -> Red 1-> Green 2->blue
r,g,b = 0,0,0 # grab RGB values from their list


for _ in range(row): # each iteration formats a single row
    for _ in range(col): #formats a single cell
        if count == 0: #checks to see if formatting red cells
            worksheet.write(cellx,celly,red_list[r]) #write to cells at the Rth value
            r +=1 #increase by one so the next R value gets written to the next cell
        elif count == 1:
            worksheet.write(cellx,celly,green_list[g])
            g+=1
        elif count == 2:
            worksheet.write(cellx,celly,blue_list[b])
            b+=1
        celly +=1
    #reset the count back to 0 on formatting the third row of a row set equal to one vertical pixels
    if count == 2:
        count = 0
    else:
        count+=1
    celly = 0
    cellx += 1
workbook.close()
