import xlsxwriter, getopt, sys, os
from PIL import Image

def main(argv):
    input_image = ""
    output_file = ""
    max_size = 256
    size = max_size
    #Handle arguments
    try:
        opts, args = getopt.getopt(argv, "hi:o:s:", ["ifile=", "ofile=","size="])
    except getopt.GetoptError:
        print("Usage: img-to-excel.py -i <input_image> -o <output_file>")
        sys.exit(2)

    for opt, arg in opts:
        if opt == "-h":
            print("img-to-excel.py -i <input_image> -o <output_file>")
            sys.exit()
        elif opt in ("-i", "--ifile"):
            input_image = arg
        elif opt in ("-o", "--ofile"):
            output_file = arg
        elif opt in ("-s", "--size"):
            size = int(arg)

    if not ".jpg" in os.path.splitext(input_image):
        print("img-to-excel.py -i <input_image> -o <output_file>")
        print("Use *.jpg image as an input")
        sys.exit()
    if not os.path.exists(input_image):
        print(input_image + " does not exist")
        sys.exit()
    if not ".xlsx" in os.path.splitext(output_file):
        print("img-to-excel.py -i <input_image> -o <output_file>")
        print("Use *.xlsx file as a output")
        sys.exit()
    if size > max_size:
        print("The size is larger than the default limit (256)")
        print("Larger files are slower and may not work")
        i = ""
        while i not in ["y","Y","n","N"]:
            i = input("Do you want to continue anyways (Y/N): ")
            if i in ["y","Y"]:
                break
            elif i in ["n","N"]:
                size = max_size
                break

    #Create the excel file
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    image = Image.open(input_image)
    image.thumbnail((size, size))
    px = image.load()

    worksheet.set_column(0,image.width,2) #Square the cells

    for x in range(0, image.width):
        for y in range(0, image.height):
            cell_format = workbook.add_format()
            r, g, b = px[x,y]
            cell_color = '#{:02x}{:02x}{:02x}'.format(r,g,b) #Creates hex color code from rgb
            cell_format.set_bg_color(cell_color)
            worksheet.write(y, x, "", cell_format)

    try:
        workbook.close()
    except:
        print("Could not write to file " + output_file)
        print("Check that it is not open")
        sys.exit()
    print("Good bye")

if __name__ == "__main__":
    main(sys.argv[1:])
