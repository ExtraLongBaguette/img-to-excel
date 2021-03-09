# img-to-excel
Small tool to convert images to colored .xlsx documents
## Dependencies
```
pip install pillow
pip install xlsxwriter
```
## Usage
```
img-to-excel.py -i <input_image> -o <output_file> [-s <size_in_xlsx>]
```
For help pass `-h`
```
img-to-excel.py -h
```
The default size is 256 cells
### Input file must be .jpg format
This tool was made just for practise so the need for other formats was not a concern<br>
Inspired by [this](https://www.reddit.com/r/InternetIsBeautiful/comments/m0h72d/convert_pictures_into_excel_spreadsheets_with/) reddit post<br>
Tested in windows, no idea if this script works in any other os
