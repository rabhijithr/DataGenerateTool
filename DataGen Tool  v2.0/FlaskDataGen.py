from flask import Flask, render_template, request, send_file, flash  # importing flask module
from faker import Faker, Factory
import xlrd
import xlwt
import os


UPLOAD_FOLDER = '/home/RaghuveerEPS/mysite/excel'

EXCEL_FOLDER = '/home/RaghuveerEPS/mysite/excel/FlaskAutoDataGenerator.xls'

# initializing a variable of Flask
app = Flask(__name__)

AU_DataGen = Faker("en_US")

AU_DataGen = Factory.create()


app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

app.config['EXCEL_FOLDER'] = EXCEL_FOLDER

Desktop = os.path.expanduser("~\\Desktop")
Dir = '{}\\DataGen_Templates'.format(Desktop)
os.makedirs(Dir, exist_ok=True)
print(Dir)

def DataGen(filenames):
    # Give the location of the file
    # filelocation = ("{}\\{}".format(Dir, filenames))
    filelocation = '/home/RaghuveerEPS/mysite/excel/{}'.format(filenames)
    # To open Workbook in read mode
    wb = xlrd.open_workbook(filelocation)
    sheet = wb.sheet_by_index(0)

    # open workbook to write data
    wb = xlwt.Workbook()
    ws = wb.add_sheet("DataGenerator", cell_overwrite_ok=True)
    style_header = "font: bold on, color black; borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color teal; align: horiz center, wrap yes,vert centre;"
    StyleHeader = xlwt.easyxf(style_header)
    style_cells = "borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white; align: horiz center, wrap yes,vert centre;"
    StyleCells = xlwt.easyxf(style_cells)
    # wb.row.height_mismatch = True
    # wb.row.height = 256*20

    StyleDateCells = xlwt.easyxf(style_cells, num_format_str='YYYY-MM-DD')

    # For row 0 and column 0
    print(sheet.cell_value(0, 0))
    print(sheet.cell_value(1, 0))

    NumberOfTimes = request.form.get('NumberOfTimes')
    # NumberOfTimes = int(sheet.cell_value(2, 0))

    OTB = request.form.get('OTB')

    print(NumberOfTimes)

    HeaderNames = sheet.col_values(0)
    # HeaderNames = [x for x in HeaderNames if x]
    HeaderNames = HeaderNames[2:]  # print list starting from 2nd element
    print(HeaderNames)
    print(HeaderNames[1])

    print(sheet.nrows)  # print number of rows in excel that have data


    # list value of 4th row in excel
    DataList = sheet.row_values(4)
    DataList = [x for x in DataList if x]
    print(DataList)

    for x in range(sheet.nrows):
        ws.col(x).width = int(20 * 260)
        ws.row(0).height_mismatch = True
        ws.row(0).height = 20 * 22

    for rowheight in range(int(NumberOfTimes)):
        ws.row(rowheight + 1).height_mismatch = True
        ws.row(rowheight + 1).height = 20 * 30

    for i in range(5, sheet.nrows):
        DataList = sheet.row_values(i)
        print(DataList[1])
        # DataList = [x for x in DataList if x]  # removes empty spaces in list
        # print(DataList[0])
        # print(DataList[0]), DataList[0] has the first column value of list and Datalist[1] is having second caloumn value.
        if 'FullName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.name(), style=StyleCells)
        else:
            pass

        if 'FirstName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, DataList[0])
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.first_name(), style=StyleCells)
        else:
            pass

        if 'LastName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.last_name(), style=StyleCells)
        else:
            pass

        if 'NumberLength' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.random_number(digits=int(DataList[2]), fix_len=DataList[3]),
                         style=StyleCells)
        else:
            pass

        if 'NumberRange' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.random.randint(int(DataList[2]), int(DataList[3])),
                         style=StyleCells)
        else:
            pass

        if 'Email' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(30 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.email(), style=StyleCells)
        else:
            pass

        if 'Safe_Mail' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(30 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.safe_email(), style=StyleCells)
        else:
            pass

        if 'Country' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(24 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.country(), style=StyleCells)
        else:
            pass

        if 'City' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.city(), style=StyleCells)
        else:
            pass

        if 'Address' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(40 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.address(), style=StyleCells)
        else:
            pass

        if 'Zipcode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.zipcode(), style=StyleCells)
        else:
            pass

        if 'CustomString' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.bothify(text=DataList[2]), style=StyleCells)
        else:
            pass

        if 'SecondaryAddress' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.secondary_address(), style=StyleCells)
        else:
            pass

        if 'State' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.state(), style=StyleCells)
        else:
            pass

        if 'StreetAddress' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.street_address(), style=StyleCells)
        else:
            pass

        if 'CountryCode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.country_code(representation="alpha-{}".format(int(DataList[2]))),
                         style=StyleCells)
        else:
            pass

        if 'LicencePlate' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.license_plate(), style=StyleCells)
        else:
            pass

        if 'BBAN' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.bban(), style=StyleCells)
        else:
            pass

        if 'IBAN' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.iban(), style=StyleCells)
        else:
            pass

        if 'EAN' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.ean(length=int(DataList[2])), style=StyleCells)
        else:
            pass

        if 'Colour' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.color_name(), style=StyleCells)
        else:
            pass

        if 'Hex_Colour' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.hex_color(), style=StyleCells)
        else:
            pass

        if 'CreditCard' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.credit_card_number(card_type=None), style=StyleCells)
        else:
            pass

        if 'CreditCard_Provider' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.credit_card_provider(card_type=None), style=StyleCells)
        else:
            pass

        if 'Currency' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.currency_name(), style=StyleCells)
        else:
            pass

        if 'CurrencyCode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.currency_code(), style=StyleCells)
        else:
            pass

        if 'CryptoCurrency' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_name(), style=StyleCells)
        else:
            pass

        if 'CryptoCurrency_Code' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_code(), style=StyleCells)
        else:
            pass

        if 'Date' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.date(pattern=DataList[2], end_datetime=None), style=StyleCells)
        else:
            pass

        if 'Time' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.time(pattern=DataList[2], end_datetime=None), style=StyleCells)
        else:
            pass

        if 'TimeZone' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.timezone(), style=StyleCells)
        else:
            pass

        if 'FileName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.file_name(category=None, extension=DataList[2]), style=StyleCells)
        else:
            pass

        if 'FileExtension' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.file_extension(category=None), style=StyleCells)
        else:
            pass

        if 'FilePath' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.file_path(depth=int(DataList[2]), category=None, extension=DataList[3]), style=StyleCells)
        else:
            pass

        if 'UpperLetter' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.random_uppercase_letter(), style=StyleCells)
        else:
            pass

        if 'CustomLetter' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.random_element(elements=list(DataList[2])), style=StyleCells)
        else:
            pass

        if 'FormatedString' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.password(length=int(DataList[2]), special_chars=DataList[3], digits=DataList[4], upper_case=DataList[5], lower_case=DataList[6]), style=StyleCells)
        else:
            pass

        if 'Latitude' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.latitude(), style=StyleCells)
        else:
            pass

        if 'Longitude' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.longitude(), style=StyleCells)
        else:
            pass

        if 'StateCode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.state_abbr(include_territories=True), style=StyleCells)
        else:
            pass

        if 'Boolean' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.boolean(chance_of_getting_true=50), style=StyleCells)
        else:
            pass

        if 'DecimalNumber' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.pydecimal(left_digits=int(DataList[2]), right_digits=int(DataList[3]), positive=int(DataList[4])), style=StyleCells)
        else:
            pass

        if 'DOB' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.date_of_birth(tzinfo=None, minimum_age=int(DataList[2]), maximum_age=int(DataList[3])), style=StyleDateCells)
        else:
            pass

        if 'FutureDate' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.future_date(end_date="+{}".format(DataList[2]), tzinfo=None), style=StyleDateCells)
        else:
            pass

        if 'StoreName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.company(), style=StyleDateCells)
        else:
            pass

        if 'StringRange' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.pystr(min_chars=int(DataList[2]), max_chars=int(DataList[3])), style=StyleDateCells)
        else:
            pass

        if 'StreetName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.street_name(), style=StyleDateCells)
        else:
            pass

        if 'BuildingNumber' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.building_number(), style=StyleDateCells)
        else:
            pass

        if 'AL_AddressLine1' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, "{} {}".format(AU_DataGen.building_number(), AU_DataGen.street_name()), style=StyleDateCells)
        else:
            pass

        if 'AL_AddressLine3' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, "{} {}".format(AU_DataGen.city(), AU_DataGen.zipcode()), style=StyleDateCells)
        else:
            pass

        if 'PhoneNumber' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.phone_number(), style=StyleDateCells)
        else:
            pass

    wb.save(app.config['EXCEL_FOLDER'])


@app.route('/')
def index():
    return render_template('DataGenerator.html')


@app.route('/', methods=['POST'])
def DataGenrator():
    if request.form['action'] == 'Generate':
        if request.method == 'POST':
            if request.form.get('OTB') == '1':
                if request.form.get('Templates') == "AL_Template":
                    listfiles = '/custom/AL_Ex_Template.xlsm'
                    DataGen(listfiles)
                    return send_file('/home/RaghuveerEPS/mysite/excel/FlaskAutoDataGenerator.xls',
                                     mimetype='text/xls',
                                     attachment_filename='DataGenOutput.xls',
                                     as_attachment=True)

                if request.form.get('Templates') == "AL_StoreImport":
                    listfiles = '/custom/AL_StoreImport.xlsm'
                    DataGen(listfiles)
                    return send_file('/home/RaghuveerEPS/mysite/excel/FlaskAutoDataGenerator.xls',
                                     mimetype='text/xls',
                                     attachment_filename='DataGenOutput.xls',
                                     as_attachment=True)

            else:
                listfiles = request.files['file']
                filenames = listfiles.filename
                listfiles.save(os.path.join(app.config['UPLOAD_FOLDER'], filenames))
                DataGen(filenames)
                os.remove('/home/RaghuveerEPS/mysite/excel/{}'.format(filenames))
                return send_file('/home/RaghuveerEPS/mysite/excel/FlaskAutoDataGenerator.xls',
                             mimetype='text/xls',
                             attachment_filename='DataGenOutput.xls',
                             as_attachment=True)


    if request.form['action'] == 'Download Template':
        if request.method == 'POST':
            gettemp = request.form.getlist('Templates')
            i=0
            for gettemp[i] in gettemp:
                if "AL_Template" in gettemp[i]:
                    return send_file('/home/RaghuveerEPS/mysite/excel/AL_Ex_Template.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='AL_EX_Template.xlsm',
                                     as_attachment=True)
                if "Default_Template" in gettemp[i]:
                    return send_file('/home/RaghuveerEPS/mysite/excel/Template.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='Default_Template.xlsm',
                                     as_attachment=True)
                if "AL_StoreImport" in gettemp[i]:
                    return send_file('/home/RaghuveerEPS/mysite/excel/AL_StoreImport.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='AL_StoreImport_Template.xlsm',
                                     as_attachment=True)


@app.errorhandler(500)
def page_not_found(e):
    # note that we set the 404 status explicitly
    return render_template('500error.html'), 500


if __name__ == "__main__":
    app.run(debug=True)