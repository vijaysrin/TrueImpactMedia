import csv
import urllib.request
from bs4 import BeautifulSoup

def open_csv():
    csv_file = open('HortonBoards.csv', 'w')
    writer = csv.writer(csv_file)
    row = ['Unique Face ID', 'Operator Face ID',    'Category',    'Media Type',  'Operator Type',  'Size (H x W, Ft/In)',    'Latitude',    'Longitude', 'Location', 'City', 'Country', 'Digital',     'Spots',     'Seconds',     'Tri-Fold',     'Rotary',     'Facing',     'Reader Side',     'Face Layout',     'Lit'    , 'Lit Hours', '    Extensions',    'DEC',     'TAB EOI',     'TAB ID',    'Rate Card',    'Rate Period', 'Tax Pct',  'Print/unit' ,  'Install/unit'  ,   'Inst. Waived'   ,  'Restrictions',    'Image URL',    'Image Date',    'Image URL 2',    ' Video URL',    'Photo Sheet URL',    'Spec Sheet URL',    'Source URL'    , 'Notes',    'Plant ID',    'Panel ID']
    writer.writerow(row)
    csv_file.close()

def output_to_csv(ID, img, read_side, facing, size, loc, x, y, illum, city, DEC, tri, digital):
    cat = 'Billboard'
    dfl = '?'
    country = 'US'
    no = 'N'
    yes = 'Y'
    space = ' '
    rate_units = '4-week'
    illum_num = dfl
    if illum == 'Y':
        illum_num = '18'
    media = get_media_type(size)
    if img == "":
        img = dfl
    if ID == "":
        ID = dfl
    if loc == "":
        loc = dfl
    if DEC == "":
        DEC = dfl
    if read_side == "":
        read_side = dfl
    csv_file = open('HortonBoards.csv', 'a')
    writer = csv.writer(csv_file)
    if digital == 'Y':
        illum_num = 24
        row = [ID, ID, cat, media, dfl, size, x, y, loc, city, country, digital, '?', '?', tri, no, facing, read_side, dfl, illum, illum_num, dfl, DEC, dfl, dfl, dfl, rate_units, dfl, dfl, dfl, dfl, dfl, img, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl]
    else:
        row = [ID, ID, cat, media, dfl, size, x, y, loc, city, country, digital, space, space, tri, no, facing, read_side, dfl, illum, illum_num, dfl, DEC, dfl, dfl, dfl, rate_units, dfl, dfl, dfl, dfl, dfl, img, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl]
    writer.writerow(row)
    csv_file.close()

def format_city(city):
    not_city_words = ["-City", "-Interstate", "-Border", "North", "South", "- Hwy 70E", "- Hwy 70W, I-10W", "- I-10/25S"]
    for n in not_city_words:
        if city.find(n) != -1:
            city = clean(city, n)
    return city

def clean_ID(id):
    not_id_words = ["Hwy 70E billboard - ", "Hwy 70W, I-10W billboard - ", "I-10/25S billboard - "]
    for i in not_id_words:
        if id.find(i) != -1:
            id = clean(id, i)
    return id

def get_media_type(size):
    sqft = 0
    size = size.replace("'", "")
    dim = size.split("x")
    try: sqft = float(dim[0].rstrip()) * float(dim[1].rstrip())
    except ValueError as v:
        print(size)
    except IndexError as e:
        print(size, dim)
    if(sqft >=800):
        return "Spectacular"
    elif(sqft >=500):
        return "Bulletin"
    elif(sqft >= 320):
        return "Junior Bulletin"
    elif(sqft >= 140):
        return "Poster"
    elif(sqft >= 30):
        return "Jr Poster"
    else:
        return "Other"

def format_size(size):
    size = size.replace("ft", "")
    size = size.replace("in", "")
    size_nums = size.split(" x ")
    for i in range (len(size_nums)):
        size_nums[i] = size_nums[i].rstrip()
        if(size_nums[i].find("x") != -1):
            size_nums[i] = size_nums[i].replace("x", "")
    if(size_nums[0].find(" ") != -1):
        temp = size_nums[0].split("  ")
        if(len(temp) == 2):
            size_nums[0] = temp[0] + "." + str(int(temp[1])//12)
        
    if len(size_nums) > 1:
        h = float(size_nums[0])
        w = float(size_nums[1])
        return str(h) + "' x " + str(w) + "'"
    else:
        return size

def getLocation(address):
    addy = address.split(', ')
    if len(addy) < 3:
        return addy[0], addy[1], 0
    if addy[2].find(' ') != -1:
        addy[2] = addy[2].split(' ')
        return addy[0], addy[1] + ', ' + addy[2][0], addy[2][1]
    return addy[0], addy[1] + ', ' + addy[2], 0

def getCoor(coordinates):
    coords = coordinates.split(', ')
    return coords[0], coords[1]

def getType(typee):
    typee = typee.split(' ')
    if typee[1] != 'Billboard':
        return '?', '?'
    if typee[0] == 'Trivision':    
        return 'Billboard', 'Y'
    return 'Billboard', 'N'

def getRead(read):
    if len(read) == 0:
        return '?'
    str1 = read.find('hand')
    return read[0:str1]

def getDimensions(dimen):
    two = dimen.split('x ')
    trash = two[0].find('.')
    two[0] = two[0][0:trash]
    trash = two[1].find('.')
    two[1] = two[1][0:trash]
    if two[0] == '0.0' or two[1]== '0.0':
        return '?'
    return two[0] + '\' x ' + two[1] + '\''

def getImpressions(eyes):
    
    if len(eyes) == 0:
        return '?'
    per = eyes.find('/')
    if per == -1:
        multiply = eyes.find('k')
        if multiply != -1:
            return str(round(float(eyes[0:multiply]) * 1000 / 7))
        return str(round(int(clean(eyes, ',')) / 7))

    return str(round(int(clean(eyes[0:per], ',')) / 7))

def clean(s ,name):
    return s.replace(name, "").rstrip()

def check_digital(read):
    if read.upper().find("DIGITAL") != -1:
        return clean(read, "Digital"), True
    return read, False

def get_read_layout(read):
    if read.upper().find("UPPER") !=-1:
        return clean(read, "Upper"), "Top"
    elif read.upper().find("LOWER") != -1:
        return clean(read, "Lower"), "Bottom"
    elif read.upper().find("LEFT") != -1:
        return "Left" + clean(read, "Left"), "Left"
    elif read.upper().find("RIGHT") != -1:
        return "Right" + clean(read, "Right"), "Right"
    else:
        return read, "Single"

def open_link(url):
    try: req = urllib.request.Request(url, data=None, headers={}, origin_req_host=None, unverifiable=False, method=None)
    except urllib.error.URLError as e:
        print(e.reason)
    except HTTPError as e:
        print(e.reason)
    else:
        with urllib.request.urlopen(req) as response:
            page = response.read()
    return page

def get_read_facing(b):
    facing_words = ["South", "North", "East", "West"]
    words = b.get_text().split(" - ")
    read = words[0]
    face = ""
    try: face = words[1]
    except IndexError as e:
        print("no dash")
    else:
        for f in facing_words:
            if(face.find(f) != -1):
                face = f
    if face == "":
        face = "?"

    if face.find("South") != -1 or face.find("FS") != -1:
        face = "S"
    elif face.find("North") != -1or face.find("FN") != -1:
        face = "N"
    elif face.find("East") != -1 or face.find("FE") != -1:
        face = "E"
    elif face.find("West") != -1 or face.find("FW") != -1:
        face = "W"

    if read.find("Right") != -1:
        read = "R"
    else:
        read = "L"
    return read, face

def get_info(k):
    link = "http://hortonboards.com/billboard/?face_id=" + str(k)
    page = open_link(link)
    soup = BeautifulSoup(page, 'html.parser')
    ID = soup.find('h2').get_text().split(' - ')
    ID = ID[1]
    if len(ID) != 0:
        digital = 'N'
        print(k)
        info = soup.find_all('p')
        strfull = info[0].get_text()
        str1 = strfull.find('Address:')
        str2 = strfull.find('Coordinates:')
        address = clean(strfull[str1:str2], 'Address: ')
        location, city, ZIP = getLocation(address)

        str3 = strfull.find('County:')
        coordinates = clean(strfull[str2:str3], 'Coordinates: ')
        latitude, longitude = getCoor(coordinates)

        strfull = info[1].get_text()
        str4 = strfull.find('Type:')
        str5 = strfull.find('Facing:')
        typee = clean(strfull[str4:str5], 'Type: ')
        if typee.find('Digital') != -1:
            digital = 'Y'
        category, tri = getType(typee)


        str6 = strfull.find('Traffic:')
        face = clean(strfull[str5:str6], 'Facing: ')
        if(len(face) == 0):
            face = '?'
        else:    
            face = face[0]
        str7 = strfull.find('Read:')
        read = getRead(clean(strfull[str7:], 'Read: '))[0]
        strfull = info[2].get_text()
        str8 = strfull.find('Illuminated:')
        dimen = getDimensions(clean(strfull[0:str8], 'Dimensions: '))
        
        str9 = strfull.find('Impressions:')
        tri = 'N'
        
        lit = 'N'
        lit = clean(strfull[str8: str9], 'Illuminated: ')
        #print(lit)
        if len(lit) == 0:
            lit = '?'
        elif lit.find('LED') != -1 or lit.find('Digital') != -1:
            digital = 'Y'
            lit = 'Y'
        elif lit.find('Static') != -1:
            lit = '?'
        elif lit.find('Tri-Vision') != -1:
            lit = '?'
            tri = 'Y'
        else:
            lit = lit[0]
        #print(lit)
        eyes = getImpressions(clean(strfull[str9:], 'Impressions: '))
        
        
        img = soup.find_all('img')
        img = img[1].get('src')
        output_to_csv(ID, img, read, face, dimen, location, latitude, longitude, lit, city, eyes, tri, digital)


def main():
    open_csv()
    bad_list = [118, 125, 166, 211]
    for k in range(1,225):
        if(k in bad_list):
            print('bad page: ', k)
            continue
        get_info(k)
main()
