import pyexcel as pe
import csv

def get_media_type(size):
    sqft = 0
    dim = size.split("x")
    x_dim = 0
    y_dim = 0
    
    if('"' in dim[0]):
        decimal = int(dim[0][(dim[0].index('"') - 1):dim[0].index('"')])
        x_dim = float(float(dim[0][0:2]) + float(decimal / 12))
    
    else:
        dim[0] = dim[0].replace("'", "")
        x_dim = float(dim[0].rstrip())

    if('"' in dim[1]):
        decimal2 = int(dim[1][(dim[1].index('"') - 1):dim[1].index('"')])
        y_dim = float(float(dim[1][0:2]) + float(decimal2 / 12))

    else:
        dim[1] = dim[1].replace("'", "")
        y_dim = float(dim[1].rstrip())

    

    #size = size.replace("'", "")

    try: sqft = x_dim * y_dim
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

def get_size(size_pro):
    sizes = size_pro.split(" ")
    result = sizes[0]
    if(sizes[1] != '0"'):
        result += sizes[1]
    result = result + " x " + sizes[3]
    if(sizes[4] != '0"'):
        result += sizes[4]
    print(result)
    return result


def get_markets(xlfile):
    total_markets = 0
    market_list = []
    for row in pe.get_array(file_name = xlfile, start_row=1):
        if(row[0] == ""):
            market_holder = row[1]
            if(':' in market_holder):
                market = market_holder.split(': ')
            else:
                market = market_holder.split(' ')
            print(market_holder)
            market_list.append(market[1])
            total_markets = total_markets + 1
    print ("There are " + str(total_markets) + " markets in: ")
    print(*market_list, sep = ", ")
    return market_list

def is_digital(xlfile):
    for row in pe.get_array(file_name = xlfile, start_row=1):
        if(row[2].split(" ")[0] == 'Digital'):
            return True
    return False

def main():

    xlfile = ""
    try:
        xlfile = input("Name of Lamar file to input: ")
    except NameError:
        print('Name Error')
    
    markets = get_markets(xlfile)
    
    full_name = ('Lamar_' + markets[0] + '.csv')
    csv_file = open(full_name, 'w')
    writer = csv.writer(csv_file)
    row = ['Unique Face ID', 'Operator Face ID',    'Category',    'Media Type',  'Operator Type',  'Size (H x W, Ft/In)',    'Latitude',    'Longitude', 'Location', 'Airport Code', 'City', 'Country', 'Digital',     'Spots',     'Seconds',     'Tri-Fold',     'Rotary',     'Facing',     'Reader Side',     'Face Layout',     'Lit'    , 'Lit Hours', 'Extensions',    'DEC',     'TAB EOI',     'TAB ID', 'Avail. Start', 'Avail. End', 'Rates', 'Qty', 'Min', 'Max', 'Rate Card',  'Proposed',  'Rate Period', 'Tax Pct',  'Print/unit' ,  'Install/unit'  ,   'Inst. Waived'   ,  'Restrictions',    'Image URL',    'Image Date',    'Image URL 2',    ' Video URL',    'Photo Sheet URL',    'Spec Sheet URL',    'Source URL'    , 'Notes',  'Metrics',  'Plant ID',    'Panel ID']
    writer.writerow(row)
    csv_file.close()

    dfl = '?'
    country = 'US'

    for row in pe.get_array(file_name = xlfile, start_row=2):
        if(row[0] != ""):
            name = str(row[0]) + "-" + str(row[3])
            city = row[2]
            category = "Billboard"
            media_type = row[2]
            digital = 'N'
            illum = row[11][0:1]
            spots = ""
            seconds = ""
            hours = row[12]
            size = get_size(row[14])
            tab_id = row[5]
            tab_eoi = row[10]
            facing = row[7][0:1]
            lat = row[8]
            lon = row[9]
            read = dfl
            location = row[6]
            start = row[15]
            end = row[16]
            rate = row[17]
            proposed = row[18]
            if(is_digital(xlfile)):
                name = str(row[0]) + "-" + str(row[3])
                city = row[2]
                category = "Billboard"
                media_type = row[2]
                digital = 'N'
                spots = ""
                seconds = ""
                if(row[2].split(' ')[0] == 'Digital'):
                    digital = 'Y'
                    spots = row[18]
                    seconds = row[17]
                facing = row[7][0:1]
                
                illum = row[11][0:1]
                hours = row[12]
                size = get_size(row[14])
                tab_id = row[5]
                tab_eoi = row[10]

            
                print(row[19])
                read = row[19][0:1]
                start = row[20]
                end = row[21]
                rate = row[22]
                proposed = row[23]
                
                lat = row[8]
                lon = row[9]
                location = row[6]
                

            csv_file = open(full_name, 'a')
            writer = csv.writer(csv_file)
            new_row = [name, name, category, media_type, dfl, size, lat, lon, location, dfl, city, 'US', digital, spots, seconds, dfl, dfl, facing, read, dfl, illum, hours, dfl, dfl, tab_eoi, tab_id, start, end, 'Unit', 1,1,1, rate, proposed, '4 week', dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl]
            writer.writerow(new_row)
            csv_file.close()



main()