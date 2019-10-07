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


def main():
    market_abbreviation = {
        'Orlando': 'ORL',
        'Albuquerque': 'ALB',
        'Atlanta': 'ATL',
        'Boston': 'BOS',
        'Chicago': 'CHI',
        'Columbus': 'CMH',
        'Dallas/Fort Worth': 'DFW',
        'Daytona Beach/Melbourne': 'MEL',
        'El Paso': 'ELP',
        'Houston': 'HOU',
        'Indianapolis': 'IND',
        'Jacksonville': 'JAX',
        'Las Vegas': 'LVG',
        'Los Angeles': 'LAX',
        'Miami/Ft. Lauderdale': 'MIA',
        'Milwaukee': 'MKE',
        'Minneapolis/St Paul': 'MSP',
        'New York/Manhattan': 'NYC',
        'Ocala/Gainesville': 'OCA',
        'Philadelphia': 'PHI',
        'Phoenix': 'PHX',
        'Sacramento': 'SAC',
        'Salisbury/Ocean City': 'SAL',
        'San Antonio': 'SAT',
        'San Diego': 'SDI',
        'San Francisco': 'OAK',
        'Tampa Bay': 'TAM',
        'Tucson': 'TUC',
        'Washington DC/Baltimore': 'BWI',
        'West Palm Beach/Ft. Pierce': 'WPB'
    }
    xlfile = ""
    try:
        xlfile = input("Name of Clear Channel file to input: ")
    except NameError:
        print('Name Error')
    
    cc_array = pe.get_array(file_name = xlfile)


    #open csv to import
    market_data = cc_array[6][5]
    #print(market_data)
    markets = market_data.split('/')
    if (markets.__len__() == 2):
        market = markets[0] + '_' + markets[1]
        #print("Here")
    else:
        market = market_data

    prefix = ""
    for i in market_abbreviation.keys():
        if market_data == i:
            prefix = market_abbreviation[i]


    #print(markets[0])
    
    full_name = ('ClearChannel_' + market + '.csv')
    csv_file = open(full_name, 'w')
    writer = csv.writer(csv_file)
    row = ['Unique Face ID', 'Operator Face ID',    'Category',    'Media Type',  'Operator Type',  'Size (H x W, Ft/In)',    'Latitude',    'Longitude', 'Location', 'Airport Code', 'City', 'Country', 'Digital',     'Spots',     'Seconds',     'Tri-Fold',     'Rotary',     'Facing',     'Reader Side',     'Face Layout',     'Lit'    , 'Lit Hours', 'Extensions',    'DEC',     'TAB EOI',     'TAB ID', 'Avail. Start', 'Avail. End', 'Rates', 'Qty', 'Min', 'Max', 'Rate Card',  'Proposed',  'Rate Period', 'Tax Pct',  'Print/unit' ,  'Install/unit'  ,   'Inst. Waived'   ,  'Restrictions',    'Image URL',    'Image Date',    'Image URL 2',    ' Video URL',    'Photo Sheet URL',    'Spec Sheet URL',    'Source URL'    , 'Notes',  'Metrics',  'Plant ID',    'Panel ID']
    writer.writerow(row)
    csv_file.close()

    dfl = '?'
    country = 'US'

    for row in pe.get_array(file_name = xlfile, start_row=6):
        csv_file = open(full_name, 'a')
        writer = csv.writer(csv_file)
        name = prefix + '-' + row[8]

        location = row[10]

        num_panels = row[13]
        rates = 'Unit'
        if num_panels != 1:
            rates = 'Package'
        rate_card = row[32]
        rate_proposed = row[33]
        weeks = float(row[30])
        periods = float(row[31])
        rate_period = str(int(weeks / periods)) + ' Week'
        print_cost = row[35]
        install_cost = row[37]
        tax = row[39]
        spots = row[137]
        seconds = row[138]
        digital = False
        dig = 'N'
        if(row[6] == 'Digital'):
            digital = True
            dig = 'Y'
        facing = row[15]
        size = row[16]
        media_type = get_media_type(size)
        category = 'Billboard'
        read = row[18]
        illum = 'N'
        hours = 0
        if(digital):
            illum = 'Y'
            hours = 24
        else:
            if(row[19].__eq__('Yes')):
                illum = 'Y'
                hours = row[20]
            else:
                illum = 'N'
                hours = 12
        city = row[21].title() + ', ' + row[23]
        lat = row[25]
        lon = row[26]
        tab_eoi = int(row[44])
        tab_ID = row[9]
        start = row[27]
        end = row[28]
        new_row = [name, name, category, media_type, dfl, size, lat, lon, location, dfl, city, country, dig, spots, seconds, dfl, dfl, facing, read, dfl, illum, hours, dfl, dfl, tab_eoi, tab_ID, start, end, rates, num_panels, num_panels, num_panels, rate_card, rate_proposed, rate_period, tax, print_cost, install_cost, dfl, dfl, "", dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl]
        writer.writerow(new_row)
        csv_file.close()



main()