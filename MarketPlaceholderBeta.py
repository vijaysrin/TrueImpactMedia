import tkinter
from tkinter import *
import csv
import pyexcel as pe
from bs4 import BeautifulSoup
import requests
import numpy as np
from geopy.geocoders import Nominatim

def main():
    def quit():
        master.destroy()

    def run():
        def add():
            city = en.get()
            state = ent.get()
            category = entr.get()
            country = entry.get()
            media = entry1.get()
            operat = entry2.get()
            zipc = entry3.get()
            write_to_csv(txt, city, state, zipc, category, country, media, operat)

        def quit():
            master.destroy()
            master2.destroy()

        txt = e.get()
        cs1 = txt + "_MP.csv"
        open_csv(cs1)
        master2 = Tk()
        master2.title("Market Placeholder")
        frame1 = Frame(master2)
        frame1.pack()
        lab = Label(frame1, text = "City: ")
        en = Entry(frame1)
        lab.pack(side='left', padx=5, pady=5)
        en.pack(side='left', padx=5, expand=True)

        frame2 = Frame(master2)
        frame2.pack()
        labe = Label(frame2, text = "State: ")
        ent = Entry(frame2)
        labe.pack(side='left', padx=5, pady=5)
        ent.pack(side='left', padx=5, expand=True)

        frame8 = Frame(master2)
        frame8.pack()
        label4 = Label(frame8, text = "Zip Code (0 if unavailable): ")
        entry3 = Entry(frame8)
        label4.pack(side='left', padx=5, pady=5)
        entry3.pack(side='left', padx=5, expand=True)

        frame5 = Frame(master2)
        frame5.pack()
        label1 = Label(frame5, text = "Country: ")
        entry = Entry(frame5)
        label1.pack(side='left', padx=5, pady=5)
        entry.pack(side='left', padx=5, expand=True)

        frame3 = Frame(master2)
        frame3.pack()
        label = Label(frame3, text = "Category: ")
        entr = Entry(frame3)
        label.pack(side='left', padx=5, pady=5)
        entr.pack(side='left', padx=5, expand=True)

        frame6 = Frame(master2)
        frame6.pack()
        label2 = Label(frame6, text = "Media Type: ")
        entry1 = Entry(frame6)
        label2.pack(side='left', padx=5, pady=5)
        entry1.pack(side='left', padx=5, expand=True)

        frame7 = Frame(master2)
        frame7.pack()
        label3 = Label(frame7, text = "Operator Type: ")
        entry2 = Entry(frame7)
        label3.pack(side='left', padx=5, pady=5)
        entry2.pack(side='left', padx=5, expand=True)

        frame4=Frame(master2)
        frame4.pack()
        add = Button(frame4, text = "Add", command = add)
        quit = Button(frame4, text= "Quit", command = quit)
        add.pack(side='left', padx=5)
        quit.pack(side='right', padx=5)

        master2.mainloop()

    master = Tk()
    master.title("Market Placeholder")
    frame1 = Frame(master)
    frame1.pack()
    el = Label(frame1, text = "Operator Name:")
    e = Entry(frame1)
    el.pack(side='left', padx=5, pady=5)
    e.pack(side='left', padx=5, expand=True)

    frame2=Frame(master)
    frame2.pack()
    runimp = Button(frame2, text = "Run", command = run)
    quit = Button(frame2, text="Quit", command = quit)
    runimp.pack(side='left', padx=5)
    quit.pack(side='right', padx=5)

    master.mainloop()

def write_to_csv(name, city, state, zipc, category, country, media, operat):
    cs1 = name + "_MP.csv"
    csv_file = open(cs1, 'a')
    writer = csv.writer(csv_file)
    id = create_id(city, media)
    dfl = '?'
    try:
        city, state, lat, longit, match = get_lat_long(city, state, zipc, country)
    except TypeError:
        print
    else:
        if(match == True):
            city_state = city + ", " + state
            row = [id, id, category, media , operat ,dfl, lat, longit, dfl, dfl, city_state, country, dfl, "", "", dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl, dfl]
            writer.writerow(row)
            csv_file.close()

def create_id(city, category):
    return city + '_' + category + '_MP'

def get_lat_long(city, state, zipc, country):
    geolocator = Nominatim()
    us_state_abbrev = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY',
    'Ontario': 'ON',
    'Alberta': 'AB',
    'British Columbia': 'BC'}
    match = True
    if(state.islower()):
        state = state.upper()
    if(len(state) == 2):
        for i in us_state_abbrev.keys():
            if state == us_state_abbrev[i]:
                state = i
    city_state = city + ", " + state
    if(zipc == '0' or len(zipc) != 5):
        location = geolocator.geocode(city_state + ", " + country)
    else:
        location = geolocator.geocode(city_state + ", " + zipc + ", " + country)
        '''
    print(location.address)
    address = location.address.split(", ")
    city = str(address[0])
    try: state2 = us_state_abbrev[str(address[2])]
    except KeyError as k:
        print('Address incorrect')
        state2 = us_state_abbrev[str(address[1])]
    state = us_state_abbrev[str(state)]
    if(state != state2):
        print("States do not match; add a zip code")
        match = False
    else:
        print(location.latitude, location.longitude)
    print
    '''
    try:
        print(location.address)
    except AttributeError:
        print('Location is unavailable.')
    else:
        print(location.latitude, location.longitude)
        print
        latit, longit = location.latitude, location.longitude
        return city, state, latit, longit, match


def open_csv(name):
    csv_file = open(name, 'w')
    writer = csv.writer(csv_file)
    row = ['Unique Face ID', 'Operator Face ID',    'Category',    'Media Type',  'Operator Type',  'Size (H x W, Ft/In)',    'Latitude',    'Longitude', 'Location', 'Airport Code', 'City', 'Country', 'Digital',     'Spots',     'Seconds',     'Tri-Fold',     'Rotary',     'Facing',     'Reader Side',     'Face Layout',     'Lit'    , 'Lit Hours', 'Extensions',    'DEC',     'TAB EOI',     'TAB ID', 'Avail. Start', 'Avail. End', 'Rates', 'Qty', 'Min', 'Max', 'Rate Card',  'Proposed',  'Rate Period', 'Tax Pct',  'Print/unit' ,  'Install/unit'  ,   'Inst. Waived'   ,  'Restrictions',    'Image URL',    'Image Date',    'Image URL 2',    ' Video URL',    'Photo Sheet URL',    'Spec Sheet URL',    'Source URL'    , 'Notes',  'Metrics',  'Plant ID',    'Panel ID']
    writer.writerow(row)
    csv_file.close()

main()
