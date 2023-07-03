import argparse
import re
import sys
import xlrd
from datetime import datetime
from pathlib import Path
from pymarc import Record, Field


def parse_title(string):
    title = []
    if ':' in string:
        t = re.sub(r'^(.*?): (.*?), (.*)$', r'\g<1>', string)
        b = re.sub(r'^(.*?): (.*?), (.*)$', r'\g<2>', string)
        v = re.sub(r'^(.*?): (.*?), (.*)$', r'\g<3>', string)
        title.append(t)
        title.append(b)
        title.append(v)
    elif ',' in string:
        t = re.sub(r'^(.*?), (.*)$', r'\g<1>', string)
        v = re.sub(r'^(.*?), (.*)$', r'\g<2>', string)
        title.append(t)
        title.append(v)
    else:
        title.append(string)
    return title


def lowercase_numbering(string):
    num = string
    num = re.sub(r'Vol', r'vol', num)
    num = re.sub(r'No\.', r'no.', num)
    return num


def subfields_from_string(string):
    subfields = []
    if '$' in string:
        string = string.split('$')
        for x in string:
            if string.index(x) == 0:
                subfields.append('a')
                subfields.append(x)
            else:
                subfields.append(x[0:1])
                subfields.append(x[1:])
    else:
        subfields.append('a')
        subfields.append(string)
    return subfields


def subfields_from_string_series(string):
    subfields = []
    if '$' in string:
        string = string.split('$')
        for x in string:
            if not x == '':
                if x[0:1] == 's' or x[0:1] =='t':
                    subfields.append(x[0:1])
                    subfields.append(x[1:])
                else:
                    subfields.append('t')
                    subfields.append(x)
    return subfields


def subfields_from_string_relator(string, relators):
    subfields = subfields_from_string(string)
    
    if '1' in subfields:
        if len(relators) == 1:
            subfields.insert(subfields.index('1'), 'e')
            subfields.insert(subfields.index('1'), relators[0] + '.')
        else:
            for i in relators[:-1]:
                subfields.insert(subfields.index('1'), 'e')
                subfields.insert(subfields.index('1'), i + ',')
            subfields.insert(subfields.index('1'), 'e')
            subfields.insert(subfields.index('1'), relators[-1] + '.')
        
    else:
        if len(relators) == 1:
            subfields.append('e')
            subfields.append(relators[0] + '.')
        else:
            for i in relators[:-1]:
                subfields.append('e')
                subfields.append(i + ',')
            subfields.append('e')
            subfields.append(relators[-1] + '.')
    
    index_subf_before_relator = subfields.index('e') - 1
    if not subfields[index_subf_before_relator].endswith(',') and not subfields[index_subf_before_relator].endswith('-'):
        subfields[index_subf_before_relator] += ','
    return subfields


def title_to_series(string):
    if 'Vol.' in string or 'No.' in string:
        series = subfields_from_string(string)
        series_dict = {series[i]: series[i + 1] for i in range(0, len(series), 2)}
        series_dict['n'] = re.sub(r'^(.*?),? \[?\d{4}\]?$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), January$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), February$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), March$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), April$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), May$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), June$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), July$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), August$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), September$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), October$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), November$', r'\g<1>', series_dict['n'])
        series_dict['n'] = re.sub(r'^(.*?), December$', r'\g<1>', series_dict['n'])
        series_dict['n'] = lowercase_numbering(series_dict['n'])
        series_dict['v'] = series_dict.pop('n')
        if 'b' in series_dict:
            series_dict['a'] = series_dict['a'].rstrip(':') + '. ' + series_dict['b'][0].upper() + series_dict['b'][1:].rstrip(',') + ';'
            series_dict.pop('b')
        else:
            series_dict['a'] = series_dict['a'].rstrip(',') +  ';'
        series = ['a', series_dict['a'], 'v', series_dict['v']]
        return series
    else:
        return None
    

def date_from_string(string):
    date = re.sub(r'^.*?(\[?\d{4}\]?).*$', r'\g<1>', string)
    return date


def year_from_date(string):
    year = re.sub(r'\[?(\d{4})\]?', r'\g<1>', string)
    return year


def country_code_from_pub_place(string):
    country_dict = {"Albania": "aa ", "Alberta": "abc", "Australian Capital Territory": "aca", "Algeria": "ae ", "Afghanistan": "af ", "Argentina": "ag ", "Armenia (Republic)": "ai ", "Azerbaijan": "aj ", "Alaska": "aku", "Alabama": "alu", "Anguilla": "am ", "Andorra": "an ", "Angola": "ao ", "Antigua and Barbuda": "aq ", "Arkansas": "aru", "American Samoa": "as ", "Australia": "at ", "Austria": "au ", "Aruba": "aw ", "Antarctica": "ay ", "AZ": "azu", "Arizona": "azu", "Bahrain": "ba ", "Barbados": "bb ", "British Columbia": "bcc", "Burundi": "bd ", "Belgium": "be ", "Bahamas": "bf ", "Bangladesh": "bg ", "Belize": "bh ", "British Indian Ocean Territory": "bi ", "Brazil": "bl ", "Bermuda Islands": "bm ", "Bosnia and Herzegovina": "bn ", "Bolivia": "bo ", "Solomon Islands": "bp ", "Burma": "br ", "Botswana": "bs ", "Bhutan": "bt ", "Bulgaria": "bu ", "Bouvet Island": "bv ", "Belarus": "bw ", "Brunei": "bx ", "Caribbean Netherlands": "ca ", "CA": "cau", "California": "cau", "Cambodia": "cb ", "China": "cc ", "Chad": "cd ", "Sri Lanka": "ce ", "Congo (Brazzaville)": "cf ", "Congo (Democratic Republic)": "cg ", "China (Republic : 1949- )": "ch ", "Croatia": "ci ", "Cayman Islands": "cj ", "Colombia": "ck ", "Chile": "cl ", "Cameroon": "cm ", "Curaçao": "co ", "Colorado": "cou", "Comoros": "cq ", "Costa Rica": "cr ", "Connecticut": "ctu", "Cuba": "cu ", "Cabo Verde": "cv ", "Cook Islands": "cw ", "Central African Republic": "cx ", "Cyprus": "cy ", "District of Columbia": "dcu", "Delaware": "deu", "Denmark": "dk ", "Benin": "dm ", "Dominica": "dq ", "Dominican Republic": "dr ", "Eritrea": "ea ", "Ecuador": "ec ", "Equatorial Guinea": "eg ", "Timor-Leste": "em ", "England": "enk", "London": "enk", "Estonia": "er ", "El Salvador": "es ", "Ethiopia": "et ", "Faroe Islands": "fa ", "French Guiana": "fg ", "Finland": "fi ", "Fiji": "fj ", "Falkland Islands": "fk ", "FL": "flu", "Florida": "flu", "Micronesia (Federated States)": "fm ", "French Polynesia": "fp ", "France": "fr ", "Terres australes et antarctiques françaises": "fs ", "Djibouti": "ft ", "Georgia": "gau", "Kiribati": "gb ", "Grenada": "gd ", "Guernsey": "gg ", "Ghana": "gh ", "Gibraltar": "gi ", "Greenland": "gl ", "Gambia": "gm ", "Gabon": "go ", "Guadeloupe": "gp ", "Greece": "gr ", "Georgia (Republic)": "gs ", "Guatemala": "gt ", "Guam": "gu ", "Guinea": "gv ", "Germany": "gw ", "Guyana": "gy ", "Gaza Strip": "gz ", "Hawaii": "hiu", "Heard and McDonald Islands": "hm ", "Honduras": "ho ", "Haiti": "ht ", "Hungary": "hu ", "Iowa": "iau", "Iceland": "ic ", "Idaho": "idu", "Ireland": "ie ", "India": "ii ", "IL": "ilu", "Illinois": "ilu", "Isle of Man": "im ", "Indiana": "inu", "Indonesia": "io ", "Iraq": "iq ", "Iran": "ir ", "Israel": "is ", "Italy": "it ", "Côte d'Ivoire": "iv ", "Iraq-Saudi Arabia Neutral Zone": "iy ", "Japan": "ja ", "Jersey": "je ", "Johnston Atoll": "ji ", "Jamaica": "jm ", "Jordan": "jo ", "Kenya": "ke ", "Kyrgyzstan": "kg ", "Korea (North)": "kn ", "Korea (South)": "ko ", "Kansas": "ksu", "Kuwait": "ku ", "Kosovo": "kv ", "Kentucky": "kyu", "Kazakhstan": "kz ", "Louisiana": "lau", "Liberia": "lb ", "Lebanon": "le ", "Liechtenstein": "lh ", "Lithuania": "li ", "Lesotho": "lo ", "Laos": "ls ", "Luxembourg": "lu ", "Latvia": "lv ", "Libya": "ly ", "Massachusetts": "mau", "Manitoba": "mbc", "Monaco": "mc ", "Maryland": "mdu", "Maine": "meu", "Mauritius": "mf ", "Madagascar": "mg ", "Michigan": "miu", "Montserrat": "mj ", "Oman": "mk ", "Mali": "ml ", "Malta": "mm ", "Minnesota": "mnu", "Montenegro": "mo ", "Missouri": "mou", "Mongolia": "mp ", "Martinique": "mq ", "Morocco": "mr ", "Mississippi": "msu", "Montana": "mtu", "Mauritania": "mu ", "Moldova": "mv ", "Malawi": "mw ", "Mexico": "mx ", "Malaysia": "my ", "Mozambique": "mz ", "Nebraska": "nbu", "NC": "ncu", "North Carolina": "ncu", "North Dakota": "ndu", "Netherlands": "ne ", "Newfoundland and Labrador": "nfc", "Niger": "ng ", "New Hampshire": "nhu", "Northern Ireland": "nik", "New Jersey": "nju", "New Brunswick": "nkc", "New Caledonia": "nl ", "New Mexico": "nmu", "Vanuatu": "nn ", "Norway": "no ", "Nepal": "np ", "Nicaragua": "nq ", "Nigeria": "nr ", "Nova Scotia": "nsc", "Northwest Territories": "ntc", "Nauru": "nu ", "Nunavut": "nuc", "Nevada": "nvu", "Northern Mariana Islands": "nw ", "Norfolk Island": "nx ", "NY": "nyu", "N.Y.": "nyu", "New York": "nyu", "New Zealand": "nz ", "Ohio": "ohu", "Oklahoma": "oku", "Ontario": "onc", "Oregon": "oru", "Mayotte": "ot ", "PA": "pau", "Pennsylvania": "pau", "Pitcairn Island": "pc ", "Peru": "pe ", "Paracel Islands": "pf ", "Guinea-Bissau": "pg ", "Philippines": "ph ", "Prince Edward Island": "pic", "Pakistan": "pk ", "Poland": "pl ", "Panama": "pn ", "Portugal": "po ", "Papua New Guinea": "pp ", "Puerto Rico": "pr ", "Palau": "pw ", "Paraguay": "py ", "Qatar": "qa ", "Queensland": "qea", "Québec (Province)": "quc", "Serbia": "rb ", "Réunion": "re ", "Zimbabwe": "rh ", "RI": "riu", "Rhode Island": "riu", "Romania": "rm ", "Russia (Federation)": "ru ", "Rwanda": "rw ", "South Africa": "sa ", "Saint-Barthélemy": "sc ", "South Carolina": "scu", "South Sudan": "sd ", "South Dakota": "sdu", "Seychelles": "se ", "Sao Tome and Principe": "sf ", "Senegal": "sg ", "Spanish North Africa": "sh ", "Singapore": "si ", "Sudan": "sj ", "Sierra Leone": "sl ", "San Marino": "sm ", "Sint Maarten": "sn ", "Saskatchewan": "snc", "Somalia": "so ", "Spain": "sp ", "Eswatini": "sq ", "Surinam": "sr ", "Western Sahara": "ss ", "Saint-Martin": "st ", "Scotland": "stk", "Saudi Arabia": "su ", "Sweden": "sw ", "Namibia": "sx ", "Syria": "sy ", "Switzerland": "sz ", "Tajikistan": "ta ", "Turks and Caicos Islands": "tc ", "Togo": "tg ", "Thailand": "th ", "Tunisia": "ti ", "Turkmenistan": "tk ", "Tokelau": "tl ", "Tasmania": "tma", "Tennessee": "tnu", "Tonga": "to ", "Trinidad and Tobago": "tr ", "United Arab Emirates": "ts ", "Turkey": "tu ", "Tuvalu": "tv ", "Texas": "txu", "Tanzania": "tz ", "Egypt": "ua ", "United States Misc. Caribbean Islands": "uc ", "Uganda": "ug ", "Ukraine": "un ", "United States Misc. Pacific Islands": "up ", "Utah": "utu", "Burkina Faso": "uv ", "Uruguay": "uy ", "Uzbekistan": "uz ", "Virginia": "vau", "British Virgin Islands": "vb ", "Vatican City": "vc ", "Venezuela": "ve ", "Virgin Islands of the United States": "vi ", "Vietnam": "vm ", "Various places": "vp ", "Victoria": "vra", "Vermont": "vtu", "WA": "wau", "Washington": "wau", "Western Australia": "wea", "Wallis and Futuna": "wf ", "Wisconsin": "wiu", "West Bank of the Jordan River": "wj ", "Wake Island": "wk ", "Wales": "wlk", "Samoa": "ws ", "WV": "wvu", "West Virginia": "wvu", "Wyoming": "wyu", "Christmas Island (Indian Ocean)": "xa ", "Cocos (Keeling) Islands": "xb ", "Maldives": "xc ", "Saint Kitts-Nevis": "xd ", "Marshall Islands": "xe ", "Midway Islands": "xf ", "Coral Sea Islands Territory": "xga", "Niue": "xh ", "Saint Helena": "xj ", "Saint Lucia": "xk ", "Saint Pierre and Miquelon": "xl ", "Saint Vincent and the Grenadines": "xm ", "North Macedonia": "xn ", "New South Wales": "xna", "Slovakia": "xo ", "Northern Territory": "xoa", "Spratly Island": "xp ", "Czech Republic": "xr ", "South Australia": "xra", "South Georgia and the South Sandwich Islands": "xs ", "Slovenia": "xv ", "No place, unknown, or undetermined": "xx ", "Canada": "xxc", "United Kingdom": "xxk", "United States": "xxu", "Yemen": "ye ", "Yukon Territory": "ykc", "Zambia": "za "}
    if "Place of publication not identified" in string:
        country = "No place, unknown, or undetermined"
    else:
        country = re.sub(r'\[', r'', string)
        country = re.sub(r'\]', r'', country)
        country = re.sub(r'^.*, ([a-zA-Z \.]*)$', r'\g<1>', country)
    if not country:
        country = "No place, unknown, or undetermined"
    country_code = country_dict[country]
    return country_code
    

def name_direct_order(string):
    if string.endswith(','):
        string = string.rstrip(',')
    last = re.sub(r'^(.*?), (.*)$', r'\g<1>', string)
    first = re.sub(r'^(.*?), (.*)$', r'\g<2>', string)
    name = first + ' ' + last
    return name


def main(arglist):
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='path to spreadsheet')
    # parser.add_argument('output', help='save directory')
    args = parser.parse_args(arglist)
    
    input = Path(args.input)
    
    # Read spreadsheet
    book_in = xlrd.open_workbook(str(input))
    sheet = book_in.sheet_by_index(0)  # get first sheet
    col_headers = sheet.row_values(0)
    
    title_col = col_headers.index('Title')
    subj_person_col = col_headers.index('Subject_Person')
    subj_topical_col = col_headers.index('Subject_Topical')
    subj_place_col = col_headers.index('Subject_Place')
    subj_corp_col = col_headers.index('Subject_Jurisdiction')
    genre_col = col_headers.index('Genre')
    pages_col = col_headers.index('Page Count')
    pub_date_col = col_headers.index('Publication Date')
    copy_date_col = col_headers.index('Copyright Date')
    pub_place_col = col_headers.index('Place of Publication')
    publisher_col = col_headers.index('Publisher')
    edition_col = col_headers.index('Edition')
    source_col = col_headers.index('Source')
    source_acq_col = col_headers.index('Source of Acquisition')
    writer_col = col_headers.index('Writer')
    penciller_col = col_headers.index('Penciller')
    inker_col = col_headers.index('Inker')
    colorist_col = col_headers.index('Colorist')
    letterer_col = col_headers.index('Letterer')
    cover_artist_col = col_headers.index('Cover Artist')
    editor_col = col_headers.index('Editor')
    # hist_note_col = col_headers.index('Historical Note')
    notes_col = col_headers.index('Notes')
    characters_col = col_headers.index('Characters')
    synopsis_col = col_headers.index('Synopsis')
    toc_col = col_headers.index('Table of Contents')
    in_series_col = col_headers.index('Is Part of Series')
    black_creators_col = col_headers.index('Black Creators (MARC 590)')
    black_chars_col = col_headers.index('Black Characters (MARC 590)')
    isbn_col = col_headers.index('ISBN')
    color_col = col_headers.index('Color?')
    series_note_col = col_headers.index('Series Note')
    copyright_holder_col = col_headers.index('Copyright holder')
    gcd_uri_col = col_headers.index('GCD URL')
    copies_col = col_headers.index('Copies')
    
    outmarc = open('records_Flota.mrc', 'wb')
    
    # Boilerplate fields
    field_ldr = '00000nam a2200000Ii 4500'
    field_040 = Field(tag = '040',
                indicators = [' ',' '],
                subfields = [
                    'a', 'VMC',
                    'b', 'eng',
                    'e', 'rda',
                    'c', 'VMC'])
    field_049 = Field(tag = '049',
                indicators = [' ',' '],
                subfields = [
                    'a', 'VMCS'])
    field_336_text = Field(tag = '336',
                    indicators = [' ',' '],
                    subfields = [
                        'a', 'text',
                        'b', 'txt',
                        '2', 'rdacontent'])
    field_336_image = Field(tag = '336',
                indicators = [' ',' '],
                subfields = [
                    'a', 'still image',
                    'b', 'sti',
                    '2', 'rdacontent'])
    field_337 = Field(tag = '337',
                indicators = [' ',' '],
                subfields = [
                    'a', 'unmediated',
                    'b', 'n',
                    '2', 'rdamedia'])
    field_338 = Field(tag = '338',
                indicators = [' ',' '],
                subfields = [
                    'a', 'volume',
                    'b', 'nc',
                    '2', 'rdacarrier'])
    field_380 = Field(tag = '380',
                indicators = [' ',' '],
                subfields = [
                    'a', 'Comic books and graphic novels.'])
    field_506 = Field(tag = '506',
                    indicators = ['1',' '],
                    subfields = [
                        'a', 'Collection open to research. Researchers must register and agree to copyright and privacy laws before using this collection. Please contact Research Services staff before visiting the James Madison University Special Collections Library to use this collection.'])
    field_541_flota1 = Field(tag = '541',
                indicators = ['1',' '],
                subfields = [
                    'c', 'Gift;',
                    'a', 'Brian Flota.'])
    field_541_flota2 = Field(tag = '541',
                indicators = ['1',' '],
                subfields = [
                    'a', "Brian Flota donated his personal collection of approximately 2,700 comic books in March 2015. In July 2016, Brian Flota donated Bradley Flota's collection of approximately 7,000 comic books."])
    field_542 = Field(tag = '542',
                indicators = [' ',' '],
                subfields = [
                    'a', 'Copyright not evaluated',
                    'u', 'http://rightsstatements.org/vocab/CNE/1.0/'])
    field_545 = Field(tag = '545',
                indicators = ['0',' '],
                subfields = [
                    'a', "This collection was assembled by Dr. Brian Flota, Humanities Librarian at James Madison University, and his father, Bradley Flota (1948-2015). Bradley Flota graduated from Mount Vernon (Illinois) Township High School in 1966 and served in the United States Army from 1970 to 1972. He was also a musician, contributing vocals and guitar intermittently to The Time Actions, who became A&M Records recording artists, Head East, from 1969 to 1973. He was a police officer for the Mount Vernon Police Department from 1974 to 1998, serving as a D.A.R.E. (Drug Abuse Resistance Education) officer for the last eight years of his career with the department. Though Bradley Flota had been collecting comic books since childhood, those collections would fall out of his possession. He gave his young son Brian a selection of about one hundred comics in 1983, which formed the basis of Brian's collection. In 1985, Bradley Flota began collecting the comics that make up this collection. He collected comics voraciously from 1985 to 1997. From that time until 2008, he collected more infrequently. Brian Flota began collecting in earnest in 1987 at the age of eleven, eventually amassing around 2,700 comics when he ceased collecting around 1998. Of those 2,700 comics, Brian Flota donated a portion of the collection to the University of California, Riverside prior to his donation to JMU Special Collections in 2015. Brian Flota holds his MS in Library and Information Sciences from the University of Illinois at Urbana-Champaign, his Ph.D. in American Literature from The George Washington University, and his B.A. in English from the University of California, Riverside. He is also the co-editor of The Politics of Post-9/11 Music, published by Ashgate in 2011."])
    field_555 = Field(tag = '555',
                indicators = ['0',' '],
                subfields = [
                    'a', 'View detailed inventory and request for use in Special Collections:',
                    'u', 'https://aspace.lib.jmu.edu/repositories/4/resources/212'])
    field_588 = Field(tag = '588',
                indicators = ['0',' '],
                subfields = [
                    'a', 'Description based on indicia and Grand Comics Database.'])
    field_655_lcgft = Field(tag = '655',
                indicators = [' ','7'],
                subfields = [
                    'a', 'Comics (Graphic works).',
                    '2', 'lcgft'])
    field_989 = Field(tag = '989',
                indicators = [' ',' '],
                subfields = [
                    'a', 'PN6728'])
    
    for row in range(1, sheet.nrows):
        print('Record ' + str(row))
        
        title = sheet.cell(row, title_col).value
        print(title)
        
        subj_person = sheet.cell(row, subj_person_col).value
        if subj_person:
            subj_person = [x.strip() for x in subj_person.split(';')]
        subj_topical = sheet.cell(row, subj_topical_col).value
        if subj_topical:
            subj_topical = [x.strip() for x in subj_topical.split(';')]
        subj_place = sheet.cell(row, subj_place_col).value
        if subj_place:
            subj_place = [x.strip() for x in subj_place.split(';')]
        subj_corp = sheet.cell(row, subj_corp_col).value
        if subj_corp:
            subj_corp = [x.strip() for x in subj_corp.split(';')]
        genre = sheet.cell(row, genre_col).value
        genre = [x.strip() for x in genre.split(';')]
        pages = str(sheet.cell(row, pages_col).value)
        
        copy_date = ''
        copy_date = str(sheet.cell(row, copy_date_col).value)
        copy_date_str = date_from_string(copy_date)
        copy_date_year = year_from_date(copy_date_str)        
        pub_date = str(sheet.cell(row, pub_date_col).value)
        if not pub_date:
            pub_date = '[' + copy_date_year + ']'
        pub_date_str = date_from_string(pub_date)
        pub_date_year = year_from_date(pub_date_str)
        
        pub_place = sheet.cell(row, pub_place_col).value
        publisher = sheet.cell(row, publisher_col).value
        edition = sheet.cell(row, edition_col).value
        source = sheet.cell(row, source_col).value
        source_acq = sheet.cell(row, source_acq_col).value
        characters = sheet.cell(row, characters_col).value
        black_creators = sheet.cell(row, black_creators_col).value
        if black_creators:
            black_creators = [x.strip() for x in black_creators.split(';')]
        black_chars = sheet.cell(row, black_chars_col).value
        if black_chars:
            black_chars = [x.strip() for x in black_chars.split(';')]
        isbn = str(sheet.cell(row, isbn_col).value)
        color = sheet.cell(row, color_col).value
        series_note = sheet.cell(row, series_note_col).value
        series_note = [x.strip() for x in series_note.split(';')]
        gcd_uri = sheet.cell(row, gcd_uri_col).value
        
        country_code = country_code_from_pub_place(pub_place)
        
        copyright_holder = []
        if sheet.cell(row, copyright_holder_col).value:
            copyright_holder = sheet.cell(row, copyright_holder_col).value
            copyright_holder = [x.strip() for x in copyright_holder.split(';')]
        writer = []
        if sheet.cell(row, writer_col).value:
            writer = sheet.cell(row, writer_col).value
            writer = [x.strip() for x in writer.split(';')]
        penciller = []
        if sheet.cell(row, penciller_col).value:
            penciller = sheet.cell(row, penciller_col).value
            penciller = [x.strip() for x in penciller.split(';')]
        inker = []
        if sheet.cell(row, inker_col).value:
            inker = sheet.cell(row, inker_col).value
            inker = [x.strip() for x in inker.split(';')]
        colorist = []
        if sheet.cell(row, colorist_col).value:
            colorist = sheet.cell(row, colorist_col).value
            colorist = [x.strip() for x in colorist.split(';')]
        letterer = []
        if sheet.cell(row, letterer_col).value:
            letterer = sheet.cell(row, letterer_col).value
            letterer = [x.strip() for x in letterer.split(';')]
        cover_artist = []
        if sheet.cell(row, cover_artist_col).value:
            cover_artist = sheet.cell(row, cover_artist_col).value
            cover_artist = [x.strip() for x in cover_artist.split(';')]
        editor = []
        if sheet.cell(row, editor_col).value:
            editor = sheet.cell(row, editor_col).value
            editor = [x.strip() for x in editor.split(';')]
        # hist_note = []
        # if sheet.cell(row, hist_note_col).value:
            # hist_note = sheet.cell(row, hist_note_col).value
        notes = []
        if sheet.cell(row, notes_col).value:
            notes = sheet.cell(row, notes_col).value
        synopsis = []
        if sheet.cell(row, synopsis_col).value:
            synopsis = sheet.cell(row, synopsis_col).value
        toc = []
        if sheet.cell(row, toc_col).value:
            toc = sheet.cell(row, toc_col).value
        in_series = sheet.cell(row, in_series_col).value
        
        contribs = {}
        if copyright_holder:
            for i in copyright_holder:
                contribs.update({i: ['producer']})
        else:
            if writer:
                for i in writer:
                    contribs.update({i: ['writer']})
            if penciller:
                for i in penciller:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('penciller')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['penciller']})
            if inker:
                for i in inker:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('inker')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['inker']})
            if colorist:
                for i in colorist:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('colorist')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['colorist']})
            if letterer:
                for i in letterer:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('letterer')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['letterer']})
            if cover_artist:
                for i in cover_artist:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('cover artist')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['cover artist']})
            if editor:
                for i in editor:
                    if i in contribs:
                        role_list = contribs[i]
                        role_list.append('editor')
                        contribs.update({i: role_list})
                    else:
                        contribs.update({i: ['editor']})
        
        record = Record()
        
        # Add boilerplate fields
        record.leader = field_ldr
        record.add_ordered_field(field_040)
        record.add_ordered_field(field_049)
        record.add_ordered_field(field_336_text)
        record.add_ordered_field(field_336_image)
        record.add_ordered_field(field_337)
        record.add_ordered_field(field_338)
        record.add_ordered_field(field_380)
        record.add_ordered_field(field_506)
        record.add_ordered_field(field_541_flota1)
        record.add_ordered_field(field_541_flota2)
        record.add_ordered_field(field_542)
        record.add_ordered_field(field_545)
        record.add_ordered_field(field_555)
        record.add_ordered_field(field_588)
        record.add_ordered_field(field_655_lcgft)
        record.add_ordered_field(field_989)        
        
        # Add other fields
        today = datetime.today().strftime('%y%m%d')
        if copy_date:
            data_008 = today + 't' + pub_date_year + copy_date_year + country_code + 'a     6    000 1 eng d'
        else:
            data_008 = today + 's' + pub_date_year + '    ' + country_code + 'a     6    000 1 eng d'
        field_008 = Field(tag = '008',
                    data = data_008)
        record.add_ordered_field(field_008)
        
        if isbn:
            field_020 = Field(tag = '020',
                        indicators = [' ',' '],
                        subfields = [
                            'a', isbn])
            record.add_ordered_field(field_020)
        
        
        subfields_099 = subfields_from_string(title)
        if 'b' in subfields_099:
            subfields_099.pop(3)
            subfields_099.pop(2)
        if 'p' in subfields_099:
            subfields_099[1] = subfields_099[1] + ' ' + subfields_099[subfields_099.index('p') + 1]  # add subfield p content to subfield a content
            subfields_099.pop(subfields_099.index('p') + 1)  # remove subfield p content
            subfields_099.pop(subfields_099.index('p'))  # remove subfield p
        if 'n' in subfields_099:
            subfields_099[subfields_099.index('n')] = 'a'
        if subfields_099[1].endswith(',') or subfields_099[1].endswith(':'):
            subfields_099[1] = subfields_099[1][:-1]
        field_099 = Field(tag = '099',
                    indicators = [' ','9'],
                    subfields = subfields_099)
        record.add_ordered_field(field_099)
        
        for i in contribs:
            if i == list(contribs.keys())[0] and 'producer' in contribs[i]: # first contributor is copyright holder
                subfield_content = subfields_from_string_relator(i, contribs[i])
                field_710 = Field(tag = '710',
                        indicators = ['2', ' '],
                        subfields = subfield_content)
                record.add_ordered_field(field_710)
            elif i == list(contribs.keys())[0] and 'writer' in contribs[i]: # first contributor is a writer
                subfield_content = subfields_from_string_relator(i, contribs[i])
                field_100 = Field(tag = '100',
                        indicators = ['1', ' '],
                        subfields = subfield_content)
                record.add_ordered_field(field_100)
            else:
                subfield_content = subfields_from_string_relator(i, contribs[i])
                if ',' not in subfield_content[1]:
                    field_710 = Field(tag = '710',
                                indicators = ['2',' '],
                                subfields = subfield_content)
                    record.add_ordered_field(field_710)
                else:
                    field_700 = Field(tag = '700',
                                indicators = ['1',' '],
                                subfields = subfield_content)
                    record.add_ordered_field(field_700)
        
        if contribs and 'writer' in contribs[list(contribs.keys())[0]]:
            f245_ind1 = 1
        else:
            f245_ind1 = 0
        
        f245_ind2 = 0
        if str.startswith(title, 'The '):
            f245_ind2 = 4
        elif str.startswith(title, 'An '):
            f245_ind2 = 3
        elif str.startswith(title, 'A '):
            f245_ind2 = 2
        
        subfields_245 = subfields_from_string(title)
        
        # If no contribs, or only copyright holder ("producer"), add 245 ending punctuation
        # Otherwise add $c with all other contributors
        if (not contribs) or ('producer' in contribs[i]):
            subfields_245[-1] = subfields_245[-1] + '.'
        else:
            subfields_245[-1] = subfields_245[-1] + ' /'
            subfields_245.append('c')
            
            # iterate through all contribs for names and roles
            contribs_245_list = []
            contribs_245 = ''
            for i in contribs:
                indiv_contrib = subfields_from_string_relator(i, contribs[i])
                for j in indiv_contrib:
                    contribs_245_list.append(j)
            for (x,y) in zip(range(len(contribs_245_list)), contribs_245_list):
                if y == 'a':
                    contribs_245 += name_direct_order(contribs_245_list[x+1])
                    contribs_245 += ', '
                if y == 'e':
                    if contribs_245_list[x+1].endswith(','):
                        contribs_245_list[x+1] += ' '
                    if contribs_245_list[x+1].endswith('.'):
                        contribs_245_list[x+1] = contribs_245_list[x+1].rstrip('.') + ', '
                    contribs_245 += contribs_245_list[x+1]
            contribs_245 = contribs_245.rstrip(', ')
            contribs_245 += '.'
            subfields_245.append(contribs_245)
            
        field_245 = Field(tag = '245',
                    indicators = [f245_ind1, f245_ind2],
                    subfields = subfields_245)
        record.add_ordered_field(field_245)
        
        if edition:
            if not edition.endswith('.'):
                edition += '.'
            field_250 = Field(tag = '250',
                    indicators = [' ', ' '],
                    subfields = [
                        'a', edition])
            record.add_ordered_field(field_250)
        
        subfields_264_1 = [
                        'a', pub_place + ' :',
                        'b', publisher + ',',
                        'c', pub_date_str + '.']
        if subfields_264_1[5].endswith('].'):
            subfields_264_1[5] = subfields_264_1[5][:-1]
        field_264_1 = Field(tag = '264',
                    indicators = [' ','1'],
                    subfields = subfields_264_1)
        record.add_ordered_field(field_264_1)
        
        if copy_date:
            field_264_4 = Field(tag = '264',
                        indicators = [' ','4'],
                        subfields = [
                            'c', '©' + copy_date_str])
            record.add_ordered_field(field_264_4)
        
        if color == 'yes' or color == 'Yes' or color == 'y' or color == 'Y':
            subfields_300 = [
                'a', pages + ' pages :',
                'b', 'color illustrations.']
        elif color == 'no' or color == 'No' or color == 'n' or color == 'N':
            subfields_300 = [
                'a', pages + ' pages :',
                'b', 'black and white illustrations.']
        
        field_300 = Field(tag = '300',
                    indicators = [' ',' '],
                    subfields = subfields_300)
        record.add_ordered_field(field_300)
        
        title_490 = subfields_from_string(title)[1]
        if title_490.endswith(',') or title_490.endswith(':'):
                title_490 = title_490[:-1]
        field_490 = Field(tag = '490',
                    indicators = ['0',' '],
                    subfields = [
                        'a', title_490])
        record.add_ordered_field(field_490)
        
        if series_note:
            # if not series_note.endswith('.'):
                # series_note += '.'
            for i in series_note:
                field_490_series_note = Field(tag = '490',
                                        indicators = ['0', ' '],
                                        subfields = ['a', i])
                record.add_ordered_field(field_490_series_note)
        
        # if hist_note:
        #     field_500_hist = Field(tag = '500',
        #                 indicators = [' ',' '],
        #                 subfields = [
        #                     'a', hist_note + '.'])
        #     record.add_ordered_field(field_500_hist)
        
        if notes:
            field_500_notes = Field(tag = '500',
                        indicators = [' ',' '],
                        subfields = [
                            'a', notes + '.'])
            record.add_ordered_field(field_500_notes)
        
        if toc:
            if not toc.endswith('.') and not toc.endswith('?') and not toc.endswith('!'):
                toc += '.'
            field_505 = Field(tag = '505',
                        indicators = ['0',' '],
                        subfields = [
                            'a', toc])
            record.add_ordered_field(field_505)
        
        if synopsis:
            if synopsis.endswith('-- Grand Comics Database.') or synopsis.endswith('-- Grand Comics Database'):
                subfield_a_520 = synopsis.rstrip('-- Grand Comics Database.')
                subfield_a_520 = synopsis.rstrip('-- Grand Comics Database')
                field_520 = Field(tag = '520',
                            indicators = [' ',' '],
                            subfields = [
                                'a', subfield_a_520,
                                'c', 'Grand Comics Database'])
            else:
                field_520 = Field(tag = '520',
                            indicators = [' ',' '],
                            subfields = [
                                'a', synopsis])
            record.add_ordered_field(field_520)
        
        if black_creators:
            for i in black_creators:
                if not i.endswith('.'):
                    i += '.'
                field_590_creators = Field(tag = '590',
                            indicators = [' ',' '],
                            subfields = [
                                'a', i])
                record.add_ordered_field(field_590_creators)
        
        if black_chars:
            for i in black_chars:
                if not i.endswith('.'):
                    i += '.'
                field_590_chars = Field(tag = '590',
                            indicators = [' ',' '],
                            subfields = [
                                'a', i])
                record.add_ordered_field(field_590_chars)
        
        if source:
            field_541_source = Field(tag = '541',
                        indicators = [' ',' '],
                        subfields = [
                            'a', source + '.'])
            record.add_ordered_field(field_541_source)
        
        if source_acq:
            field_541_source_acq = Field(tag = '541',
                        indicators = [' ',' '],
                        subfields = [
                            'a', source_acq + '.'])
            record.add_ordered_field(field_541_source_acq)
        
        if subj_person:
            for i in subj_person:
                i_subfields = subfields_from_string(i)
                
                # Set first indicator based on presence of comma in $a
                if 'a' in i_subfields:
                    if ',' in i_subfields[i_subfields.index('a') + 1]:
                        field_600_ind1 = '1'
                    else:
                        field_600_ind1 = '0'
                
                if '1' in i_subfields:
                    last_except_subf1 = i_subfields.index('1') - 1
                else:
                    last_except_subf1 = len(i_subfields) - 1
                
                if i_subfields[last_except_subf1].endswith(','):
                    i_subfields[last_except_subf1] = re.sub(r'^(.*),$', r'\g<1>.', i_subfields[last_except_subf1])
                if not i_subfields[last_except_subf1].endswith('.') and not i_subfields[last_except_subf1].endswith(')') and not i_subfields[last_except_subf1].endswith('?') and not i_subfields[last_except_subf1].endswith('-'):
                    i_subfields[last_except_subf1] += '.'
                
                field_600 = Field(tag = '600', 
                            indicators = [field_600_ind1,'0'],
                            subfields = i_subfields)
                record.add_ordered_field(field_600)
        
        if subj_topical:
            for i in subj_topical:
                i_subfields = subfields_from_string(i)
                if not i_subfields[-1].endswith('.') and not i_subfields[-1].endswith(')'):
                    i_subfields[-1] += '.'
                field_650 = Field(tag = '650',
                            indicators = [' ','0'],
                            subfields = i_subfields)
                record.add_ordered_field(field_650)
        
        if subj_place:
            for i in subj_place:
                i_subfields = subfields_from_string(i)
                if not i_subfields[-1].endswith('.') and not i_subfields[-1].endswith(')'):
                    i_subfields[-1] += '.'
                field_651 = Field(tag = '651',
                        indicators = [' ','0'],
                        subfields = i_subfields)
                record.add_ordered_field(field_651)
        
        if subj_corp:
            for i in subj_corp:
                i_subfields = subfields_from_string(i)
                if not i_subfields[-1].endswith('.') and not i_subfields[-1].endswith(')'):
                    i_subfields[-1] += '.'
                field_610 = Field(tag = '610',
                        indicators = ['1','0'],
                        subfields = i_subfields)
                record.add_ordered_field(field_610)
        
        if genre:
            for i in genre:
                if not i.endswith('.') and not i.endswith(')'):
                    i += '.'
                field_655 = Field(tag = '655',
                        indicators = [' ','7'],
                        subfields = [
                            'a', i,
                            '2', 'lcgft'])
                record.add_ordered_field(field_655)
        
        if characters:
            field_500_chars = Field(tag = '500',
                        indicators = [' ', ' '],
                        subfields = [
                            'a', characters])
            record.add_ordered_field(field_500_chars)
        
        if gcd_uri:
            title_758 = subfields_from_string(title)[1]
            if title_758.endswith(',') or title_758.endswith(':'):
                title_758 = title_758[:-1]
            field_758 = Field(tag = '758',
                        indicators = [' ',' '],
                        subfields = [
                            '4', 'http://rdaregistry.info/Elements/m/P30135',
                            'i', 'Has work manifested:',
                            'a', title_758,
                            '1', gcd_uri])
            record.add_ordered_field(field_758)
        
        if in_series:
            subfields_773 = subfields_from_string_series(in_series)
            field_773 = Field(tag = '773',
                        indicators = ['0','8'],
                        subfields = subfields_773)
            record.add_ordered_field(field_773)
        
        subfields_852 = [
            'b', 'CARRIER',
            'c', 'carrspec']
        if len(subfields_099) == 4:
            subfields_852.append('h')
            subfields_852.append(subfields_099[1])
            subfields_852.append('i')
            subfields_852.append(subfields_099[3])
        if len(subfields_099) == 2:
            subfields_852.append('h')
            subfields_852.append(subfields_099[1])
        if edition:
            if edition.endswith('.'):
                edition = edition[:-1]
            subfields_852.append('z')
            subfields_852.append(edition)
        
        field_852 = Field(tag = '852',
                    indicators = ['8',' '],
                    subfields = subfields_852)
        record.add_ordered_field(field_852)
        
        outmarc.write(record.as_marc())
        print()
    outmarc.close()
    
    
if __name__ == '__main__':
    main(sys.argv[1:])
