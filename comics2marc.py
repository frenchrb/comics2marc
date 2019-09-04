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


def lowercase_title(string):
    title = string
    title = re.sub(r'Vol', r'vol', title)
    title = re.sub(r'No\.', r'no.', title)
    return title


def subfields_from_string(string):
    # print('STRING IS:', string)
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
    # print()
    return subfields


def subfields_from_string_relator(string, relator):
    # print('STRING IS:', string)
    subfields = []
    if '$' in string:
        string = string.split('$')
        # print(string)
        # print(len(string) - 1)
        # print(string[len(string) - 1])
        if not string[len(string) - 2].endswith(',') and not string[len(string) - 2].endswith('-'):
            string[len(string) - 2] += ','
        string.insert(len(string) - 1, 'e' + relator + '.')
        # print(i)
        for x in string:
            if string.index(x) == 0:
                subfields.append('a')
                subfields.append(x)
            else:
                subfields.append(x[0:1])
                subfields.append(x[1:])
    else:
        if not string.endswith(',') and not string.endswith('-'):
            string += ','
        subfields.append('a')
        subfields.append(string)
        subfields.append('e')
        subfields.append(relator + '.')
    # print()
    return subfields


def name_direct_order(string):
    last = re.sub(r'^(.*?), (.*),$', r'\g<1>', string)
    first = re.sub(r'^(.*?), (.*),$', r'\g<2>', string)
    name = first + ' ' + last
    return name


def main(arglist):
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='path to spreadsheet')
    # parser.add_argument('output', help='save directory')
    # parser.add_argument('--production', help='production DOIs', action='store_true')
    args = parser.parse_args(arglist)
    
    input = Path(args.input)
    
    # Read spreadsheet
    book_in = xlrd.open_workbook(str(input))
    sheet = book_in.sheet_by_index(0)  # get first sheet
    col_headers = sheet.row_values(0)
    # print(col_headers)
    # print()
    
    title_col = col_headers.index('Title')
    subj_col = col_headers.index('Subject')
    genre_col = col_headers.index('Genre')
    pages_col = col_headers.index('Pages')
    date_col = col_headers.index('Date')
    pub_place_col = col_headers.index('Pub_Place')
    publisher_col = col_headers.index('Publisher')
    source_col = col_headers.index('Source')
    writer_col = col_headers.index('Writer')
    penciller_col = col_headers.index('Penciller')
    inker_col = col_headers.index('Inker')
    colorist_col = col_headers.index('Colorist')
    letterer_col = col_headers.index('Letterer')
    cover_artist_col = col_headers.index('Cover Artist')
    editor_col = col_headers.index('Editor')
    hist_note_col = col_headers.index('Historical Note')
    note_col = col_headers.index('Note')
    characters_col = col_headers.index('Characters')
    story_arc_col = col_headers.index('Story Arc')
    toc_col = col_headers.index('Table of Contents')
    series_col = col_headers.index('Is Part of Series')
    
    outmarc = open('records.mrc', 'wb')
    
    # Boilerplate fields
    field_ldr = '00000nam  2200000Ii 4500'
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
                    'a', 'VMCM'])
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
    field_542 = Field(tag = '542',
                indicators = [' ',' '],
                subfields = [
                    'a', 'Copyright not evaluated',
                    'u', 'http://rightsstatements.org/vocab/CNE/1.0/'])
    field_588 = Field(tag = '588',
                indicators = ['0',' '],
                subfields = [
                    'a', 'Description based on indicia and Grand Comics Database.'])
    field_989 = Field(tag = '989',
                indicators = [' ',' '],
                subfields = [
                    'a', 'PN6728'])
    
    for row in range(1, sheet.nrows):
        print('Record ' + str(row))
        
        title = sheet.cell(row, title_col).value
        print(title)
        lower_title = parse_title(lowercase_title(title))
        title = parse_title(sheet.cell(row, title_col).value)
        has_part_title = False
        if len(title) == 3:
            has_part_title = True
        
        subj = sheet.cell(row, subj_col).value
        subj = [x.strip() for x in subj.split(';')]
        genre = sheet.cell(row, genre_col).value
        genre = [x.strip() for x in genre.split(';')]
        pages = sheet.cell(row, pages_col).value
        date = sheet.cell(row, date_col).value[0:4]
        pub_place = sheet.cell(row, pub_place_col).value
        publisher = sheet.cell(row, publisher_col).value
        source = sheet.cell(row, source_col).value
        # writer = sheet.cell(row, writer_col).value
        
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
            # print(colorist)
            # print('COLORIST FROM SHEET=' + colorist + '=END')
            # print(bool(colorist))
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
        hist_note = []
        if sheet.cell(row, hist_note_col).value:
            hist_note = sheet.cell(row, hist_note_col).value
        note = []
        if sheet.cell(row, note_col).value:
            note = sheet.cell(row, note_col).value
        characters = []
        if sheet.cell(row, characters_col).value:
            characters  = sheet.cell(row, characters_col).value
            characters = [x.strip() for x in characters.split(';')]
        story_arc = []
        if sheet.cell(row, story_arc_col).value:
            story_arc = sheet.cell(row, story_arc_col).value
        toc = []
        if sheet.cell(row, toc_col).value:
            toc = sheet.cell(row, toc_col).value
        series = sheet.cell(row, series_col).value
        
        # print(cover_artist)
        # print(characters)
        # print(writer)
        # print(subfields_from_string(writer[0]))
        # print(name_direct_order(subfields_from_string(writer[0])[1]))
        # print(title)
        # print(parse_title(title))
        
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
        record.add_ordered_field(field_542)
        record.add_ordered_field(field_588)
        record.add_ordered_field(field_989)        
        
        # Add other fields
        today = datetime.today().strftime('%y%m%d')
        data_008 = today + 't' + date + date + 'xx a     6    000 1 eng d'
        field_008 = Field(tag = '008',
                    data = data_008)
        record.add_ordered_field(field_008)
        
        subfields_099 = []
        if has_part_title:
            subfields_099 = [
                'a', title[0] + ': ' + title[1],
                'a', title[2]]
        else:
            subfields_099 = [
                'a', title[0],
                'a', title[1]]
        field_099 = Field(tag = '099',
                    indicators = [' ','9'],
                    subfields = subfields_099)
        record.add_ordered_field(field_099)
                  
        if writer:
            # Add 100 for first writer
            subfield_content = subfields_from_string_relator(writer[0], 'writer')
            field_100 = Field(tag = '100',
                    indicators = ['1', ' '],
                    subfields = subfield_content)
            record.add_ordered_field(field_100)
            # Multiple writers
            if len(writer)>1:
                # Add 700s for all writers after the first
                for i in writer[1:]:
                    subfield_content = subfields_from_string_relator(i, 'writer')
                    field_700 = Field(tag = '700',
                                indicators = ['1',' '],
                                subfields = subfield_content)
                    record.add_ordered_field(field_700)                    
        
        if writer:
            f245_ind1 = 1
        else:
            f245_ind1 = 0
        
        f245_ind2 = 0
        if str.startswith(title[0], 'The '):
            f245_ind2 = 4
        elif str.startswith(title[0], 'An '):
            f245_ind2 = 3
        elif str.startswith(title[0], 'A '):
            f245_ind2 = 2
        
        subfields_245 = []
        if has_part_title:
            subfields_245 = [
                'a', title[0] + '.',
                'p', title[1] + ',',
                'n', title[2]]
        else:
            subfields_245 = [
                'a', title[0] + ',',
                'n', title[1]]
        # If writer exists, add $c
        if writer:
            subfields_245[-1] = subfields_245[-1] + ' /'
            subfields_245.append('c')
            subfields_245.append(name_direct_order(subfields_from_string(writer[0])[1]) + ', writer.')
        else:
            # If no writer, add 245 ending punctuation
            subfields_245[-1] = subfields_245[-1] + '.'
        field_245 = Field(tag = '245',
                    indicators = [f245_ind1, f245_ind2],
                    subfields = subfields_245)
        record.add_ordered_field(field_245)
        
        field_264_1 = Field(tag = '264',
                    indicators = [' ','1'],
                    subfields = [
                        'a', pub_place + ' :',
                        'b', publisher + ',',
                        'c', date + '.'])
        record.add_ordered_field(field_264_1)
        
        field_264_4 = Field(tag = '264',
                    indicators = [' ','4'],
                    subfields = [
                        'c', '©' + date])
        record.add_ordered_field(field_264_4)
        
        field_300 = Field(tag = '300',
                    indicators = [' ',' '],
                    subfields = [
                        'a', pages + ' pages :',
                        'b', 'chiefly color illustrations.'])
        record.add_ordered_field(field_300)
        
        subfields_490 = []
        if has_part_title:
            subfields_490 = [
                'a', lower_title[0] + '. ' + lower_title[1] + ' ;',
                'v', lower_title[2]]
        else:
            subfields_490 = [
                'a', lower_title[0] + ' ;',
                'v', lower_title[1]]
        field_490 = Field(tag = '490',
                    indicators = ['1',' '],
                    subfields = subfields_490)
        record.add_ordered_field(field_490)
        
        if hist_note:
            field_500_hist = Field(tag = '500',
                        indicators = [' ',' '],
                        subfields = [
                            'a', hist_note + '.'])
            record.add_ordered_field(field_500_hist)
        
        if note:
            field_500_note = Field(tag = '500',
                        indicators = [' ',' '],
                        subfields = [
                            'a', note + '.'])
            record.add_ordered_field(field_500_note)
        
        if toc:
            if not toc.endswith('.') and not toc.endswith('?') and not toc.endswith('!'):
                toc += '.'
            field_505 = Field(tag = '505',
                        indicators = ['0',' '],
                        subfields = [
                            'a', toc])
            record.add_ordered_field(field_505)
        
        if story_arc:
            field_520 = Field(tag = '520',
                        indicators = [' ',' '],
                        subfields = [
                            'a', '"' + story_arc + '" -- Grand Comics Database.'])
            record.add_ordered_field(field_520)
        
        field_561 = Field(tag = '561',
                    indicators = [' ',' '],
                    subfields = [
                        'a', source + '.'])
        record.add_ordered_field(field_561)
        
        for i in subj:
            if not i.endswith('.') and not i.endswith(')'):
                i += '.'
            field_650 = Field(tag = '650',
                    indicators = [' ','0'],
                    subfields = [
                        'a', i])
            record.add_ordered_field(field_650)
        
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
            print(characters)
            subfield_content = 'Characters: '
            for i in characters[:-1]:
                subfield_content += i + '; '
            subfield_content += characters[-1] + '.'
            field_500 = Field(tag = '500',
                        indicators = [' ', ' '],
                        subfields = [
                            'a', subfield_content])
            record.add_ordered_field(field_500)
            
            # Create 600 and 650 for "Fictitious character" entries
            # TODO check for existing 650 and don't add if a duplicate
            if any('Fictitious character' in c for c in characters):
                fic_chars = [c for c in characters if 'Fictitious character' in c]
                for i in fic_chars:
                    fic_char_name = re.sub(r'^(.*?) (\(Fictitious character.*\))$', r'\g<1>', i)
                    fic_char_c = re.sub(r'^(.*?) (\(Fictitious character.*\))$', r'\g<2>', i)
                    field_600 = Field(tag = '600',
                                indicators = ['0', '0'],
                                subfields = [
                                    'a', fic_char_name,
                                    'c', fic_char_c])
                    record.add_ordered_field(field_600)
                    
                    field_650 = Field(tag = '650',
                                indicators = [' ', '0'],
                                subfields = [
                                    'a', i])
                    record.add_ordered_field(field_650)
        
        if penciller:
            for i in penciller:
                subfield_content = subfields_from_string_relator(i, 'penciller')
                field_700 = Field(tag = '700',
                            indicators = ['1',' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
        
        if inker:
            for i in inker:
                subfield_content = subfields_from_string_relator(i, 'inker')
                field_700 = Field(tag = '700',
                            indicators = ['1', ' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
        
        if colorist:
            for i in colorist:
                subfield_content = subfields_from_string_relator(i, 'colorist')
                field_700 = Field(tag = '700',
                            indicators = ['1', ' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
        
        if letterer:
            for i in letterer:
                subfield_content = subfields_from_string_relator(i, 'letterer')
                field_700 = Field(tag = '700',
                            indicators = ['1', ' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
              
        if cover_artist:
            for i in cover_artist:
                subfield_content = subfields_from_string_relator(i, 'cover artist')
                field_700 = Field(tag = '700',
                            indicators = ['1',' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
        
        if editor:
            for i in editor:
                subfield_content = subfields_from_string_relator(i, 'editor')
                field_700 = Field(tag = '700',
                            indicators = ['1', ' '],
                            subfields = subfield_content)
                record.add_ordered_field(field_700)
        
        # field_700 = Field(tag = '700',
                    # indicators = ['7',' '],
                    # subfields = [
                        # 'a', doi,
                        # '2', 'doi'])
        
        subfields_773 = subfields_from_string(series)
        field_773 = Field(tag = '773',
                    indicators = ['0','8'],
                    subfields = subfields_773)
        record.add_ordered_field(field_773)
        
        subfields_830 = []
        if has_part_title:
            subfields_830 = [
                'a', lower_title[0] + '.',
                'p', lower_title[1] + ' ;',
                'v', lower_title[2] + '.']
        else:
            subfields_830 = [
                'a', lower_title[0] + ' ;',
                 'v', lower_title[1] + '.']
        field_830 = Field(tag = '830',
                    indicators = [' ','0'],
                    subfields = subfields_830)
        record.add_ordered_field(field_830)
        
        outmarc.write(record.as_marc())
        print()
    outmarc.close()
    
    
if __name__ == '__main__':
    main(sys.argv[1:])
