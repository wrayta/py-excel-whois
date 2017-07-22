from openpyxl import Workbook, load_workbook
import whois
import datetime

def remove_all_but_domain(cell_text):
    # take out 'http://', 'www.', etc.
    return

def get_domain(w):
    if isinstance(w['domain_name'], list):
        domain = w['domain_name'][0]
    else:
        domain = w['domain_name']
    return domain

def check_if_atl_owned(w):
    atl_owned = '---'

    if w['registrar'] > 0:
        if 'CSC' in w['registrar']:
            atl_owned = 'CSC'
        elif 'Mark' in w['registrar']:
            atl_owned = 'MARK'

    elif w['emails'] > 0:
        if '@atlanticrecords.com' in w['emails']:
            atl_owned = w['registrar']

    elif w['org'] > 0:
        if 'Warner Music' in w['org']:
            atl_owned = w['registrar']

    return atl_owned

def get_exp_date(w):

    if isinstance(w['expiration_date'], list):
        exp_date = w['expiration_date'][0]
    else:
        exp_date = w['expiration_date']
    return exp_date


def do_notes_and_registrar(w):
    info = 'Registrant: '

    if w['org'] > 0:
        if 'Warner Music' in w['org']:
            info += 'Warner Music Group'
        elif 'Atlantic Record' in w['org']:
            info += 'Atlantic Records'
        else:
            info += w['org']

    elif w['name'] > 0:
        info += w['name']

    return info

def create_write_to_workbook():
    wb = Workbook()

    ws1 = wb.active
    ws1.title = 'ACTIVE_NEW'

    ws1['A1'] = 'Artist'
    ws1['B1'] = 'Digital Marketer'
    ws1['C1'] = 'Domain'
    ws1['D1'] = 'ATL Own?'
    ws1['E1'] = 'Expiration'
    ws1['F1'] = 'Notes/Registrar'

    return wb

def add_whois_info(wb, cell_row_index, info):
    cell_col_index = 3 # starts at Excel cell col 'C'

    ws_name = wb.sheetnames[0]
    ws1 = wb.get_sheet_by_name(ws_name)

    for i in info:
        ws1.cell(column=cell_col_index, row=cell_row_index, value='{}'.format(i))
        cell_col_index += 1

    return wb

def add_non_whois_info(old_sheet, wb):
    ws_name = wb.sheetnames[0]
    new_sheet = wb.get_sheet_by_name(ws_name)

    for row in old_sheet.iter_rows('A{}:B{}'.format(old_sheet.min_row, old_sheet.max_row)):
        for cell in row:
            new_sheet['{}{}'.format(cell.column, cell.row)] = cell.value
    return wb

def verify_artists_and_digital_marketers(wb):
    artist_roster = load_workbook(filename = 'DigitalRoster.xlsx')
    roster_sheet_name = artist_roster.sheetnames[0]
    roster_sheet = artist_roster.get_sheet_by_name(roster_sheet_name)

    new_domain_list_sheet_name = wb.sheetnames[0]
    new_domain_list_sheet = wb.get_sheet_by_name(new_domain_list_sheet_name)

    # First, store all of the artists and dms
    artist_to_dm_dict = {} # data structure to keep track of 'digital roster' artists -> dm
    for artist_cell in roster_sheet['C']:
        artist_to_dm_dict[artist_cell.value] = roster_sheet['B{}'.format(artist_cell.row)].value
    #     for dm_cell in roster_sheet['B']:
    #          artist_to_dm_dict[artist_cell.value] = dm_cell.value

    domain_artist_to_dm_dict = {} # data structure to keep track of 'new domain book' artists -> dm
    for artist_cell in new_domain_list_sheet['A']:
        domain_artist_to_dm_dict[artist_cell.value] = new_domain_list_sheet['B{}'.format(artist_cell.row)].value
        # for dm_cell in new_domain_list_sheet['B']:
        #     domain_artist_to_dm_dict[artist_cell.value] = dm_cell.value

    # Second, loop through and see if artists in 'artist_to_dm_dict' are also contained in 'domain_artist_to_dm_dict'
    for artist, artist_dm in artist_to_dm_dict.iteritems():
        if domain_artist_to_dm_dict.get(artist) == None:
            # Add artist and dm to new workbook sheet
            max_row = new_domain_list_sheet.max_row
            new_domain_list_sheet.cell(row=max_row + 1, column=1, value='{}'.format(artist))
            new_domain_list_sheet.cell(row=max_row + 1, column=2, value='{}'.format(artist_dm))

        elif domain_artist_to_dm_dict.get(artist) != artist_dm:
            # Update new workbook sheet with dm
            print('Update DM 1')
            for artist_cell in new_domain_list_sheet['A']:
                if artist_cell.value == artist:
                    print('Update DM 2')
                    new_domain_list_sheet.cell(row=artist_cell.row, column=2, value='{}'.format(artist_dm))

    return wb

def add_sort_conditions(wb, cell_row_end):
    ws_name = wb.sheetnames[0]
    ws1 = wb.get_sheet_by_name(ws_name)

    ws1.auto_filter.ref = "A1:F{}".format(cell_row_end)
    ws1.auto_filter.add_sort_condition("E2:E{}".format(cell_row_end))

    return wb

def main():
    loaded_wb = load_workbook(filename = 'DomainsJuly2017.xlsx')
    ws_name = loaded_wb.sheetnames[0]
    ws1 = loaded_wb.get_sheet_by_name(ws_name)
    cell_row_index = 2 # starts at Excel cell row '2'

    write_to_wb = create_write_to_workbook()

    for cell in ws1['C']:
        if cell.value == "Domain":
            continue
        # remove_all_but_domain(cell.value)
        w = whois.whois(cell.value)
        # print(w)

        domain = get_domain(w)
        org = check_if_atl_owned(w)
        expiration_date = get_exp_date(w)
        notes = do_notes_and_registrar(w)

        info = [] # structure to hold whois info
        info.extend((domain, org, expiration_date, notes))

        # print('Domain: {}; ATL Own? {}; Expiration: {}; Notes/Registrar: {}'.format(domain, org, expiration_date, notes))
        write_to_wb = add_whois_info(write_to_wb, cell_row_index, info)
        cell_row_index += 1

    write_to_wb = add_non_whois_info(ws1, write_to_wb)
    write_to_wb = verify_artists_and_digital_marketers(write_to_wb)
    write_to_wb = add_sort_conditions(write_to_wb, cell_row_index)
    write_to_wb.save(filename = 'new_domain_book.xlsx')


if __name__ == '__main__':
    main()
