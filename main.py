import gspread
import time

sa = gspread.service_account()

# print('Rows: ', wks.row_count)
# print(wks.acell('A9').value)
# print(wks.cell(3, 4))
# print(wks.get('A7:E9'))
# print(wks.get_all_records())
# wks.update("A1", "Hello")

# ADD NEW RSVP SHEETS HERE AND SHARE SHEET WITH DATABASE EMAIL
sheets = ['MYCC Climate HW Due Youth Lobby Week - 1/24/22 (Responses)', 'Green Future Act Lobby Week MAR 2021',
          'Humanities Workshop Climate Summit & Lobby Day RSVP (Responses)',
          'Digital Lobby Week Planning Sheet (includes RSVPs)', 'Jan 29 Lobby Day - RSVPs, Meetings, Outreach']


def get_constituents(sheet, legislator):
    constituents = []
    sh = sa.open(sheet)
    wks = sh.worksheet("Form Responses")
    rsvp = wks.get('A2:K{}'.format(wks.row_count))
    for i in range(len(rsvp)):
        for j in range(len(rsvp[i])):
            rsvp[i][j] = rsvp[i][j].title()
            if legislator in rsvp[i][j]:
                constituents.append(rsvp[i])
    return constituents

def search_all_rsvps(sheets, legislator):
    constituents = []
    for i in range(len(sheets)):
        for item in get_constituents(sheets[i], legislator):
            constituents.append(constituents.append(item))
    return constituents

def update_outreach(legislator):
    constituents0 = search_all_rsvps(sheets, legislator)
    sh = sa.open("{} Outreach Team".format(legislator))
    wks = sh.worksheet('Sheet1')
    constituents1 = []
    for i in range(len(constituents0)):
        if constituents0[i] is not None:
            constituents1.append(constituents0[i])
    constituents2 = []
    for i in range(len(constituents1)):
        if len(constituents2) == 0:
            constituents2.append(constituents1[i])
        for j in range(len(constituents2)):
            if constituents1[i][0] == constituents2[j][0] and constituents1[i][1] == constituents2[j][1]:
                break
            elif j == len(constituents2) - 1:
                constituents2.append(constituents1[i])
    for i in range(len(constituents2)):
        if len(constituents2[i]) < 11:
            del constituents2[i][-2:]
            break
        elif 'High School' in constituents2[i][10]:
            constituents2[i][8] = True
        elif 'College' in constituents2[i][10]:
            constituents2[i][8] = True
        elif '18' in constituents2[i][10]:
            constituents2[i][8] = True
        elif 'Middle School' in constituents2[i][10]:
            constituents2[i][8] = True
        elif 'Youth' in constituents2[i][10]:
            constituents2[i][8] = True
        else:
            constituents2[i][8] = False
        del constituents2[i][-2:]
    wks.batch_clear(['A2:H25'])
    wks.update('A2', constituents2)
    # wks.format("A2:AB2", {
    #     "backgroundColor": {
    #         "red": 255,
    #         "green": 0,
    #         "blue": 0}})
    # wks.update('A2:H2', constituents2[0])
    # wks.update('A2:H{}'.format(len(constituents2)+1), constituents2)
    # print(constituents2)
    return constituents2

# FUNCTION TO CLEAR CHECKS FOR LOBBY DAYS
def clear_boxes(legislator):
    sh = sa.open("{} Outreach Team".format(legislator))
    wks = sh.worksheet('Sheet1')
    for i in range(14):
        for item in ['L', 'M', 'N', 'O', 'P']:
            if wks.get('{}{}'.format(item, i + 2)):
                wks.update('{}{}'.format(item, i + 2), False)
        time.sleep(5)

list_of_leg = ['Arciero', 'Ashe', 'Ayers', 'Balser', 'Barber', 'Barrett', 'Barrows', 'Belsito', 'Berthiaume', 'Biele',
               'Blais', 'Boldyga', 'Cabral', 'Cahill', 'Campbell', 'Capano', 'Carey', 'Cassidy', 'Chan', 'Ciccolo', 'Connolly', 'Consalvo',
               'Coppinger', 'Cronin', 'Cusack', 'Cutler', 'D\'Emilia', 'Day', 'Decker', 'DeLeo', 'Devers', 'Diggs', 'Doherty', 'Domb',
               'Donahue', 'Donato', 'Dooley', 'Driscoll', 'DuBois', 'Duffy', 'Durant', 'Dykema', 'Ehrlich', 'Elugardo', 'Farley-Bouvier',
               'Ferguson', 'Fernandes','Ferrante', 'Finn', 'Fiola', 'Fluker-Oakley', 'Frost', 'Galvin', 'Garballey', 'Garlick', 'Garry',
               'Gentile', 'Giannino', 'Gifford', 'Golden', 'Gonzalez', 'Gordon', 'Gouveia', 'Gregoire', 'Haddad', 'Haggerty', 'Harrington', 'Hawkins',
               'Hendricks', 'Higgins','Hogan', 'Holmes', 'Honan', 'Howard', 'Howitt', 'Hunt', 'Jones', 'Kane', 'Kearney', 'Keefe', 'Kelcourse',
               'Kerans', 'Khan', 'Kilcoyne','Kushmerek', 'LaNatra', 'Lawn', 'LeBoeuf', 'Lewis', 'Linsky', 'Lipper-Garabedian', 'Livingstone',
               'Lombardo', 'Madaro', 'Mahoney', 'Malia', 'Mariano', 'Mark', 'Markey', 'McGonagle', 'McKenna','McMurtry', 'Meschino', 'Michlewitz',
               'Minicucci','Miranda', 'Mirra', 'Mom', 'Moran', 'Moran', 'Muradian', 'Muratore', 'Murphy', 'Murray', 'Nguyen','O\'Day',
               'Oliveira', 'Orrall', 'Owens', 'Parisella', 'Peake', 'Pease', 'Peisch', 'Philips', 'Pignatelli','Puppolo', 'Ramos', 'Robertson',
               'Robinson','Rogers', 'Roy', 'Ryan', 'Sabadosa', 'Santiago', 'Scanlon', 'Schmid', 'Sena', 'Silvia', 'Smola', 'Soter',
               'Stanley', 'Straus', 'Sullivan','Tucker', 'Turco', 'Tyler', 'Ultrino', 'Uyterhoeven', 'Vargas', 'Vieira', 'Vitolo', 'Wagner', 'Walsh',
               'Whelan', 'Whipps', 'Williams', 'Wong','Xiarhos', 'Zlotnik']
list_of_senators = ['Barrett', 'Boncore', 'Brady', 'Brownsberger', 'Chandler', 'Chang-Diaz', 'Collins', 'Comerford', 'Creem', 'Crighton', 'Cronin',
                    'Cyr', 'DiDomenico', 'DiZoglio', 'Eldridge', 'Fattman', 'Feeney', 'Finegold', 'Friedman', 'Gobi', 'Gomez', 'Hinds', 'Jehlen',
                    'Keenan', 'Kennedy', 'Lesser', 'Lewis', 'Lovely', 'Montigny', 'Moore', 'Moran', "O'Connor", 'Oliveira', 'Pacheco', 'Rausch',
                    'Rodrigues', 'Rush', 'Spilka', 'Tarr', 'Timilty', 'Velis']


# 'Chan', 'Cutler','D\'Emilia','DeCoste','Domb',

for item in list_of_senators:
    # update_outreach(item)
    clear_boxes(item)
    # NEED SLEEP TIME ON FREE VERSION
    time.sleep(60)
