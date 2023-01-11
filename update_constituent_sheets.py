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
sheets = ['MYCC Youth Lobby Week 2023 (Responses)']


def get_constituents(sheet, legislator):
    constituents = []
    sh = sa.open(sheet)
    wks = sh.worksheet("Form Responses")
    rsvp = wks.get('A2:K{}'.format(wks.row_count))
    for i in range(len(rsvp)):
        for j in range(8, len(rsvp[i])):
            rsvp[i][j] = rsvp[i][j].title()
            legislator_comma = legislator + ","
            if legislator_comma in rsvp[i][j].split():
                constituents.append(rsvp[i])
            if legislator in rsvp[i][j].split():
                constituents.append(rsvp[i])
    return constituents

def search_all_rsvps(sheets, legislator):
    constituents = []
    for i in range(len(sheets)):
        for item in get_constituents(sheets[i], legislator):
            constituents.append(constituents.append(item))
    # print(constituents)
    return constituents

# Gets constituents already in the document
def get_current_constituents(legislator):
    sh = sa.open("{} Outreach Team".format(legislator))
    wks = sh.worksheet('Sheet1')
    current_constituents = wks.get('A2:B{}'.format(wks.row_count))
    # print(current_constituents)
    return current_constituents

get_current_constituents("Coppinger")

# UPDATES SHEET WITH RSVPS
def update_outreach(legislator):
    print(legislator)
    constituents0 = search_all_rsvps(sheets, legislator)
    # print(constituents0)
    sh = sa.open("{} Outreach Team".format(legislator))
    wks = sh.worksheet('Sheet1')
    constituents1 = []
    for i in range(len(constituents0)):
        if constituents0[i] is not None:
            constituents1.append(constituents0[i])

    # Constituents2 removes duplicates
    constituents2 = []
    for i in range(len(constituents1)):
        if len(constituents2) == 0:
            constituents2.append(constituents1[i])
        for j in range(len(constituents2)):
            if constituents1[i][0].lower() == constituents2[j][0].lower() and constituents1[i][1].lower() == constituents2[j][1].lower():
                break
            elif j == len(constituents2) - 1:
                constituents2.append(constituents1[i])
    print(constituents2)

    # Checks if they are already in the outreach sheet
    constituents3 = []
    current_constituents = get_current_constituents(legislator)
    for i in range(len(constituents2)):
        for j in range(len(current_constituents)):
            if constituents2[i][0] == current_constituents[j][0] and constituents2[i][1] == current_constituents[j][1]:
                break
            elif j == len(current_constituents) - 1:
                constituents3.append(constituents2[i])
    print(constituents3)

    # Checks for Youth
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
    # wks.batch_clear(['A2:H25'])
    wks.update('A{}'.format(len(current_constituents)+2), constituents3)
    return constituents2

# update_outreach("Coppinger")
# 'Arciero', 'Ashe', 'Ayers', 'Balser', 'Barber','Barrett', 'Barrows', 'Belsito', 'Berthiaume', 'Biele',
# 'Blais', 'Boldyga', 'Cabral', 'Cahill', 'Campbell', 'Capano', 'Carey', 'Cassidy', 'Chan', 'Ciccolo', 'Connolly', 'Consalvo',
# 'Coppinger', 'Cronin', 'Cusack','Cutler', 'D\'Emilia', 'Day', 'Decker', 'DeLeo', 'Devers', 'Diggs', 'Doherty', 'Domb',
# 'Donahue', 'Donato', 'Dooley', 'Driscoll', 'DuBois', 'Duffy', 'Durant', 'Dykema', 'Ehrlich', 'Elugardo', 'Farley-Bouvier',
# 'Ferguson', 'Fernandes','Ferrante', 'Finn', 'Fiola', 'Fluker-Oakley', 'Frost', 'Galvin','Garballey', 'Garlick', 'Garry',
#                'Gentile', 'Giannino', 'Gifford', 'Golden', 'Gonzalez', 'Gordon', 'Gouveia', 'Gregoire', 'Haddad', 'Haggerty', 'Harrington', 'Hawkins',
#                'Hendricks', 'Higgins','Hogan', 'Holmes', 'Honan', 'Howard', 'Howitt', 'Hunt', 'Jones', 'Kane', 'Kearney', 'Keefe', 'Kelcourse',
#                'Kerans', 'Khan', 'Kilcoyne','Kushmerek', 'LaNatra', 'Lawn', 'LeBoeuf', 'Lewis', 'Linsky', 'Lipper-Garabedian', 'Livingstone',
#                'Lombardo', 'Madaro', 'Mahoney', 'Malia', 'Mariano','Mark', 'Markey', 'McGonagle', 'McKenna','McMurtry', 'Meschino', 'Michlewitz',
#                'Minicucci','Miranda', 'Mirra', 'Mom', 'Moran', 'Moran', 'Muradian', 'Muratore', 'Murphy', 'Murray', 'Nguyen','O\'Day',
#                'Oliveira', 'Orrall', 'Owens', 'Parisella', 'Peake', 'Pease', 'Peisch', 'Philips', 'Pignatelli','Puppolo', 'Ramos', 'Robertson',
#                'Robinson','Rogers', 'Roy', 'Ryan', 'Sabadosa', 'Santiago', 'Scanlon', 'Schmid', 'Sena', 'Silvia', 'Smola', 'Soter',
#                'Stanley', 'Straus', 'Sullivan','Tucker', 'Turco', 'Tyler', 'Ultrino', 'Uyterhoeven', 'Vargas', 'Vieira',
list_of_leg = [
               'Vitolo', 'Wagner', 'Walsh',
               'Whelan', 'Whipps', 'Williams', 'Wong','Xiarhos', 'Zlotnik']
list_of_senators = ['Barrett', 'Boncore', 'Brady', 'Brownsberger', 'Chandler', 'Chang-Diaz', 'Collins', 'Comerford', 'Creem', 'Crighton', 'Cronin',
                    'Cyr', 'DiDomenico', 'DiZoglio', 'Eldridge', 'Fattman', 'Feeney', 'Finegold', 'Friedman', 'Gobi', 'Gomez', 'Hinds', 'Jehlen',
                    'Keenan', 'Kennedy', 'Lesser', 'Lewis', 'Lovely', 'Montigny', 'Moore', 'Moran', "O'Connor", 'Oliveira', 'Pacheco', 'Rausch',
                    'Rodrigues', 'Rush', 'Spilka', 'Tarr', 'Timilty', 'Velis']

# 'Chan', 'Cutler','D\'Emilia','DeCoste','Domb',

for item in list_of_leg:
    update_outreach(item)
    # clear_boxes(item)
    # NEED SLEEP TIME ON FREE VERSION
    time.sleep(30)



# update_outreach("Coppinger")