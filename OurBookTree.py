import os.path
import copy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from geopy.geocoders import Nominatim
from geopy.distance import great_circle

SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("our-book-tree-5e94c77c0c5f.json", SCOPES)
connection = gspread.authorize(credentials)
gmail_user = 'bookabookasap@gmail.com'
gmail_password = 'ankithsucks'

################################################################################################################################
#Reg

def reg():
    global worksheetReg
    worksheetReg = connection.open('Registration').worksheet('Reg')
    global valuesReg
    valuesReg = worksheetReg.get('A2:F')
    for i in range(len(valuesReg)):
        if 'Bangalore' not in valuesReg[i][-2] and 'Bengaluru' not in valuesReg[i][-2]:
            valuesReg[i][-2] += ' Bangalore'
    global RegEmailList
    RegEmailList = [i[1] for i in valuesReg]


if __name__ == '__main__':
    reg()

def regDeniedEmail(ToEmail):
    server3 = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server3.ehlo()
    server3.login(gmail_user, gmail_password)
    sent_from = gmail_user
    to = [ToEmail]
    subject = "Our Book Tree: Access Denied."
    body = "Hello,\n\nUnfortunately your entry was not recorded as you have not registered with our website.\nRegister here: https://forms.gle/DDRK6nwn7diDkUzD6 and submit your response again.\n\n\nThank You,\nOur Book Tree"
    email_text = """\
    From: %s\nTo: %s\nSubject: %s\n%s
    
    """ % (sent_from, ", ".join(to), subject, body)
    server3.sendmail(sent_from, to, email_text)
    server3.close()
################################################################################################################################
#Distance
def mdistance(TakeLoc, GiveLoc):
    geolocator = Nominatim(user_agent="our-book-tree")
    try:
        location2 = geolocator.geocode(TakeLoc[0]) 
        (x1,y1) = (location2.latitude, location2.longitude)
    except:
        location2 = geolocator.geocode(str(TakeLoc[1])) 
        (x1,y1) = (location2.latitude, location2.longitude)
    distList = []
    for i in GiveLoc:
        try:
            location1 = geolocator.geocode(i[-2]) 
            (x2,y2) = (location1.latitude, location1.longitude)
        except:
            location1 = geolocator.geocode(str(i[-1])) 
            (x2,y2) = (location1.latitude, location1.longitude)
        distList += [i + [great_circle((x1, y1), (x2, y2)).kilometers]]
    mindist = distList[0][-1]
    minL = distList[0]
    for i in distList:
        if i[-1] < mindist:
            mindist = i[-1]
            minL = i

    return minL

################################################################################################################################
#Email Order Give
server1 = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server1.ehlo()
server1.login(gmail_user, gmail_password)
server2 = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server2.ehlo()
server2.login(gmail_user, gmail_password)
def EmailSend(keyTuple, valueList, FlagGT):
    sent_from = gmail_user
    to = [keyTuple[0]]
    
    body = 'Hello %s,\n\nWe have found a match for your books.\n\n'%(keyTuple[1])
    if FlagGT == 'Give':
        subject = 'Give: We have found a match for your books!'
        for i in valueList: 
            body+= "%s has requested for the following books:\n%s\n\n    Contact Information: \n    Email ID: %s Phone Number: %s \n\nWe have found the best possible match for you using our optimization algorithms.\nIf you need to deliver your books, you can use: https://www.dunzo.com/bangalore/send-packages \n\n"%(i[2], "\n\n".join(i[0]), i[1], i[3])
    elif FlagGT == 'Take':
        subject = 'Take: We have found a match for your books!'
        for i in valueList:
            body+="%s has offered the following books:\n%s\n\n    Contact Information: \n    Email ID: %s Phone Number: %s \n\nWe have found the best possible match for you using our optimization algorithms.\nIf you need to deliver your books, you can use: https://www.dunzo.com/bangalore/send-packages \n\n"%(i[2], "\n\n".join(i[0]), i[1], i[3])
    else:
        print("Error")
    body+="You can help end hunger in classrooms. Visit: https://www.ourbooktree.org/donate-now\n\n\nThank You,\nOur Book Tree"
    email_text = """\
    From: %s\nTo: %s\nSubject: %s\n%s

    """ % (sent_from, ", ".join(to), subject, body)

    if FlagGT == 'Give':
        server1.sendmail(sent_from, to, email_text)
    elif FlagGT == 'Take':
        server2.sendmail(sent_from, to, email_text)
### Deleting rows
def main3(y, ws):
    ws.delete_rows(y)

#################################################################################################################################
#ALL Books Table
def main3sub():

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    global worksheet3
    global values3
    worksheet3 = connection.open("BooksToGive").worksheet('AllEntries')
    values3 = worksheet3.get('A2:P')
    for i in range(len(values3)):
        values3[i][13] = int(values3[i][13])
    Lindex = []
    
    for m in range(len(values3)):
        if values3[m][2] not in RegEmailList:
            Lindex+=[m+1]


    LRegEmail3=[]
    for i in Lindex:
        i = i - Lindex.index(i)
        main3(i+1, worksheet3)
        if values3[i-1][2] not in LRegEmail3:
            LRegEmail3 += [values3[i-1][2]]
        values3.remove(values3[i-1])
        
    for i in LRegEmail3:
        regDeniedEmail(i)

if __name__ == '__main__':
    main3sub()




#################################################################################################################################
# Orders Table COMPLETE

def main1():
    global worksheet1
    worksheet1 = connection.open("BooksToTake").worksheet('FormResponses')
    global values1
    values1 = worksheet1.get('A2:O')

    Lindex = []
    
    for m in range(len(values1)):
        if values1[m][12] not in RegEmailList:
            Lindex+=[m+1]


    LRegEmail1=[]
    for i in Lindex:
        i = i - Lindex.index(i)
        main3(i+1, worksheet1)
        if values1[i-1][12] not in LRegEmail1:
            LRegEmail1+= [values1[i-1][12]]
        values1.remove(values1[i-1])
        
    for i in LRegEmail1:
        regDeniedEmail(i)
        
    for i in range(len(values1)):
        if len(values1[i]) == 13:
            s = ''.join(values1[i][4:10])
            s.strip(", ,")
            values1[i] += [s]
            if values1[i][-1] != '0' and values1[i][-1] != '1':
                values1[i] += ['0']



    if not values1:
        print('No data found.')
    else:
        global LOrder
        LOrder = []
        for row in values1:
            if ';' in row[-2]:
                w = row[-2].split(';,')
                for i in range(len(w)):
                    w[i] = [w[i].strip('; ')]
                    w[i] = w[i][0].split(':')
                    w[i][1] = int(w[i][1])

            else:
                w = row[-2].split(',')
                for i in range(len(w)):
                    w[i] = [w[i].strip()] + [0]

            LOrder += [[row[0], w, row[-1],row[12],row[10],row[11],row[1]]]


if __name__ == '__main__':
    main1()


#################################################################################################################################
# Consolidated Books Table WORKING


def main2():
    worksheet2 = connection.open("BooksToGive").worksheet('BookQuan')
    global values2
    values2 = worksheet2.get('A3:C')
    for i in range(len(values2)):
        values2[i][0] = values2[i][0].strip()



    global LProduct

    if not values2:
        print('No data found.')
    else:
        for i in range(len(values2)):
            values2[i][1] = int(values2[i][1])
        LProduct = copy.deepcopy(values2)
    global BookL
    BookL = []
    for i in LOrder:
        if int(i[2]) != len(i[1]):
            for j in i[1]:
                for k in range(len(LProduct)):
                    if j[0] == LProduct[k][0] and j[1] == 0 and i[-1] == LProduct[k][-1]:
                        if LProduct[k][1] > 0:
                            LProduct[k][1] -= 1
                            j[1] = 1
                            for m in valuesReg:
                                if i[3] == m[1] and i[-1] == LProduct[k][-1]:  
                                    BookL+=[[i[3],m[2],m[3],m[4],m[5],i[-1],[j[0]]]]

    for i in LOrder:
        ctr = 0
        for j in i[1]:
            if j[1] == 1:
                ctr+=1
        i[2] = str(ctr)

    for i in LOrder:
        str2 = ''
        str1 = ''
        for j in i[1]:
            str1 = ':'.join([str(x) for x in j])
            str2 += str1 + ";, "
        str2 = str2.strip(', ')
        i[1] = str2

    LOrderf = []
    for i in values1:
            for k in LOrder:
                if i[0] == k[0]:
                    LOrderf.append(i[:-2] + k[1:3])

    worksheet1.update('A2:O', LOrderf)


if __name__ == '__main__':
    main2()


#################################################################################################################################
#Categorized


def main4():
    worksheet4 = connection.open("BooksToGive").worksheet('Categorized')
    global values4
    values4 = worksheet4.get('A2:H')
    for i in range(len(values4)):
        values4[i][1] = values4[i][1].strip()


if __name__ == '__main__':
    main4()

##################################################################################

#ALLBOOKSUPDATE


#LProduct = deepcopy of values2
#values1 = FormResponses
#values2 = (Give)BookQuan
#values3 = (Give)AllEntries
#values4 = (Give)Categorised


def allbooksupdate():
    for i in BookL:
        locGive = []
        for j in values3:
            if [''.join(j[7:13])] == i[6] and i[5] == j[4]:
                for k in valuesReg:
                    if k[1] == i[0]:
                        locTake = k[4:]
                    if k[1] == j[2]:
                        locGive+= [j+k[4:]]
                
        
        Lopt = mdistance(locTake,locGive)
        for p in valuesReg:
            if p[1] == Lopt[2]:
                i += [Lopt[2], p[2], p[3]]
        i[6][0] += '\n    Condition of Book: ' + Lopt[14] + '\tYear of Publishing: ' + Lopt[15]
        for q in values3:
            if q == Lopt[:16]:
                q[13] = int(q[13]) - 1
                worksheet3.update('A2:P', values3)
                        

    for i in BookL:
        for j in BookL:
            if i != j and i[:6] == j[:6] and i[7:] == j[7:]:
                i[6] += j[6]
                BookL.remove(j)
    
    for i in BookL:
        for j in range(len(i[6])):
            i[6][j] = str(j+1) + '. ' + i[6][j]

    DGive = {}
    DTake = {}
    
    for i in BookL:
        if tuple(i[7:]) in DGive:
            DGive[tuple(i[7:])] += [[i[6]] + i[:6]]
        else:
            DGive[tuple(i[7:])] = [[i[6]] + i[:6]]

    for i in BookL:
        if tuple(i[:6]) in DTake:
            DTake[tuple(i[:6])] += [[i[6]] + i[7:]]
        else:
            DTake[tuple(i[:6])] = [[i[6]] + i[7:]]
    
    for i in DGive:
        EmailSend(i, DGive[i], 'Give')

    for i in DTake:
        EmailSend(i, DTake[i], 'Take')
    
    Lm = []
    for m in range(len(values3)):
        if values3[m][13] == 0:
            Lm+=[m+1]
    for i in Lm:
        i = i - Lm.index(i)
        main3(i+1, worksheet3)

if __name__ == '__main__':
    allbooksupdate()

server1.close()
server2.close()
#########################################################

exec(open('OurBookTree910.py').read())







