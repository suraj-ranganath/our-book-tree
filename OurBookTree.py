import os.path
import copy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib

SCOPES = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("our-book-tree-5e94c77c0c5f.json", SCOPES)
connection = gspread.authorize(credentials)

################################################################################################################################
#Email Order Give
def EmailSend(SName, RName, BookList, Email, Phn, ToEmail, FlagGT):
    gmail_user = 'bookabookasap@gmail.com'
    gmail_password = 'ankithsucks'

    sent_from = gmail_user
    to = [ToEmail]
    
    if FlagGT == 'Give':
        subject = 'Give: We have found a match for your books!'
        body = 'Hello %s,\n\nWe have found a match for your books.\n\n%s has requested for the following books:\n%s\n\nContact Information: \nEmail ID: %s Phone Number: %s \n\nYou can help to end hunger in classrooms. Visit: https://www.ourbooktree.org/donate-now\n\n\nThank You,\nOur Book Tree '%(SName, RName, BookList, Email, Phn)
    elif FlagGT == 'Take':
        subject = 'Take: We have found a match for your books!'
        body = 'Hello %s,\n\nWe have found a match for your books.\n\n%s has offered the following books:\n%s\n\nContact Information: \nEmail ID: %s Phone Number: %s \n\nYou can help to end hunger in classrooms. Visit: https://www.ourbooktree.org/donate-now\n\n\nThank You,\nOur Book Tree '%(SName, RName, BookList, Email, Phn)
    else:
        print("Error")
    #print(sent_from)
    email_text = """\
    From: %s\nTo: %s\nSubject: %s\n%s

    """ % (sent_from, ", ".join(to), subject, body)

    if FlagGT == 'Give':
        server1 = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server1.ehlo()
        server1.login(gmail_user, gmail_password)
        server1.sendmail(sent_from, to, email_text)
        server1.close()
    elif FlagGT == 'Take':
        server2 = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server2.ehlo()
        server2.login(gmail_user, gmail_password)
        server2.sendmail(sent_from, to, email_text)
        server2.close()        
    #print("SUCCESS! ankith sucks1!")

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
    for i in values3:
        i[13] = int(i[13])



if __name__ == '__main__':
    main3sub()

def main3(y):
    worksheet3.delete_rows(y)


#################################################################################################################################
# Orders Table COMPLETE

def main1():
    global worksheet1
    worksheet1 = connection.open("BooksToTake").worksheet('FormResponses')
    global values1
    values1 = worksheet1.get('A2:O')
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

            LOrder += [[row[0], w, row[-1],row[12],row[10],row[11]]]


if __name__ == '__main__':
    main1()


#################################################################################################################################
# Consolidated Books Table WORKING


def main2():
    worksheet2 = connection.open("BooksToGive").worksheet('BookQuan')
    global values2
    values2 = worksheet2.get('A3:B')



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
                    if j[0] == LProduct[k][0] and j[1] == 0:
                        if LProduct[k][1] > 0:
                            LProduct[k][1] -= 1
                            j[1] = 1
                            BookL+=[[i[3],i[4],i[5],[j[0]]]]

    global LTaken
    LTaken = copy.deepcopy(LOrder)
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


if __name__ == '__main__':
    main4()

##################################################################################

#ALLBOOKSUPDATE


#LProduct = (Give)BookQuan
#values1 = FormResponses
#values2 = deepcopy of LProduct
#values3 = (Give)AllEntries
#values4 = (Give)Categorised


def allbooksupdate():
    global LGiven
    LGiven = []
    for i in LProduct:
        for j in values2:
            if i[1] != j[1] and i[0]==j[0]:
                for m in range(len(values3)):
                    if values4[m][1] == i[0]:
                        values3[m][13] = int(values3[m][13]) - 1
                        LGiven += [values3[m]]
                        worksheet3.update('A2:P', values3)


    
    for i in LGiven:
        s = ''.join(i[7:13])
        i += [s]

    for i in BookL:
        for j in LGiven:
            if [j[-1]] ==  i[3]:
                i+= [j[2],j[1],j[3]]
                i[3][0] += '\n    Edition: ' + j[-3] + '\tYear of Publishing: ' + j[-2]
    for i in BookL:
        for j in BookL:
            if i != j and i[:3] == j[:3] and i[4:] == j[4:]:
                i[3] += j[3]
                BookL.remove(j)
    #print(BookL)
    for i in BookL:
        for j in range(len(i[3])):
            i[3][j] = str(j+1) + '. ' + i[3][j]
    for q in BookL:
        BookStr = '\n\n'.join(q[3])
        EmailSend(q[5], q[1], BookStr, q[0], q[2], q[4], 'Give')
        EmailSend(q[1], q[5], BookStr, q[4], q[6], q[0], 'Take')
    Lm = []
    for m in range(len(values3)):
        if values3[m][13] == 0:
            Lm+=[m+1]
    for i in Lm:
        i = i - Lm.index(i)
        main3(i+1)


if __name__ == '__main__':
    allbooksupdate()

#########################################################









