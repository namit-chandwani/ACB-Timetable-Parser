import pandas as pd


df = pd.read_excel('Timetable_inp.xlsx', sheet_name='ftp-tincy Ma\'am TT',

                   header=2)

fac_psrn = pd.read_excel('Timetable_inp2.xlsx', sheet_name='faculty psrn',

                         header=0)

phd_psrn = pd.read_excel('PHD SCHOLARS SYS ID.xlsx', sheet_name='phd scholars',

                         header=0)

room_id = pd.read_excel('Timetable_inp2.xlsx', sheet_name='rooms',

                        header=0)

components = pd.read_excel('Timetable_inp2.xlsx', sheet_name='components',

                           header=1)

df = df.rename(columns={"COM CODE": "Course ID"})


def getRoom(x):

    s = ""

    if(pd.isnull(x)):

        return x

    x = str(x)

    if(x[0] == "A" or x[0] == "B" or (x[0] == "C" and x[1] != "C") or (x[0] == "D" and x[1] == "L") or x[0] == "W" or (x[0] == "L" and x[1].isdigit())):

        s = x[0]+"-"+x[1:]

    elif((x[0] == "D" and x[1].isdigit()) or (x[0] == "L" and x[1] == "T")):

        s = x

    elif((x[0] == "L" and x[1] == "C") or (x[0] == "C" and x[1] == "C") or (x[0] == "E" and x[1] == "L")):

        s = x[:2]+"-"+x[2:]

    return s

#     for i in range(len(x)):

#         c = x[i]

#         if(c.isalpha() == False):

#             return x[:i]+"-"+x[i:]


df["ROOM"] = df["ROOM"].apply(getRoom)

df["Course ID"] = df["Course ID"].apply(lambda x: "0"+str(x))

new = df["COURSENO"].str.split(" ", n=1, expand=True)

df["Subject"] = new[0]

df["Catalog"] = new[1]

df.drop(columns=["COURSENO"], inplace=True)

df["compocode"] = df["STAT"].apply(

    lambda x: "01" if x == 'L' else ("02" if x == 'T' else "03"))

df["SEC"] = df["SEC"].apply(lambda x: str(x))

df["Class Code"] = df["Course ID"] + df["compocode"] + df["SEC"]

df["Class Nbr"] = "####"

df["Section"] = df["STAT"] + df["SEC"].map(str)

df.drop(columns=["STAT", "SEC"], inplace=True)

df["Course_admin"] = "G####"

df["Role"] = "##"

df["Instructor ID"] = "###"

new = df["COMPRE DATE"].str.split(" ", n=1, expand=True)

df["Exam Date"] = new[0]

new2 = df["Exam Date"].str.split("/", n=1, expand=True)

df["Exam Tm Cd"] = "G" + new[1].str.slice(start=1, stop=3) + new2[0]

df.drop(columns=["COMPRE DATE", "compocode"], inplace=True)

df["Career"] = df["Catalog"].str.slice(start=0, stop=1).apply(

    lambda x: "First Deg" if x == "F" else "Higher Deg")
df = (df.set_index(df.columns.drop("INSTRUCTOR IN CHARGE/Instructor", 1).tolist())["INSTRUCTOR IN CHARGE/Instructor"].str.split(',', expand=True)

      .stack()

      .reset_index()

      .rename(columns={0: "INSTRUCTOR IN CHARGE/Instructor"})

      .loc[:, df.columns]

      )

df = (df.set_index(df.columns.drop("INSTRUCTOR IN CHARGE/Instructor", 1).tolist())["INSTRUCTOR IN CHARGE/Instructor"].str.split('/', expand=True)

      .stack()

      .reset_index()

      .rename(columns={0: "INSTRUCTOR IN CHARGE/Instructor"})

      .loc[:, df.columns]

      )

df = df.rename(columns={"INSTRUCTOR IN CHARGE/Instructor": "Display Name"})

df["Display Name"] = df["Display Name"].apply(

    lambda x: str(x)[1:] if str(x)[0] == " " else str(x))


def getTime(x):

    if(pd.isnull(x)):

        return x

    if(str(x)[-1] == "?"):

        x = str(x)

        x = x.replace(" ", "")

        for j in range(len(x)):

            if(x[j].isalpha() == False):

                x = x[:j]+" "+x[j]+"-"+str(int(x[j])+1)+"?"

                break

        return x

    x = str(x).split(' ')

    s1 = ""

    s2 = ""

    flag = False

    n = len(x)

    for i in range(n):

        c = x[i]

        if(c == ""):

            continue

        if(c.isalpha()):

            if(flag):

                s2 = s1

                s1 = ""

                flag = False

            s1 += c

        else:

            if(i+1 < n and x[i+1].isalpha() == False):

                s1 += (" "+str(int(c)+7)+"-"+str(int(c)+9))

            elif(x[i-1].isalpha()):

                s1 += (" "+str(int(c)+7)+"-"+str(int(c)+8))

            flag = True

    return s1+","+s2


df["DAYS/ H"] = df["DAYS/ H"].apply(getTime)

df["DAYS/ H"] = df["DAYS/ H"].apply(lambda x: str(x)

                                    [:-1] if str(x)[-1] == "," else str(x))


# for i in range(len(pos_list)):

#     pos_list[i] = pos_list[i]+4


df = (df.set_index(df.columns.drop("DAYS/ H", 1).tolist())["DAYS/ H"].str.split(',', expand=True)

      .stack()

      .reset_index()

      .rename(columns={0: "DAYS/ H"})

      .loc[:, df.columns]

      )

new = df["DAYS/ H"].str.split(" ", n=1, expand=True)

df["Class Pattern"] = new[0]

df["DAYS/ H"] = new[1]

df["DAYS/ H"] = df["DAYS/ H"].apply(lambda x: (str(x).split(' ')[0].split('-')[0]+"-"+str(

    x).split(' ')[1].split('-')[1]) if (len(str(x).split(' ')) == 2) else str(x))

new2 = df["DAYS/ H"].str.split("-", n=1, expand=True)

df["Mtg Start"] = new2[0]+":00:00"

df["End time"] = new2[1].apply(lambda x: str(

    x)[:-1]+":30:00" if str(x)[-1] == "?" else str(x)+":00:00")


df.drop(

    columns=["DAYS/ H", "CREDIT                          L P U"], inplace=True)

df["Mtg Start"] = df["Mtg Start"].apply(

    lambda x: None if str(x) == "None:00:00" else str(x))

df["End time"] = df["End time"].apply(

    lambda x: None if str(x) == "None:00:00" else str(x))

df['Mtg Start'] = pd.to_datetime(df['Mtg Start'], format='%H:%M:%S').dt.time

df['End time'] = pd.to_datetime(df['End time'], format='%H:%M:%S').dt.time

df["Mtg Start"] = df["Mtg Start"].apply(lambda x: str(x) if pd.isnull(x) else (str(int(
    str(x)[0:2])-12)+":00:00 PM" if int(str(x)[0:2]) > 12 else str(int(str(x)[0:2]))+":00:00 AM"))

# df["Mtg Start"] = df["Mtg Start"].apply(

#     lambda x: "0"+str(x) if len(str(x)) == 10 else x)

df["End time"] = df["End time"].apply(lambda x: str(x) if pd.isnull(x) else (str(int(str(

    x)[0:2])-12)+":00:00 PM" if int(str(x)[0:2]) > 12 else str(int(str(x)[0:2]))+":00:00 AM"))

# df["End time"] = df["End time"].apply(

#     lambda x: "0"+str(x) if len(str(x)) == 10 else x)

# print(df['Mtg Start'])
# print(str(df['Mtg Start'][923])+"AM")

df["Mtg Start"] = df["Mtg Start"].apply(lambda x: x if str(x) == "NaT" else (

    str(x)[:-2]+"PM" if (int(str(x)[0]) >= 12 and int(str(x)[0]) <= 7) else str(x)))

df["End time"] = df["End time"].apply(lambda x: x if str(x) == "NaT" else (

    str(x)[:-2]+"PM" if (int(str(x)[0]) >= 12 and int(str(x)[0]) <= 7) else str(x)))

# print(df['Mtg Start'])


df["Mtg Start"] = df["Mtg Start"].apply(lambda x: x if str(x) == "NaT" else (

    str(x)[:-2]+"AM" if (int(str(x)[0]) >= 8 and int(str(x)[0]) <= 11) else str(x)))

df["End time"] = df["End time"].apply(lambda x: x if str(x) == "NaT" else (

    str(x)[:-2]+"AM" if (int(str(x)[0]) >= 8 and int(str(x)[0]) <= 11) else str(x)))


# print(df['Mtg Start'])


# df["End time"] = df["End time"].apply(lambda x: str(x)[:-2]+"PM" if (x!="NaT" and (int(str(x)[0]) < 8)) else str(x))


# for i in range(len(df["Mtg Start"])):

#     if(str(df["Mtg Start"][i])=='NaT'):

#         df["Mtg Start"][i]="Hello"

#         print(i)


df["Mtg Start"] = df["Mtg Start"].apply(

    lambda x: "12:00:00 PM" if str(x) == "12:00:00 AM" else str(x))

df["End time"] = df["End time"].apply(

    lambda x: "12:00:00 PM" if str(x) == "12:00:00 AM" else str(x))

# print(df['Mtg Start'])


# for x in pos_list:

#     df["Mtg Start"][x] = df["Mtg Start"][x][:-2]+"PM"

#     df["End time"][x] = df["End time"][x][0:3]+"3"+df["End time"][x][4:8]+" PM"


df["Role"] = df["Display Name"].apply(

    lambda x: "PI" if str(x).isupper() == True else "SI")

df = df[['Course ID', 'Subject', 'Catalog', 'COURSETITLE', 'Class Nbr', 'Section', 'ROOM', 'Class Pattern', 'Mtg Start', 'End time', 'Instructor ID', 'Display Name',

         'Role', 'Exam Tm Cd',

         'Exam Date', 'Course_admin',  'Career', 'Class Code']]


df2 = df[["Course ID", "Section", "Class Pattern",

          "Mtg Start", "End time", "Role", "ROOM"]]

df2 = df2.rename(columns={"Section": "Class Section", "Class Pattern": "Meeting Pattern",

                          "Mtg Start": "Meeting Start Time", "End time": "Meeting End Time", "Role": "Instructor Role"})

df2["Access"] = df2["Instructor Role"].apply(

    lambda x: "N" if str(x) == "SI" else "Y")

df2["Institution"], df2["Semester"], df2["Class Type"], df2["Campus"], df2["Session"], df2["Facility ID"], df2["Sequence Number"], df2[

    "Instructors ID"], df2["IC PSRN"] = "BITS", "####", "#", "GOAON", 1, "GLEC####", "#", "###", "G####"

df2["Component"] = df2["Class Section"].apply(lambda x: "THE" if x[0:2] == "TH" else ("LEC" if x[0] == "L" else (

    "PRO" if x[0] == "R" else ("LAB" if x[0] == "P" else ("TUT" if x[0] == "T" else "IND")))))


def getPsrn(x):

    s = ""

    x = str(x).upper()

    if(fac_psrn.index[fac_psrn["name"] == x].size != 0):

        s = fac_psrn.iloc[fac_psrn.index[fac_psrn["name"] == x][0]]["PSRN"]

    elif(phd_psrn.index[phd_psrn["name"] == x].size != 0):

        s = phd_psrn.iloc[phd_psrn.index[phd_psrn["name"] == x]

                          [0]]["SYSTEM ID"]

#         year=s[:4]

#         idno=s[-5:-1]

#         s="###"+year+idno

    return s


df2["Instructors ID"] = df["Display Name"].apply(getPsrn)


def getRoomId(x):

    s = ""

    x = str(x)

    if(room_id.index[room_id["Room"] == x].size != 0):

        s = room_id.iloc[room_id.index[room_id["Room"] == x][0]]["Facility ID"]

    return s


df2["Facility ID"] = df2["ROOM"].apply(getRoomId)


df2.drop(columns=["ROOM"], inplace=True)

df2 = df2[['Institution', 'Semester', 'Course ID', 'Class Section', 'Component', 'Class Type', 'Campus', 'Session', 'Facility ID', 'Meeting Pattern', 'Meeting Start Time', 'Meeting End Time', 'Sequence Number',

           'Instructors ID', 'Access',

           'Instructor Role', 'IC PSRN']]


df2["ROOM"], df2["COURSETITLE"], df2["Display Name"], df2["Subject"], df2["Catalog"] = df[

    "ROOM"], df["COURSETITLE"], df["Display Name"], df["Subject"], df["Catalog"]


IcMap = {}

for i in range(len(df2["Access"])):

    if(df2["Access"][i] == "Y"):

        #         print(df2["Instructors ID"][i]+" => "+df2["Access"][i])

        if(df2["Instructors ID"][i] != ""):

            IcMap[df2["Course ID"][i]] = df2["Instructors ID"][i]


df2["IC PSRN"] = df2["Course ID"].apply(lambda x: IcMap.get(x, ""))


df["Instructor ID"] = df2["Instructors ID"]

df["Course_admin"] = df2["IC PSRN"]


df.to_excel("output_Format1.xlsx", index=False)


df2["Sequence Number"][0] = 1

for i in range(1, len(df2["Class Section"])):

    if(df2["Class Section"][i] == df2["Class Section"][i-1]):

        df2["Sequence Number"][i] = df2["Sequence Number"][i-1]+1

    else:

        df2["Sequence Number"][i] = 1


components["Course ID"] = components["Course ID"].apply(lambda x: "0"+str(x))


def getPrimaryComponent(x):

    s = ""

    x = str(x)

    if(components.index[components["Course ID"] == x].size != 0):

        s = components.iloc[components.index[components["Course ID"]

                                             == x][0]]["Primary Component"]

    return s


df2["Class Type"] = df2["Course ID"].apply(getPrimaryComponent)

for i in range(len(df2["Class Type"])):

    if(df2["Class Type"][i] == df2["Component"][i]):

        df2["Class Type"][i] = "E"

    else:

        df2["Class Type"][i] = "N"


df2["Component"] = df2["Component"].apply(

    lambda x: "LEC" if str(x) == "TUT" else str(x))


df2.to_excel("output_Format2.xlsx", index=False)
