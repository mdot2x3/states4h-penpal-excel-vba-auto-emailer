VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Sub testnumber()
'
'     Dim outlookapp As outlook.Application
'
'     Dim I As Long
'
'
'     Set outlookapp = CreateObject("Outlook.Application")
'
'
'     For I = 1 To outlookapp.Session.Accounts.Count
'         MsgBox outlookapp.Session.Accounts.Item(I) & " : This is account number " & I
'     Next I
'
' End Sub

'Sub Send_email_fromexcel()
'Dim edress As String
'Dim subj As String
'Dim message As String
''Dim filename As String
'Dim outlookapp As Object
'Dim outlookmailitem As Object
''Dim myAttachments As Object
''Dim path As String
'Dim lastrow As Integer
'Dim attachment As String
'Dim x As Integer
'
'
'
'    x = 3
'
'Do While Sheet1.Cells(x, 11) <> ""
'
'    Set outlookapp = CreateObject("Outlook.Application")
'    Set outlookmailitem = outlookapp.CreateItem(0)
'    'Set myAttachments = outlookmailitem.Attachments
'    'path = "C:\Users\M\Doc~~~~~~~"
'    edress = Sheet1.Cells(x, 11)
'
'    subj = "States' 4-H Pen Pal Match Information"
'    'filename = Sheet1.Cells(x, 3)
'    'attachment = path + filename
'
'
'
'        outlookmailitem.To = edress
'        outlookmailitem.CC = ""
'        outlookmailitem.BCC = ""
'        outlookmailitem.Subject = subj
'        outlookmailitem.Body = "Please find your 2020 Summer Pen Pal Program match information below." & vbCrLf & "Best regards,"
'
'
'        'myAttachments.Add (attachment)
'        outlookmailitem.Display
'        Set outlookmailitem.SendUsingAccount = outlookapp.Session.Accounts.Item(3)
'
'        lastrow = lastrow + 1
'        edress = ""
'    x = x + 1
'
'    Loop
'
'    Set outlookapp = Nothing
'    Set outlookmailitem = Nothing
'
' End Sub

Sub congradletter()
Dim edress As String
Dim subj As String
Dim message As String
Dim filename As String
Dim outlookapp As Object
Dim outlookmailitem As Object
Dim myAttachments As Object
Dim path As String
Dim x As Integer
Dim erow As Integer
Dim Subject, name, matchname, age, gender, country, email1, email2, contact1, contact2, contact3, contact4, interests, notes As String

x = 3
erow = Sheet1.Cells(x, 1).CurrentRegion.Rows.Count

For x = 3 To erow

    Set outlookapp = CreateObject("Outlook.Application")
    Set outlookmailitem = outlookapp.CreateItem(0)
    Set myAttachments = outlookmailitem.Attachments
    
    path = "C:\Users\MS\Dropbox (S4IEP)\Company Team Folder\New Program Ideas\Summer 2021 Virtual Programs\2021 Summer Pen Pal\Monthly Topics\"
    filename = "July - Topic of Month.pdf"
    attachment = path + filename
    
    edress = Sheet1.Cells(x, 2) & ";" & Sheet1.Cells(x, 3)
    Subject = "States' 4-H Summer Pen Pal Program - Match Information"
    name = Sheet1.Cells(x, 4)
    matchname = Sheet1.Cells(x, 18)
    age = Sheet1.Cells(x, 19)
    gender = Sheet1.Cells(x, 20)
    country = Sheet1.Cells(x, 21)
    email1 = Sheet1.Cells(x, 16)
    email2 = Sheet1.Cells(x, 17)
    contact1 = Sheet1.Cells(x, 22)
    contact2 = Sheet1.Cells(x, 23)
    contact3 = Sheet1.Cells(x, 24)
    contact4 = Sheet1.Cells(x, 25)
    interests = Sheet1.Cells(x, 26)
    notes = Sheet1.Cells(x, 27)
    
    outlookmailitem.Body = "Dear " & name & "," & vbCrLf & vbCrLf & "Welcome to the States' 4-H Summer Pen Pal Program! We are very excited to match you with an international pen pal. Please see your matching information below." & vbCrLf & vbCrLf & _
    "Name:  " & matchname & vbCrLf & _
    "Age:  " & age & vbCrLf & _
    "Gender:  " & gender & vbCrLf & _
    "Country:  " & country & vbCrLf & _
    "Interests:  " & interests & vbCrLf & _
    "Notes:  " & notes & vbCrLf & _
    "Primary Email:  " & email1 & vbCrLf & _
    "Secondary Email:  " & email2 & vbCrLf & _
    "Additional Contact Information Below (if blank, they did not provide additional information):  " & vbCrLf & _
    contact1 & vbCrLf & _
    contact2 & vbCrLf & _
    contact3 & vbCrLf & _
    contact4 & vbCrLf & _
    "You are now prepared to start communication with your match! Feel free to begin writing when you feel ready after July 1st. Also, please be sure to add '4-H Pen Pal' somewhere in your Email Subject Line. We have instructed our international pen pals to do the same so it will be easy to identify the emails. If you don't see contact information you were hoping for, or need clarification on what they listed, please reach out to your pen pal match via email first and see if they can assist you. If they do not want to share other communication methods, please be flexible!" & vbCrLf & vbCrLf & _
    "Thank you again for registering for our pen pal program. We hope you have a nice time getting to know someone new, and we look forward to hearing about your experience at the end of the program in late September! Expect another email around this time next month with our next Topic of the Month suggestion." & vbCrLf & _
    "Best regards," & vbCrLf & _
    "The States' 4-H International Exchange Programs Team"
    
    outlookmailitem.To = edress
    outlookmailitem.Subject = Subject
    Set outlookmailitem.SendUsingAccount = outlookapp.Session.Accounts.Item(3)
    myAttachments.Add (attachment)
        outlookmailitem.Display
        'outlookmailitem.send
        Application.DisplayAlerts = True
        Set outlookapp = Nothing
        Set outlookmailitem = Nothing
      Next
End Sub




