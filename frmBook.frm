VERSION 5.00
Begin VB.Form frmPhone 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "myBooks 2.0"
   ClientHeight    =   3840
   ClientLeft      =   1920
   ClientTop       =   2175
   ClientWidth     =   7395
   Icon            =   "frmBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7395
   Begin VB.ComboBox cbo1 
      BeginProperty Font 
         Name            =   "Serpentine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmBook.frx":08CA
      Left            =   1680
      List            =   "frmBook.frx":08CC
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "Serpentine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Nex&t"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Serpentine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Serpentine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Serpentine"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CyberBot Books"
      BeginProperty Font 
         Name            =   "Broadway BT"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Returned"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Books Name"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label lblComments 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lender Name"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Noyes
Dim mPerson As PersonInfo
Dim mFileNum As Integer
Dim mRecordLen As Long
Dim mCurrentRecord As Long
Dim mLastRecord As Long
Dim ShiftTest As Integer

'Add the following code to the form's MouseDown event:
Private Sub Form_MouseDown(Button As Integer, _
      Shift As Integer, X As Single, Y As Single)
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 5
      MsgBox "P H O N E  :  8 8 0  - 0 2  -  9 6 6 0 0 8 3 ", vbInformation, "My Phone"
      Case 6
       MsgBox "ehasan@citechco.net, ehasan@yahoo.com" & vbCrLf & vbCrLf & "       emran_the_cracker@yahoo.com", vbInformation, "My E-mail"
      Case 7
        MsgBox "FLAT # 3-A AZIZ CO-OPERATIVE HOUSING COMPLEX," & vbCrLf & vbCrLf & "   SHAHBAG, RAMNA, DHAKA-1000, BANGLADESH.", vbInformation, "My Address"
      End Select
End Sub

Public Sub SaveCurrentRecord()
'assaign data to valid variable
mPerson.BookID = txtName.Text
mPerson.BookName = txtPhone.Text
mPerson.LenderName = txtComments.Text
mPerson.Date = txtDate.Text
mPerson.Returned = cbo1.ListIndex
'put data to file
Put #mFileNum, mCurrentRecord, mPerson
End Sub
Public Sub ShowCurrentRecord()
'get current record from file
Get #mFileNum, mCurrentRecord, mPerson
'display current record
txtName.Text = Trim(mPerson.BookID)
txtPhone.Text = Trim(mPerson.BookName)
txtComments.Text = Trim(mPerson.LenderName)
txtDate.Text = Trim(mPerson.Date)
If mPerson.Returned = "1" Then
cbo1.ListIndex = 1
Else
cbo1.ListIndex = 0
End If
'show record number
frmPhone.Caption = "CyberBot Books 2.0" + Str(mCurrentRecord) + "/" + Str(mLastRecord)
End Sub

Private Sub cmdExit_Click()
SaveCurrentRecord
Unload Me
End Sub

Private Sub Form_Load()
'add necessary item to the combobox
cbo1.AddItem "Yes"
cbo1.AddItem "No"
cbo1.ListIndex = 1

' Calculate the length of a record.
mRecordLen = Len(mPerson)

' Get the next available file number.
mFileNum = FreeFile

' Open the file for random-access. If the file

' does not exist then it is created.

Open "PHONE.DAT" For Random As mFileNum Len = mRecordLen

' Update gCurrentRecord.
mCurrentRecord = 1

' Find what is the last record number of

' the file.

mLastRecord = FileLen("PHONE.DAT") / mRecordLen

' If the file was just created

' (i.e. mLastRecord=0) then update mLastRecord

' to 1.

If mLastRecord = 0 Then

mLastRecord = 1

End If

' Display the current record.
ShowCurrentRecord

End Sub

Private Sub cmdNew_Click()

' Save the current record.
SaveCurrentRecord

' Add a new blank record.

mLastRecord = mLastRecord + 1

mPerson.BookID = ""

mPerson.BookName = ""

mPerson.LenderName = ""

mPerson.Date = ""
mPerson.Returned = ""

Put #mFileNum, mLastRecord, mPerson
' Update gCurrentRecord.
mCurrentRecord = mLastRecord

' Display the record that was just created^.
ShowCurrentRecord

' Give the focus to the txtName field.
txtName.SetFocus

End Sub


Private Sub cmdNext_Click()

' If the current record is the last record,
' beep and display an error message. Otherwise,
' save the current record and skip to the
' next record.
If mCurrentRecord = mLastRecord Then

Beep

MsgBox "It's the last record !", vbExclamation, "CyberBot Studental 2.0"
Else

SaveCurrentRecord

mCurrentRecord = mCurrentRecord + 1

ShowCurrentRecord
End If

' Give the focus to the txtName field.
txtName.SetFocus

End Sub


Private Sub cmdback_Click()

' If the current record is the first record,
' beep and display an error message. Otherwise,
' save the current record and go to the
' previous record.
If mCurrentRecord = 1 Then
Beep

MsgBox "It's the first record !", vbExclamation, "CyberBot Studental 2.0"
Else

SaveCurrentRecord

mCurrentRecord = mCurrentRecord - 1

ShowCurrentRecord
End If

' Give the focus to the txtName field.
txtName.SetFocus

End Sub

Private Sub Form_Paint()
'draw the shadow
Shadow Me, txtName, 2, vbBlack
Shadow Me, txtPhone, 2, vbBlack
Shadow Me, txtComments, 2, vbBlack
Shadow Me, txtDate, 2, vbBlack
Shadow Me, cbo1, 2, vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveCurrentRecord
End Sub
Private Sub cmdSearch_Click()

Dim NameToSearch As String

Dim Found As Integer

Dim RecNum As Long

Dim TmpPerson As PersonInfo

' Get the name to search from the user.
NameToSearch = InputBox("Search for Book ID :", "Search")

' If the user did not enter a name, exit
' from this procedure.
If NameToSearch = "" Then

' Give the focus to the txtName field.

txtName.SetFocus

' Exit this procedure.

Exit Sub
End If

' Convert the name to be searched to upper case.
NameToSearch = UCase(NameToSearch)

' Initialize the Found flag to False.
Found = False

' Search for the name that the user entered.
For RecNum = 1 To mLastRecord
Get #mFileNum, RecNum, TmpPerson

If NameToSearch = UCase(Trim(TmpPerson.BookID)) Then
Found = True
Exit For
End If
Next

' If the name was found, display the record
' of the found name.
If Found = True Then

SaveCurrentRecord

mCurrentRecord = RecNum

ShowCurrentRecord
Else

MsgBox "The Book ID you entered could not be found!" & vbCrLf & vbCrLf & "Tip : Enter the correct Book ID to search !", vbInformation, "CyberBot Studental 2.0"


End If

' Give the focus to the txtName field.
txtName.SetFocus

End Sub
Private Sub cmdDelete_Click()
Dim DirResult
Dim TmpFileNum
Dim TmpPerson As PersonInfo
Dim RecNum As Long
Dim TmpRecNum As Long

' Before deleting get a confirmation from the user.
If MsgBox("Delete the current record?", vbYesNo + vbCritical, "CyberBot Studental 2.0") = vbNo Then

' Give the focus to the txtName field.

txtName.SetFocus

' Exit the procedure without deleting.

Exit Sub
End If

' To physically delete the current record of PHONE.DAT,

' all the records of PHONE.DAT, except the

' current record, are copied into a temporary file

' (PHONE.TMP) and then the file PHONE.TMP is copied into

' PHONE.DAT:

' Make sure that PHONE.TMP does not exist.
If Dir("PHONE.TMP") = "PHONE.TMP" Then

Kill "PHONE.TMP"
End If

' Create PHONE.TMP with the same format

' as PHONE.DAT.

TmpFileNum = FreeFile

Open "PHONE.TMP" For Random As TmpFileNum Len = mRecordLen

' Copy all the records from PHONE.DAT
' to PHONE.TMP, except the current record.
RecNum = 1
TmpRecNum = 1
Do While RecNum < mLastRecord + 1
If RecNum <> mCurrentRecord Then
Get #mFileNum, RecNum, TmpPerson
Put #TmpFileNum, TmpRecNum, TmpPerson
TmpRecNum = TmpRecNum + 1
End If
RecNum = RecNum + 1
Loop

' Delete PHONE.DAT.
Close mFileNum
Kill "PHONE.DAT"

' Rename PHONE.TMP into PHONE.DAT.

Close TmpFileNum

Name "PHONE.TMP" As "PHONE.DAT"

' Re-open PHONE.DAT.

mFileNum = FreeFile

Open "PHONE.DAT" For Random As mFileNum Len = mRecordLen

' Update the value of LastRecord.
mLastRecord = mLastRecord - 1

' Make sure that gLastRecord is not 0
If mLastRecord = 0 Then mLastRecord = 1

If mCurrentRecord > mLastRecord Then
   mCurrentRecord = mLastRecord
End If
ShowCurrentRecord
txtName.SetFocus
End Sub

