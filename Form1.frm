VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuotes 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get yahoo.com stock quotes"
   ClientHeight    =   4650
   ClientLeft      =   1785
   ClientTop       =   1515
   ClientWidth     =   6975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6975
   Begin VB.Timer tmrAutoUpdate 
      Enabled         =   0   'False
      Left            =   90
      Top             =   4275
   End
   Begin VB.Frame fraKeepUpdating 
      BackColor       =   &H80000012&
      Caption         =   "AutoUpdate"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   1350
      Width           =   3165
      Begin MSComctlLib.Slider Slider1 
         Height          =   285
         Left            =   1485
         TabIndex        =   7
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CheckBox chkAutoUpdate 
         BackColor       =   &H80000012&
         Caption         =   "Auto Update"
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   435
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   765
      Width           =   3165
   End
   Begin VB.ListBox lstCompanies 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1980
      Left            =   135
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin MSComctlLib.ListView lsvDisplayQuotes 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   2295
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16744448
      BackColor       =   16777152
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtQuotes 
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGetQuotes 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Get Quotes"
      Height          =   435
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   3165
   End
   Begin InetCtlsObjects.Inet inetQuotes 
      Left            =   5160
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j10 As Integer
Dim arrSelCompanies(100, 2) As String ' change the dimension of this array if you add more than 100 companies to the text file
Dim arrCompanies(100, 2) As String ' change the dimension of this array if you add more than 100 companies to the text file
Dim strQueryURl As String
Private response As Variant
Private z As Integer
Private arrSplit() As String


Private Sub chkAutoUpdate_Click()
    If chkAutoUpdate.Value = 1 Then
       tmrAutoUpdate.Interval = 60000
       tmrAutoUpdate.Enabled = True
    Else
       tmrAutoUpdate.Enabled = False
    End If
       
End Sub

Private Sub cmdGetQuotes_Click()
Dim not_first_symbol As Boolean
Dim symbol As String
Dim query_url As String
Dim i As Integer
Dim response As Variant
Dim intNoofRecords As Integer
Dim intTimeToCall As Integer
    'chkAutoUpdate.Value = 0
    cmdGetQuotes.Enabled = False
    lsvDisplayQuotes.ListItems.Clear
    MousePointer = vbHourglass
    txtQuotes.Text = ""
    DoEvents
    Call GetSelectedCompanies
    If arrSelCompanies(1, 1) = "" Then
        MsgBox "Hey, you have to select at least one company."
        MousePointer = vbDefault
        cmdGetQuotes.Enabled = True
        Exit Sub
    End If
    Call BuildQueryURL
    response = inetQuotes.OpenURL(strQueryURl)
    txtQuotes.Text = response
    arrSplit = Split(response, ",")
    intNoofRecords = (UBound(arrSplit) + (z * 1))
    intTimeToCall = intNoofRecords / 9
    For i = 1 To intTimeToCall
        Call PopulateListBox(arrSelCompanies(i, 1), arrSelCompanies(i, 2), arrSplit(1 + (8 * (i - 1))), arrSplit(2 + (8 * (i - 1))), arrSplit(3 + (8 * (i - 1))), arrSplit(4 + (8 * (i - 1))))
    Next i
    cmdGetQuotes.Enabled = True
    MousePointer = vbDefault
'    Call PopulateListBox(intNoofRecords)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdGetQuotes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdGetQuotes.Font.Size = 11
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdQuit.Font.Size = 11
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdQuit.Font.Size = 8
    cmdGetQuotes.Font.Size = 8
End Sub
Private Sub Form_Load()
Dim j As Integer
    Open App.Path & "/Companies.csv" For Input As #1
    Do While Not EOF(1)
       j = j + 1
       Input #1, arrCompanies(j, 1), arrCompanies(j, 2)
       lstCompanies.AddItem (arrCompanies(j, 1))
    Loop
    Close #1
    Slider1.Value = 5
    Call BuildTvwHeadings
End Sub


Private Sub Form_Unload(Cancel As Integer)
    inetQuotes.Cancel
End Sub

Private Sub PopulateListBox(strName As String, strTag As String, strCurrPrice As String, strDate As String, strTime As String, strChange As String)
Dim j As Integer
Dim j1 As Integer
Dim newItem As ListItem
    Set newItem = lsvDisplayQuotes.ListItems.Add(1, , strName)
    newItem.SubItems(1) = strTag
    newItem.SubItems(2) = strCurrPrice
    newItem.SubItems(3) = strChange
    newItem.SubItems(4) = strTime
    newItem.SubItems(5) = strDate
End Sub

Private Sub GetSelectedCompanies()
Dim i As Integer
    z = 0
    For i = 0 To lstCompanies.ListCount - 1
        If lstCompanies.Selected(i) Then
           z = z + 1
           arrSelCompanies(z, 1) = arrCompanies(i + 1, 1)
           arrSelCompanies(z, 2) = arrCompanies(i + 1, 2)
        End If
    Next i
    
End Sub
Private Sub BuildQueryURL()
Dim strCompaniesContribution As String

Dim i As Integer
    For i = 1 To z
        If i = 1 Then
            strCompaniesContribution = arrSelCompanies(i, 2)
        Else
            strCompaniesContribution = strCompaniesContribution + "," + arrSelCompanies(i, 2)
        End If
    Next i
    strQueryURl = "http://quote.yahoo.com/d/quotes.csv?s=" & strCompaniesContribution & "&f=sl1d1t1c1ohgv&e=.csv"

End Sub
Private Sub BuildTvwHeadings()
    lsvDisplayQuotes.ColumnHeaders.Add , , "Name", 1500
    lsvDisplayQuotes.ColumnHeaders.Add , , "Tag", 900
    lsvDisplayQuotes.ColumnHeaders.Add , , "Current Price", 1200
    lsvDisplayQuotes.ColumnHeaders.Add , , "Change", 900
    lsvDisplayQuotes.ColumnHeaders.Add , , "Time", 1000
    lsvDisplayQuotes.ColumnHeaders.Add , , "Date", 1100
    
End Sub

Private Sub tmrAutoUpdate_Timer()
    j10 = j10 + 1
    If j10 >= Slider1.Value Then
        j10 = 0
        Call cmdGetQuotes_Click
    End If
End Sub
