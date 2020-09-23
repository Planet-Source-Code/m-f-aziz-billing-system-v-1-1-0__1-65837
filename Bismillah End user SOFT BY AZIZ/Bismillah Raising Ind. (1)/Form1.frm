VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Bismillah Raising Ind"
   ClientHeight    =   6570
   ClientLeft      =   225
   ClientTop       =   1305
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "About"
      Height          =   375
      Left            =   8400
      TabIndex        =   63
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3840
      TabIndex        =   62
      Top             =   7200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker8 
      Height          =   285
      Left            =   1440
      TabIndex        =   61
      Top             =   6240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker7 
      Height          =   285
      Left            =   6720
      TabIndex        =   60
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   285
      Left            =   6720
      TabIndex        =   59
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   285
      Left            =   6720
      TabIndex        =   58
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   285
      Left            =   1440
      TabIndex        =   57
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   285
      Left            =   1440
      TabIndex        =   56
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   1440
      TabIndex        =   55
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   5880
      TabIndex        =   54
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarForeColor=   8388608
      CalendarTitleBackColor=   -2147483634
      CalendarTitleForeColor=   16711680
      CalendarTrailingForeColor=   0
      Format          =   20316161
      CurrentDate     =   38896
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Total"
      Height          =   315
      Left            =   3600
      TabIndex        =   52
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   8880
      TabIndex        =   51
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H80000018&
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   6360
      TabIndex        =   48
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H80000018&
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   1440
      TabIndex        =   47
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   9360
      TabIndex        =   45
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   3960
      TabIndex        =   44
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   9360
      TabIndex        =   39
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3960
      TabIndex        =   37
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   9360
      TabIndex        =   36
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   3960
      TabIndex        =   27
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Total Amount"
      Height          =   315
      Left            =   9480
      TabIndex        =   26
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6765
      TabIndex        =   25
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   5205
      TabIndex        =   24
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   9000
      TabIndex        =   22
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   9
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   6
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8880
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label29 
      Caption         =   "Date"
      Height          =   255
      Left            =   360
      TabIndex        =   53
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "Total"
      Height          =   255
      Left            =   8160
      TabIndex        =   50
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "CP + Damage Cloth"
      Height          =   255
      Left            =   4800
      TabIndex        =   49
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label26 
      Caption         =   "Total Meters"
      Height          =   255
      Left            =   360
      TabIndex        =   46
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label25 
      Caption         =   "Meters"
      Height          =   255
      Left            =   8520
      TabIndex        =   43
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   42
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Meters"
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Bismillah Raising Industry  Faisalabad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label22 
      Caption         =   "Date"
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "Meters"
      Height          =   255
      Left            =   8520
      TabIndex        =   35
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   34
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Meters"
      Height          =   255
      Left            =   3360
      TabIndex        =   33
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "Date"
      Height          =   255
      Left            =   600
      TabIndex        =   32
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Meters"
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Meters"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Date"
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Total Amount"
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Rate"
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Width"
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Net Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Ail Short"
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Sent Pieces"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Sent Cloth"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Received Pieces"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Received Cloth"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Quality"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Party Name"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CN As New ADODB.Connection
Private rs As New Recordset
Private rs1 As New Recordset
Private dbpath As String

Private Sub Command2_Click()
Dim htmltext1 As String
Dim htmltext2 As String
Dim htmltext3 As String
Dim htmltext4 As String
Dim htmltext5 As String
Dim htmltext6 As String
Dim htmltext7 As String
Dim htmltext8 As String
Dim htmltext9 As String
Dim htmltext10 As String

Dim htmlfinal As String


Dim ie As New InternetExplorer


htmltext1 = "<html><head><title>Bismillah Raising Industry (press Ctrl+P to print)</title></head><body><h1 align='center'>Bismillah Raising Ind. Faisalabad <br></h1><center><table width='582' height='92'> <tr><td width='284' height='19'><b> Company: </b>" & Text1.Text & "<td width='284' height='19'><b>DATE: </b>" & DTPicker1 & "</tr><br><tr><td width='284' height='19'><b>Quality: </b>" & Text3.Text & "<td width='284' height='19'><b>Received Cloth: </b>" & Text4.Text & "<td width='284' height='19'><b>Received Piece: </b>" & Text5.Text & "</tr>"
htmltext2 = htmltex1 & "<tr><td width='284' height='19'><b> Sent Cloth: </b>" & Text6.Text & "<td width='284' height='19'><b>Sent Pieces: </b>" & Text7.Text & "<td width='284' height='19'><b>Ail Short: </b>" & Text8.Text & "</td></tr>"
htmltext3 = "<tr><td width='284' height='19'><b> Net Quantity: </b>" & Text9.Text & "<td width='284' height='19'><b>Width: </b>" & Text10.Text & "<td width='284' height='19'><b>Rate: </b>" & Text11.Text & "</td></tr><br>"
htmltext4 = "<tr></tr><tr><td width='284' height='19'></td><td width='284' height='19'><b> Total Amount: </b>" & Text12.Text & "</td></tr></table>"

htmltext5 = "<br><table border='1'><tr><td width='284' height='19'><b>Date:</b>" & DTPicker2 & "</td><td width='284' height='19'><b>Amount:" & Text14.Text & "</td><td width='284' height='19'><b>Date:</b>" & DTPicker5 & "</td><td width='284' height='19'><b>Amount:</b>" & Text16.Text & "</td></tr>"
htmltext6 = "<br><tr><td width='284' height='19'><b>Date:</b>" & DTPicker3 & "</td><td width='284' height='19'><b>Amount:" & Text18.Text & "</td><td width='284' height='19'><b>Date:</b>" & DTPicker6 & "</td><td width='284' height='19'><b>Amount:</b>" & Text20.Text & "</td></tr>"
htmltext7 = "<br><tr><td width='284' height='19'><b>Date:</b>" & DTPicker4 & "</td><td width='284' height='19'><b>Amount:" & Text22.Text & "</td><td width='284' height='19'><b>Date:</b>" & DTPicker7 & "</td><td width='284' height='19'><b>Amount:</b>" & Text22.Text & "</td></tr></table>"

htmltext8 = "<br><table><tr><td width='284' height='19'><b>Total Meters:</b>" & Text25.Text & "</td><td width='284' height='19'><b>CP+Damage Cloth:</b>" & Text26.Text & "</td><td width='284' height='19'><b>Total:</b>" & Text27.Text & "</td></tr><tr><td width='284' height='19'><b>Date:</b>" & DTPicker8 & "</td></tr> <tr><td width='800' height='19' align='right'><b>Signature:</b>________________" & "</td></tr></table></body></html>"



htmlfinal = htmltext1 & htmltext2 & htmltext3 & htmltext4 & htmltext5 & htmltext6 & htmltext7 & htmltext8
Open App.Path + "/empty.html" For Output As #1
Print #1, htmlfinal
Close #1

ie.Visible = True
ie.Navigate App.Path + "/empty.html"

End Sub

Private Sub Command6_Click()
                    
                    Text1.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                    Text5.Text = ""
                    Text6.Text = ""
                    Text7.Text = ""
                    Text8.Text = ""
                    Text12.Text = ""
                    Text16.Text = ""
                    Text18.Text = ""
                    Text20.Text = ""
                    Text22.Text = ""
                    Text24.Text = ""
                    
                    
                    Text25.Text = ""
                    Text26.Text = ""
                    Text27.Text = ""
End Sub

Private Sub Command7_Click()
MsgBox "AZM TECHNOLOGY GROUP", vbInformation, "About"



End Sub

Private Sub Form_Load()
                    
                    
                    
                    Text9.Text = 1
                    Text10.Text = 1
                    Text11.Text = 1
                    
                   'Open the database
                    
                    dbpath = App.Path & "\db1.mdb"
                    
                    With CN
                            .CommandTimeout = 5
                            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & dbpath & "              "
                            .CursorLocation = adUseClient
                            .Open
                            
                    
                    'Persist Security Info=false;Jet OLedb:Database password=azm"
                    
                    rs.Open "Select * from Party", CN, adOpenStatic, adLockOptimistic
                    'rs1.Open "Select * from ClothDetail", CN, adOpenStatic, adLockOptimistic
                    
                    End With
                     
End Sub


Private Sub Command1_Click()
   If Text1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Then
   
    MsgBox "One of your field is empty..!", vbExclamation, "Empty Field"
    
Else
    With rs
                    .AddNew
                    .Fields(0) = Text1
                    .Fields(1) = DTPicker1
                    .Fields(2) = Text3
                    .Fields(3) = Text4
                    .Fields(4) = Text5
                    .Fields(5) = Text6
                    .Fields(6) = Text7
                    .Fields(7) = Text8
                    .Fields(8) = Text9
                    .Fields(9) = Text10
                    .Fields(10) = Text11
                    .Fields(11) = Text12
                    
                    .Fields(12) = Text25
                    .Fields(13) = DTPicker8
                    .Fields(14) = Text26
                    .Fields(15) = Text27
                    '.Fields(16) = Text28
                    
                    
                    .Update
                
     
        MsgBox "Record is added successfully...!", vbinformtaion, "Bismillah Raising Ind."
                    

End With
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
On Error Resume Next
                    
                    
                    Dim a As Double
                    a = Text9 * Text10 * Text11
                    Text12 = a
                    
End Sub

Private Sub Command5_Click()
Dim a As Double
On Error Resume Next

a = Val(Text14) + Val(Text16) + Val(Text18) + Val(Text20) + Val(Text22) + Val(Text24)
Text25.Text = a

End Sub



Private Sub mnuparty_Click(Index As Integer)

End Sub

Private Sub mnusentcloth_Click(Index As Integer)
Form2.Show vbModal

End Sub

Private Sub Form_Unload(Cancel As Integer)

     CN.Close
     End
     
End Sub



