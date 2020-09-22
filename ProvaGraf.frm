VERSION 5.00
Begin VB.Form Graph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Graph"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Graph"
      Height          =   280
      Left            =   1680
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "By: CiTRiX"
      Height          =   255
      Left            =   1200
      TabIndex        =   28
      Top             =   3600
      Width           =   855
   End
   Begin VB.Line DLine 
      BorderColor     =   &H00FF0000&
      Index           =   0
      Visible         =   0   'False
      X1              =   1320
      X2              =   1800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label20 
      Caption         =   "Sun:"
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Sat:"
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "Fri:"
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "Thu:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "Wed:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "Tue:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Mon:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   480
      Y1              =   1540
      Y2              =   250
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   2880
      Y1              =   1550
      Y2              =   1550
   End
   Begin VB.Label Label14 
      Caption         =   "Sun"
      Height          =   255
      Left            =   2610
      TabIndex        =   20
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Sat"
      Height          =   255
      Left            =   2250
      TabIndex        =   19
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "Fri"
      Height          =   255
      Left            =   1950
      TabIndex        =   18
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "Thu"
      Height          =   255
      Left            =   1540
      TabIndex        =   17
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "Wed"
      Height          =   255
      Left            =   1150
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Tue"
      Height          =   255
      Left            =   810
      TabIndex        =   15
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Mon"
      Height          =   255
      Left            =   440
      TabIndex        =   14
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0-"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "1-"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "2-"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3-"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "4-"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "5-"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   2880
      Y1              =   350
      Y2              =   350
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   2880
      Y1              =   590
      Y2              =   590
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   2880
      Y1              =   830
      Y2              =   830
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   2880
      Y1              =   1070
      Y2              =   1070
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00C0C0C0&
      X1              =   480
      X2              =   2880
      Y1              =   1310
      Y2              =   1310
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l As Integer
Dim Mon, Tue, Wed, Thu, Fri, Sat, Sun As Long

Private Sub Command1_Click()
'Return the error if the number in the txt box isn't from 0 to 5
If Val(Text1.Text) < 0 Or Val(Text1.Text) > 5 Then GoTo ERROR
If Val(Text2.Text) < 0 Or Val(Text2.Text) > 5 Then GoTo ERROR
If Val(Text3.Text) < 0 Or Val(Text3.Text) > 5 Then GoTo ERROR
If Val(Text4.Text) < 0 Or Val(Text4.Text) > 5 Then GoTo ERROR
If Val(Text5.Text) < 0 Or Val(Text5.Text) > 5 Then GoTo ERROR
If Val(Text6.Text) < 0 Or Val(Text6.Text) > 5 Then GoTo ERROR
If Val(Text7.Text) < 0 Or Val(Text7.Text) > 5 Then GoTo ERROR
'Place the line DLine(0) on the graph
DLine(0).X1 = Mon
DLine(0).X2 = Tue
DLine(0).Y1 = Val(1550) - Val(240) * Val(Text1.Text)
DLine(0).Y2 = Val(1550) - Val(240) * Val(Text2.Text)
DLine(0).Visible = True
'Place the line DLine(1) on the graph
DLine(1).X1 = Tue
DLine(1).X2 = Wed
DLine(1).Y1 = Val(1550) - Val(240) * Val(Text2.Text)
DLine(1).Y2 = Val(1550) - Val(240) * Val(Text3.Text)
DLine(1).Visible = True
'Place the line DLine(2) on the graph
DLine(2).X1 = Wed
DLine(2).X2 = Thu
DLine(2).Y1 = Val(1550) - Val(240) * Val(Text3.Text)
DLine(2).Y2 = Val(1550) - Val(240) * Val(Text4.Text)
DLine(2).Visible = True
'Place the line DLine(3) on the graph
DLine(3).X1 = Thu
DLine(3).X2 = Fri
DLine(3).Y1 = Val(1550) - Val(240) * Val(Text4.Text)
DLine(3).Y2 = Val(1550) - Val(240) * Val(Text5.Text)
DLine(3).Visible = True
'Place the line DLine(4) on the graph
DLine(4).X1 = Fri
DLine(4).X2 = Sat
DLine(4).Y1 = Val(1550) - Val(240) * Val(Text5.Text)
DLine(4).Y2 = Val(1550) - Val(240) * Val(Text7.Text)
DLine(4).Visible = True
'Place the line DLine(5) on the graph
DLine(5).X1 = Sat
DLine(5).X2 = Sun
DLine(5).Y1 = Val(1550) - Val(240) * Val(Text7.Text)
DLine(5).Y2 = Val(1550) - Val(240) * Val(Text6.Text)
DLine(5).Visible = True
Exit Sub
'The error
ERROR:
MsgBox "Check the numbers in txt boxes!" & vbCrLf & "Max number is 5" & vbCrLf & "Min number is 0", vbCritical, "ERROR"
Exit Sub

End Sub

Private Sub Form_Load()
'Set how far is the middle of the labels from the left
Mon = Val(600)
Tue = Val(960)
Wed = Val(1320)
Thu = Val(1680)
Fri = Val(2040)
Sat = Val(2400)
Sun = Val(2760)
'load 5 lines and set them in front of the gray lines (and other objects)
For l = 1 To 5
Load DLine(l)
DLine(l).ZOrder Front
Next l
End Sub

'SORY FOR MY BAD ENGLISH :D
Private Sub Label21_Click()

End Sub
