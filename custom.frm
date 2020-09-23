VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   " Custom Database"
   ClientHeight    =   3000
   ClientLeft      =   6465
   ClientTop       =   5805
   ClientWidth     =   5220
   Icon            =   "custom.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   120
      TabIndex        =   24
      Top             =   1230
      Width           =   1605
      Begin VB.OptionButton Option2 
         Caption         =   "Modify Current"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Create New DB"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   210
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   1
      Left            =   270
      TabIndex        =   23
      Top             =   2520
      Width           =   1125
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2880
      TabIndex        =   21
      Top             =   2550
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2880
      TabIndex        =   20
      Top             =   2310
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2880
      TabIndex        =   19
      Top             =   2070
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   2880
      TabIndex        =   18
      Top             =   1830
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2880
      TabIndex        =   17
      Top             =   1590
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2880
      TabIndex        =   16
      Top             =   1350
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2880
      TabIndex        =   15
      Top             =   1110
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2880
      TabIndex        =   14
      Top             =   870
      Visible         =   0   'False
      Width           =   2020
   End
   Begin VB.ComboBox cboCols 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   870
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&OK"
      Height          =   300
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   2190
      Width           =   1125
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   630
      Width           =   2020
   End
   Begin VB.TextBox Headers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   210
      Width           =   3525
   End
   Begin VB.Label Label2 
      Caption         =   "Num Columns:"
      Height          =   225
      Left            =   240
      TabIndex        =   22
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 9"
      Height          =   180
      Index           =   9
      Left            =   2160
      TabIndex        =   13
      Top             =   2580
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 8"
      Height          =   180
      Index           =   8
      Left            =   2160
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 7"
      Height          =   180
      Index           =   7
      Left            =   2160
      TabIndex        =   11
      Top             =   2100
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 6"
      Height          =   180
      Index           =   6
      Left            =   2160
      TabIndex        =   10
      Top             =   1860
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 5"
      Height          =   180
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 4"
      Height          =   180
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   1380
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 3"
      Height          =   180
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Top             =   1140
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 2"
      Height          =   180
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      Top             =   900
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCols 
      Caption         =   "Column 1"
      Height          =   180
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   660
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Main Caption"
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   1245
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click(Index As Integer)
    Dim i As Long
    If (Index = 0) Then
        frmCustDB.Caption = " " & Headers(0).Text
        lColCnt = CLng(cboCols.Text)
        If (lColFlag > lColCnt) Then lColFlag = lColCnt
        Call frmCustDB.ListViewHeaders
        If Option1.Value Then lItemCnt = 0
        Call frmCustDB.FillListView
        For i = 1 To lColCnt
            frmCustDB.ListView1.ColumnHeaders.Item(i).Text = Headers(i).Text
            frmCustDB.Labels(i - 1).Caption = "&" & Headers(i).Text
        Next i
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    Headers(0).SelStart = 0
    Headers(0).SelLength = Len(Headers(0).Text)
    Headers(0).SetFocus
End Sub

Private Sub Form_Initialize()
    cboCols.AddItem "2"
    cboCols.AddItem "3"
    cboCols.AddItem "4"
    cboCols.AddItem "5"
    cboCols.AddItem "6"
    cboCols.AddItem "7"
    cboCols.AddItem "8"
    cboCols.AddItem "9"
End Sub

Private Sub cboCols_Click()
    Dim i As Long
    Do While (i < CLng(cboCols.Text))
        i = i + 1
        Headers(i).Visible = True
        lblCols(i).Visible = True
    Loop
    i = i + 1
    Do While (i < Headers.Count)
        Headers(i).Visible = False
        lblCols(i).Visible = False
        i = i + 1
    Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = (UnloadMode = 0)
    Me.Hide
End Sub
