VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Resizable Controls"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HSB 
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3600
      Width           =   4455
   End
   Begin VB.ListBox LST 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "FRMMAIN.frx":0000
      Left            =   4560
      List            =   "FRMMAIN.frx":0007
      TabIndex        =   10
      Top             =   3000
      Width           =   4455
   End
   Begin VB.ComboBox CMB 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   9
      Text            =   "Hello, I am CMB"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.OptionButton OPT 
      Caption         =   "Hello, I am OPT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton CMD 
      Caption         =   "Hello, I am CMD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CheckBox CHK 
      Caption         =   "Hello, I am CHK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Frame FRAM 
      Caption         =   "Hello, I am FRAM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   4335
   End
   Begin VB.PictureBox PIC 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   120
      Picture         =   "FRMMAIN.frx":001C
      ScaleHeight     =   1815
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   1080
      Width           =   4335
   End
   Begin VB.ComboBox cmbObject 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FRMMAIN.frx":686E
      Left            =   120
      List            =   "FRMMAIN.frx":688D
      TabIndex        =   3
      Text            =   "PIC"
      Top             =   120
      Width           =   8895
   End
   Begin VB.TextBox TXT 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Hello, I am TXT"
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton cmdFalse 
      Caption         =   "FALSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdTrue 
      Caption         =   "TRUE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdFalse_Click()
    Select Case cmbObject.Text
        Case "PIC" 'Picturebox
            CanResize Me, PIC, False
            
        Case "TXT" 'Textbox
            CanResize Me, TXT, False
            
        Case "FRAM" 'Frame - Changes state, though I can not get it to resize yet
            CanResize Me, FRAM, False
            
        Case "CHK" 'Checkbox
            CanResize Me, CHK, False
            
        Case "CMD" 'Command Button
            CanResize Me, CMD, False
            
        Case "OPT" 'Option Button
            CanResize Me, OPT, False
            
        Case "CMB" 'Combobox
            CanResize Me, CMB, False
            
        Case "LST" 'Listbox
            CanResize Me, LST, False
            
        Case "HSB" 'Horizontal Scroll Bar
            CanResize Me, HSB, False
    End Select
End Sub

Private Sub cmdTrue_Click()
    Select Case cmbObject.Text
        Case "PIC" 'Picturebox
            InitialState Me, PIC
            CanResize Me, PIC, True
            
        Case "TXT" 'Textbox
            InitialState Me, TXT
            CanResize Me, TXT, True
            
        Case "FRAM" 'Frame - Changes state, though I can not get it to resize yet
            InitialState Me, FRAM
            CanResize Me, FRAM, True
            
        Case "CHK" 'Checkbox
            InitialState Me, CHK
            CanResize Me, CHK, True
            
        Case "CMD" 'Command Button
            InitialState Me, CMD
            CanResize Me, CMD, True
            
        Case "OPT" 'Option Button
            InitialState Me, OPT
            CanResize Me, OPT, True
            
        Case "CMB" 'Combobox
            InitialState Me, CMB
            CanResize Me, CMB, True
            
        Case "LST" 'Listbox
            InitialState Me, LST
            CanResize Me, LST, True
            
        Case "HSB" 'Horizontal Scroll Bar
            InitialState Me, HSB
            CanResize Me, HSB, True
    End Select
End Sub
