VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txt_contact1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "david@b0ff.co.uk"
      Top             =   2715
      Width           =   1815
   End
   Begin VB.TextBox memo_aboutdesc 
      BackColor       =   &H8000000F&
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frm_about.frx":0000
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lbl_contact1 
      Caption         =   "Contact:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lbl_auth2 
      Caption         =   "David Lowe"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label lbl_auth1 
      AutoSize        =   -1  'True
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   630
   End
   Begin VB.Line ln_div 
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lbl_hotel 
      Caption         =   "Hotel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ok_Click()

    'Hide the about window
    Me.Hide

End Sub
