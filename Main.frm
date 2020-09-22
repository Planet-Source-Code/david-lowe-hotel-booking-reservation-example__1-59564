VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton About 
      Caption         =   "&About"
      Height          =   495
      Left            =   9720
      TabIndex        =   18
      Top             =   120
      Width           =   2055
   End
   Begin VB.Timer tmr_time 
      Interval        =   800
      Left            =   10320
      Top             =   5280
   End
   Begin VB.Frame frame_time 
      Caption         =   "Time"
      Height          =   1095
      Left            =   9720
      TabIndex        =   15
      Top             =   3480
      Width           =   2055
      Begin VB.Label lbl_time 
         Alignment       =   2  'Center
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frame_selroom 
      Caption         =   "Selected Room Information"
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   9495
      Begin VB.TextBox txt_status 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "????"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox memo_desc 
         BackColor       =   &H8000000F&
         Height          =   1560
         HideSelection   =   0   'False
         Left            =   1605
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1215
         Width           =   3615
      End
      Begin VB.Label lbl_status 
         Caption         =   "Status:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbl_desc 
         Caption         =   "Description:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lbl_rmname2 
         Caption         =   " "
         Height          =   255
         Left            =   1590
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lbl_rmname1 
         Caption         =   "Room Name:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl_rmnum2 
         Caption         =   "0"
         Height          =   255
         Left            =   1590
         TabIndex        =   8
         Top             =   495
         Width           =   4215
      End
      Begin VB.Label lbl_rmnum1 
         Caption         =   "Room Number:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frm_ViewOpt 
      Caption         =   "Rooms View Options"
      Height          =   2535
      Left            =   9720
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      Begin VB.OptionButton Opt_InUse 
         Caption         =   "Rooms In Use"
         Height          =   255
         Left            =   360
         MaskColor       =   &H8000000F&
         TabIndex        =   4
         Top             =   1365
         Width           =   1455
      End
      Begin VB.OptionButton Opt_Free 
         Caption         =   "Free Rooms"
         Height          =   255
         Left            =   360
         MaskColor       =   &H8000000F&
         TabIndex        =   3
         Top             =   1005
         Width           =   1335
      End
      Begin VB.OptionButton Opt_All 
         Caption         =   "All Rooms"
         Height          =   255
         Left            =   360
         MaskColor       =   &H8000000F&
         TabIndex        =   2
         Top             =   645
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lbl_optheader 
         Caption         =   "Only Show:"
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
         Left            =   180
         TabIndex        =   5
         Top             =   345
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid RoomGrid 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
      _Version        =   393216
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   1
      SelectionMode   =   1
   End
   Begin VB.Label lbl_hotelname 
      Caption         =   "Hotel Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   150
      Width           =   9375
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                   RANDOM EXAMPlE
'                   By David Lowe
'                   E: david@b0ff.co.uk
'
'
' I in no way suggest this is a well written piece of code
' and expect it is not as i am a very messy & inexperienced coder
'
'
' This example could be built on or help anyone trying to
' make a hotel booking/reservation system.
'
' This example was written to help and show how some of
' Visual Basic 's controls can be used.
'
' This example includes the use of:
'  * MSFlexGrid
'  * Database (Access DB's)
'  * And more…
'
' Feel free to contact me about anything.



' Required DB variables.
    Dim DBstate As Integer
    Dim DB As Database
    Dim DBRecord As Recordset
    Dim WS As Workspace
    Dim max As Integer
    
' Random variables
    Dim Rooms As Integer
    Dim LineCounter As Integer
    


Public Function RenderGrid()

    LineCounter = 1

    RoomGrid.Clear
    RoomGrid.Cols = 1
    RoomGrid.Rows = 1
    RoomGrid.FixedCols = 0
    RoomGrid.FixedRows = 0

    RoomGrid.Cols = 4
    
    'Until max record for rooms
    RoomGrid.Rows = recordcount + 1
    
    RoomGrid.FixedCols = 1
    RoomGrid.FixedRows = 1
    
    RoomGrid.ColWidth(0) = 1000
    RoomGrid.ColWidth(1) = 2000
    RoomGrid.ColWidth(2) = 4650
    RoomGrid.ColWidth(3) = 1500
    
 
    For i = 0 To 3
   
        RoomGrid.Row = 0
        RoomGrid.Col = i
        RoomGrid.CellFontBold = True
    
    Next i
    
    'Bold the room numbers
    For i = 1 To recordcount
    
        RoomGrid.Row = i
        RoomGrid.Col = 0
        RoomGrid.Text = i
        RoomGrid.CellFontBold = True
        RoomGrid.CellAlignment = 3
    
    Next i
    
    'Set Alignment for each row.
    For i = 1 To recordcount
    
        RoomGrid.Row = i
        RoomGrid.Col = 1
        RoomGrid.CellAlignment = 0
        RoomGrid.Col = 2
        RoomGrid.CellAlignment = 0
        RoomGrid.Col = 3
        RoomGrid.CellAlignment = 3
        
        
    Next i

 
 
    RoomGrid.TextMatrix(0, 0) = "Room No"
    RoomGrid.TextMatrix(0, 1) = "Room Name"
    RoomGrid.TextMatrix(0, 2) = "Room Type"
    RoomGrid.TextMatrix(0, 3) = "   Room Status"
    
    
    Rooms = recordcount
    
    
    
    For i = 1 To recordcount
    

        
        If DBRecord.Fields("RoomStatus").Value = False Then

            RoomGrid.TextMatrix(LineCounter, 0) = DBRecord.Fields("RoomID").Value
            RoomGrid.TextMatrix(LineCounter, 1) = DBRecord.Fields("RoomName").Value
            RoomGrid.TextMatrix(LineCounter, 2) = DBRecord.Fields("RoomDesc").Value

            RoomGrid.Row = LineCounter
            RoomGrid.Col = 3
            RoomGrid.Text = "IN USE"
            RoomGrid.CellFontBold = True
            RoomGrid.CellBackColor = &HC0C0FF
            RoomGrid.CellForeColor = vbWhite
        
        Else

            RoomGrid.TextMatrix(LineCounter, 0) = DBRecord.Fields("RoomID").Value
            RoomGrid.TextMatrix(LineCounter, 1) = DBRecord.Fields("RoomName").Value
            RoomGrid.TextMatrix(LineCounter, 2) = DBRecord.Fields("RoomDesc").Value

            RoomGrid.Row = LineCounter
            RoomGrid.Col = 3
            RoomGrid.Text = "Free"
            RoomGrid.CellFontBold = True
            RoomGrid.CellBackColor = &HC0FFC0
            
        End If
        
        LineCounter = LineCounter + 1
        DBRecord.MoveNext
        
    
    Next i
    

    Call CloseConn


End Function


    
Public Function opentable(dbname As String, tblname As String, ShowType As String)
    
    Set WS = DBEngine.Workspaces(0)
    Set DB = WS.OpenDatabase(App.Path & "\" & dbname)
    
    If ShowType = "ALL" Then
    
        Set DBRecord = DB.OpenRecordset(tblname, dbOpenTable)
    
    End If
    
    If ShowType = "FREE" Then
    
        Set DBRecord = DB.OpenRecordset("Select * from " & tblname & " where RoomStatus = TRUE")
    
    End If
    
    If ShowType = "INUSE" Then
    
        Set DBRecord = DB.OpenRecordset("Select * from " & tblname & " where RoomStatus = FALSE")
    
    End If
   
   
    DBRecord.MoveFirst
    DBRecord.MoveLast
    
    max = DBRecord.recordcount
    DBRecord.MoveFirst
    
    DBstate = 1
    
End Function


Public Function RM_Selected(dbname As String, tblname As String, SelRecord As Integer)
    
    Set WS = DBEngine.Workspaces(0)
    Set DB = WS.OpenDatabase(App.Path & "\" & dbname)
    
    Set DBRecord = DB.OpenRecordset("Select * from " & tblname & " where RoomID = " & SelRecord)
    
    DBRecord.MoveFirst
    DBRecord.MoveLast
    
    max = DBRecord.recordcount
    DBRecord.MoveFirst
    
    DBstate = 1
    
    lbl_rmnum2 = DBRecord.Fields("RoomID").Value
    lbl_rmname2 = DBRecord.Fields("RoomName").Value
    memo_desc = DBRecord.Fields("RoomDesc").Value
    
    If DBRecord.Fields("RoomStatus").Value = False Then
    
        txt_status.Text = "INUSE"
        txt_status.BackColor = "&HC0C0FF"
            
    
    Else
    
        txt_status.Text = "FREE"
        txt_status.BackColor = "&HC0FFC0"
    
    End If
    
End Function


Public Function recordcount()

    DBRecord.MoveFirst
    DBRecord.MoveLast
    
    recordcount = DBRecord.recordcount
    DBRecord.MoveFirst

End Function

Public Function CloseConn()

    DBRecord.Close
    DB.Close
    WS.Close
    
    Set WS = Nothing
    Set DB = Nothing
    Set DBRecord = Nothing
    
    DBstate = 0

End Function





Private Sub Command1_Click()

    RenderGrid

End Sub

Private Sub About_Click()

    'Show the about form
    frm_about.Show

End Sub

Private Sub Form_Load()

    Me.Top = 10
    Me.Left = 10

    Call opentable("DB.mdb", "TBL_Rooms", "ALL")
    RenderGrid

End Sub


Private Sub Opt_All_Click()
    
    Call opentable("DB.mdb", "TBL_Rooms", "ALL")
    RenderGrid
    
End Sub

Private Sub Opt_Free_Click()

    Call opentable("DB.mdb", "TBL_Rooms", "FREE")
    RenderGrid

End Sub

Private Sub Opt_InUse_Click()

    Call opentable("DB.mdb", "TBL_Rooms", "INUSE")
    RenderGrid

End Sub

Private Sub RoomGrid_Click()

    RoomGrid.Row = RoomGrid.RowSel
    RoomGrid.Col = 0
    
    Call RM_Selected("DB.mdb", "TBL_Rooms", RoomGrid.Text)
    
    Call CloseConn

End Sub

Private Sub tmr_time_Timer()

    lbl_time.Caption = Time

End Sub
