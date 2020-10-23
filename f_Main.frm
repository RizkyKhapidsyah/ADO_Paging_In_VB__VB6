VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form f_Main 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   Picture         =   "f_Main.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LV_Results 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4210752
      BackColor       =   15066597
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Currency Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Country Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apha 2 Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Currency Code A"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Currency Code N"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox T_Search 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   4470
   End
   Begin VB.Label Top_Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "ADO Paging through a recordset in Visual Basic [ Example ] by XIII"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   30
      Width           =   4815
   End
   Begin VB.Label B_Close 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5880
      MouseIcon       =   "f_Main.frx":4B27
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Label B_Search 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5040
      MouseIcon       =   "f_Main.frx":4E31
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   435
      Width           =   855
   End
   Begin VB.Label t_Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label T_Page 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   960
      MouseIcon       =   "f_Main.frx":513B
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4710
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label T_Results 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   975
      TabIndex        =   5
      Top             =   5265
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   5265
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Showing :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   5265
      Width           =   705
   End
   Begin VB.Label T_Showing_Records 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto page :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   4710
      Width           =   840
   End
End
Attribute VB_Name = "f_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Source Code dimulai dari sini
'Created by Rizky Khapidsyah

Private CN As ADODB.Connection
Private RS As ADODB.Recordset

Private Sub B_Close_Click()
    Unload Me
End Sub

Private Sub B_Search_Click()
    If Len(Trim(T_Search)) < 1 Then
        MsgBox "Please enter some text in a search TextBox."
        Exit Sub
    End If
    
    Temp_String = "WHERE (((t_Cur.Cur_Name) Like '%" & T_Search & "%'  )) OR (((t_Cur.Country_Name) Like '%" & T_Search & "%'  )) OR (((t_Cur.Alpha2_Code) Like '%" & T_Search & "%'  )) OR (((t_Cur.Currency_CodeA) Like '%" & T_Search & "%'  )) OR (((t_Cur.Currency_CodeN) Like '%" & T_Search & "%'  ));"

    Me.MousePointer = 11
    t_Status.Caption = "Checking connection..."
    
    If RS.State <> adStateClosed Then
        RS.Close
    End If

    RS.Open "Select * from t_Cur " & Temp_String, CN, adOpenStatic, adLockReadOnly
    
    t_Status.Caption = "Searching..."

    If RS.RecordCount > 0 Then
        RS.MoveLast
        RS.MoveFirst
    End If
    
    t_Status.Caption = "Building reults list..."

    Call Build_Results
    
        
    t_Status.Caption = "..."
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()
    Call Round_Corners(Me)
    Call Make_On_Top(Me.HWND, True)
    
    Set CN = New ADODB.Connection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\VB_Pages_2000.mdb;Persist Security Info=False"
    Set RS = New ADODB.Recordset
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Make_On_Top(Me.HWND, False)
    Set RS = Nothing
    CN.Close
    Set CN = Nothing
End Sub

Private Sub Build_Results(Optional Start_From = 0)
    
On Error GoTo Err_1
    
    Dim LI As ListItem   ' ListItem object
    Dim Temp_Counter As Long
    Dim Last_Page As Long ' Last page in current recordset
    Dim Start_Page As Long ' The page we will start from [ Start from=21 , Start_Page = 20 ]
    Dim X As Long
    
    
    LV_Results.ListItems.Clear
    Temp_Counter = 0
    
    With RS
        If .RecordCount > 0 Then
            .Move Start_From * 13, 1
        End If
        
        
        Do While Not .EOF And Temp_Counter < 13
           ' DoEvents
            Set LI = LV_Results.ListItems.Add(, "K" & !Cur_ID, !Cur_Name)
            LI.SubItems(1) = !Country_Name
            LI.SubItems(2) = IIf(IsNull(!Alpha2_Code) = True, " ", !Alpha2_Code)
            LI.SubItems(3) = IIf(IsNull(!Currency_CodeA) = True, " ", !Currency_CodeA)
            LI.SubItems(4) = IIf(IsNull(!Currency_CodeN) = True, " ", !Currency_CodeN)
            .MoveNext
            Temp_Counter = Temp_Counter + 1
        Loop
        
        T_Results.Caption = CStr(.RecordCount)
        
        ' Calculating Showing_Records value
        If .RecordCount > 0 Then
            T_Showing_Records.Caption = (Start_From * 13) + 1 & " - "
            If (Start_From * 13) + 1 + 13 >= .RecordCount Then
                T_Showing_Records.Caption = T_Showing_Records.Caption & .RecordCount
            Else
                T_Showing_Records.Caption = T_Showing_Records.Caption & (Start_From * 13) + 13
            End If
        Else
            T_Showing_Records.Caption = "0"
        End If

        
        
        ' Removing old page navigators
        For T = 1 To T_Page.Count - 1
            Unload T_Page(T)
        Next
            
        ' Getting last page in current recordset
        If .RecordCount Mod 13 > 0 Then
            Last_Page = Int(.RecordCount / 13) + 1
        Else
            Last_Page = Int(.RecordCount / 13)
        End If
   
        ' Geting first page we will show [ Start_Page ]
        For y = 1 To Last_Page Step 6
            If Start_From + 1 >= y And Start_From + 1 <= y + 5 Then
                Exit For
            End If
        Next
   
        Start_Page = y
        X = 1
            
        ' If we are showing pages not from first 20... <<- [ Previous ]
        If y > 1 Then
            Load T_Page(T_Page.Count)
            T_Page(T_Page.Count - 1).Caption = "<<-"
            T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
            T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
            T_Page(T_Page.Count - 1).Visible = True
        End If
            
        For T = Start_Page To Last_Page
            Load T_Page(T_Page.Count)
            If X > 6 Then ' If there are more pages then we can show... ->> [ Next ]
                T_Page(T_Page.Count - 1).Caption = "->>"
                T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
                T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
                T_Page(T_Page.Count - 1).Visible = True
                Exit For
            Else
                T_Page(T_Page.Count - 1).Caption = CStr(T)
                T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
                T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
                If T = Start_From + 1 Then ' If this is a current page
                    T_Page(T_Page.Count - 1).ForeColor = &HFF&
                End If
                T_Page(T_Page.Count - 1).Visible = True
            End If
            X = X + 1
        Next
    End With
    
    
Exit_Sub:
   Exit Sub
    
Err_1:
    MsgBox Err.Description, vbOKOnly + vbCritical + vbApplicationModal, "StaCS : System error # " & Err.Number
    Resume Exit_Sub
    
End Sub

Private Sub T_Page_Click(Index As Integer)
    
On Error GoTo Err_1
    
    Me.MousePointer = 11
    Me.AutoRedraw = False
    
    If T_Page(Index).Caption = "->>" Then
            Call Build_Results(Val(T_Page(Index - 1).Caption))
    ElseIf T_Page(Index).Caption = "<<-" Then
        Call Build_Results(Val(T_Page(Index + 1).Caption) - 2)
    Else
        Call Build_Results(Val(T_Page(Index).Caption) - 1)
    End If
    
    Me.AutoRedraw = True
    Me.MousePointer = 0
Exit_Sub:
   Exit Sub
    
Err_1:
    Resume Exit_Sub
End Sub

Private Sub Top_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub
