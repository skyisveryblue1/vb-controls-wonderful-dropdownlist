VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFlexGrid 
   Caption         =   "Test Grid with dropdown"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picUnchecked 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMFG.frx":0000
      ScaleHeight     =   169
      ScaleMode       =   0  'User
      ScaleWidth      =   169
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picChecked 
      Height          =   255
      Left            =   8640
      Picture         =   "frmMFG.frx":00EA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid mfgMain 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gridSuggestion As clsGridWithDropDown
Dim nTargetCol As Integer, i As Integer
Dim nCheckedRow As Integer

Private Sub mfgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = &HD Or KeyCode = &HA) Then
        'If mfgMain.Col <> nTargetCol Then Exit Sub
        
        gridSuggestion.EnterEdit mfgMain.Row, mfgMain.Col, ""
    End If
End Sub

Private Sub mfgMain_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 Then
        'If mfgMain.Col <> nTargetCol Then Exit Sub
        
        gridSuggestion.EnterEdit mfgMain.Row, mfgMain.Col, Chr(KeyAscii)
    End If
End Sub

Private Sub mfgMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If mfgMain.Col <> nTargetCol Then Exit Sub
    
    gridSuggestion.EnterEdit mfgMain.MouseRow, mfgMain.MouseCol, ""
End Sub
Private Sub Form_Activate()
    Dim strDetailFields(3) As String
    strDetailFields(0) = "ContactName"
    strDetailFields(1) = "ContactTitle"
    strDetailFields(2) = "Address"
    
    gridSuggestion.InitComponent Me, mfgMain, "CompanyName", strDetailFields, _
       "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb", _
        "SELECT * FROM Customers"
        
    nTargetCol = 1
End Sub

Private Sub mfgMain_Click()
    Dim nCurRow As Integer
    With mfgMain
        If .Col = 0 Then
            nCurRow = .Row
            
            If nCheckedRow <> 0 Then
                .Row = nCheckedRow
                Set .CellPicture = picUnchecked
            End If
            
            .Row = nCurRow
            If .CellPicture = picChecked Then
                Set .CellPicture = picUnchecked
                nCheckedRow = 0
            Else
                Set .CellPicture = picChecked
                nCheckedRow = .Row
            End If
           
        End If
    End With
End Sub

Private Sub Form_Load()
    picChecked.Visible = False
    picUnchecked.Visible = False
    
    With mfgMain
        .Cols = 3
        .Rows = 10
        .ColWidth(0) = 250 ' CheckBox column
        .ColWidth(1) = 2500
        .ColWidth(2) = 2000
        .ZOrder 1
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            'Align the checkbox
            .CellPictureAlignment = 4
            ' Set the default checkbox picture to the empty box
            Set .CellPicture = picUnchecked.Picture
        Next
    End With
    Set gridSuggestion = New clsGridWithDropDown
End Sub




