VERSION 5.00
Begin VB.Form frmTextBox 
   Caption         =   "Test Textbox with dropdown"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   4
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox txtContactName 
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Tag             =   "2"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtCompanyName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Tag             =   "1"
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Contact Name:"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Company Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim suggestionForCompanyName As clsTextBoxWithGridDropDown
Dim suggestionForContactName As clsTextBoxWithGridDropDown
Dim i As Integer

Private Sub Form_Activate()
    Dim strDetailFields(3) As String
    strDetailFields(0) = "ContactName"
    strDetailFields(1) = "ContactTitle"
    strDetailFields(2) = "Address"
    
    suggestionForCompanyName.InitComponent Me, txtCompanyName, "CompanyName", strDetailFields, _
       "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb", _
        "SELECT * FROM Customers", "1"
        
    suggestionForContactName.InitComponent Me, txtContactName, "ContactName", strDetailFields, _
       "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.mdb", _
        "SELECT * FROM Customers", "2"
        
    txtCompanyName.SetFocus

End Sub

Private Sub Form_Load()
    Set suggestionForCompanyName = New clsTextBoxWithGridDropDown
    Set suggestionForContactName = New clsTextBoxWithGridDropDown
End Sub


