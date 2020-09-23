VERSION 5.00
Begin VB.Form frmMAin 
   Caption         =   "Singlelize"
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   1350
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton BtnDupes 
      Caption         =   "Singlelize"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox CBDupes 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Scam"
      Top             =   135
      Width           =   2475
   End
End
Attribute VB_Name = "frmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnDupes_Click()
    Singlelize CBDupes
End Sub

Private Sub Form_Load()
With CBDupes
     .AddItem "Scam"
     .AddItem "scam"
     .AddItem "Cheese"
     .AddItem "cheese"
     .AddItem "Rodney"
     .AddItem "godfried"
     .AddItem "rgodfried"
     .AddItem "RGODFRIED"
     .AddItem "RODNEY"
     .AddItem "GOdfrIed"
     .AddItem "Chello"
     .AddItem "CHELLO"
     .AddItem "ChElO"
     .AddItem "cheLo"
     .AddItem "SinGLELize"
     .AddItem "Singlelize"
End With
End Sub
