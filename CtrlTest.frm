VERSION 5.00
Object = "*\Avb6Ruler.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CtrlTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruler OCX Test Form"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   FillColor       =   &H00FFFFFF&
   Icon            =   "CtrlTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBlocca 
      Caption         =   "Blocca righello"
      Height          =   405
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1140
      Width           =   1305
   End
   Begin vb6Ruler.Ruler Ruler1 
      Height          =   2955
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5212
      BorderStyle     =   0
      Orientation     =   1
      MargineDx       =   5
      RulerDx         =   0
   End
   Begin VB.CommandButton cmdFissaMargini 
      Caption         =   "Fissa margini"
      Height          =   405
      Left            =   6270
      TabIndex        =   4
      Top             =   600
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Esci"
      Height          =   405
      Left            =   6270
      TabIndex        =   3
      Top             =   2910
      Width           =   1305
   End
   Begin VB.CheckBox chkIndicatore 
      Caption         =   "Abilita indicatore"
      Height          =   405
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   1305
   End
   Begin vb6Ruler.Ruler Righello 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10345
      _ExtentY        =   582
      BorderStyle     =   0
      MarginColor     =   16711680
      Minimo          =   -1
      MargineSx       =   1
      MargineDx       =   7
      RulerSxTop      =   1
      RulerDx         =   7.5
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2895
      Left            =   390
      TabIndex        =   1
      Top             =   420
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"CtrlTest.frx":038A
   End
End
Attribute VB_Name = "CtrlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkIndicatore_Click()
    Righello.position = (chkIndicatore.Value = vbChecked)
    Ruler1.position = (chkIndicatore.Value = vbChecked)
End Sub

Private Sub chkBlocca_Click()
    Righello.Locked = (chkBlocca.Value = vbChecked)
End Sub

Private Sub cmdFissaMargini_Click()
    Righello.MargineSx = Righello.RulerSxBot
    Righello.MargineDx = Righello.RulerDx
    Text1.Text = Righello.MargineSx & " <--> " & Righello.MargineDx
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    chkIndicatore.Value = IIf(Righello.position, vbChecked, vbUnchecked)
    chkBlocca.Value = IIf(Righello.Locked, vbChecked, vbUnchecked)
    Text1.Text = "This is the OFFICIAL Test Form :p FFICIAL Test Form :pFFICIAL Test Form :pFFICIAL Test Form :p"
    Righello.FormattaTesto
End Sub

Private Sub Righello_rulerDxChanged(X As Long)
    Text1.SelRightIndent = Text1.Width - X
End Sub

Private Sub Righello_rulerSxbotChanged(X As Long)
    Text1.SelHangingIndent = X - Text1.SelIndent
End Sub

Private Sub Righello_rulerSxTopChanged(X As Long)
    Text1.SelIndent = X
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Righello.MouseMoved X, Y
    Ruler1.MouseMoved X, Y
End Sub
