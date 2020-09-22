VERSION 5.00
Begin VB.UserControl Ruler 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "Ruler.ctx":0000
   ScaleHeight     =   315
   ScaleMode       =   0  'User
   ScaleWidth      =   4815
   ToolboxBitmap   =   "Ruler.ctx":003B
   Begin VB.PictureBox picRulerSxBot 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1560
      Picture         =   "Ruler.ctx":034D
      ScaleHeight     =   225
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   120
      Width           =   165
   End
   Begin VB.PictureBox picRulerDx 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   2310
      Picture         =   "Ruler.ctx":048B
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   3
      Top             =   120
      Width           =   165
   End
   Begin VB.PictureBox picRulerSxTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   870
      Picture         =   "Ruler.ctx":0575
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   1
      Top             =   30
      Width           =   165
   End
   Begin VB.PictureBox picRighello 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   2010
         X2              =   2010
         Y1              =   60
         Y2              =   270
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Scala"
      Visible         =   0   'False
      Begin VB.Menu mnuMode 
         Caption         =   "Centimetri"
         Index           =   0
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Inch"
         Index           =   1
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Pixel"
         Index           =   2
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Twip"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Ruler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

''Basato su codice di MystikalSoft, riadattato da
''Daniele Binaghi, www.pecorElettriche.it:
''- aggiunti cursori, con relative proprietà ed eventi
''- modificato indicatore di posizione, sostituendo con oggetto linea
''- aggiunte pagine proprietà generale e cursori
''- aggiunta gestione del minimo per la scala

Public Enum rlrBorderStyle
    rlrNoBorder = 0
    rlrSunken = 1
    rlrSunkenOuter = 2
    rlrRaised = 3
    rlrRaisedInner = 4
    rlrBump = 5
    rlrEtched = 6
End Enum
Public Enum rlrOrientationConstants
    rlrHorizontal = 0
    rlrVertical = 1
End Enum

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  tabella colori in RGB
Private Const pixR As Integer = 3
Private Const pixG As Integer = 2
Private Const pixB As Integer = 1

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum RulerModeConst
    Millimetri = 0
    Inch = 1
    Pixel = 2
    Twips = 3
End Enum

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Valori predefiniti proprietà:
Const m_def_Minimo = 0
Const m_def_Locked = False
Const m_def_Position = False
Const m_def_RulerScaleMode = 0
Const m_def_BorderStyle = 2
Const m_Def_Orientation = 0
'Variabili proprietà:
Dim m_Minimo As Single
Dim m_Locked As Boolean
Dim m_Position As Boolean
Dim m_MarginColor As OLE_COLOR
Dim m_lMargineSx As Long
Dim m_lMargineDx As Long
Dim m_lRulerSxTop As Long
Dim m_lRulerSxBot As Long
Dim m_lRulerDx As Long
Dim m_BorderStyle As rlrBorderStyle
Dim m_bPositionVisible As Boolean
Dim m_Orientation As rlrOrientationConstants
Dim m_RulerScaleMode As Variant
Dim X1 As Single
Dim RScale As Long
'Dichiarazioni di eventi:
Event Click()
Event DblClick()
Event MargineSxChanged(X As Long)
Event MargineDxChanged(X As Long)
Event RulerSxBotChanged(X As Long) 'MappingInfo=picRulerSxBot,picRulerSxBot,-1,MouseMove
Event RulerSxBotMouseDown()
Event RulerSxBotMouseUp()
Event RulerSxTopChanged(X As Long) 'MappingInfo=picRulerSxTop,picRulerSxTop,-1,MouseMove
Event RulerSxTopMouseDown()
Event RulerSxTopMouseUp()
Event RulerDxChanged(X As Long) 'MappingInfo=picRulerDx,picRulerDx,-1,MouseMove
Event RulerDxMouseDown()
Event RulerDxMouseUp()

Public Property Get BorderStyle() As rlrBorderStyle
Attribute BorderStyle.VB_Description = "Restituisce o imposta lo stile del bordo di un oggetto"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(New_Val As rlrBorderStyle)
    m_BorderStyle = New_Val
    picRighello.Cls
    EdgeSubClass picRighello.hWnd, New_Val
    pDrawRuler
    PropertyChanged "BorderStyle"
End Property

Public Property Get Position() As Boolean
Attribute Position.VB_Description = "Visualizza o nasconde l'indicatore di posizione del cursore"
Attribute Position.VB_ProcData.VB_Invoke_Property = "Generale"
    Position = m_bPositionVisible
End Property

Public Property Let Position(Visible As Boolean)
    m_bPositionVisible = Visible
    Line1.Visible = Visible
    PropertyChanged "Position"
End Property

Public Property Get RulerSxTop() As Single
Attribute RulerSxTop.VB_ProcData.VB_Invoke_Property = "Cursori"
    RulerSxTop = m_lRulerSxTop / RScale
End Property

Public Property Let RulerSxTop(X As Single)
Dim lX As Long
    lX = CLng(X * RScale)
    If lX > m_lRulerDx Then lX = m_lRulerDx
    picRulerSxTop.Left = lX - m_Minimo * RScale - 8
    m_lRulerSxTop = lX
    PropertyChanged "RulerSxTop"
End Property

Public Property Get MargineSx() As Single
Attribute MargineSx.VB_ProcData.VB_Invoke_Property = "Cursori"
    MargineSx = m_lMargineSx / RScale
End Property
    
Public Property Let MargineSx(X As Single)
    On Error Resume Next
    m_lMargineSx = CLng(X * RScale)
    If Err Then m_lMargineSx = UserControl.Width
    On Error GoTo 0
    Call pDrawRuler(True)
    PropertyChanged "MargineSx"
End Property

Public Property Get MargineDx() As Single
Attribute MargineDx.VB_ProcData.VB_Invoke_Property = "Cursori"
    MargineDx = m_lMargineDx / RScale
End Property

Public Property Let MargineDx(X As Single)
    On Error Resume Next
    m_lMargineDx = CLng(X * RScale)
    If Err Then m_lMargineDx = UserControl.Width
    On Error GoTo 0
    Call pDrawRuler(True)
    PropertyChanged "MargineDx"
End Property

Public Property Get RulerSxBot() As Single
Attribute RulerSxBot.VB_ProcData.VB_Invoke_Property = "Cursori"
    RulerSxBot = m_lRulerSxBot / RScale
End Property

Public Property Let RulerSxBot(X As Single)
Dim lX As Long
    lX = CLng(X * RScale)
    If lX > m_lRulerDx Then lX = m_lRulerDx
    picRulerSxBot.Left = lX - m_Minimo * RScale - 8
    m_lRulerSxBot = lX
    PropertyChanged "RulerSxBot"
End Property

Public Property Get RulerDx() As Single
Attribute RulerDx.VB_ProcData.VB_Invoke_Property = "Cursori"
    RulerDx = m_lRulerDx / RScale
End Property

Public Property Let RulerDx(X As Single)
Dim lX As Long
    lX = CLng(X * RScale)
    If lX < m_lRulerSxTop Or lX < m_lRulerSxBot Then
        lX = m_lRulerSxTop
        If lX < m_lRulerSxBot Then lX = m_lRulerSxBot
    End If
    picRulerDx.Left = lX - m_Minimo * RScale - 8
    m_lRulerDx = lX
    PropertyChanged "RulerDx"
End Property

Public Sub FormattaTesto()
    RaiseEvent RulerSxTopChanged(m_lRulerSxTop - m_Minimo * RScale)
    RaiseEvent RulerSxBotChanged(m_lRulerSxBot - m_Minimo * RScale)
    RaiseEvent RulerDxChanged(m_lRulerDx - m_Minimo * RScale)
End Sub

Private Sub pDrawRuler(Optional Clear As Boolean)
Dim Sincr As Single 'Scalemode is in TWIPS 1440 per inch
Dim I As Integer 'Number of segment across form
Dim iMinimo As Integer 'Parte intera del minimo
Dim iDecimale As Integer 'Parte decimale del minimo
    Sincr = RScale / 10
    iMinimo = Fix(m_Minimo)
    iDecimale = Int(Abs(m_Minimo - iMinimo) * 10) + 1
    With picRighello
        If Clear Then .Cls
        If m_Orientation = rlrVertical Then
            picRighello.Line (0, 0)-(picRighello.ScaleWidth, m_lMargineSx - m_Minimo * RScale), m_MarginColor, BF
            picRighello.Line (0, m_lMargineDx - m_Minimo * RScale)-(picRighello.ScaleWidth, picRighello.ScaleHeight), m_MarginColor, BF
            Do While Sincr < .ScaleHeight
                'Number of sections
                For I = iDecimale To 10
                    'Size of Tics
                    If I = 10 Then
                        picRighello.Line (0, Sincr)-(.ScaleHeight, Sincr)
                        .CurrentX = 0
                        picRighello.Print CStr(Int(Sincr / RScale) + iMinimo)
                        iDecimale = 1
                    ElseIf I = Int(10 * 0.5) Then '50%
                        picRighello.Line (.ScaleWidth - (.ScaleWidth * 0.5), Sincr)-(.ScaleWidth, Sincr)
                    Else
                        picRighello.Line (.ScaleWidth - (.ScaleWidth * 0.125), Sincr)-(.ScaleWidth, Sincr)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        Else
            picRighello.Line (0, 0)-(m_lMargineSx - m_Minimo * RScale, picRighello.ScaleHeight), m_MarginColor, BF
            picRighello.Line (m_lMargineDx - m_Minimo * RScale, 0)-(picRighello.ScaleWidth, picRighello.ScaleHeight), m_MarginColor, BF
            Do While Sincr < .ScaleWidth
                'Number of sections
                For I = iDecimale To 10
                    'Size of Tics
                    If I = 10 Then
                        picRighello.Line (Sincr, 0)-(Sincr, .ScaleHeight)
                        .CurrentY = 0
                        picRighello.Print " " + CStr(Int(Sincr / RScale) + iMinimo)
                        iDecimale = 1
                    ElseIf I = Int(10 * 0.5) Then '50%
                        picRighello.Line (Sincr, .ScaleHeight - (.ScaleHeight * 0.5))-(Sincr, .ScaleHeight)
                    Else
                        picRighello.Line (Sincr, .ScaleHeight - (.ScaleHeight * 0.125))-(Sincr, .ScaleHeight)
                    End If
                    Sincr = Sincr + (RScale / 10)
                Next
            Loop
        End If
    End With
End Sub

Public Sub MouseMoved(X As Single, Y As Single)
    picRighello.AutoRedraw = False
    If m_bPositionVisible Then
        If m_Orientation = rlrHorizontal Then
            Line1.X1 = X
            Line1.X2 = X
        Else
            Line1.Y1 = Y
            Line1.Y2 = Y
        End If
    End If
    Line1.Visible = m_bPositionVisible
    picRighello.AutoRedraw = True
End Sub

Private Sub picRulerDx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerDxMouseDown
End Sub

Private Sub picRulerDx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PosCur As POINTAPI
Dim PosPic As POINTAPI
Dim lX As Long
    If Button > 0 Then
        GetCursorPos PosCur
        ClientToScreen picRighello.hWnd, PosPic
        lX = (PosCur.X - PosPic.X) * Screen.TwipsPerPixelX
        If picRulerDx.Left = lX Then Exit Sub
        If lX < m_lRulerSxTop - m_Minimo * RScale Then lX = m_lRulerSxTop - m_Minimo * RScale
        If lX < m_lRulerSxBot - m_Minimo * RScale Then lX = m_lRulerSxBot - m_Minimo * RScale
        If lX > UserControl.Width - 196 Then lX = UserControl.Width - 196
        picRulerDx.Left = lX - 8
        m_lRulerDx = lX + m_Minimo * RScale
        picRighello.Refresh
        RaiseEvent RulerDxChanged(lX)
    End If
End Sub

Private Sub picRulerDx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerDxMouseUp
End Sub

Private Sub picRulerSxTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerSxTopMouseDown
End Sub

Private Sub picRulerSxTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PosCur As POINTAPI
Dim PosPic As POINTAPI
Dim lX As Long
    If Button > 0 Then
        GetCursorPos PosCur
        ClientToScreen picRighello.hWnd, PosPic
        lX = (PosCur.X - PosPic.X) * Screen.TwipsPerPixelX
        If picRulerSxTop.Left = lX Then Exit Sub
        If lX < 0 Then lX = 0
        If lX > m_lRulerDx - m_Minimo * RScale Then lX = m_lRulerDx - m_Minimo * RScale
        picRulerSxTop.Left = lX - 8
        m_lRulerSxTop = lX + m_Minimo * RScale
        picRighello.Refresh
        RaiseEvent RulerSxTopChanged(lX)
        RaiseEvent RulerSxBotChanged(picRulerSxBot.Left)
    End If
End Sub

Private Sub picRulerSxTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerSxTopMouseUp
End Sub

Private Sub mnuMode_Click(Index As Integer)
    RulerScaleMode = Index
End Sub

Private Sub picRulerSxBot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerSxBotMouseDown
End Sub

Private Sub picRulerSxBot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PosCur As POINTAPI
Dim PosPic As POINTAPI
Dim lX As Long
Dim lSopraX As Long
Dim lMin As Long
    If Button > 0 Then
        GetCursorPos PosCur
        ClientToScreen picRighello.hWnd, PosPic
        lX = (PosCur.X - PosPic.X) * Screen.TwipsPerPixelX
        If picRulerSxBot.Left = lX Then Exit Sub
        lMin = m_Minimo * RScale
        If lX < 0 Then lX = 0
        If lX > m_lRulerDx - lMin Then lX = m_lRulerDx - lMin
        If Y > 6 Then
            lSopraX = lX - m_lRulerSxBot + lMin
            lSopraX = m_lRulerSxTop + lSopraX
            If lSopraX < 0 + lMin Then lSopraX = 0 + lMin
            If lSopraX > m_lRulerDx - lMin Then
                lSopraX = m_lRulerDx - lMin
                lX = lSopraX - (m_lRulerSxTop - m_lRulerSxBot)
            End If
            picRulerSxTop.Left = lSopraX - lMin - 8
            m_lRulerSxTop = lSopraX
        End If
        picRulerSxBot.Left = lX - 8
        m_lRulerSxBot = lX + lMin
        picRighello.Refresh
        UserControl.Refresh
        RaiseEvent RulerSxTopChanged(lSopraX - lMin)
        RaiseEvent RulerSxBotChanged(lX)
    End If
End Sub

Private Sub picRulerSxBot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent RulerSxBotMouseUp
End Sub

Private Sub picRighello_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim I As Integer
        For I = 0 To mnuMode.Count - 1
            mnuMode(I).Checked = False
        Next I
        mnuMode(RulerScaleMode).Checked = True
        UserControl.PopupMenu mnuMenu
    End If
End Sub

Private Sub UserControl_Initialize()
    RScale = 570
    picRulerSxTop.ScaleMode = vbPixels
    pTrimPicture picRulerSxTop, vbRed
    picRulerSxBot.ScaleMode = vbPixels
    pTrimPicture picRulerSxBot, vbRed
    picRulerDx.ScaleMode = vbPixels
    pTrimPicture picRulerDx, vbRed
End Sub

Private Sub UserControl_InitProperties()
    m_RulerScaleMode = m_def_RulerScaleMode
    m_BorderStyle = m_def_BorderStyle
    m_Position = m_def_Position
    m_Locked = m_def_Locked
    m_Orientation = m_Def_Orientation
    m_Minimo = m_def_Minimo
End Sub

Private Sub UserControl_Resize()
    picRighello.Move 64, 0, UserControl.Width - 128, UserControl.ScaleHeight - 92
    If m_Orientation = rlrHorizontal Then
        picRulerSxTop.Top = 0
        picRulerSxBot.Top = UserControl.ScaleHeight - 208
        picRulerDx.Top = picRulerSxBot.Top
        Line1.X2 = Line1.X1
        Line1.Y1 = 0
        Line1.Y2 = picRighello.ScaleHeight
    Else
        Line1.X1 = 0
        Line1.X2 = picRighello.ScaleWidth
        Line1.Y2 = Line1.Y1
    End If
    UserControl.Cls
    Call pDrawRuler(True)
End Sub

Private Sub UserControl_Show()
    Call pDrawRuler(True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    picRighello.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    picRighello.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_MarginColor = PropBag.ReadProperty("MarginColor", vbButtonFace)
    Orientation = PropBag.ReadProperty("Orientation", m_Def_Orientation)
    Minimo = PropBag.ReadProperty("Minimo", m_def_Minimo)
    RulerScaleMode = PropBag.ReadProperty("RulerScaleMode", m_def_RulerScaleMode)
    MargineSx = PropBag.ReadProperty("MargineSx", 0)
    MargineDx = PropBag.ReadProperty("MargineDx", UserControl.Width / RScale)
    RulerSxBot = PropBag.ReadProperty("RulerSxBot", 0)
    RulerSxTop = PropBag.ReadProperty("RulerSxTop", 0)
    RulerDx = PropBag.ReadProperty("RulerDx", UserControl.Width / RScale)
    Position = PropBag.ReadProperty("Position", m_def_Position)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Locked", Not m_def_Locked)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BackColor", picRighello.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", picRighello.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MarginColor", m_MarginColor, vbButtonFace)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_Def_Orientation)
    Call PropBag.WriteProperty("Minimo", m_Minimo, m_def_Minimo)
    Call PropBag.WriteProperty("RulerScaleMode", m_RulerScaleMode, m_def_RulerScaleMode)
    Call PropBag.WriteProperty("MargineSx", m_lMargineSx / RScale, 0)
    Call PropBag.WriteProperty("MargineDx", m_lMargineDx / RScale, UserControl.Width / RScale)
    Call PropBag.WriteProperty("RulerSxBot", m_lRulerSxBot / RScale, 0)
    Call PropBag.WriteProperty("RulerSxTop", m_lRulerSxTop / RScale, 0)
    Call PropBag.WriteProperty("RulerDx", m_lRulerDx / RScale, UserControl.Width / RScale)
    Call PropBag.WriteProperty("Position", m_Position, m_def_Position)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
End Sub

Private Sub pTrimPicture(ByVal pic As PictureBox, ByVal transparent_color As Long)
' Restrict the form to its "transparent" pixels.
Const RGN_OR = 2
Dim bitmap_info As BITMAPINFO
Dim pixels() As Byte
Dim bytes_per_scanLine As Integer
Dim pad_per_scanLine As Integer
Dim transparent_r As Byte
Dim transparent_g As Byte
Dim transparent_b As Byte
Dim wid As Integer
Dim hgt As Integer
Dim X As Integer
Dim Y As Integer
Dim start_x As Integer
Dim stop_x As Integer
Dim combined_rgn As Long
Dim new_rgn As Long
    ' Prepare the bitmap description.
    With bitmap_info.bmiHeader
        .biSize = 40
        .biWidth = pic.ScaleWidth
        ' Use negative height to scan top-down.
        .biHeight = -pic.ScaleHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
    End With

    ' Load the bitmap's data.
    wid = pic.ScaleWidth
    hgt = pic.ScaleHeight
    ReDim pixels(1 To 4, 0 To wid - 1, 0 To hgt - 1)
    GetDIBits pic.hDC, pic.Image, 0, pic.ScaleHeight, pixels(1, 0, 0), bitmap_info, DIB_RGB_COLORS

    ' Break the transparent color into its components.
    pUnRGB transparent_color, transparent_r, transparent_g, transparent_b

    ' Create the PictureBox's regions.
    For Y = 0 To hgt - 1
        ' Create a region for this row.
        X = 1
        Do While X < wid
            start_x = 0
            stop_x = 0

            ' Find the next non-transparent column.
            Do While X < wid
                If pixels(pixR, X, Y) <> transparent_r Or _
                   pixels(pixG, X, Y) <> transparent_g Or _
                   pixels(pixB, X, Y) <> transparent_b _
                Then
                    Exit Do
                End If
                X = X + 1
            Loop
            start_x = X

            ' Find the next transparent column.
            Do While X < wid
                If pixels(pixR, X, Y) = transparent_r And _
                   pixels(pixG, X, Y) = transparent_g And _
                   pixels(pixB, X, Y) = transparent_b _
                Then
                    Exit Do
                End If
                X = X + 1
            Loop
            stop_x = X

            ' Make a region from start_x to stop_x.
            If start_x < wid Then
                If stop_x >= wid Then stop_x = wid - 1

                ' Create the region.
                new_rgn = CreateRectRgn(start_x, Y, stop_x, Y + 1)

                ' Add it to what we have so far.
                If combined_rgn = 0 Then
                    combined_rgn = new_rgn
                Else
                    CombineRgn combined_rgn, combined_rgn, new_rgn, RGN_OR
                    DeleteObject new_rgn
                End If
            End If
        Loop
    Next Y
    
    ' Restrict the PictureBox to the region.
    SetWindowRgn pic.hWnd, combined_rgn, True
    DeleteObject combined_rgn
End Sub

Private Sub pUnRGB(ByRef Color As Long, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    r = Color And &HFF&
    g = (Color And &HFF00&) \ &H100&
    b = (Color And &HFF0000) \ &H10000
End Sub
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=picRighello,picRighello,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
    BackColor = picRighello.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picRighello.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call pDrawRuler(True)
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=14,0,0,0
Public Property Get RulerScaleMode() As RulerModeConst
    RulerScaleMode = m_RulerScaleMode
End Property

Public Property Let RulerScaleMode(ByVal New_RulerScaleMode As RulerModeConst)
    m_RulerScaleMode = New_RulerScaleMode
    Select Case m_RulerScaleMode
        Case 0
            RScale = 570
        Case 1
            RScale = 1440
        Case 2
            RScale = Screen.TwipsPerPixelX * 100
        Case 3
            RScale = 1000
    End Select
    Call pDrawRuler(True)
    PropertyChanged "RulerScaleMode"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=10,0,0,vbButtonFace
Public Property Get MarginColor() As OLE_COLOR
Attribute MarginColor.VB_Description = "Restituisce o imposta il colore delle aree del righello fuori dai margini"
    MarginColor = m_MarginColor
End Property

Public Property Let MarginColor(ByVal New_MarginColor As OLE_COLOR)
    m_MarginColor = New_MarginColor
    PropertyChanged "MarginColor"
    Call pDrawRuler(True)
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Restituisce o imposta il tipo di puntatore del mouse visualizzato quando il puntatore si trova su una parte specifica di un oggetto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Imposta un'icona personalizzata per il puntatore del mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=picRighello,picRighello,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Restituisce o imposta il colore di primo piano utilizzato per la visualizzazione di testo e grafica in un oggetto."
    ForeColor = picRighello.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picRighello.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Call pDrawRuler(True)
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,0
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    UserControl.Enabled = Not m_Locked
    PropertyChanged "Locked"
End Property

Public Property Get Orientation() As rlrOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(New_Val As rlrOrientationConstants)
    If m_Orientation <> New_Val Then
        m_Orientation = New_Val
        UserControl.Cls
        picRighello.Cls
        pChangeOrientSizes
        picRulerDx.Visible = (m_Orientation = rlrHorizontal)
        picRulerSxTop.Visible = (m_Orientation = rlrHorizontal)
        picRulerSxBot.Visible = (m_Orientation = rlrHorizontal)
        Call pDrawRuler
        PropertyChanged "Orientation"
    End If
End Property

Private Sub pChangeOrientSizes()
Dim rlrWidth As Long
Dim rlrHeight As Long
    rlrWidth = UserControl.Width
    rlrHeight = UserControl.Height
    UserControl.Height = rlrWidth
    UserControl.Width = rlrHeight
    UserControl.Cls
    Call pDrawRuler(True)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=12,0,0,0
Public Property Get Minimo() As Single
Attribute Minimo.VB_Description = "Restituisce o imposta il punto iniziale della scala del righello"
    Minimo = m_Minimo
End Property

Public Property Let Minimo(ByVal New_Minimo As Single)
    m_Minimo = New_Minimo
    Call pDrawRuler(True)
    Call pPosizionaCursori
    PropertyChanged "Minimo"
End Property

Private Sub pPosizionaCursori()
    picRulerSxTop.Left = m_lRulerSxTop - m_Minimo * RScale
    picRulerSxBot.Left = m_lRulerSxBot - m_Minimo * RScale
    picRulerDx.Left = m_lRulerDx - m_Minimo * RScale
End Sub
