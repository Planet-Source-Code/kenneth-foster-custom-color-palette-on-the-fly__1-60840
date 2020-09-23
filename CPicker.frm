VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Create Your Own Color Palette...DEMO"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   217
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6465
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   990
      Width           =   375
   End
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6000
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   24
      Top             =   990
      Width           =   375
   End
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5520
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   23
      Top             =   990
      Width           =   375
   End
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   5040
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   22
      Top             =   990
      Width           =   375
   End
   Begin VB.CommandButton cmdLoadDef 
      Caption         =   "Load Default Palette"
      Height          =   420
      Left            =   195
      TabIndex        =   20
      Top             =   2775
      Width           =   2040
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Color Palette"
      Height          =   435
      Left            =   2880
      TabIndex        =   18
      Top             =   2250
      Width           =   2055
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Color Palette"
      Height          =   435
      Left            =   210
      TabIndex        =   17
      Top             =   2250
      Width           =   2025
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   195
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   442
      TabIndex        =   11
      Top             =   120
      Width           =   6660
   End
   Begin VB.Frame fraControls 
      Caption         =   "Color Control Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   4785
      Begin VB.Label lblCol 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   10
         Left            =   4155
         TabIndex        =   16
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   9
         Left            =   3735
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   8
         Left            =   3330
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   7
         Left            =   2925
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   6
         Left            =   2535
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   2145
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCap 
         BackStyle       =   0  'Transparent
         Caption         =   "Left-click to change color   Right-click to de/activate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   420
         Width           =   4515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   1740
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   1335
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   945
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   555
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Left click to select  box then go Select color          Right click box for values"
      Height          =   705
      Left            =   5070
      TabIndex        =   26
      Top             =   1410
      Width           =   1845
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4950
      TabIndex        =   21
      Top             =   630
      Width           =   1905
   End
   Begin VB.Label Label5 
      Caption         =   "Note: File is saved in app folder. Filename is ColorPal."
      Height          =   570
      Left            =   2865
      TabIndex        =   19
      Top             =   2730
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Select a color by clicking or dragging"
      Height          =   195
      Left            =   1875
      TabIndex        =   14
      Top             =   615
      Width           =   2835
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5100
      TabIndex        =   13
      Top             =   2625
      Width           =   1755
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5100
      TabIndex        =   12
      Top             =   2280
      Width           =   1755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Original Source code by
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77
' modified by Ken Foster

Option Explicit

Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim color As Long
Dim lcolor As Long

Dim ValMain(11) As Long   'backcolors of gradient
Dim ValBColor(3) As Long  'backcolors for custom palette


Private Sub Form_Load()
   Dim X As Integer
   Dim Y As Integer
   
   Fillit  'allows custom colors from modColorOnly module
   color = RGB(255, 0, 0) 'to prevent false color reading on startup
   'Preload array with colors
   For X = 0 To 10
      ValMain(X) = lblCol(X).BackColor  'load colors into array to get the show on the road
   Next X
   
   For Y = 0 To UBound(ValBColor)
      ValBColor(Y) = vbWhite   'load arrray with white,otherwise an ugly black shows up
   Next Y
End Sub

Private Sub MultiGrad()
   
   Dim bRet    As Boolean
   
   bRet = Gradient(picSel.hDC, 0, 0, picSel.ScaleWidth, picSel.ScaleHeight, _
   0, 0, 0, _
   IIf(lblCol(0).BorderStyle, lblCol(0).BackColor, -1), _
   IIf(lblCol(1).BorderStyle, lblCol(1).BackColor, -1), _
   IIf(lblCol(2).BorderStyle, lblCol(2).BackColor, -1), _
   IIf(lblCol(3).BorderStyle, lblCol(3).BackColor, -1), _
   IIf(lblCol(4).BorderStyle, lblCol(4).BackColor, -1), _
   IIf(lblCol(5).BorderStyle, lblCol(5).BackColor, -1), _
   IIf(lblCol(6).BorderStyle, lblCol(6).BackColor, -1), _
   IIf(lblCol(7).BorderStyle, lblCol(7).BackColor, -1), _
   IIf(lblCol(8).BorderStyle, lblCol(8).BackColor, -1), _
   IIf(lblCol(9).BorderStyle, lblCol(9).BackColor, -1), _
   IIf(lblCol(10).BorderStyle, lblCol(10).BackColor, -1))
   
   If Not bRet Then MsgBox "Gradient failed. Must have at least two colors.", vbExclamation, "Error"
End Sub

Private Sub picPal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Sell As Long
  Dim dg As Integer
  
   Sell = picPal(Index).BackColor  'remembers backcolor in case user changes mind
   'allow only one selection at a time
   For dg = 0 To UBound(ValBColor)
      If dg <> Index Then
        picPal(dg).Appearance = 0
        picPal(dg).BackColor = ValBColor(dg)
      End If
   Next dg
  'flip flops picturebox appearance
   If Button = 1 Then
      If picPal(Index).Appearance = 1 Then
         picPal(Index).Appearance = 0
      Else
         picPal(Index).Appearance = 1
      End If
   End If
   
    picPal(Index).BackColor = Sell
    lcolor = Index
    
   'get values of selected color in picturebox
   If Button = 2 Then
      color = picPal(Index).BackColor
      ColorValues
   End If
End Sub

Private Sub picSel_Click()
   If picPal(lcolor).Appearance <> 1 Then Exit Sub
   picPal(lcolor).Appearance = 0  'reset appearance to default
   picPal(lcolor).BackColor = color  'set new color
   ValBColor(lcolor) = color
End Sub

Private Sub picSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo ErrExit
   If Button = 1 Then
      color = picSel.Point(X, Y)
      If color = -1 Then Exit Sub  'Prevents false RGB values
      ColorValues  'color selected so go get values for the labels.
   End If
ErrExit:
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo ErrExit
   If Button = 1 Then
      color = picSel.Point(X, Y)
      If color = -1 Then Exit Sub  'Prevents false RGB values
      ColorValues  'color selected so go get values for the labels.
   End If
ErrExit:
End Sub

Private Sub picSel_Paint()
   MultiGrad  'draw gradient in picturebox
   ColorValues  'go get values for the labels.
End Sub

Private Sub lblCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Sure As Long
   
   On Error GoTo ErrExit
   
   If Button = vbLeftButton Then
      Sure = ShowColor
      If Sure = -1 Then Exit Sub  'Cancel was clicked in ShowColor
      lblCol(Index).BackColor = Sure
      ValMain(Index) = Sure
   Else
      ' Toggle border style.
      lblCol(Index).BorderStyle = lblCol(Index).BorderStyle Xor 1
   End If
   MultiGrad
   Exit Sub
   
ErrExit:
End Sub

Private Sub ColorValues()
   Label1.BackColor = color
   'convert color to rgb
   R = color And 255
   G = (color \ 256) And 255
   B = (color \ 65536) And 255
   Label2.Caption = "RGB:  " & R & "," & G & "," & B
   'convert color to hex
   Label3.Caption = "HEX:  " & Hex(color)
End Sub

Private Sub cmdload_Click()
   Dim FileName As String
   Dim Y As Integer
   Dim X As Long
   
   FileName = App.Path & "\" & "ColorPal"
   
   If FileName <> "" Then
      Dim Free As Long
      
      Free = FreeFile
      
      Open FileName For Binary As #Free
      Get #Free, , ValMain
      Close #Free
   End If
   
   For Y = 0 To 10
      X = ValMain(Y)
      lblCol(Y).BackColor = X
   Next Y
   
   MultiGrad
End Sub

Private Sub cmdLoadDef_Click()
   Dim FileName As String
   Dim Y As Integer
   Dim X As Long
   
   FileName = App.Path & "\" & "DefaultPalette"
   
   If FileName <> "" Then
      Dim Free As Long
      
      Free = FreeFile
      
      Open FileName For Binary As #Free
      Get #Free, , ValMain
      Close #Free
   End If
   
   For Y = 0 To 10
      X = ValMain(Y)
      lblCol(Y).BackColor = X
   Next Y
   
   MultiGrad
End Sub

Private Sub cmdSave_Click()
   Dim FileName As String
   
   FileName = App.Path & "\" & "ColorPal"
   
   If FileName <> "" Then
      Dim Free As Long
      
      Free = FreeFile
      
      Open FileName For Binary As #Free
      Put #Free, , ValMain
      Close #Free
   End If
   
End Sub
