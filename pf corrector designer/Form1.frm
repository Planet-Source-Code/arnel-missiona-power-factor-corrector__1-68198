VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Power Factor Corrector"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll3 
      Height          =   3975
      Left            =   9120
      Max             =   50
      Min             =   1
      TabIndex        =   26
      Top             =   4320
      Value           =   1
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   17
      Left            =   6600
      Top             =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   8835
      TabIndex        =   20
      Top             =   4320
      Width           =   8895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   17
      Left            =   7800
      Top             =   240
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1335
      Left            =   7440
      Max             =   100
      Min             =   1
      TabIndex        =   13
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   6720
      Max             =   3000
      Min             =   1
      TabIndex        =   12
      Top             =   2400
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox txtpf 
      Height          =   285
      Left            =   6720
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Click to connect capacitor"
      Height          =   435
      Left            =   3840
      TabIndex        =   8
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMULATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtload 
      Height          =   285
      Left            =   6720
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtcap 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtammeter 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   690
      Width           =   855
   End
   Begin VB.TextBox txtsource 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   8280
      TabIndex        =   27
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "PF corrector"
      Height          =   255
      Left            =   2880
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Legend:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Power Wave From"
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   1080
      X2              =   2880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label11 
      Caption         =   "Current Wave Form"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   1080
      X2              =   2880
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label10 
      Caption         =   "Voltage Wave Form"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   1080
      X2              =   2880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   19
      Top             =   1680
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   720
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "microF"
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Line Line10 
      X1              =   3960
      X2              =   3960
      Y1              =   2040
      Y2              =   3360
   End
   Begin VB.Line Line9 
      BorderWidth     =   3
      X1              =   3480
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line8 
      BorderWidth     =   3
      X1              =   3480
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line7 
      X1              =   3960
      X2              =   3960
      Y1              =   840
      Y2              =   1920
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7800
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "PF"
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Load "
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1215
      Left            =   5760
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "power factor"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Load, watts"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ammeter"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label source 
      Caption         =   "SOURCE"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   6120
      X2              =   6120
      Y1              =   3360
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   6120
      X2              =   6120
      Y1              =   840
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   6120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1080
      Y1              =   2280
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   6120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   1080
      Y1              =   2400
      Y2              =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Arnel Missiona
'www.geocities.com/arnelmissiona

Private OLDpf As Double
Private corCAP As Double
Private x As Double
Private prevx1 As Double
Private prevy1  As Double
Private prevcy1 As Double
Private prevpy1 As Double
Private mult As Double
Private center As Integer
Private xmult As Integer
Dim angle As Double
Private angleLAG As Double


Private Sub Check1_Click()
On Error GoTo err
Dim Capvar As Double

If Check1.Value = 1 Then
OLDpf = txtpf
xc = Round(1 / (2 * 3.1416 * 60 * txtcap) * 1000000, 2)
Capvar = (txtsource ^ 2) / xc

va2 = (CDbl(txtload) / CDbl(txtpf)) ^ 2
p2 = CDbl(txtload) ^ 2
var = va2 - p2
indvar = Sqr(var)

newvar = indvar - Capvar

newVA = Sqr(CDbl(txtload) ^ 2 + newvar ^ 2)
newpf = txtload / newVA
txtpf = Round(newpf, 4)
Else
txtpf = OLDpf
End If
err:
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer2.Enabled = False
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
source = "SOURCE: 220V"
txtsource = 220
txtload = 1
txtpf = 1
mult = 1 * 0.2
center = pic.Height / 2
xmult = 20
End Sub

Private Function ampere(pf As Double, power As Double, volts As Double) As Double
ampere = Round(power / (pf * volts), 2)
End Function

Private Sub Timer1_Timer()
Dim pf As Double
Dim volts As Double
Dim power As Double
Dim var As Double


pf = txtpf
volts = txtsource
power = txtload

txtammeter = ampere(pf, power, volts)

var = Sqr((txtload / txtpf) ^ 2 - txtload ^ 2)
Label8 = "VAR: " & Round(var, 2)
If Check1.Value = 0 Then
txtcap = Round(capacitor(txtsource, var) * 1000000, 0)
End If
End Sub

Private Sub Timer2_Timer()
On Error GoTo err
Dim radANGLE As Double
Dim radlagANGLE As Double

pic.DrawWidth = 1
pic.Scale (0, pic.Height)-(pic.Width, 0)

xpf = txtpf
If xpf <> 1 Then
angleLAG = (Atn(-xpf / Sqr(-xpf * xpf + 1)) + 2 * Atn(1)) * (180 / 3.1416)
Else
angleLAG = 0
End If


x = (x + 1) + xmult 'scale
angle = angle + 10 ' convert this to radians first

radANGLE = angle * (3.1416 / 180)  ' in radians

yval = Sin(radANGLE) * txtsource * mult + center
cvalA = angle - angleLAG

radlagANGLE = cvalA * (3.1416 / 180) ' lagging angle in radians

lag = "Angle:" & angle & " anglelag:" & angleLAG & " LAG o:" & cvalA

cval = Sin(radlagANGLE) * CDbl(txtammeter) * (mult * 4) + center
power = ((Sin(radANGLE) * txtsource * Sin(radlagANGLE) * CDbl(txtammeter)) * (mult / 2)) + center

If (x - xmult) = 1 Then
pic.Line (0, yval)-(0, yval), vbGreen
pic.Line (0, cval)-(0, cval), vbGreen
pic.Line (0, power)-(0, power), vbGreen
prevx1 = x
prevy1 = yval
prevcy1 = cval
prevpy1 = power
Else

pic.Line (0, center)-(pic.Width, center), vbRed
pic.Line (prevx1, prevy1)-(x, yval), vbGreen
pic.Line (prevx1, prevcy1)-(x, cval), vbWhite
pic.Line (prevx1, prevpy1)-(x, power), vbYellow


prevx1 = x
prevy1 = yval
prevcy1 = cval
prevpy1 = power
End If

If x > pic.Width Then x = 0: pic.Cls

err:
'MsgBox err.Description
End Sub

Private Sub VScroll1_Scroll()
txtload = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
txtpf = VScroll2.Value / 100
End Sub

Private Sub VScroll2_Scroll()
txtpf = VScroll2.Value / 100
End Sub

Private Function capacitor(volts As Double, var As Double)
On Error GoTo err
Dim x As Double

x = volts ^ 2 / var
capacitor = 1 / (2 * 3.1416 * 60 * x)
err:
End Function

Private Sub VScroll3_Change()
mult = VScroll3.Value * 0.2
Label15 = "ZOOM : " & VScroll3.Value
End Sub

Private Sub VScroll3_Scroll()
mult = VScroll3.Value * 0.2
Label15 = "ZOOM : " & VScroll3.Value
End Sub
