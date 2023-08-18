VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock - By Hevanafa (Aug 2023"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRefresh 
      Interval        =   100
      Left            =   1650
      Top             =   1650
   End
   Begin VB.PictureBox pbClock 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   0
      Top             =   0
      Width           =   2115
   End
   Begin VB.Label lblDayOfWeek 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   2250
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

DefSng A-Z
Const PI# = 3.14159265358979

Const ClockWidth = 140
Const ClockHeight = 140
Const Cornflowerblue = &HED9564

Dim LastDay%


Private Sub Form_Load()

tmrRefresh.Enabled = True
'ScaleWidth = ClockWidth
'ScaleHeight = ClockHeight

End Sub


Function GetHour%()

GetHour = Timer \ 3600

End Function

Function GetMinutes%()

GetMinutes = (Timer Mod 3600) \ 60

End Function

Function GetSeconds%()

GetSeconds = Timer Mod 60

End Function


Function GetSecondsFloat!()

' Modulo operator returns integer
GetSecondsFloat = (Timer * 100 Mod 6000) / 100

End Function

Function PadStart(text$, length%, s$) As String

Dim result$
result$ = text

Do While Len(result) < length

result = s & result

Loop

PadStart = result

End Function


Function StrTrim$(n As Variant)

If IsNumeric(n) Then
    StrTrim = LTrim(Str(n))
Else
    StrTrim = "" & n
End If

End Function


Function GetDigitalStr$()

GetDigitalStr = PadStart(StrTrim(GetHour), 2, "0") & ":" & PadStart(StrTrim(GetMinutes), 2, "0") & ":" & PadStart(StrTrim(GetSeconds), 2, "0")

End Function


Function Deg2Rad#(deg#)

Deg2Rad = deg / 180 * PI

End Function


Sub RedrawClock()

Dim half_w%, half_h%
half_w = ClockWidth \ 2
half_h = ClockHeight \ 2

Dim a%

Dim cx%, cy%
cx = ClockWidth \ 2
cy = ClockHeight \ 2

pbClock.Line (cx - half_w, cy - half_h)-(cx + half_w, cy + half_h), vbWhite, BF


pbClock.Circle (cx, cy), ClockWidth \ 2, ClockHeight \ 2
pbClock.Circle (cx, cy), ClockWidth \ 2 - 10, ClockHeight \ 2 - 10

Dim angle!
Dim x1, y1, x2, y2

For a = 1 To 12

angle = Deg2Rad(a * 30)

x1 = cx + Sin(angle) * 53
y1 = cy + Cos(angle) * 53
x2 = cx + Sin(angle) * 58
y2 = cy + Cos(angle) * 58

pbClock.Line (x1, y1)-(x2, y2), Cornflowerblue

Next


' Hour hand
angle = Deg2Rad((GetHour Mod 12) * 30 + GetMinutes / 2)  ' / 60 * 30

x2 = cx + Sin(angle) * 30
y2 = cy - Cos(angle) * 30

pbClock.Line (cx, cy)-(x2, y2), vbBlack


' Minute hand
angle = Deg2Rad(GetMinutes * 6)

x2 = cx + Sin(angle) * 40
y2 = cy - Cos(angle) * 40

pbClock.Line (cx, cy)-(x2, y2), vbBlack


' Second hand
angle = Deg2Rad(GetSecondsFloat * 6)

' pbClock.Print Str(angle)

x2 = cx + Sin(angle) * 50
y2 = cy - Cos(angle) * 50

pbClock.Line (cx, cy)-(x2, y2), vbRed

pbClock.Circle (cx, cy), 2, 2, vbBlack

' pbClock.Print Str(GetSecondsFloat)

End Sub


Function GetDayOfWeek%()

GetDayOfWeek = DatePart("w", Now)

End Function


Function GetDayOfWeekAbbr$()

GetDayOfWeekAbbr = Format(DatePart("w", Now), "ddd")

End Function



Sub RefreshDayDisplay()

If LastDay <> Day(Now) Then
    LastDay = Day(Now)
    lblDayOfWeek.Caption = Str(GetDayOfWeek)
    lblDayOfWeek.Visible = False
End If

pbClock.ForeColor = IIf(GetDayOfWeek = vbSunday, vbRed, vbBlack)
pbClock.CurrentX = (ClockWidth - lblDayOfWeek.Width) \ 2
pbClock.CurrentY = ClockHeight \ 2 + 20

pbClock.Print GetDayOfWeekAbbr

End Sub


Private Sub tmrRefresh_Timer()

RedrawClock

RefreshDayDisplay

' pbClock.Print GetDigitalStr ' Timer

End Sub
