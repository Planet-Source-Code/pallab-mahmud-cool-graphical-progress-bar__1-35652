Attribute VB_Name = "modFunction"
'--------------------------------------------------'
'|Progress bar [cool!!]                           |'
'|------------------------------------------------|'
'|Written by Pallab Mahmud                        |'
'|© Copyright 2001 by Pallab Mahmud               |'
'|email: pallmahmud@yahoo.com                     |'
'|                                                |'
'|This sample code is a FREEWARE. Use it in your  |'
'|own project as it fits You but do not re-sale   |'
'|this code or destroy the original authors name. |'
'|                                                |'
'|Warning: No warranty is provided with this set  |'
'|of code so use it in your own risk. The author  |'
'|is not responsible for the Damage caused by     |'
'|this code.                                      |'
'--------------------------------------------------'
'--------------------------------------------------'
'Comments:This is a cool progress bar.You can change
'it base and bar picture whatever you want.It uses
'only one api call.I think it is great.What do you think?
'Hey,listen i am new in programing and i am 14 years old
'So,don't mind and Please please........vote for me
'--------------------------------------------------'
Option Explicit
Public intPercent As String
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Sub curPercent(ByVal c_Perc As Integer, picMain As PictureBox, picBar As PictureBox)
    If c_Perc > 100 Or c_Perc < 0 Then Exit Sub
    BitBlt picMain.hDC, 0, 0, (c_Perc / 100) * (picMain.Width / Screen.TwipsPerPixelX), (picMain.Height / Screen.TwipsPerPixelY), picBar.hDC, 0, 0, &HCC0020
    picMain.Refresh
    picBar.Refresh
    intPercent = c_Perc & "%"
End Sub
