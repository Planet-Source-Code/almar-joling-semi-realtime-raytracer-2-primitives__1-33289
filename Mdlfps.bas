Attribute VB_Name = "mdlFPS"
Option Explicit

'//Realtime raytracer [version 2]
'//Original (c++) version and other nice
'//Raytrace versions (with shadows, cilinders, etc)
'//Can be found at http://www.2tothex.com/
'//VB port by Almar Joling / quadrantwars@quadrantwars.com
'//Websites: http://www.quadrantwars.com (my game)
'//          http://vbfibre.digitalrice.com (Many VB speed tricks with benchmarks)

'//This code is highly optimized. If you manage to gain some more FPS
'//I'm always interested =-)

'//Finished @ 01/04/2002
'//Feel free to post this code anywhere, but please leave the above info
'//and author info intact. Thank you.


Private Declare Function GetTickCount Lib "kernel32" () As Long

Private lngTimer As Long
Private intFPSCounter As Integer
Private intFPS As Integer

Public Function GetFPS() As Long
    '//Count FPS
    If lngTimer + 1000 <= GetTickCount Then
        lngTimer = GetTickCount
        intFPS = intFPSCounter + 1
        intFPSCounter = 0
    Else
        intFPSCounter = intFPSCounter + 1
    End If
    
    '//Return FPS
    GetFPS = intFPS
End Function
