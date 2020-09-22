Attribute VB_Name = "mdlScene"
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


'//To Get/Set pixel data
Public Type BITMAPINFOHEADER
    biSize           As Long
    biWidth          As Long
    biHeight         As Long
    biPlanes         As Integer
    biBitCount       As Integer
    biCompression    As Long
    biSizeImage      As Long
    biXPelsPerMeter  As Long
    biYPelsPerMeter  As Long
    biClrUsed        As Long
    biClrImportant   As Long
End Type

'//Convert Picture to Array and back
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'//Bitmapinfo type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
End Type
Public Const DIB_RGB_COLORS As Long = 0

'//Our bitmap bits array
Public Bits() As RGBQUAD

'//Bitmapinfo
Public BInfo As BITMAPINFO

'//Handle variables
Public lngHDC As Long
Public lngImageHandle As Long

'//Timing
Public lngStart As Long, lngEnd As Double, lngCurrTime As Single
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'//For the triangle types
Public Vertices(4) As Vector3D
Private lngCount As Long

Public Sub SetupScene()
    Dim I As Long
    ReDim LightSources(0)
    
    lngCount = -1
    For I = 0 To 3
        If frmMain.chkRender(I).Value = vbChecked Then lngCount = lngCount + 1
    Next I
    
    I = 0
    If lngCount = -1 Then Exit Sub
    ReDim Primitives(lngCount)
    
    '//Sphere enabled?
    If frmMain.chkRender(0).Value = vbChecked Then
        
        '-------SPHERE
        Primitives(I).Type = SPHERE_TYPE
        With Primitives(I).Sphere
            .vecCenter.x = 10
            .vecCenter.y = 50
            .vecCenter.z = 0
            .sngRadius = 75
        End With
        I = I + 1
    End If
    
    '-------TRIANGLES
    '//Enabled?
    If frmMain.chkRender(1).Value = vbChecked Then
        Call VectorSetXYZ(Vertices(0), 0, -100, 50)
        Call VectorSetXYZ(Vertices(1), 100, 100, 0)
        Call VectorSetXYZ(Vertices(2), 0, 100, 250)
        Call VectorSetXYZ(Vertices(3), -100, 100, 0)
        
        Primitives(I).Type = TRIANGLE_TYPE
        With Primitives(I).Triangle
            .v1 = Vertices(0)
            .v2 = Vertices(1)
            .v3 = Vertices(2)
        End With
        I = I + 1
    End If
    
    '--------PLANE
    '//Enabled?
    If frmMain.chkRender(2).Value = vbChecked Then
        Dim Normal As Vector3D
        Call VectorSetXYZ(Normal, 0.15, -1, 0)
        Primitives(I).Type = PLANE_TYPE
        Primitives(I).Plane.vecNormal = Normal
        Primitives(I).Plane.sngDisplacement = 75   '//Width of plane
        I = I + 1
    End If
    
    '--------CYLINDER
    '//Enabled?
    If frmMain.chkRender(3).Value = vbChecked Then
        Primitives(I).Type = CYLINDER_TYPE
        Primitives(I).Cilinder.Axis = Y_INFINITE
        Call VectorSetXYZ(Primitives(I).Cilinder.Center, -200, 0, 150)
        I = I + 1
    End If
    
    '//Light location
    With LightSources(0)
        .Color.R = 0.5
        .Color.G = 0.3
        .Color.b = 0.81
        .location.x = 100
        .location.y = 200
        .location.z = -400
    End With

'    With LightSources(1)
'        .Color.R = 0.81
'        .Color.G = 0.2
'        .Color.b = 0.4
'        .location.x = -100
'        .location.y = -200
'        .location.z = 400
'    End With

End Sub

Public Sub Main()
    Dim I As Long
    
    '//Quick method..easy to remove as well
    If LCase(Command$) = "uncompiled" Then
        If MsgBox("You are running the raytracer uncompiled. Do you really to continue? " & vbCrLf & "Compiling is recommended!", vbYesNo, "Uncompiled") = vbNo Then
            Unload frmMain
            End
        End If
    End If
    
    frmMain.Show
    DoEvents
    
    '//Allocate the ray direction lookup table
    DirectionTable = GenerateRayDirectionTable
    
    '//Setup stuff
    Call mdlScene.SetupScene

    '//Device Context
    lngHDC = frmMain.picRay.hdc

    '//Image handle
    lngImageHandle = frmMain.picRay.Image.Handle
    
    '//Set bitmap ino and create our pixel array
    With BInfo.bmiHeader
       .biSize = 40
       .biWidth = frmMain.picRay.ScaleWidth
       .biHeight = frmMain.picRay.ScaleHeight
       .biPlanes = 1
       .biBitCount = 32
       .biCompression = 0
       .biClrUsed = 0
       .biClrImportant = 0
       .biSizeImage = frmMain.picRay.ScaleWidth * frmMain.picRay.ScaleHeight
    End With
    
    '//Redim our array to the size of the picturebox
    With frmMain.picRay
        ReDim Bits(0 To BInfo.bmiHeader.biWidth - 1, 0 To BInfo.bmiHeader.biHeight)
    End With
    
    
    '//Main loop
    
    lngStart = timeGetTime
        
    Do
        
        lngEnd = timeGetTime
        lngCurrTime = (lngEnd - lngStart) \ 1000
        
        '//Only if there is atleast one item:
        If lngCount > -1 Then
            '//Update the scene
            Call UpdateScene
        
            Call TraceScene(Primitives, UBound(Primitives), LightSources, UBound(LightSources))
        End If
        
        '//FPS counter
        frmMain.Caption = "RayTrace 2 :: " & GetFPS & "fps"
        
        '//Picture output: Call SavePicture(frmMain.picRay.Image, App.Path & "\" & lngEnd & ".bmp")
        DoEvents
    Loop

    
End Sub

Public Sub UpdateScene()
    Dim I As Long
    
    For I = 0 To UBound(Primitives)
        '//Change color of primitives
        With Primitives(I).Surface
            .BaseColor.R = 128 * Sin(lngCurrTime * I) + 128
            .BaseColor.G = 128 * Sin(lngCurrTime * I + 3 + Sin(lngCurrTime)) + 128
            .BaseColor.b = 128 * Sin(lngCurrTime * I + 2 + 1.5 * Sin(lngCurrTime)) + 128
            .sngReflectivity = 0.5 + 0.5 * Sin(lngCurrTime + I)
        End With
        
        With Primitives(I)
            Select Case .Type
                Case SPHERE_TYPE
                
                    Call VectorSetXYZ(.Sphere.vecCenter, 120 * Sin(lngCurrTime + I * 2 + 0.56), 120 * Sin(lngCurrTime + I * 2 + 3 * Sin(lngCurrTime * 0.2)), 120 * Sin(lngCurrTime + I * 5 + 1 + Sin(lngCurrTime * 0.1)))
                    .Sphere.sngRadius = 30 * Sin(lngCurrTime * 0.4 + Sin(lngCurrTime + I)) + 40
                
                Case TRIANGLE_TYPE
                
                    '//Update the vertices for the triangles
                    Dim lngCalc1 As Single, lngCalc2 As Single
                    lngCalc1 = lngCurrTime * 1.4
                    lngCalc2 = lngCurrTime * 1.2352
                    Call Rotate(.Triangle.v1, lngCalc1, lngCurrTime, lngCalc2)
                    Call Rotate(.Triangle.v2, lngCalc1, lngCurrTime, lngCalc2)
                    Call Rotate(.Triangle.v3, lngCalc1, lngCurrTime, lngCalc2)
                
                Case PLANE_TYPE
                    Call Rotate(.Plane.vecNormal, 0, lngCurrTime, 0)

                Case CYLINDER_TYPE
                    .Cilinder.sngRadius = 60 + 40 * Sin(lngCurrTime)
            End Select
        End With

    Next I
    
    '//Update lightsources
    For I = 0 To UBound(LightSources)
        With LightSources(I)
            .location.x = 250 * Sin(lngCurrTime * 0.5 + I)
            .location.y = -100
            .location.z = 250 * Cos(lngCurrTime * 0.5 + I)
            .Color.R = 0.15 * Sin(lngCurrTime + I) + 0.4
            .Color.G = 0.15 * Sin(lngCurrTime + 3 + Sin(lngCurrTime + I)) + 0.4
            .Color.b = 0.15 * Sin(lngCurrTime * I + 2 + 1.5 * Sin(lngCurrTime)) + 0.4
        End With
    Next I
End Sub
