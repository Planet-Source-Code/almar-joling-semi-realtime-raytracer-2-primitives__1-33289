Attribute VB_Name = "mdlVectorMath"
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


Public Function VectorSetXYZ(V As Vector3D, x As Single, y As Single, z As Single)
    V.x = x
    V.y = y
    V.z = z
End Function

Public Function VectorAdd(a As Vector3D, b As Vector3D) As Vector3D      '// result = a + b
    Dim result As Vector3D
    
    With result
        .x = a.x + b.x
        .y = a.y + b.y
        .z = a.z + b.z
    End With
    
    VectorAdd = result
End Function

Public Function VectorSub(a As Vector3D, b As Vector3D) As Vector3D    '// result = a - b
    Dim result As Vector3D
    
    With result
        .x = a.x - b.x
        .y = a.y - b.y
        .z = a.z - b.z
    End With
    
    VectorSub = result
End Function

Public Function VectorDot(a As Vector3D, b As Vector3D) As Single
    VectorDot = (a.x * b.x) + (a.y * b.y) + (a.z * b.z)
End Function

Public Function VectorCross(a As Vector3D, b As Vector3D) As Vector3D
    Dim C As Vector3D
    With a
        C.x = .y * b.z - .z * b.y
        C.y = .z * b.x - .x * b.z
        C.z = .x * b.y - .y * b.x
    End With
    
    VectorCross = C
End Function

Public Sub VectorNormalize(V As Vector3D)
    Dim sngScaleFactor As Single
    sngScaleFactor = 1 / Sqr((V.x * V.x) + (V.y * V.y) + (V.z * V.z))  '// this is definitely not the fastest way to do this. the processor by default computes 1/sqrt and then flips it.
                                                                      '          // i dont know of a way to get at that with math.h. when i optimize for altivec, i will take advantage of this.
    V.x = V.x * sngScaleFactor
    V.y = V.x * sngScaleFactor
    V.z = V.x * sngScaleFactor
End Sub

' //I havent yet bothered to impliment matrix based rotation. this is used so infrequently that it hardly matters though.
Public Sub Rotate(ByRef V As Vector3D, ByRef ax As Single, ByRef ay As Single, ByRef az As Single)
    Dim Temp As Vector3D
    Dim sngCosX As Single, sngCosY As Single, sngCosZ As Single
    Dim sngSinX As Single, sngSinY As Single, sngSinZ As Single
    
    '//The less Sin/Cos...the better. Are very slow functions
    '//A lookup table might be used, sacrificing precision
    '//Note: Taylor series do not make it much faster either..
    sngCosX = Cos(ax)
    sngSinX = Sin(ax)
    sngCosY = Cos(ay)
    sngSinY = Sin(ay)
    sngCosZ = Cos(az)
    sngSinZ = Sin(az)
    
    With V
        Temp.y = .y
        .y = (.y * sngCosX - .z * sngSinX)
        .z = (.z * sngCosX + Temp.y * sngSinX)
    
        Temp.z = .z
        .z = (.z * sngCosY - .x * sngSinY)
        .x = (.x * sngCosY + Temp.z * sngSinY)
    
        Temp.x = .x
        .x = (.x * sngCosZ - .y * sngSinZ)
        .y = (.y * sngCosZ + Temp.x * sngSinZ)
    End With
End Sub

