Attribute VB_Name = "mdlMath"
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

Public DirectionTable() As Vector3D
Public Primitives() As Primitive
Public LightSources() As LightSource

Public Function GenerateRayDirectionTable() As Vector3D()
    Dim Direction(640& * 480&) As Vector3D
    Dim currDirection As Vector3D
    Dim x As Long, y As Long
    Dim lngPosition As Long
    
    '//Inline should be faster...
    Dim sngScaleFactor As Single
    
    '//Create lookup table...Only used once
    For y = 0 To 480 - 1
        For x = 0 To 640 - 1
            lngPosition = x + (y * 640)
            currDirection = Direction(lngPosition)
            currDirection.x = x - 320
            currDirection.y = y - 240
            
            '//This value is fairly arbitrary and can basically be interpreted as field of view
            currDirection.z = 255
            Direction(lngPosition) = currDirection
            
            '// This is definitely not the fastest way to do this. the processor by default computes 1/sqrt and then flips it.
            With Direction(lngPosition)
                sngScaleFactor = 1 / Sqr((.x * .x) + (.y * .y) + (.z * .z))
                .x = .x * sngScaleFactor
                .y = .y * sngScaleFactor
                .z = .z * sngScaleFactor
            End With
        Next x
    Next y
    
    '//Return array
    GenerateRayDirectionTable = Direction
End Function

Public Function CalculateIntersection(myRay As Ray, sngDistance As Single) As Vector3D
    Dim Intersection As Vector3D
    
    '//Calculate the location of the intersection between the primitive and the ray
    With myRay
        Intersection.x = .Origin.x + sngDistance * .Direction.x
        Intersection.y = .Origin.y + sngDistance * .Direction.y
        Intersection.z = .Origin.z + sngDistance * .Direction.z
    End With
    
    '//Return result
    CalculateIntersection = Intersection
End Function


Public Function CalculateLightRay(lightLoc As Vector3D, Intersection As Vector3D, rayToLight As Ray) As Single
    '//Puts the resulting vector in lightDir
    Dim LightDir As Vector3D
    Dim lightDirMagnitudeReciprocal As Single
    Dim distanceToLight As Single
    
    LightDir = VectorSub(lightLoc, Intersection)
    
    '//Because we need the distance to the light for the shadow calculations, we will normalize lightDir manually
    distanceToLight = Sqr((LightDir.x * LightDir.x) + (LightDir.y * LightDir.y) + (LightDir.z * LightDir.z))

    lightDirMagnitudeReciprocal = 1 / distanceToLight
    
    With LightDir
        .x = .x * lightDirMagnitudeReciprocal
        .y = .y * lightDirMagnitudeReciprocal
        .z = .z * lightDirMagnitudeReciprocal
    End With
    
    '//Change the values BYREF
    With rayToLight
        .Origin = Intersection
        .Direction = LightDir
    End With
    
    '//Return distance
    CalculateLightRay = distanceToLight
End Function

Public Function CalculateLightingCoef(IsShadowed As Boolean, directionToLight As Vector3D, Normal As Vector3D) As Single
    Dim lightCoef As Single
    If IsShadowed = True Then
        '//No light reaches the intersection
        CalculateLightingCoef = 0
    Else
        '//Only calculate how much light reaches the intersection if it is not in shadow
        lightCoef = VectorDot(directionToLight, Normal)

        If lightCoef < 0 Then lightCoef = 0
        CalculateLightingCoef = lightCoef
    End If
End Function

Public Function CalculateReflection(myRay As Ray, Intersection As Vector3D, Normal As Vector3D, currPrimitiveNum As Long, Primitives() As Primitive, numPrimitives As Long, LightSources() As LightSource, numLightSources As Long, Depth As Long) As RGBQUAD
    Dim reflectedRay As Ray
    Dim nDotI As Single
    '// R = I - 2(N.I)*N
    
    nDotI = 2 * ((Normal.x * myRay.Direction.x) + (Normal.y * myRay.Direction.y) + (Normal.z * myRay.Direction.z))
    With reflectedRay
        .Direction.x = myRay.Direction.x - (nDotI * Normal.x)
        .Direction.y = myRay.Direction.y - (nDotI * Normal.y)
        .Direction.z = myRay.Direction.z - (nDotI * Normal.z)
        .Origin = Intersection
    End With
    
   CalculateReflection = TraceRay(currPrimitiveNum, reflectedRay, Primitives, numPrimitives, LightSources, numLightSources, Depth + 1)
End Function


Public Function Shade(myPrimitive As Primitive, currPrimitiveNum As Long, myRay As Ray, Distance As Single, LightSources() As LightSource, numLightSources As Long, Primitives() As Primitive, numPrimitives As Long, Depth As Long) As RGBQUAD
    Dim returnColor As ColorFloat, reflectedColor As RGBQUAD
    Dim Quad As RGBQUAD
    Dim Intersection As Vector3D
    Dim Normal As Vector3D
    Dim distanceToLight As Single, lightCoef As Single
    Dim rayToLight As Ray
    Dim I As Long
    Dim sngReflect As Single
    
    Intersection = CalculateIntersection(myRay, Distance)
    
    Normal = CalculateNormal(myPrimitive, Intersection)
    
    '//Add specular components
    For I = 0 To numLightSources
        '//Sets rayToLight
        distanceToLight = CalculateLightRay(LightSources(I).location, Intersection, rayToLight)
        
        '//-->Shadows enabled?
        If frmMain.chkShadows.Value = vbChecked Then
            lightCoef = CalculateLightingCoef(IsShadowed(currPrimitiveNum, rayToLight, distanceToLight, Primitives, numPrimitives), rayToLight.Direction, Normal)
        Else
            lightCoef = CalculateLightingCoef(False, rayToLight.Direction, Normal)
        End If
        
        '//Try checking first if lightCoef is 0... the check will probably be amortized over the cost of all these multiplications
        With returnColor
            .R = .R + myPrimitive.Surface.BaseColor.R * lightCoef * LightSources(I).Color.R
            .G = .G + myPrimitive.Surface.BaseColor.G * lightCoef * LightSources(I).Color.G
            .b = .b + myPrimitive.Surface.BaseColor.b * lightCoef * LightSources(I).Color.b
        End With
        
        '//Add reflective components --> If enabled
        If frmMain.chkReflections.Value = vbChecked Then
            If myPrimitive.Surface.sngReflectivity <> 0 Then
                reflectedColor = CalculateReflection(myRay, Intersection, Normal, currPrimitiveNum, Primitives, numPrimitives, LightSources, numLightSources, Depth)
                With returnColor
                    sngReflect = myPrimitive.Surface.sngReflectivity
                    .R = .R + reflectedColor.rgbRed * sngReflect
                    .G = .G + reflectedColor.rgbGreen * sngReflect
                    .b = .b + reflectedColor.rgbBlue * sngReflect
                End With
            End If
        End If
    Next I
    
    With returnColor
        If .R > 255 Then .R = 255
        If .G > 255 Then .G = 255
        If .b > 255 Then .b = 255
    End With

    
    With Quad
        .rgbRed = returnColor.R
        .rgbGreen = returnColor.G
        .rgbBlue = returnColor.b
    End With
    
    '//Return
    Shade = Quad
End Function

Public Sub TraceScene(Primitives() As Primitive, numPrimitives As Long, LightSources() As LightSource, numLightSources As Long)
    Dim primaryRay As Ray
    Dim xDiff As Long, yDiff As Long
    Dim y As Long, x As Long, lngY As Long
    Dim Color As RGBQUAD
    '//Setup view rays
    With primaryRay.Origin
        .x = 0
        .y = 0
        .z = -frmMain.scrFov.Value
    End With
    
    xDiff = 80 '//Equal to 1/2 of the vertical screen size
    yDiff = 60
        
    For y = 160 To 320
        lngY = y * 640
        For x = (320 - xDiff) To 320 + xDiff
            primaryRay.Direction = DirectionTable(x + lngY)             ' // implimenting the direction table added 20.125 fps @ 240x180....
                                                                        ' // this could be changed so that the pointer is incrimented after each pixel, rather than recalcing the whole thing each time
                                                                         '// this would save a few adds and bitshifts. same with the buffer set at the bottom
            Color = TraceRay(-1, primaryRay, Primitives, numPrimitives, LightSources, numLightSources, 0)
            
            With Color
                If (.rgbRed > 255) Then .rgbRed = 255
                If (.rgbGreen > 255) Then .rgbGreen = 255
                If (.rgbBlue > 255) Then .rgbBlue = 255
            End With
            Bits(x, y) = Color
        Next x
    Next y

    With frmMain.picRay
        '//Set the bits back to the picture
        SetDIBits lngHDC, lngImageHandle, 0, BInfo.bmiHeader.biHeight, Bits(0, 0), BInfo, DIB_RGB_COLORS
        
        '//Refresh
        .Refresh
    End With
End Sub



Public Function TraceRay(ignoreNum As Long, myRay As Ray, Primitives() As Primitive, numPrimitives As Long, LightSources() As LightSource, numLightSources As Long, Depth As Long) As RGBQUAD
    Dim returnColor As RGBQUAD
    Dim closestIntersectionDistance As Single
    Dim closestIntersectedPrimitiveNum As Long
    Dim currResult As TraceResult
    Dim currPrimitiveNum As Long
    
    If Depth > 4 Then '// prevent infinite reflection
        With returnColor
            .rgbRed = 0
            .rgbGreen = 0
            .rgbBlue = 0
            
            TraceRay = returnColor
            Exit Function
        End With
    End If
    
    closestIntersectionDistance = 100000    '// an impossibly large value
    closestIntersectedPrimitiveNum = -1
    
    '//Cycle through all of the spheres to find the closest interesction
    For currPrimitiveNum = 0 To numPrimitives
        '//Check if this primitive is enabled by the user
            If currPrimitiveNum <> ignoreNum Then
                currResult = IntersectPrimitive(Primitives(currPrimitiveNum), myRay)
            
                If currResult.Hit = True Then
                    If (currResult.Distance < closestIntersectionDistance) Then
                        closestIntersectionDistance = currResult.Distance
                        closestIntersectedPrimitiveNum = currPrimitiveNum
                    End If
                End If
            End If
    Next currPrimitiveNum
    
    If closestIntersectedPrimitiveNum = -1 Then  '// nothing was intersected
        With returnColor
            .rgbRed = 0
            .rgbGreen = 0
            .rgbBlue = 0
        End With
    Else
        returnColor = Shade(Primitives(closestIntersectedPrimitiveNum), closestIntersectedPrimitiveNum, myRay, closestIntersectionDistance, LightSources, numLightSources, Primitives, numPrimitives, Depth)
    End If
    
    TraceRay = returnColor
End Function

Public Function IntersectPrimitive(myPrimitive As Primitive, myRay As Ray) As TraceResult
    Select Case myPrimitive.Type
        Case SPHERE_TYPE
            IntersectPrimitive = IntersectSphere(myPrimitive.Sphere, myRay)
        
        Case PLANE_TYPE
            IntersectPrimitive = IntersectPlane(myPrimitive.Plane, myRay)
        
        Case CYLINDER_TYPE
            IntersectPrimitive = IntersectCylinder(myPrimitive.Cilinder, myRay)
        
        Case TRIANGLE_TYPE
            IntersectPrimitive = IntersectTriangle(myPrimitive.Triangle, myRay)
    End Select
End Function

Public Function CalculateNormal(myPrimitive As Primitive, Intersection As Vector3D) As Vector3D
    Select Case myPrimitive.Type
        Case SPHERE_TYPE
            CalculateNormal = SphereNormal(myPrimitive.Sphere, Intersection)
        
        Case PLANE_TYPE
            CalculateNormal = PlaneNormal(myPrimitive.Plane)
        
        Case CYLINDER_TYPE
            CalculateNormal = CylinderNormal(myPrimitive.Cilinder, Intersection)
        
        Case TRIANGLE_TYPE
            CalculateNormal = TriangleNormal(myPrimitive.Triangle)
    End Select
End Function
