Attribute VB_Name = "mdlPrimitives"
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


'//Vector
Public Type Vector3D
    x As Single
    y As Single
    z As Single
End Type

'//1 Ray
Public Type Ray
    Origin As Vector3D
    Direction As Vector3D
End Type

'//Color
Public Type ColorFloat
    R As Single
    G As Single
    b As Single
End Type

'//Result of raytrace (for one ray)
Public Type TraceResult
    Hit As Boolean
    Distance As Single
End Type

'//Surface properties
Public Type udtSurface
    BaseColor As ColorFloat
    sngReflectivity As Single
End Type

'//Light source
Public Type LightSource
    location As Vector3D
    Color As ColorFloat
End Type

'-------------------------------

'//Primitives
Public Enum EnumAxis
    X_INFINITE
    Y_INFINITE
    Z_INFINITE
End Enum

Public Type udtCylinder
    Axis As EnumAxis
    Center As Vector3D
    sngRadius As Single
End Type


Public Type udtPlane
    sngDisplacement As Single
    vecNormal As Vector3D
End Type

Public Type udtSphere
    vecCenter As Vector3D
    sngRadius As Single
End Type

Public Type udtTriangle
    v1 As Vector3D
    v2 As Vector3D
    v3 As Vector3D
End Type
'\\

'-------------------------------

'//Collection of primitives
Public Enum EnumPrimitives
    SPHERE_TYPE = 0
    PLANE_TYPE = 1
    CYLINDER_TYPE = 2
    TRIANGLE_TYPE = 3
End Enum

Public Type Primitive
    Cilinder As udtCylinder
    Plane As udtPlane
    Sphere As udtSphere
    Triangle As udtTriangle
    Surface As udtSurface
    Type As EnumPrimitives
End Type
'\\




'//Math

Public Function CylinderNormal(Cylinder As udtCylinder, Intersection As Vector3D) As Vector3D
    Dim Normal As Vector3D
    Dim oneOverRadius As Single
    oneOverRadius = 1 / Cylinder.sngRadius ' // might be faster to precalc this and store it in the sphere data structure, but also might be slower because it takes longer to load 32 bits of
                                              '// data than to calculate 1 division (i think)
    Select Case Cylinder.Axis
        Case X_INFINITE
            With Normal
                .x = 0
                .y = (Intersection.y - Cylinder.Center.y) * oneOverRadius
                .z = (Intersection.z - Cylinder.Center.z) * oneOverRadius
                CylinderNormal = Normal
            End With
            
        Case Y_INFINITE
            With Normal
                .x = (Intersection.x - Cylinder.Center.x) * oneOverRadius
                .y = 0
                .z = (Intersection.z - Cylinder.Center.z) * oneOverRadius
                CylinderNormal = Normal
            End With
            
        Case Z_INFINITE
            With Normal
                .x = (Intersection.x - Cylinder.Center.x) * oneOverRadius
                .y = (Intersection.y - Cylinder.Center.y) * oneOverRadius
                .z = 0
                CylinderNormal = Normal
            End With
    End Select
    
End Function

Public Function PlaneNormal(Plane As udtPlane) As Vector3D
    PlaneNormal = Plane.vecNormal
End Function

Public Function SphereNormal(Sphere As udtSphere, Intersection As Vector3D) As Vector3D
    Dim Normal As Vector3D
    Dim sngOneOverRadius As Single
    '// calculate the normal of the sphere at the point of interesction
    sngOneOverRadius = 1 / Sphere.sngRadius   '// might be faster to precalc this and store it in the sphere data structure, but also might be slower because it takes longer to load 32 bits of
                                           '// data than to calculate 1 division (i think)
    With Sphere.vecCenter
        Normal.x = (Intersection.x - .x) * sngOneOverRadius  ' // same as ( intersection.x - sphere.center.x ) / sphere.radiu
        Normal.y = (Intersection.y - .y) * sngOneOverRadius
        Normal.z = (Intersection.z - .z) * sngOneOverRadius
    End With
    
    SphereNormal = Normal
End Function

Public Function TriangleNormal(Triangle As udtTriangle) As Vector3D
    Dim vecEdge1 As Vector3D, vecEdge2 As Vector3D
    Dim vecNormal As Vector3D
    
    vecEdge1 = VectorSub(Triangle.v2, Triangle.v1)
    vecEdge2 = VectorSub(Triangle.v3, Triangle.v1)
    vecNormal = VectorCross(vecEdge1, vecEdge2)
    Call VectorNormalize(vecNormal)
    
    TriangleNormal = vecNormal
End Function

'---------------------------------------------------------------------------------
'INTERSECT FUNCTIONS
'---------------------------------------------------------------------------------

Public Function IntersectCylinder(Cylinder As udtCylinder, myRay As Ray) As TraceResult
    Dim myTraceResult As TraceResult
    Dim a As Single, b As Single, C As Single
    Dim Discriminant As Single
    
    Select Case Cylinder.Axis
        Case X_INFINITE
            With myRay
                a = .Direction.y * .Direction.y + .Direction.z * .Direction.z
                b = 2 * (.Direction.y * (.Origin.y - Cylinder.Center.y) + .Direction.z * (.Origin.z - Cylinder.Center.z))
                C = (.Origin.y - Cylinder.Center.y) * (.Origin.y - Cylinder.Center.y) + (.Origin.z - Cylinder.Center.z) * (.Origin.z - Cylinder.Center.z) - Cylinder.sngRadius * Cylinder.sngRadius
            End With
            
        Case Y_INFINITE
            With myRay
                a = .Direction.x * .Direction.x + .Direction.z * .Direction.z
                b = 2 * (.Direction.x * (.Origin.x - Cylinder.Center.x) + .Direction.z * (.Origin.z - Cylinder.Center.z))
                C = (.Origin.x - Cylinder.Center.x) * (.Origin.x - Cylinder.Center.x) + (.Origin.z - Cylinder.Center.z) * (.Origin.z - Cylinder.Center.z) - Cylinder.sngRadius * Cylinder.sngRadius
            End With
            
        Case Z_INFINITE:
            With myRay
                a = .Direction.x * .Direction.x + .Direction.y * .Direction.y
                b = 2 * (.Direction.x * (.Origin.x - Cylinder.Center.x) + .Direction.y * (.Origin.y - Cylinder.Center.y))
                C = (.Origin.x - Cylinder.Center.x) * (.Origin.x - Cylinder.Center.x) + (.Origin.y - Cylinder.Center.y) * (.Origin.y - Cylinder.Center.y) - Cylinder.sngRadius * Cylinder.sngRadius
            End With
    End Select
    
    Discriminant = b * b - 4 * a * C
    If Discriminant < 0 Then
        myTraceResult.Hit = False
        IntersectCylinder = myTraceResult
        Exit Function
    End If
    
    myTraceResult.Distance = (-b - Sqr(Discriminant)) / (2 * a)
    If myTraceResult.Distance < 0 Then
        myTraceResult.Hit = False
        IntersectCylinder = myTraceResult
        Exit Function
    End If
        
    '//Return true
    myTraceResult.Hit = True
    IntersectCylinder = myTraceResult

End Function

Public Function IntersectPlane(Plane As udtPlane, myRay As Ray) As TraceResult
    Dim myTraceResult As TraceResult
    Dim t As Single
    On Error Resume Next
    
    t = -(Plane.vecNormal.x * myRay.Origin.x + Plane.vecNormal.y * myRay.Origin.y + Plane.vecNormal.z * myRay.Origin.z + Plane.sngDisplacement) / (Plane.vecNormal.x * myRay.Direction.x + Plane.vecNormal.y * myRay.Direction.y + Plane.vecNormal.z * myRay.Direction.z)
    
    If t < 0 Then
        myTraceResult.Hit = False
        IntersectPlane = myTraceResult
        Exit Function
    End If
    
    myTraceResult.Hit = True
    myTraceResult.Distance = t
    IntersectPlane = myTraceResult
End Function

Public Function IntersectSphere(Sphere As udtSphere, myRay As Ray) As TraceResult
    Dim myTraceResult As TraceResult
    Dim rayToSphereCenter As Vector3D
    Dim lengthRTSC2 As Single, closestApproach As Single, halfCord2 As Single
    
    rayToSphereCenter = VectorSub(Sphere.vecCenter, myRay.Origin)
    lengthRTSC2 = VectorDot(rayToSphereCenter, rayToSphereCenter)   ' // lengthRTSC2 = length of the ray from the ray's origin to the sphere's center squared
      
    closestApproach = VectorDot(rayToSphereCenter, myRay.Direction)
    If closestApproach < 0 Then '// the intersection is behind the ray
        myTraceResult.Hit = False
        IntersectSphere = myTraceResult
        Exit Function
    End If
    
    '// halfCord2 = the distance squared from the closest approach of the ray to a perpendicular to the ray through the center of the sphere to the place where the ray actually intersects the sphere
    halfCord2 = (Sphere.sngRadius * Sphere.sngRadius) - lengthRTSC2 + (closestApproach * closestApproach)  '// sphere.radius * sphere.radius could be precalced, but it might take longer to load it
                                                                                                            '// than to calculate it
    If halfCord2 < 0 Then '// the ray misses the sphere
        myTraceResult.Hit = False
        IntersectSphere = myTraceResult
        Exit Function
    End If
    
    myTraceResult.Hit = True
    myTraceResult.Distance = closestApproach - Sqr(halfCord2)
    IntersectSphere = myTraceResult
End Function

Public Function IntersectTriangle(Triangle As udtTriangle, myRay As Ray) As TraceResult
    Dim myTraceResult As TraceResult
    Dim u As Single, V As Single
    Dim edge1 As Vector3D, edge2 As Vector3D, tvec As Vector3D, pvec As Vector3D, qvec As Vector3D
    Dim det As Single, invDet As Single
    
    edge1 = VectorSub(Triangle.v2, Triangle.v1)
    edge2 = VectorSub(Triangle.v3, Triangle.v1)
    pvec = VectorCross(myRay.Direction, edge2)
    
    det = VectorDot(edge1, pvec)
    
    With myTraceResult
        If det > -0.000001 And det < 0.000001 Then
            .Hit = False
            IntersectTriangle = myTraceResult
            Exit Function
        End If
        
        invDet = 1 / det
        
        tvec = VectorSub(myRay.Origin, Triangle.v1)
        
        u = VectorDot(tvec, pvec) * invDet
        
        If u < 0 Or u > 1 Then
            .Hit = False
            IntersectTriangle = myTraceResult
            Exit Function
        End If
        
        qvec = VectorCross(tvec, edge1)
        
        V = VectorDot(myRay.Direction, qvec) * invDet
        If (V < 0 Or (u + V) > 1) Then
            .Hit = False
            IntersectTriangle = myTraceResult
            Exit Function
        End If
        
        .Distance = VectorDot(edge2, qvec) * invDet
        If (.Distance < 0) Then
            .Hit = False
            .Hit = False
            IntersectTriangle = myTraceResult
            Exit Function
        End If
        
        .Hit = True
        IntersectTriangle = myTraceResult
    End With
End Function

'// to optimize, write special shadow functions for each primitive. add a surface property: shadowed
'// so it doesnt shadow itself
Public Function IsShadowed(currPrimitiveNum As Long, rayToLight As Ray, distanceToLight As Single, Primitives() As Primitive, numPrimitives As Long)
    Dim myTraceResult As TraceResult
    Dim I As Long
    myTraceResult.Hit = False
    
    '// check every other primitive
    For I = 0 To numPrimitives
        If I <> currPrimitiveNum Then  '// dont self-shadow
            
            Select Case Primitives(I).Type
                Case SPHERE_TYPE:
                    myTraceResult = IntersectSphere(Primitives(I).Sphere, rayToLight)
                
                Case PLANE_TYPE:
                    myTraceResult = IntersectPlane(Primitives(I).Plane, rayToLight)
                
                Case CYLINDER_TYPE:
                    myTraceResult = IntersectCylinder(Primitives(I).Cilinder, rayToLight)
                
                Case TRIANGLE_TYPE:
                    myTraceResult = IntersectTriangle(Primitives(I).Triangle, rayToLight)
            End Select
            
            If (myTraceResult.Hit = True And (myTraceResult.Distance < distanceToLight)) Then
                    IsShadowed = True
                    Exit Function
            End If
        End If
    Next I

    IsShadowed = False

End Function
