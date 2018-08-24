Imports System.Math

Public Module Geometry
    ''' <summary>
    ''' Calculate the heading of the line to join the two given point
    ''' </summary>
    ''' <param name="point1X">Start point X</param>
    ''' <param name="point1Y">Start point Y</param>
    ''' <param name="point2X">End point X</param>
    ''' <param name="point2Y">End point Y</param>
    ''' <returns>Heading value in radians</returns>
    Public Function CalcHeading(ByVal point1X As Double, ByVal point1Y As Double,
                                ByVal point2X As Double, ByVal point2Y As Double) As Double
        Dim DeltaX, DeltaY, hdg As Double
        DeltaX = point2X - point1X
        DeltaY = point2Y - point1Y
        hdg = Math.Atan2(DeltaY, DeltaX)
        If hdg < 0 Then hdg = (2 * Math.PI) + hdg
        Return hdg
    End Function

    ''' <summary>
    ''' Calculate the heading of the line to join the two given point
    ''' </summary>
    ''' <param name="point1">Start point</param>
    ''' <param name="point2">End point</param>
    ''' <returns>Heading value in radians</returns>
    Public Function CalcHeading(ByVal point1 As Point3D, ByVal point2 As Point3D) As Double
        Dim DeltaX, DeltaY, hdg As Double
        If point1 Is Nothing Or point2 Is Nothing Then
            Return Double.NaN
        End If
        DeltaX = point2.X - point1.X
        DeltaY = point2.Y - point1.Y
        hdg = Math.Atan2(DeltaY, DeltaX)
        If hdg < 0 Then hdg = (2 * Math.PI) + hdg
        Return hdg
    End Function

    ''' <summary>
    ''' Calculate the coordinate of a point from the given point.
    ''' </summary>
    ''' <param name="startPointX">Start point X</param>
    ''' <param name="startPointY">Start point Y</param>
    ''' <param name="range">Distance between start point and end point</param>
    ''' <param name="bearingR">Bearing in radians between start point and end point</param>
    ''' <returns>Computed End point</returns>
    Public Function RangeBearing(ByVal startPointX As Double, ByVal startPointY As Double, ByVal range As Double, ByVal bearingR As Double) As Point3D
        Dim tmppoint As New Point3D
        bearingR = AngleFit2pi(bearingR)
        Select Case bearingR
            Case Is = 0
                tmppoint.X = startPointX + range
                tmppoint.Y = startPointY
            Case Is = Math.PI / 2
                tmppoint.X = startPointX
                tmppoint.Y = startPointY + range
            Case Is = Math.PI
                tmppoint.X = startPointX - range
                tmppoint.Y = startPointY
            Case Is = 3 * Math.PI / 2
                tmppoint.X = startPointX
                tmppoint.Y = startPointY - range
            Case Is = 2 * Math.PI
                tmppoint.X = startPointX + range
                tmppoint.Y = startPointY
            Case Is < Math.PI / 2
                tmppoint.X = startPointX + If(range > 0, Math.Abs(range * Math.Cos(bearingR)), -Math.Abs(range * Math.Cos(bearingR)))
                tmppoint.Y = startPointY + If(range > 0, Math.Abs(range * Math.Sin(bearingR)), -Math.Abs(range * Math.Sin(bearingR)))
            Case Is > 3 * Math.PI / 2
                tmppoint.X = startPointX + If(range > 0, Math.Abs(range * Math.Cos(bearingR)), -Math.Abs(range * Math.Cos(bearingR)))
                tmppoint.Y = startPointY - If(range > 0, Math.Abs(range * Math.Sin(bearingR)), -Math.Abs(range * Math.Sin(bearingR)))
            Case Else
                If bearingR > Math.PI / 2 And bearingR < Math.PI Then
                    tmppoint.X = startPointX - If(range > 0, Math.Abs(range * Math.Cos(bearingR)), -Math.Abs(range * Math.Cos(bearingR)))
                    tmppoint.Y = startPointY + If(range > 0, Math.Abs(range * Math.Sin(bearingR)), -Math.Abs(range * Math.Sin(bearingR)))
                Else
                    If bearingR > Math.PI And bearingR < 3 * Math.PI / 2 Then
                        tmppoint.X = startPointX - If(range > 0, Math.Abs(range * Math.Cos(bearingR)), -Math.Abs(range * Math.Cos(bearingR)))
                        tmppoint.Y = startPointY - If(range > 0, Math.Abs(range * Math.Sin(bearingR)), -Math.Abs(range * Math.Sin(bearingR)))
                    Else
                        tmppoint.X = startPointX
                        tmppoint.Y = startPointY
                    End If
                End If
        End Select
        Return tmppoint
    End Function

    ''' <summary>
    ''' Calculate the coordinate of a point from the given point.
    ''' </summary>
    ''' <param name="startPoint">Start point</param>
    ''' <param name="range">Distance between start point and end point</param>
    ''' <param name="bearingR">Bearing in radians between start point and end point</param>
    ''' <returns>Computed End point</returns>
    Public Function RangeBearing(ByVal startPoint As Point3D, ByVal range As Double, ByVal bearingR As Double) As Point3D
        Dim tmppoint As New Point3D
        If startPoint IsNot Nothing Then
            tmppoint = RangeBearing(startPoint.X, startPoint.Y, range, bearingR)
        End If
        Return tmppoint
    End Function

    ''' <summary>
    ''' Calculate the 2D distance between two points
    ''' </summary>
    ''' <param name="point1X">Start point X</param>
    ''' <param name="point1Y">Start point Y</param>
    ''' <param name="point2X">End point X</param>
    ''' <param name="point2Y">End point Y</param>
    ''' <returns>Computed distance</returns>
    Public Function Distance2D(ByVal point1X As Double, ByVal point1Y As Double,
                             ByVal point2X As Double, ByVal point2Y As Double) As Double
        Return Math.Sqrt((point1X - point2X) ^ 2 + (point1Y - point2Y) ^ 2)
    End Function

    ''' <summary>
    ''' Calculate the 2D distance between two points
    ''' </summary>
    ''' <param name="point1">Start point</param>
    ''' <param name="point2">End point</param>
    ''' <returns>Computed distance</returns>
    Public Function Distance2D(ByVal point1 As Point3D, ByVal point2 As Point3D) As Double
        If point1 Is Nothing Or point2 Is Nothing Then
            Return Double.NaN
        Else
            Return Math.Sqrt((point1.X - point2.X) ^ 2 + (point1.Y - point2.Y) ^ 2)
        End If
    End Function

    ''' <summary>
    ''' Calculate the 3D distance between two points
    ''' </summary>
    ''' <param name="point1X">Start point X</param>
    ''' <param name="point1Y">Start point Y</param>
    ''' <param name="point1Z">Start point Z</param>
    ''' <param name="point2X">End point X</param>
    ''' <param name="point2Y">End point Y</param>
    ''' <param name="point2Z">End point Z</param>
    ''' <returns>Computed distance</returns>
    Public Function Distance3D(ByVal point1X As Double, ByVal point1Y As Double, ByVal point1Z As Double,
                             ByVal point2X As Double, ByVal point2Y As Double, ByVal point2Z As Double) As Double
        Return Math.Sqrt((point1X - point2X) ^ 2 + (point1Y - point2Y) ^ 2 + (point1Z - point2Z) ^ 2)
    End Function

    ''' <summary>
    ''' Calculate the 3D distance between two points
    ''' </summary>
    ''' <param name="point1">Start point</param>
    ''' <param name="point2">End point</param>
    ''' <returns>Computed distance</returns>
    Public Function Distance3D(ByVal point1 As Point3D, ByVal point2 As Point3D) As Double
        If point1 Is Nothing Or point2 Is Nothing Then
            Return Double.NaN
        Else
            Return Math.Sqrt((point1.X - point2.X) ^ 2 + (point1.Y - point2.Y) ^ 2 + (point1.Z - point2.Z) ^ 2)
        End If
    End Function

End Module
