Public Class Point3D
    Dim dblX, dblY, dblZ As Double

    Property X As Double
        Set(value As Double)
            dblX = value
        End Set
        Get
            Return dblX
        End Get
    End Property

    Property Y As Double
        Set(value As Double)
            dblY = value
        End Set
        Get
            Return dblY
        End Get
    End Property

    Property Z As Double
        Set(value As Double)
            dblZ = value
        End Set
        Get
            Return dblZ
        End Get
    End Property

    Public Sub New()
        dblX = Double.NaN
        dblY = Double.NaN
        dblZ = Double.NaN
    End Sub

    Public Sub New(valueX As Double, valueY As Double)
        dblX = valueX
        dblY = valueY
        dblZ = Double.NaN
    End Sub

    Public Sub New(valueX As Double, valueY As Double, valueZ As Double)
        dblX = valueX
        dblY = valueY
        dblZ = valueZ
    End Sub

    Public Shared Operator +(point1 As Point3D, point2 As Point3D) As Point3D
        Dim point3 As New Point3D
        If point1 Is Nothing Or point2 Is Nothing Then Return point3
        If IsFinite(point1.X) Then
            If IsFinite(point2.X) Then
                point3.X = point1.X + point2.X
            Else
                point3.X = point1.X
            End If
        Else
            If IsFinite(point2.X) Then
                point3.X = point2.X
            End If
        End If

        If IsFinite(point1.Y) Then
            If IsFinite(point2.Y) Then
                point3.Y = point1.Y + point2.Y
            Else
                point3.Y = point1.Y
            End If
        Else
            If IsFinite(point2.Y) Then
                point3.Y = point2.Y
            End If
        End If

        If IsFinite(point1.Z) Then
            If IsFinite(point2.Z) Then
                point3.Z = point1.Z + point2.Z
            Else
                point3.Z = point1.Z
            End If
        Else
            If IsFinite(point2.Z) Then
                point3.Z = point2.Z
            End If
        End If

        Return point3

    End Operator

    Public Shared Function Add(point1 As Point3D, point2 As Point3D) As Point3D
        Return point1 + point2
    End Function

    Public Shared Operator -(point1 As Point3D, point2 As Point3D) As Point3D
        Dim point3 As New Point3D
        If point1 Is Nothing Or point2 Is Nothing Then Return point3
        If IsFinite(point1.X) Then
            If IsFinite(point2.X) Then
                point3.X = point1.X - point2.X
            Else
                point3.X = point1.X
            End If
        Else
            If IsFinite(point2.X) Then
                point3.X = point2.X
            End If
        End If

        If IsFinite(point1.Y) Then
            If IsFinite(point2.Y) Then
                point3.Y = point1.Y - point2.Y
            Else
                point3.Y = point1.Y
            End If
        Else
            If IsFinite(point2.Y) Then
                point3.Y = point2.Y
            End If
        End If

        If IsFinite(point1.Z) Then
            If IsFinite(point2.Z) Then
                point3.Z = point1.Z - point2.Z
            Else
                point3.Z = point1.Z
            End If
        Else
            If IsFinite(point2.Z) Then
                point3.Z = point2.Z
            End If
        End If

        Return point3

    End Operator

    Public Shared Function Subtract(point1 As Point3D, point2 As Point3D) As Point3D
        Return point1 - point2
    End Function

    Public Shared Operator *(point1 As Point3D, point2 As Point3D) As Point3D
        Dim point3 As New Point3D
        If point1 Is Nothing Or point2 Is Nothing Then Return point3
        If IsFinite(point1.X) Then
            If IsFinite(point2.X) Then
                point3.X = point1.X * point2.X
            Else
                point3.X = point1.X
            End If
        Else
            If IsFinite(point2.X) Then
                point3.X = point2.X
            End If
        End If

        If IsFinite(point1.Y) Then
            If IsFinite(point2.Y) Then
                point3.Y = point1.Y * point2.Y
            Else
                point3.Y = point1.Y
            End If
        Else
            If IsFinite(point2.Y) Then
                point3.Y = point2.Y
            End If
        End If

        If IsFinite(point1.Z) Then
            If IsFinite(point2.Z) Then
                point3.Z = point1.Z * point2.Z
            Else
                point3.Z = point1.Z
            End If
        Else
            If IsFinite(point2.Z) Then
                point3.Z = point2.Z
            End If
        End If

        Return point3

    End Operator

    Public Shared Function Multiply(point1 As Point3D, point2 As Point3D) As Point3D
        Return point1 * point2
    End Function

    Public Shared Operator /(point1 As Point3D, point2 As Point3D) As Point3D
        Dim point3 As New Point3D
        If point1 Is Nothing Or point2 Is Nothing Then Return point3
        If IsFinite(point1.X) Then
            If IsFinite(point2.X) Then
                point3.X = point1.X / point2.X
            Else
                point3.X = point1.X
            End If
        Else
            If IsFinite(point2.X) Then
                point3.X = point2.X
            End If
        End If

        If IsFinite(point1.Y) Then
            If IsFinite(point2.Y) Then
                point3.Y = point1.Y / point2.Y
            Else
                point3.Y = point1.Y
            End If
        Else
            If IsFinite(point2.Y) Then
                point3.Y = point2.Y
            End If
        End If

        If IsFinite(point1.Z) Then
            If IsFinite(point2.Z) Then
                point3.Z = point1.Z / point2.Z
            Else
                point3.Z = point1.Z
            End If
        Else
            If IsFinite(point2.Z) Then
                point3.Z = point2.Z
            End If
        End If

        Return point3

    End Operator

    Public Shared Function Divide(point1 As Point3D, point2 As Point3D) As Point3D
        Return point1 / point2
    End Function

    Public Shared Operator =(point1 As Point3D, point2 As Point3D) As Boolean
        Dim pointX, pointY, pointZ As Boolean
        If point1 Is Nothing Or point2 Is Nothing Then Return False

        pointX = (point1.X = point2.X)
        pointY = (point1.Y = point2.Y)
        pointZ = (point1.Z = point2.Z)

        If pointX And pointY And pointZ Then
            Return True
        Else
            Return False
        End If

    End Operator

    Public Shared Operator <>(point1 As Point3D, point2 As Point3D) As Boolean
        If point1 Is Nothing Or point2 Is Nothing Then
            Return False
        Else
            Return Not (point1 = point2)
        End If
    End Operator

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        Else
            Dim pnt As Point3D = CType(obj, Point3D)
            Return Me = pnt
        End If
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function
End Class
