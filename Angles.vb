Imports System.Math

Public Module Angles

    ''' <summary>
    ''' Convert the degree value in to the equivalent radian value
    ''' </summary>
    ''' <param name="angleD">Angle to convert.</param>
    ''' <returns>Equivalent angle in radians.</returns>
    Public Function DegRad(ByVal angleD As Double) As Double
        DegRad = angleD / (180 / Math.PI)
    End Function

    ''' <summary>
    ''' Convert the radian value in to the equivalent degree value
    ''' </summary>
    ''' <param name="angleR">Angle to convert.</param>
    ''' <returns>Equivalent angle in degrees.</returns>
    Public Function RadDeg(ByVal angleR As Double) As Double
        RadDeg = angleR * (180 / Math.PI)
    End Function

    ''' <summary>
    ''' Function to shrink the angle value between 0 and 2pi
    ''' </summary>
    ''' <param name="angleR">Angle to check</param>
    ''' <returns></returns>
    Public Function AngleFit2Pi(ByVal angleR As Double) As Double
        Dim TmpAngle As Double = angleR
        Do
            If TmpAngle > (Math.PI * 2) Then TmpAngle = TmpAngle - (Math.PI * 2)
            If TmpAngle < 0 Then TmpAngle = (Math.PI * 2) + TmpAngle
        Loop Until (TmpAngle <= (Math.PI * 2) And TmpAngle >= 0)
        Return TmpAngle
    End Function

    ''' <summary>
    ''' Function to shrink the angle value between -pi and +pi
    ''' </summary>
    ''' <param name="angleR">Angle to check</param>
    ''' <returns></returns>
    Public Function AngleFit1Pi(ByVal angleR As Double) As Double
        Dim TmpAngle As Double = angleR
        Do
            If TmpAngle > Math.PI Then TmpAngle = TmpAngle - (Math.PI * 2)
            If TmpAngle < -Math.PI Then TmpAngle = (Math.PI * 2) + TmpAngle
        Loop Until (TmpAngle <= Math.PI And TmpAngle >= -Math.PI)
        Return TmpAngle
    End Function

    ''' <summary>
    ''' Convert an angle from the mathematical to the coordinates convention.
    ''' </summary>
    ''' <param name="angleR">Angle to convert</param>
    ''' <returns>Converted angle</returns>
    Public Function AngleToBearing(ByVal angleR As Double) As Double
        Return AngleFit2Pi(((2 * Math.PI) - angleR) + (Math.PI / 2))
    End Function

    ''' <summary>
    ''' Convert an angle from the coordinates to the mathematical convention.
    ''' </summary>
    ''' <param name="angleR">Angle to convert</param>
    ''' <returns>Converted angle</returns>
    Public Function BearingToAngle(ByVal angleR As Double) As Double
        Return AngleFit2Pi((2 * Math.PI) - (angleR - (Math.PI / 2)))
    End Function

    ''' <summary>
    ''' Represent the secant of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Secant</returns>
    Public Function Sec(ByVal value As Double) As Double
        Sec = 1 / Cos(value)
    End Function

    ''' <summary>
    ''' Represent the cosecant of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Cosecant</returns>
    Public Function Csc(ByVal value As Double) As Double
        Csc = 1 / Sin(value)
    End Function

    ''' <summary>
    ''' Represent the cotangent of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Cotangent</returns>
    Public Function CTan(ByVal value As Double) As Double
        CTan = 1 / Tan(value)
    End Function

    ''' <summary>
    ''' Retrive the angle from the sine value
    ''' </summary>
    ''' <param name="value">Sine of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ASin(ByVal value As Double) As Double
        ASin = Atan(value / Sqrt(-value * value + 1))
    End Function

    ''' <summary>
    ''' Retrive the angle from the cosine value
    ''' </summary>
    ''' <param name="value">Cosine of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACos(ByVal value As Double) As Double
        ACos = Atan(-value / Sqrt(-value * value + 1)) + 2 * Atan(1)
    End Function

    ''' <summary>
    ''' Retrive the angle from the secant value
    ''' </summary>
    ''' <param name="value">Secant of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ASec(ByVal value As Double) As Double
        ASec = 2 * Atan(1) - Atan(Sign(value) / Sqrt(value * value - 1))
    End Function

    ''' <summary>
    ''' Retrive the angle from the cosecant value
    ''' </summary>
    ''' <param name="value">Cosecant of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACsc(ByVal value As Double) As Double
        ACsc = Atan(Sign(value) / Sqrt(value * value - 1))
    End Function

    ''' <summary>
    ''' Retrive the angle from the cotangent value
    ''' </summary>
    ''' <param name="value">Cotangent of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACot(ByVal value As Double) As Double
        ACot = 2 * Atan(1) - Atan(value)
    End Function

    ''' <summary>
    ''' Represent the hyperbolic sine of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic Sine</returns>
    Public Function SinH(ByVal value As Double) As Double
        SinH = (Exp(value) - Exp(-value)) / 2
    End Function

    ''' <summary>
    ''' Represent the hyperbolic cosine of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic cosine</returns>
    Public Function CosH(ByVal value As Double) As Double
        CosH = (Exp(value) + Exp(-value)) / 2
    End Function

    ''' <summary>
    ''' Represent the hyperbolic tangent of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic tangent</returns>
    Public Function TanH(ByVal value As Double) As Double
        TanH = (Exp(value) - Exp(-value)) / (Exp(value) + Exp(-value))
    End Function

    ''' <summary>
    ''' Represent the hyperbolic secant of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic secant</returns>
    Public Function SecH(ByVal value As Double) As Double
        SecH = 2 / (Exp(value) + Exp(-value))
    End Function

    ''' <summary>
    ''' Represent the hyperbolic cosecant of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic cosecant</returns>
    Public Function CscH(ByVal value As Double) As Double
        CscH = 2 / (Exp(value) - Exp(-value))
    End Function

    ''' <summary>
    ''' Represent the hyperbolic cotangent of an angle
    ''' </summary>
    ''' <param name="value">Angle</param>
    ''' <returns>Hyperbolic cotangent</returns>
    Public Function CotH(ByVal value As Double) As Double
        CotH = (Exp(value) + Exp(-value)) / (Exp(value) - Exp(-value))
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic sine value
    ''' </summary>
    ''' <param name="value">Hyperbolic sine of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ASinH(ByVal value As Double) As Double
        ASinH = Log(value + Sqrt(value * value + 1))
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic cosine value
    ''' </summary>
    ''' <param name="value">Hyperbolic cosine of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACosH(ByVal value As Double) As Double
        ACosH = Log(value + Sqrt(value * value - 1))
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic tangent value
    ''' </summary>
    ''' <param name="value">Hyperbolic tangent of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ATanH(ByVal value As Double) As Double
        ATanH = Log((1 + value) / (1 - value)) / 2
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic secant value
    ''' </summary>
    ''' <param name="value">Hyperbolic secant of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ASecH(ByVal value As Double) As Double
        ASecH = Log((Sqrt(-value * value + 1) + 1) / value)
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic cosecant value
    ''' </summary>
    ''' <param name="value">Hyperbolic cosecant of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACscH(ByVal value As Double) As Double
        ACscH = Log((Sign(value) * Sqrt(value * value + 1) + 1) / value)
    End Function

    ''' <summary>
    ''' Retrive the angle from the hyperbolic cotangent value
    ''' </summary>
    ''' <param name="value">Hyperbolic cotangent of the angle</param>
    ''' <returns>Angle</returns>
    Public Function ACotH(ByVal value As Double) As Double
        ACotH = Log((value + 1) / (value - 1)) / 2
    End Function

    ''' <summary>
    ''' Retrieve the hypotenuse from X and Y dimensions
    ''' </summary>
    ''' <param name="valueX">X dimension</param>
    ''' <param name="valueY">Y dimension</param>
    ''' <returns>Hypotenuse</returns>
    Public Function Hypot(ByVal valueX As Double, ByVal valueY As Double) As Double
        Dim a, b As Double
        valueX = Abs(valueX)
        valueY = Abs(valueY)
        a = Max(valueX, valueY)
        b = Min(valueX, valueY) / (If(a = 0, 1, a))
        Hypot = a * Sqrt(1 + b * b)
    End Function

End Module