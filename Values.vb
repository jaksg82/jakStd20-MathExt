Imports System.Math

Public Module Values

    ''' <summary>
    ''' Calculate the Nth root of a number
    ''' </summary>
    ''' <param name="value">Value</param>
    ''' <param name="root">Root degree</param>
    ''' <returns>Computed rooted value</returns>
    Public Function NthRoot(ByVal value As Double, ByVal root As Double) As Double
        NthRoot = Pow(value, (1 / root))
    End Function

    ''' <summary>
    ''' Check if the value are a finite number
    ''' </summary>
    ''' <param name="value">Value to check</param>
    ''' <returns>Return true if is finite</returns>
    Public Function IsFinite(ByVal value As Double) As Boolean
        If Double.IsInfinity(value) Then
            Return False
        ElseIf Double.IsNaN(value) Then
            Return False
        ElseIf Double.MinValue < value And value < Double.MaxValue Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Check if the value is a multiple of the base
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="base"></param>
    ''' <returns></returns>
    Public Function MultipleOf(value As Double, base As Double) As Boolean
        Dim Res1, Res2 As Double
        Res1 = value / base
        Res2 = Truncate(Res1)
        If Res1 = Res2 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Check if the value is a multiple of the base
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="base"></param>
    ''' <returns></returns>
    Public Function MultipleOf(value As Integer, base As Integer) As Boolean
        Dim Res1, Res2 As Double
        Res1 = value / base
        Res2 = Truncate(Res1)
        If Res1 = Res2 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Check if the value is a multiple of the base
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="base"></param>
    ''' <returns></returns>
    Public Function MultipleOf(value As Single, base As Single) As Boolean
        Dim Res1, Res2 As Double
        Res1 = value / base
        Res2 = Truncate(Res1)
        If Res1 = Res2 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Check if an integer value is even
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>True if even, false if odd</returns>
    Public Function IsEven(value As Integer) As Boolean
        If value Mod 2 <> 0 Then
            Return False 'Odd number
        Else
            Return True 'Even number
        End If
    End Function

    ''' <summary>
    ''' Calculate the average of given array of data.
    ''' </summary>
    ''' <param name="data">Array of values</param>
    ''' <returns></returns>
    Public Function Average(data() As Double) As Double
        Dim avg As Double = 0
        If data Is Nothing Then Return Double.NaN

        For d = 0 To data.Count - 1
            avg = avg + data(d)
        Next
        avg = avg / data.Count

        Return avg
    End Function
    Public Function Average(data As List(Of Double)) As Double
        Dim avg As Double = 0
        If data Is Nothing Then Return Double.NaN

        For d = 0 To data.Count - 1
            avg = avg + data(d)
        Next
        avg = avg / data.Count

        Return avg
    End Function

    ''' <summary>
    ''' Calculate the trend of given array of data.
    ''' </summary>
    ''' <param name="data">Array of values</param>
    ''' <returns></returns>
    Public Function Trend(data() As Double) As Double
        If data Is Nothing Then Return Double.NaN
        Dim tmp(data.Count - 1) As Double

        For t = 0 To data.Count - 2
            tmp(t) = data(t + 1) - data(t)
        Next

        Return Average(tmp)

    End Function
    Public Function Trend(data As List(Of Double)) As Double
        If data Is Nothing Then Return Double.NaN
        Dim tmp(data.Count - 1) As Double

        For t = 0 To data.Count - 2
            tmp(t) = data(t + 1) - data(t)
        Next

        Return Average(tmp)

    End Function

End Module
