Public Class clsComputation

    Public Function SumArray(ByVal DataArray() As Double) As Double

        SumArray = DataArray(0)
        For i As Integer = 1 To DataArray.Length - 1
            SumArray = SumArray + DataArray(i)
        Next i

        Return SumArray

    End Function

    Public Function AverageArray(ByVal DataArray() As Double, Optional Length As Long = -1) As Double

        Dim ThisArray() As Double

        If Length <> -1 Then
            If DataArray.Length > Length - 1 Then
                ReDim Preserve ThisArray(0 To Length - 1)

            End If
        Else
            ThisArray = DataArray
        End If

        AverageArray = SumArray(ThisArray)
        AverageArray = AverageArray / ThisArray.Length

        Return AverageArray

    End Function

    Public Function MaxArray(ByVal DataArray() As Double) As Double

        MaxArray = DataArray(0)
        For i As Integer = 1 To DataArray.Length - 1
            MaxArray = Math.Max(MaxArray, DataArray(i))
        Next i

        Return MaxArray

    End Function

    Public Function MinArray(ByVal DataArray() As Double) As Double

        MinArray = DataArray(0)
        For i As Integer = 1 To DataArray.Length - 1
            MinArray = Math.Min(MinArray, DataArray(i))
        Next i

        Return MinArray

    End Function

    Public Function StandardDevArray(ByVal DataArray() As Double) As Double

        Dim Mean As Double = DataArray.Average()
        Dim sumDeviation As Double = 0
        Dim dataSize As Integer = DataArray.Length

        For i As Integer = 0 To dataSize - 1
            Mean += DataArray(i)
        Next

        Mean = Mean / dataSize

        For i As Integer = 0 To dataSize - 1
            sumDeviation += (DataArray(i) - Mean) * (DataArray(i) - Mean)
        Next

        Return Math.Sqrt(sumDeviation / dataSize)

    End Function

End Class
