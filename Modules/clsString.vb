Public Class clsString

    Public Function JoinDeletedSpaces(ByVal InputData() As String, Optional Delimeter As String = " ") As String

        JoinDeletedSpaces = vbNullString

        For x As Long = 0 To UBound(InputData)
            If InputData(x).Trim <> vbNullString Then
                If JoinDeletedSpaces = vbNullString Then
                    JoinDeletedSpaces = InputData(x).Trim
                Else
                    JoinDeletedSpaces = JoinDeletedSpaces & Delimeter & InputData(x).Trim
                End If
            End If
        Next x

    End Function

    Public Function SeparateUnits(ByVal InputString As String, ByRef NumVal As String, ByRef UnitVal As String)

        Dim CharArr() As Char
        Dim TempIndex As Long

        If InputString = vbNullString Then Exit Function

        CharArr = InputString.ToCharArray
        For x As Long = UBound(CharArr) To 0 Step -1
            If IsNumeric(CharArr(x)) Then
                TempIndex = x
                Exit For
            End If
        Next x

        If CharArr.Length = 1 Or TempIndex = CharArr.Length - 1 Then
            NumVal = InputString
            Exit Function
        End If

        NumVal = InputString.Remove(TempIndex + 1)
        UnitVal = InputString.Remove(0, TempIndex + 1)

    End Function

    Public Function SeparateCharNum(ByVal InputString As String, ByRef CharPart As String, ByRef NumPart As String, Optional StringFirst As Boolean = True)

        Dim CharArr() As Char
        Dim TempIndex As Long

        CharArr = InputString.ToCharArray

        If StringFirst = True Then
            For x As Long = 0 To UBound(CharArr)
                If IsNumeric(CharArr(x)) Then
                    TempIndex = x
                    Exit For
                End If
            Next x
            CharPart = InputString.Remove(TempIndex)
            NumPart = InputString.Remove(0, TempIndex)
        Else
            For x As Long = UBound(CharArr) To 0 Step -1
                If IsNumeric(CharArr(x)) Then
                    TempIndex = x
                    Exit For
                End If
            Next x

            NumPart = InputString.Remove(TempIndex + 1)
            CharPart = InputString.Remove(0, TempIndex + 1)
        End If

    End Function

    Public Function SplitJoin(ByVal InputString As String, Optional Delimeter As String = " ") As String

        Dim TempArr() As String

        TempArr = Split(InputString, Delimeter)
        If TempArr.Length = 1 Then
            SplitJoin = InputString
            Exit Function
        End If

        For x As Long = 0 To UBound(TempArr)
            If x = 0 Then
                If TempArr(x).Trim <> vbNullString Then
                    SplitJoin = TempArr(x).Trim
                End If
            Else
                If TempArr(x).Trim <> vbNullString Then
                    If SplitJoin <> vbNullString Then
                        SplitJoin = SplitJoin & "_" & TempArr(x).Trim
                    Else
                        SplitJoin = TempArr(x).Trim
                    End If
                End If
            End If
        Next x

    End Function

End Class
