Public Module StringFunctions

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

    Public Function RemoveRight(ByVal InputString As String, ByVal SearchString As String, Optional IsLastIndex As Boolean = False) As String

        Dim TempIndex As Long

        InputString = InputString.Replace("" & vbTab, "")

        InputString = Trim(InputString)

        If IsLastIndex = False Then
            TempIndex = InStr(UCase(InputString), UCase(SearchString))
        Else
            TempIndex = InStrRev(UCase(InputString), UCase(SearchString))
        End If
        If TempIndex <> 0 Then
            RemoveRight = Trim(Right(InputString, (Len(InputString)) - (TempIndex + (Len(SearchString) - 1))))
        Else
            RemoveRight = Trim(InputString)
        End If

    End Function

    Public Function RemoveLeft(ByVal InputString As String, ByVal SearchString As String, Optional IsLastIndex As Boolean = False) As String

        Dim TempIndex As Long

        InputString = InputString.Replace("" & vbTab, "")

        InputString = Trim(InputString)

        If IsLastIndex = False Then
            TempIndex = InStr(UCase(InputString), UCase(SearchString))
        Else
            TempIndex = InStrRev(UCase(InputString), UCase(SearchString))
        End If
        If TempIndex <> 0 Then
            RemoveLeft = Trim(Left(InputString, TempIndex - 1))
        Else
            RemoveLeft = Trim(InputString)
        End If

    End Function

    Public Function IsInString(ByVal RefString As String, ByVal SearchVal As String) As Boolean

        If InStr(Trim(UCase(RefString)), Trim(UCase(SearchVal))) > 0 Then
            IsInString = True
        Else
            IsInString = False
        End If

    End Function

    Public Function IsInString_MultiSearch(ByVal RefString As String, ByVal SearchVal As String) As Boolean

        Dim SearchCounter As Long
        Dim SearchArr() As String
        Dim retVal As Boolean

        SearchArr = Split(SearchVal, "|")

        For SearchCounter = 0 To UBound(SearchArr)
            If InStr(Trim(UCase(RefString)), Trim(UCase(SearchArr(SearchCounter)))) > 0 Then
                retVal = True
                Exit For
            Else
                retVal = False
            End If

        Next SearchCounter

        IsInString_MultiSearch = retVal
    End Function

    Public Function IsFirstCharNum(ByVal InputString As String) As Boolean

        IsFirstCharNum = False

        If Len(InputString) > 1 Then

            IsFirstCharNum = IsNumeric(Left(InputString, 1))

        End If

    End Function

    Public Function EvaluateFormula(ByVal Formula As String) As Double

        Dim ReturnVal As Double

        Dim ThisTable As New DataTable()
        Dim ThisObj As Object = ThisTable.Compute(Formula, "")
        ReturnVal = Convert.ToDouble(ThisObj)

        EvaluateFormula = ReturnVal

    End Function

End Module
