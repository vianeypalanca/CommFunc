Public Module ConversionFunctions

    Public Function ConvertNumSys(ByVal InputVal As String, ByVal FromRadix As Long, ByVal ToRadix As Long) As Long

        Select Case ToRadix
            Case 2
            Case 8
            Case 10
                Select Case FromRadix
                    Case 2
                        ConvertNumSys = BinToDec(InputVal)
                    Case 8
                        ConvertNumSys = OctToDec(InputVal)
                    Case 10
                        ConvertNumSys = CLng(InputVal)
                    Case 16
                        ConvertNumSys = HexToDec(InputVal)
                End Select
            Case 16
        End Select

    End Function

    Public Function BinToDec(ByVal InputVal As String) As Long

        Dim TempArr() As String = SeparateCharacters(InputVal)

        If TempArr.Length > 1 Then
            For x As Long = 0 To UBound(TempArr)
                If TempArr(x).Trim = vbNullString Or TempArr(x) = "0" Then
                    TempArr(x) = vbNullString
                Else
                    Exit For
                End If
            Next x
            InputVal = Join(TempArr, "").Trim
            If InputVal = vbNullString Then InputVal = "0"
        End If

        BinToDec = Convert.ToInt32(InputVal, 2)

    End Function

    Public Function OctToDec(ByVal InputVal As String) As Long

        Dim TempArr() As String = SeparateCharacters(InputVal)

        If TempArr.Length > 1 Then
            For x As Long = 0 To UBound(TempArr)
                If TempArr(x).Trim = vbNullString Or TempArr(x) = "0" Then
                    TempArr(x) = vbNullString
                Else
                    Exit For
                End If
            Next x
            InputVal = Join(TempArr, "").Trim
            If InputVal = vbNullString Then InputVal = "0"
        End If

        OctToDec = Convert.ToInt32(InputVal, 8)

    End Function

    Public Function HexToDec(ByVal InputVal As String) As Long

        Dim TempArr() As String = SeparateCharacters(InputVal)

        If TempArr.Length > 1 Then
            For x As Long = 0 To UBound(TempArr)
                If TempArr(x).Trim = vbNullString Or TempArr(x) = "0" Then
                    TempArr(x) = vbNullString
                Else
                    Exit For
                End If
            Next x
            InputVal = Join(TempArr, "").Trim
            If InputVal = vbNullString Then InputVal = "0"
        End If

        HexToDec = Convert.ToInt32(InputVal, 16)

    End Function

End Module
