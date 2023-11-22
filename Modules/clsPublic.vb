Public Class clsPublic

    Public Function Delay(ByVal Milliseconds As Integer)

        'Milliseconds - delay process time by parameter in milliseconds
        'This Function delays the process by Milliseconds based from parameter input

        Dim SW2 As New Stopwatch

        SW2.Start()
        Do
        Loop Until SW2.ElapsedMilliseconds >= Milliseconds

    End Function

    Public Function CheckTime(ByVal T1 As Date, ByVal T2 As Date) As Boolean

        'T1 - begin time
        'T2 - end time
        'This function checks the current time if it is greater than T1 and less than T2 and returns a boolean whether this condition is met
        'Ex: If CheckTime("9:00:00 AM", "12:01:00 PM") = True Then

        CheckTime = False

        Dim CurrentTime As DateTime = Convert.ToDateTime(DateTime.Now)

        If CurrentTime.TimeOfDay.Ticks >= T1.Ticks And CurrentTime.TimeOfDay.Ticks <= T2.Ticks Then
            CheckTime = True
        End If

        Return CheckTime

    End Function

    Public Function CharCount(ByVal InputString As String, ByVal CharToCount As String) As Long

        Dim TempArr() As String

        CharCount = 0

        If InputString.Contains(CharToCount) Then
            TempArr = Split(InputString, CharToCount)
            CharCount = TempArr.Length - 1
        End If

    End Function

End Class
