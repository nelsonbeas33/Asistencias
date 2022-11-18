Public Class transform

    Public Shared Function ToDate(_date As String) As DateTime

        Dim format() = {"MM/dd/yyyy", "dd/MM/yyyy"}

        Dim expenddt As DateTime
        DateTime.TryParseExact(_date, format,
            System.Globalization.DateTimeFormatInfo.InvariantInfo,
            Globalization.DateTimeStyles.None, expenddt)

        Return expenddt
    End Function

End Class
