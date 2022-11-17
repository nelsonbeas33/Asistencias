Imports Finisar.SQLite

Public Class DB

    Public Const cadenaConnection As String = "Data Source=asistencias.db;Version=3;"
    Public Shared objConn As SQLiteConnection = New SQLiteConnection(cadenaConnection)

    Shared Sub New()
        objConn.Open()
    End Sub

End Class
