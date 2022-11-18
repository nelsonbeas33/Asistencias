Imports Finisar.SQLite

Public Class DB

    Public Const cadenaConnection As String = "Data Source=asistencias.db;Version=3;"
    Public Shared objConn As SQLiteConnection = New SQLiteConnection(cadenaConnection)

    Shared Sub New()
        objConn.Open()
    End Sub

    Public Shared Sub refresh()
        objConn.Close()
        objConn.Open()
    End Sub

    Public Shared Sub WriteTextFile(query As String)
        Dim textFileStream As New IO.FileStream("queryTMP.txt", IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite, IO.FileShare.None)

        Dim myFileWriter As New IO.StreamWriter(textFileStream)

        myFileWriter.WriteLine(query)

        myFileWriter.Close()
        textFileStream.Close()
    End Sub

End Class
