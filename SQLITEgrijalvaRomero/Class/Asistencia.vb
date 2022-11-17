Imports Finisar.SQLite

Public Class Asistencia

    Dim Id As Integer
    Public Fecha As DateTime
    Public Llegada As DateTime
    Public Salida As DateTime
    Public ClvEmpleado As Integer

    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"

    Public Sub Inicializar(
        _Fecha As DateTime,
        _Llegada As DateTime,
        _Salida As DateTime,
        _ClvEmpleado As Integer)

        Fecha = _Fecha
        Llegada = _Llegada
        Salida = _Salida
        ClvEmpleado = _ClvEmpleado
    End Sub

    Public Function InsertarAsitencia() As Boolean
        Dim AsistenciaBuscada = New Asistencia()

        Dim ComandText As String
        AsistenciaBuscada = BuscarAsistencia(ClvEmpleado, Fecha)

        If AsistenciaBuscada.ClvEmpleado = 0 Then
            ComandText = "INSERT INTO Asistencia (Fecha, Llegada, ClvEmpleado) VALUES ("
            ComandText += "'" + Date.Today() + "',"
            ComandText += "'" + Format(DateTime.Now(), "hh:mm:ss") + "',"
            ComandText += ClvEmpleado.ToString()
            ComandText += ");"
            Dim Query As New SQLiteCommand(ComandText, DB.objConn)
            Query.ExecuteNonQuery()

            Return True
        End If

        Return False
    End Function

    Public Shared Function BorrarAsistencia(_ClvEmpleado As Integer, _Fecha As String) As Boolean

        Dim ComandText As String

        ComandText = "DELETE FROM Asistencia WHERE ClvEmpleado = " + _ClvEmpleado.ToString + " AND Fecha = '" + _Fecha + "'"
        Dim Query As New SQLiteCommand(ComandText, DB.objConn)

        Query.ExecuteNonQuery()

        Return False
    End Function

    Public Function InsertarSalida() As Boolean

        Dim ComandText As String

        ComandText = "update Asistencia set Salida = '" + Format(DateTime.Now(), "hh:mm:ss") + "' WHERE ClvEmpleado = 2042777" 'falta fecha
        Dim Query As New SQLiteCommand(ComandText, DB.objConn)

        Query.ExecuteNonQuery()
    End Function

    Public Shared Function BuscarAsistencia(_ClvEmpleado As String, _Fecha As String) As Asistencia
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim AsistenciaTMp As Asistencia


        objcommand = DB.objConn.CreateCommand()
        objcommand.CommandText = "SELECT * FROM Asistencia 
            WHERE ClvEmpleado = " + _ClvEmpleado.ToString() + " AND Fecha = '" + _Fecha + "'"
        objReader = objcommand.ExecuteReader()
        AsistenciaTMp = New Asistencia()

        If (Not objReader.Read()) Then
            Return AsistenciaTMp
        End If

        AsistenciaTMp.Id = objReader("Id")
        AsistenciaTMp.ClvEmpleado = objReader("ClvEmpleado")
        AsistenciaTMp.Fecha = Convert.ToDateTime(objReader("Fecha"))
        AsistenciaTMp.Llegada = Convert.ToDateTime(objReader("Llegada"))

        If (objReader("Salida") <> "") Then
            AsistenciaTMp.Salida = Convert.ToDateTime(objReader("Salida"))
        End If

        Return AsistenciaTMp

    End Function

    Public Shared Function BuscarPorClv(_ClvEmpleado As String) As Asistencia

        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim AsistenciaTMp As Asistencia


        objcommand = DB.objConn.CreateCommand()
        objcommand.CommandText = "SELECT * FROM Asistencia WHERE ClvEmpleado = " + _ClvEmpleado.ToString()
        objReader = objcommand.ExecuteReader()
        AsistenciaTMp = New Asistencia()

        If (Not objReader.Read()) Then
            Return AsistenciaTMp
        End If

        AsistenciaTMp.Id = objReader("Id")
        AsistenciaTMp.ClvEmpleado = objReader("ClvEmpleado")
        AsistenciaTMp.Fecha = Convert.ToDateTime(objReader("Fecha"))
        AsistenciaTMp.Llegada = Convert.ToDateTime(objReader("Llegada"))

        If (objReader("Salida") <> "") Then
            AsistenciaTMp.Salida = Convert.ToDateTime(objReader("Salida"))
        End If

        Return AsistenciaTMp

    End Function

End Class
