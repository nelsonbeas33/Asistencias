﻿Imports Finisar.SQLite

Public Class Asistencia

    Dim Id As Integer
    Public Fecha As Date
    Public Llegada As DateTime
    Public Salida As DateTime
    Public ClvEmpleado As Integer

    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"

    Public Sub Inicializar(
        _Fecha As Date,
        _Llegada As DateTime,
        _Salida As DateTime,
        _ClvEmpleado As Integer)

        Fecha = _Fecha
        Llegada = _Llegada
        Salida = _Salida
        ClvEmpleado = _ClvEmpleado
    End Sub

    Public Function InsertarAsitencia() As Boolean
        Dim objConn As SQLiteConnection
        objConn = New SQLiteConnection(cadena_conexion)
        Dim AsistenciaBuscada = New Asistencia()

        objConn.Open()
        Dim ComandText As String
        AsistenciaBuscada = BuscarAsistencia(ClvEmpleado)

        If AsistenciaBuscada.ClvEmpleado = 0 Then
            ComandText = "INSERT INTO Asistencia (Fecha, Llegada, ClvEmpleado) VALUES ("
            ComandText += "'" + Date.Today() + "',"
            ComandText += "'" + Format(DateTime.Now(), "hh:mm:ss") + "',"
            ComandText += ClvEmpleado.ToString()
            ComandText += ");"
            Dim Query As New SQLiteCommand(ComandText, objConn)
            Query.ExecuteNonQuery()

            objConn.Close()
            Return True
        End If

        objConn.Close()
        Return False
    End Function

    Public Function InsertarSalida() As Boolean
        Dim objConn As SQLiteConnection
        objConn = New SQLiteConnection(cadena_conexion)
        Dim ComandText As String
        objConn.Open()

        Dim AsistenciaEmpleado = BuscarAsistencia(ClvEmpleado)

        If AsistenciaEmpleado.ClvEmpleado <> 0 And Format(AsistenciaEmpleado.Salida, "d/m/yy") = "1/0/01" Then

            ComandText = "delete * FROM Asistencia WHERE ClvEmpleado = " + ClvEmpleado.ToString()
            Dim Query As New SQLiteCommand(ComandText, objConn)
            Query.ExecuteNonQuery()
            objConn.Close()
            Return True
        End If

        objConn.Close()
        Return False
    End Function

    Public Shared Function BuscarAsistencia(_ClvEmpleado As String) As Asistencia
        Dim objConn As SQLiteConnection
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim AsistenciaTMp As Asistencia

        objConn = New SQLiteConnection(cadena_conexion)
        objConn.Open()
        objcommand = objConn.CreateCommand()
        objcommand.CommandText = "SELECT * FROM Asistencia 
            WHERE ClvEmpleado = " + _ClvEmpleado.ToString() + ""
        objReader = objcommand.ExecuteReader()
        AsistenciaTMp = New Asistencia()

        If (Not objReader.Read()) Then
            Return AsistenciaTMp
        End If

        AsistenciaTMp.Id = objReader("Id")
        AsistenciaTMp.ClvEmpleado = objReader("ClvEmpleado")
        AsistenciaTMp.Fecha = objReader("Fecha")
        AsistenciaTMp.Llegada = objReader("Llegada")

        If (objReader("Salida") <> "") Then
            AsistenciaTMp.Salida = objReader("Salida")
        End If

        Return AsistenciaTMp

        objConn.Close()

    End Function

End Class