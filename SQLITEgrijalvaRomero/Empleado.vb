Imports Finisar.SQLite

Public Class Empleado

    Dim Id As Integer
    Public ClvEmpleado As Integer
    Public Nombre As String
    Public ApellidoPaterno As String
    Public ApellidoMaterno As String
    Public Direccion As String
    Public Cargo As String
    Public Rfc As String
    Public Curp As String
    Public Telefono As String
    Public HorasSemana As Integer

    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"

    Public Sub Inicializar(
        _Id As Integer,
        _ClvEmpleado As Integer,
        _Nombre As String,
        _ApellidoPaterno As String,
        _ApellidoMaterno As String,
        _Direccion As String,
        _Cargo As String,
        _Rfc As String,
        _Curp As String,
        _Telefono As String,
        _HorasSemana As Integer)

        Id = _Id
        ClvEmpleado = _ClvEmpleado
        Nombre = _Nombre
        ApellidoPaterno = _ApellidoPaterno
        ApellidoMaterno = _ApellidoMaterno
        Direccion = _Direccion
        Cargo = _Cargo
        Rfc = _Rfc
        Curp = _Curp
        Telefono = _Telefono
        HorasSemana = _HorasSemana

    End Sub


    Public Shared Function BuscarEmpleadoClv(_ClvEmpleado As String) As Empleado
        Dim objConn As SQLiteConnection
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim EmpleadoBuscado As Empleado

        Try
            objConn = New SQLiteConnection(cadena_conexion)
            objConn.Open()
            objcommand = objConn.CreateCommand()
            objcommand.CommandText = "SELECT * from Empleado WHERE ClvEmpleado = " + _ClvEmpleado
            objReader = objcommand.ExecuteReader()
            EmpleadoBuscado = New Empleado()

            If (Not objReader.Read()) Then
                Return EmpleadoBuscado
            End If

            EmpleadoBuscado.Nombre = objReader("Nombre")
            Return EmpleadoBuscado

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If Not IsNothing(objConn) Then
                objConn.Close()
            End If
        End Try

    End Function

    Public Function InsertarEmpleado() As Boolean
        Dim objConn As SQLiteConnection
        objConn = New SQLiteConnection(cadena_conexion)

        objConn.Open()

        Dim ComandText As String

        ComandText = "insert into Empleado (
            ClvEmpleado, 
            Nombre, 
            ApellidoPaterno,
            ApellidoMaterno, 
            Direccion, Cargo, Rfc, Curp, Telefono, HorasSemana) VALUES ("

        ComandText += ClvEmpleado.ToString()
        ComandText += ",'"
        ComandText += Nombre
        ComandText += "','"
        ComandText += ApellidoPaterno
        ComandText += "','"
        ComandText += ApellidoMaterno
        ComandText += "','"
        ComandText += Direccion
        ComandText += "','"
        ComandText += Cargo
        ComandText += "','"
        ComandText += Rfc
        ComandText += "','"
        ComandText += Curp
        ComandText += "','"
        ComandText += Telefono
        ComandText += "','"
        ComandText += HorasSemana.ToString()
        ComandText += "');"
        Dim Query As New SQLiteCommand(ComandText, objConn)
        Query.ExecuteNonQuery()

        objConn.Close()
        Return True
    End Function

End Class
