Imports System.Reflection.Emit
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
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim EmpleadoBuscado As Empleado

        objcommand = DB.objConn.CreateCommand()
        objcommand.CommandText = "SELECT * FROM Empleado WHERE ClvEmpleado = " + _ClvEmpleado
        objReader = objcommand.ExecuteReader()
        EmpleadoBuscado = New Empleado()

        If (Not objReader.Read()) Then
            Return EmpleadoBuscado
        End If

        EmpleadoBuscado.Id = objReader("Id")
        EmpleadoBuscado.ClvEmpleado = objReader("ClvEmpleado")
        EmpleadoBuscado.Nombre = objReader("Nombre")
        EmpleadoBuscado.ApellidoPaterno = objReader("ApellidoPaterno")
        EmpleadoBuscado.ApellidoMaterno = objReader("ApellidoMaterno")
        EmpleadoBuscado.Direccion = objReader("Direccion")
        EmpleadoBuscado.Cargo = objReader("Cargo")
        EmpleadoBuscado.Rfc = objReader("Rfc")
        EmpleadoBuscado.Curp = objReader("Curp")
        EmpleadoBuscado.Telefono = objReader("Telefono")
        EmpleadoBuscado.HorasSemana = objReader("HorasSemana")

        Return EmpleadoBuscado

    End Function

    Public Shared Function GetEmpleados() As List(Of Empleado)

        Dim Empleadotmp As Empleado
        Dim Empleados As List(Of Empleado) = New List(Of Empleado)()

        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader

        objcommand = DB.objConn.CreateCommand()
        objcommand.CommandText = "SELECT * FROM Empleado"
        objReader = objcommand.ExecuteReader()

        While (objReader.Read())

            Empleadotmp = New Empleado()

            Empleadotmp.Id = objReader("Id")
            Empleadotmp.ClvEmpleado = objReader("ClvEmpleado")
            Empleadotmp.Nombre = objReader("Nombre")
            Empleadotmp.ApellidoPaterno = objReader("ApellidoPaterno")
            Empleadotmp.ApellidoMaterno = objReader("ApellidoMaterno")
            Empleadotmp.Direccion = objReader("Direccion")
            Empleadotmp.Cargo = objReader("Cargo")
            Empleadotmp.Rfc = objReader("Rfc")
            Empleadotmp.Curp = objReader("Curp")
            Empleadotmp.Telefono = objReader("Telefono")
            Empleadotmp.HorasSemana = objReader("HorasSemana")

            Empleados.Add(Empleadotmp)
        End While

        Return Empleados

    End Function

    Public Function InsertarEmpleado() As Boolean

        Dim EmpleadoBuscado As Empleado = Empleado.BuscarEmpleadoClv(ClvEmpleado)

        If EmpleadoBuscado.ClvEmpleado = 0 Then
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
            Dim Query As New SQLiteCommand(ComandText, DB.objConn)
            DB.WriteTextFile(ComandText)
            Query.ExecuteNonQuery()
            Return True
        Else
            Return False
        End If


    End Function

    Public Function EditarEmpleado() As Boolean

        Dim ComandText As String

        ComandText = "DELETE FROM Empleado WHERE ClvEmpleado = " + ClvEmpleado.ToString()
        Dim QueryDelete As New SQLiteCommand(ComandText, DB.objConn)
        QueryDelete.ExecuteNonQuery()

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

        Dim QueryInsert As New SQLiteCommand(ComandText, DB.objConn)
        QueryInsert.ExecuteNonQuery()

        Return True
    End Function

    Public Function BorrarEmpleado(_ClveEmpleado As Integer) As Boolean
        Dim ComandText As String

        ComandText = "DELETE FROM Empleado WHERE ClvEmpleado = " + _ClveEmpleado.ToString
        Dim Query As New SQLiteCommand(ComandText, DB.objConn)
        Query.ExecuteNonQuery()

        Return True
    End Function

End Class
