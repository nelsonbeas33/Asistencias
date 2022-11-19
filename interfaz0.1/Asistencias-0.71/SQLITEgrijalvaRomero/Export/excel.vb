Public Class excel

    Public Shared Function WriteEmpleadosCSV(path As String, data As List(Of Empleado), delimitador As String) As Boolean

        Try
            My.Computer.FileSystem.DeleteFile(path)
        Catch ex As Exception

        End Try
        Dim fileWriter As New System.IO.StreamWriter(path, True)

        Dim ColumName As String = ""
        ColumName += "ClvEmpleado"
        ColumName += delimitador
        ColumName += "Nombre"
        ColumName += delimitador
        ColumName += "ApellidoPaterno"
        ColumName += delimitador
        ColumName += "ApellidoMaterno"
        ColumName += delimitador
        ColumName += "Direccion"
        ColumName += delimitador
        ColumName += "Cargo"
        ColumName += delimitador
        ColumName += "Rfc"
        ColumName += delimitador
        ColumName += "Curp"
        ColumName += delimitador
        ColumName += "Telefono"
        ColumName += delimitador
        ColumName += "HorasSemana"
        fileWriter.WriteLine(ColumName)

        For Each empleado As Empleado In data
            Dim record As String = ""
            record += empleado.ClvEmpleado.ToString()
            record += delimitador
            record += empleado.Nombre
            record += delimitador
            record += empleado.ApellidoPaterno
            record += delimitador
            record += empleado.ApellidoMaterno
            record += delimitador
            record += empleado.Direccion
            record += delimitador
            record += empleado.Cargo
            record += delimitador
            record += empleado.Rfc
            record += delimitador
            record += empleado.Curp
            record += delimitador
            record += empleado.Telefono
            record += delimitador
            record += empleado.HorasSemana.ToString()
            fileWriter.WriteLine(record)
        Next empleado

        fileWriter.Close()
        Return True
    End Function


    Public Shared Function WriteAsistenciasCSV(path As String, data As List(Of Asistencia), delimitador As String) As Boolean

        Try
            My.Computer.FileSystem.DeleteFile(path)
        Catch ex As Exception

        End Try
        Dim fileWriter As New System.IO.StreamWriter(path, True)

        Dim ColumName As String = ""
        ColumName += "Fecha"
        ColumName += delimitador
        ColumName += "Llegada"
        ColumName += delimitador
        ColumName += "Salida"
        ColumName += delimitador
        ColumName += "ClvEmpleado"
        ColumName += delimitador
        fileWriter.WriteLine(ColumName)

        For Each asistencia As Asistencia In data
            Dim record As String = ""
            record += asistencia.Fecha.ToShortDateString()
            record += delimitador
            record += Format(asistencia.Llegada, "hh:mm:ss")
            record += delimitador
            record += Format(asistencia.Salida, "hh:mm:ss")
            record += delimitador
            record += asistencia.ClvEmpleado.ToString()
            record += delimitador
            fileWriter.WriteLine(record)
        Next asistencia

        fileWriter.Close()
        Return True
    End Function

End Class
