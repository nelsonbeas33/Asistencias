Imports Finisar.SQLite
Imports SQLITEgrijalvaRomero.Empleado

Public Class getNames
    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim objConn As SQLiteConnection
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader
        Dim EmpleadoBuscado As Empleado

        Try
            objConn = New SQLiteConnection(cadena_conexion)
            objConn.Open()
            objcommand = objConn.CreateCommand()
            objcommand.CommandText = "SELECT * from Empleado"
            objReader = objcommand.ExecuteReader()

            ListBox1.Items.Clear()
            While (objReader.Read())
            ListBox1.Items.Add(objReader("Nombre"))
            End While

            EmpleadoBuscado = Empleado.BuscarEmpleadoClv("1922777")
            Label1.Text = EmpleadoBuscado.Nombre

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            If Not IsNothing(objConn) Then
                objConn.Close()
            End If
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim NuevoEmpleado As Empleado


        NuevoEmpleado = New Empleado
        NuevoEmpleado.Inicializar(
            0,
            1942777,
            "Nombre1",
            "apellido1",
            "Apellido1.1",
            "direccion 1",
            "123",
            "123",
            "123",
            "8117201353",
            40
        )

        NuevoEmpleado.InsertarEmpleado()


    End Sub
End Class
