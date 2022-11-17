Imports Finisar.SQLite
Imports SQLITEgrijalvaRomero

Public Class getNames
    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim objConn As SQLiteConnection
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader

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
        Dim EmpleadoBuscado As Empleado


        NuevoEmpleado = New Empleado
        NuevoEmpleado.Inicializar(
            0,
            2042777,
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


        EmpleadoBuscado = Empleado.BuscarEmpleadoClv(2042777)

        If EmpleadoBuscado.ClvEmpleado = 0 Then
            NuevoEmpleado.InsertarEmpleado()
            Label1.Text = NuevoEmpleado.Nombre
        Else
            Label1.Text = "clv ya existente"
        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim NuevoEmpleado As Empleado
        Dim EmpleadoBuscado As Empleado


        NuevoEmpleado = New Empleado
        NuevoEmpleado.Inicializar(
            0,
            2042777,
            "Nombre2",
            "apellido1",
            "Apellido1.1",
            "direccion 1",
            "123",
            "123",
            "123",
            "8117201353",
            40
        )


        EmpleadoBuscado = Empleado.BuscarEmpleadoClv(2042777)
        If EmpleadoBuscado.ClvEmpleado = 0 Then
            Label1.Text = "Empleado no existe"
        Else
            NuevoEmpleado.EditarEmpleado()
            Label1.Text = NuevoEmpleado.Nombre
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim AsistenciaEmpleado = New Asistencia With {
            .ClvEmpleado = 2042777
        }

        If AsistenciaEmpleado.InsertarAsitencia() Then
            Label1.Text = "insertada con exito"
        Else
            Label1.Text = "no insertada"
        End If

        Label1.Text = AsistenciaEmpleado.Salida

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim AsistenciaEmpleado = New Asistencia With {
            .ClvEmpleado = 2042777
        }

        If AsistenciaEmpleado.InsertarSalida() Then
            Label1.Text = "salida con exito"
        Else
            Label1.Text = "no salida"
        End If

    End Sub
End Class
