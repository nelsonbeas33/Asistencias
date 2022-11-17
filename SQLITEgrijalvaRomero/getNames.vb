Imports Finisar.SQLite
Imports SQLITEgrijalvaRomero

Public Class getNames
    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim objcommand As SQLiteCommand
        Dim objReader As SQLiteDataReader

        Try

            objcommand = DB.objConn.CreateCommand()
            objcommand.CommandText = "SELECT * from Empleado"
            objReader = objcommand.ExecuteReader()

            ListBox1.Items.Clear()
            While (objReader.Read())
            ListBox1.Items.Add(objReader("Nombre"))
            End While

        Catch ex As Exception
            MsgBox(ex.ToString)

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
            .ClvEmpleado = 2042777,
            .Fecha = Format(DateAndTime.Now(), "dd/MM/yyyy")
        }

        If AsistenciaEmpleado.InsertarAsitencia() Then
            Label1.Text = "insertada con exito"
        Else
            Label1.Text = "no insertada"
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim AsistenciaEmpleado = New Asistencia With {
            .ClvEmpleado = 2042777,
            .Fecha = Format(DateAndTime.Now(), "dd/MM/yyyy")
        }

        Dim resultado As Boolean = AsistenciaEmpleado.InsertarSalida()
        If resultado Then
            Label1.Text = "salida con exito"
        Else
            Label1.Text = "no salida"
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim AsistenciaEmpleado = New Asistencia With {
            .ClvEmpleado = 2042777,
            .Fecha = Format(DateAndTime.Now(), "dd/MM/yyyy")
        }

        Asistencia.BorrarAsistencia(AsistenciaEmpleado.ClvEmpleado, AsistenciaEmpleado.Fecha)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim asistenciatmp As Asistencia = Asistencia.BuscarAsistencia(2042777, "17/11/2022")

        If asistenciatmp.ClvEmpleado = 0 Then
            Label1.Text = "no existe"
        Else
            Label1.Text = "existe"
        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim Empleadotmp As Empleado = New Empleado()
        Empleadotmp.BorrarEmpleado(2042777)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim Asistemciatmp As Asistencia = Asistencia.BuscarPorClv(2042777)
        Label1.Text = Asistemciatmp.Fecha.ToShortDateString()
    End Sub
End Class
