Imports Finisar.SQLite
Imports SQLITEgrijalvaRomero

Public Class getNames
    Const cadena_conexion As String = "Data Source=asistencias.db;Version=3;"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Empleados As List(Of Empleado) = Empleado.GetEmpleados()

        ListBox1.Items.Clear()
        For Each tmpEmpleado As Empleado In Empleados
            ListBox1.Items.Add(tmpEmpleado.Nombre)
        Next tmpEmpleado

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim NuevoEmpleado As Empleado


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

        If NuevoEmpleado.InsertarEmpleado() Then
            Label1.Text = "insertado"
        Else
            Label1.Text = "no insertado"
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
            .Fecha = DateAndTime.Now()
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
            .Fecha = DateAndTime.Now()
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
            .Fecha = DateAndTime.Now()
        }

        Asistencia.BorrarAsistencia(AsistenciaEmpleado.ClvEmpleado, Format(AsistenciaEmpleado.Fecha, "MM/dd/yyyy"))
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Dim asistenciatmp As Asistencia = Asistencia.BuscarAsistencia(2042777, "11/17/2022")

        If asistenciatmp.ClvEmpleado = 0 Then
            Label1.Text = "no existe"
            ListBox1.Items.Clear()
        Else
            Label1.Text = "existe"
            ListBox1.Items.Clear()

            Dim texto As String
            texto = asistenciatmp.ClvEmpleado.ToString()
            texto += " - " + asistenciatmp.Fecha.ToShortDateString()
            texto += " - " + asistenciatmp.Llegada.TimeOfDay.ToString()

            If Format(asistenciatmp.Salida, "MM/dd/yyyy") <> "01/01/0001" Then
                texto += " - " + asistenciatmp.Salida.TimeOfDay.ToString()
            End If

            ListBox1.Items.Add(texto)
        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim Empleadotmp As Empleado = New Empleado()
        Empleadotmp.BorrarEmpleado(2042777)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Dim listaAsistencias As List(Of Asistencia) = Asistencia.BuscarPorClv(2042777)

        ListBox1.Items.Clear()
        For Each Asistenciatmp As Asistencia In listaAsistencias

            Dim texto As String
            texto = Asistenciatmp.ClvEmpleado.ToString()
            texto += " - " + Asistenciatmp.Fecha.ToShortDateString()
            texto += " - " + Asistenciatmp.Llegada.TimeOfDay.ToString()

            If Format(Asistenciatmp.Salida, "MM/dd/yyyy") <> "01/01/0001" Then
                texto += " - " + Asistenciatmp.Salida.TimeOfDay.ToString()
            End If

            ListBox1.Items.Add(texto)
        Next Asistenciatmp

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim listaAsistencias As List(Of Asistencia) = Asistencia.BuscarPorFecha("12/17/2022")

        ListBox1.Items.Clear()

        For Each Asistenciatmp As Asistencia In listaAsistencias

            Dim texto As String
            texto = Asistenciatmp.ClvEmpleado.ToString()
            texto += " - " + Asistenciatmp.Fecha.ToShortDateString()
            texto += " - " + Asistenciatmp.Llegada.TimeOfDay.ToString()

            If Format(Asistenciatmp.Salida, "MM/dd/yyyy") <> "01/01/0001" Then
                texto += " - " + Asistenciatmp.Salida.TimeOfDay.ToString()
            End If

            ListBox1.Items.Add(texto)
        Next Asistenciatmp
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim listaAsistencias As List(Of Asistencia) = Asistencia.BuscarPorRangoFecha("11/03/2022", "11/04/2022")
        ListBox1.Items.Clear()

        For Each Asistenciatmp As Asistencia In listaAsistencias

            Dim texto As String
            texto = Asistenciatmp.ClvEmpleado.ToString()
            texto += " - " + Asistenciatmp.Fecha.ToShortDateString()
            texto += " - " + Asistenciatmp.Llegada.TimeOfDay.ToString()

            If Format(Asistenciatmp.Salida, "MM/dd/yyyy") <> "01/01/0001" Then
                texto += " - " + Asistenciatmp.Salida.TimeOfDay.ToString()
            End If

            ListBox1.Items.Add(texto)
        Next Asistenciatmp
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        FolderBrowserDialog1.ShowDialog()
        Dim path As String = FolderBrowserDialog1.SelectedPath() + "\empleados.csv"
        excel.WriteEmpleadosCSV(path, Empleado.GetEmpleados(), ",")
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        FolderBrowserDialog1.ShowDialog()
        Dim path As String = FolderBrowserDialog1.SelectedPath() + "\asistencias.csv"
        excel.WriteAsistenciasCSV(path, Asistencia.GetAsistencias(), ",")
    End Sub
End Class
