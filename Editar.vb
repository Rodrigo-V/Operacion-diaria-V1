
Imports System.Data
Imports System.Data.SqlClient

Public Class Editar

    Dim guardar As New SqlCommand
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")


    Private Sub Editar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
        DateTimePicker2.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
        DateTimePicker3.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or DateTimePicker1.Text = "" Or DateTimePicker2.Text = "" Or DateTimePicker3.Text = "" Or TextBox8.Text = "" Or TextBox7.Text = "" Then


            MsgBox("Faltan datos, revise el formulario y vuelva a interntarlo", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)


        Else
            guardar.Parameters.Clear()
            guardar.CommandType = CommandType.StoredProcedure
            guardar.CommandText = "modificar"
            guardar.Parameters.Add("@ID_REGISTRO", SqlDbType.Int).Value = (TextBox1.Text)
            guardar.Parameters.Add("@FECHA", SqlDbType.Date).Value = (TextBox2.Text)
            guardar.Parameters.Add("@TIPO_DIA", SqlDbType.NChar).Value = (TextBox3.Text)
            guardar.Parameters.Add("@HORA", SqlDbType.Time).Value = (DateTimePicker1.Text)
            guardar.Parameters.Add("@EMPRESA", SqlDbType.NVarChar).Value = (TextBox5.Text)
            guardar.Parameters.Add("@SERVICIO", SqlDbType.NVarChar).Value = (TextBox6.Text)
            guardar.Parameters.Add("@PATENTE", SqlDbType.NVarChar).Value = (TextBox11.Text)
            guardar.Parameters.Add("@HORA_INI", SqlDbType.Time).Value = (DateTimePicker2.Text)
            guardar.Parameters.Add("@HORA_TER", SqlDbType.Time).Value = (DateTimePicker3.Text)
            guardar.Parameters.Add("@CLASIFICACION", SqlDbType.NVarChar).Value = (TextBox8.Text)
            guardar.Parameters.Add("@DETALLE_CMB", SqlDbType.NVarChar).Value = (TextBox7.Text)
            guardar.Parameters.Add("@OBSERVACIONES", SqlDbType.NVarChar).Value = (TextBox12.Text)


            guardar.Connection = (conexion)
            conexion.Open()
            guardar.ExecuteNonQuery()

            MsgBox("Registro modificado con exito", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
            conexion.Close()
        End If
        Me.Close()
        MDInicio.Show()
    End Sub
End Class