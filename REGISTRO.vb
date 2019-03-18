Imports System.Data
Imports System.Data.SqlClient


Public Class REGISTRO
    Public Property mask As String
    Dim tipodia As String
    Dim comando2 As New SqlCommand
    Dim guardar As New SqlCommand
    Dim clasif As New SqlCommand
    Dim detalle As New SqlCommand
    Dim empre As New SqlCommand
    Dim serv As New SqlCommand
    Dim edi As New SqlCommand
    Dim variable As SqlDataReader
    Dim variable5 As SqlDataReader
    Dim variable3 As SqlDataReader
    Dim variable4 As SqlDataReader
    Dim variable2 As SqlDataReader
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")


    Private Sub REGISTRO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'muestra fecha en el form
        TextBox1.Text = DateValue(Now)
        TextBox1.Enabled = False

        'formato inicial de horas 

        DateTimePicker4.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)

        'seleccion de toda la fila en el datagrid
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Me.ComboBox1.Text = "Seleccione una clasificacion"
        'llenar combo con nombres de clasificacion
        clasif.CommandType = CommandType.Text
        clasif.CommandText = ("select * from CLASIFICACION")
        clasif.Connection = (conexion)
        conexion.Open()
        variable = clasif.ExecuteReader()

        'ciclo que llena el combobox 
        While variable.Read = True
            ComboBox1.Items.Add(variable(1))

        End While
        conexion.Close()


        'llenar combo con nombres de empresa
        empre.CommandType = CommandType.Text
        empre.CommandText = ("select * from EMPRESA1")
        empre.Connection = (conexion)
        conexion.Open()
        variable3 = empre.ExecuteReader()

        'ciclo que llena el combobox 
        While variable3.Read = True
            ComboBox4.Items.Add(variable3(1))

        End While
        conexion.Close()

        armardatatable()
        llenargrilla()

        DataGridView1.Columns(0).Width = 35
        DataGridView1.Columns(1).Width = 75
        DataGridView1.Columns(2).Width = 45
        DataGridView1.Columns(3).Width = 70
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).Width = 60
        DataGridView1.Columns(6).Width = 60
        DataGridView1.Columns(7).Width = 60
        DataGridView1.Columns(8).Width = 60
        DataGridView1.Columns(9).Width = 130
        DataGridView1.Columns(10).Width = 130
        DataGridView1.Columns(11).Width = 220

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        ComboBox2.Items.Clear()
        Me.ComboBox2.Text = ""
        'llenar combo con DETALLES 
        detalle.CommandType = CommandType.Text
        detalle.CommandText = ("select NOM_DETALLE from DETALLE where ID_CLASIFICACION= (SELECT ID_CLASIFICACION FROM CLASIFICACION WHERE NOM_CLASIFICACION= '" & ComboBox1.SelectedItem & "')")
        detalle.Connection = (conexion)
        conexion.Open()
        variable2 = detalle.ExecuteReader()

        'ciclo que llena el combobox 
        While variable2.Read = True
            ComboBox2.Items.Add(variable2(0))
        End While
        conexion.Close()
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        ComboBox3.Items.Clear()
        Me.ComboBox3.Text = ""
        'llenar combo con DETALLES 
        serv.CommandType = CommandType.Text
        serv.CommandText = ("select NOM_SERVICIO from SERVICIO where ID_EMPRESA= (SELECT ID_EMPRESA FROM EMPRESA1 WHERE NOM_EMPRESA= '" & ComboBox4.SelectedItem & "')")
        serv.Connection = (conexion)
        conexion.Open()
        variable4 = serv.ExecuteReader()

        'ciclo que llena el combobox 
        While variable4.Read = True
            ComboBox3.Items.Add(variable4(0))
        End While
        conexion.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim res As MsgBoxResult


        If RadioButton1.Checked Then
            tipodia = "DL"
        ElseIf RadioButton2.Checked Then
            tipodia = "DS"
        ElseIf RadioButton3.Checked Then
            tipodia = "DF"
        End If

        guardar.Parameters.Clear()
        guardar.CommandType = CommandType.StoredProcedure
        guardar.CommandText = "agregar"
        guardar.Parameters.Add("@fecha", SqlDbType.Date).Value = (DateTimePicker1.Text)
        guardar.Parameters.Add("@tipodia", SqlDbType.NChar).Value = (tipodia)
        guardar.Parameters.Add("@hora", SqlDbType.Time).Value = (DateTimePicker2.Text)
        guardar.Parameters.Add("@empresa", SqlDbType.NVarChar).Value = (ComboBox4.Text)
        guardar.Parameters.Add("@servicio", SqlDbType.NVarChar).Value = (ComboBox3.Text)
        guardar.Parameters.Add("@patente", SqlDbType.NVarChar).Value = (TextBox2.Text)
        guardar.Parameters.Add("@horaini", SqlDbType.Time).Value = (DateTimePicker3.Text)
        guardar.Parameters.Add("@horater", SqlDbType.Time).Value = (DateTimePicker4.Text)
        guardar.Parameters.Add("@clasificacion", SqlDbType.NVarChar).Value = (ComboBox1.Text)
        guardar.Parameters.Add("@detalle", SqlDbType.NVarChar).Value = (ComboBox2.Text)
        guardar.Parameters.Add("@observaciones", SqlDbType.NVarChar).Value = (TextBox9.Text)

        If tipodia = "" Or ComboBox1.SelectedItem = "" Or ComboBox2.SelectedItem = "" Then

            MsgBox("Faltan datos, revise el formulario y vuelva a interntarlo", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)

        Else
            guardar.Connection = (conexion)
            conexion.Open()
            guardar.ExecuteNonQuery()
            conexion.Close()
            llenargrilla()
            DataGridView1.Columns(0).Width = 35
            DataGridView1.Columns(1).Width = 75
            DataGridView1.Columns(2).Width = 45
            DataGridView1.Columns(3).Width = 70
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 60
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(7).Width = 60
            DataGridView1.Columns(8).Width = 60
            DataGridView1.Columns(9).Width = 130
            DataGridView1.Columns(10).Width = 130
            DataGridView1.Columns(11).Width = 220


            res = MsgBox("Registro guardado con exito, ¿desea guardar o borrar un registro?", MsgBoxStyle.YesNo + MsgBoxStyle.Question)
            If res = MsgBoxResult.Yes Then

                ComboBox1.SelectedIndex = -1
                ComboBox2.SelectedIndex = -1
                ComboBox3.SelectedIndex = -1
                ComboBox4.SelectedIndex = -1
                DateTimePicker3.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                DateTimePicker4.Value = New DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0)
                TextBox2.Clear()
                TextBox9.Clear()
            Else
                If res = MsgBoxResult.No Then
                    Me.Close()
                    MDInicio.Show()
                End If
            End If

        End If
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Hide()
        MDInicio.Show()

    End Sub

  
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim res As String
        res = MsgBox("Seguro que desea eliminar este registro de la base de datos", MsgBoxStyle.YesNo, "Eliminar registro")
        If res = vbYes And DataGridView1.RowCount > 1 Then
            Dim sqlquery As String = "delete from REGISTRO where ID_REGISTRO= '" & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & "'"
            comando2.CommandText = sqlquery
            comando2.Connection = (conexion)
            conexion.Open()
            comando2.ExecuteNonQuery()
            conexion.Close()
            MsgBox("Registro eliminado con exito!!!", MsgBoxStyle.ApplicationModal)

            llenargrilla()
        End If
    End Sub

    
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim res As String
        res = MsgBox("Va a modificar el registro seleccionado en la grilla, si continua se abrira una ventana para edicion", MsgBoxStyle.YesNo, "Eliminar registro")
        If res = vbYes And DataGridView1.RowCount > 1 Then
            Editar.Show()


            edi.CommandType = CommandType.Text
            edi.CommandText = "select * from REGISTRO where ID_REGISTRO= '" & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & "'"
            edi.Connection = (conexion)
            conexion.Open()
            variable5 = edi.ExecuteReader()

            If variable5.Read = True Then

                Editar.TextBox1.Text = variable5.Item("ID_REGISTRO")
                Editar.TextBox2.Text = variable5.Item("FECHA")
                Editar.TextBox3.Text = variable5.Item("TIPO_DIA")
                Editar.DateTimePicker1.Text = variable5.Item("HORA").ToString
                Editar.TextBox5.Text = variable5.Item("EMPRESA")
                Editar.TextBox6.Text = variable5.Item("SERVICIO")
                Editar.TextBox11.Text = variable5.Item("PATENTE")
                Editar.DateTimePicker2.Text = variable5.Item("HORA_INI").ToString
                Editar.DateTimePicker3.Text = variable5.Item("HORA_TER").ToString
                Editar.TextBox8.Text = variable5.Item("CLASIFICACION")
                Editar.TextBox7.Text = variable5.Item("DETALLE_CMB")
                Editar.TextBox12.Text = variable5.Item("OBSERVACIONES")
                Me.Close()

            End If
        End If
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click

    End Sub
End Class
