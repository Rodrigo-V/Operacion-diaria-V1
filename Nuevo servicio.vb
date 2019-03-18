Imports System.Data
Imports System.Data.SqlClient

Public Class Nuevo_servicio
    Dim buscar As New SqlCommand
    Dim guardar As New SqlCommand
    Dim variable As SqlDataReader
    Dim variable1 As SqlDataReader
    Dim variable2 As SqlDataReader
    Dim clasif As New SqlCommand
    Dim clasif2 As New SqlCommand
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")
    Private Sub Nuevo_servicio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.CharacterCasing = CharacterCasing.Upper

        clasif.CommandType = CommandType.Text
        clasif.CommandText = ("select * from EMPRESA1")
        clasif.Connection = (conexion)
        conexion.Open()
        variable = clasif.ExecuteReader()

        'ciclo que llena el combobox 
        While variable.Read = True

            ComboBox1.Items.Add(variable(1))
        End While
        conexion.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        clasif2.CommandType = CommandType.Text
        clasif2.CommandText = ("select * from EMPRESA1 WHERE NOM_EMPRESA='" & ComboBox1.SelectedItem & "'")
        clasif2.Connection = (conexion)
        conexion.Open()
        variable2 = clasif2.ExecuteReader()

        'ciclo que llena el combobox 
        If variable2.Read = True Then

            TextBox2.Text = (variable2(0))

        End If
        conexion.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        Mantenedores.Show()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        conexion.Close()
        If TextBox1.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("debe seleccionar una clasificacion y escribir un nuevo detalle para continuar", MsgBoxStyle.Critical)
        Else

            buscar.CommandType = CommandType.Text
            buscar.CommandText = ("select NOM_SERVICIO from SERVICIO WHERE NOM_SERVICIO  = '" & TextBox1.Text & "'")
            buscar.Connection = (conexion)
            conexion.Open()
            variable1 = buscar.ExecuteReader()


            If variable1.Read = False Then
                conexion.Close()

                guardar.Parameters.Clear()
                guardar.CommandType = CommandType.StoredProcedure
                guardar.CommandText = "agrega_servicio"
                guardar.Parameters.Add("@ser", SqlDbType.NVarChar).Value = (TextBox1.Text)
                guardar.Parameters.Add("@id", SqlDbType.Int).Value = (TextBox2.Text)
                guardar.Connection = (conexion)
                conexion.Open()
                guardar.ExecuteNonQuery()


                MsgBox("Registro Guardado con exito", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            Else
                conexion.Close()
                MsgBox("Registro ya se encuentra en la base de datos", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            End If
            conexion.Close()
        End If
    End Sub
End Class