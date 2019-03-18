Imports System.Data
Imports System.Data.SqlClient


Public Class Nuevo_detalle
    Dim buscar As New SqlCommand
    Dim guardar As New SqlCommand
    Dim variable As SqlDataReader
    Dim variable10 As SqlDataReader
    Dim variable2 As SqlDataReader
    Dim detal As New SqlCommand
    Dim clasif As New SqlCommand
    Dim clasif2 As New SqlCommand
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")

    Private Sub Nuevo_detalle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.CharacterCasing = CharacterCasing.Upper

        clasif.CommandType = CommandType.Text
        clasif.CommandText = ("select ID_CLASIFICACION, NOM_CLASIFICACION from CLASIFICACION")
        clasif.Connection = (conexion)
        conexion.Open()
        variable = clasif.ExecuteReader()

        'ciclo que llena el combobox 
        While variable.Read = True

            ComboBox1.Items.Add(variable(1))
        End While
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
            buscar.CommandText = ("select NOM_DETALLE from DETALLE WHERE NOM_DETALLE  = '" & TextBox1.Text & "'")
            buscar.Connection = (conexion)
            conexion.Open()
            variable10 = buscar.ExecuteReader()


            If variable10.Read = False Then
                conexion.Close()

                guardar.Parameters.Clear()
                guardar.CommandType = CommandType.StoredProcedure
                guardar.CommandText = "agrega_detalle"
                guardar.Parameters.Add("@det", SqlDbType.NVarChar).Value = (TextBox1.Text)
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
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        clasif2.CommandType = CommandType.Text
        clasif2.CommandText = ("select * from CLASIFICACION WHERE NOM_CLASIFICACION='" & ComboBox1.SelectedItem & "'")
        clasif2.Connection = (conexion)
        conexion.Open()
        variable2 = clasif2.ExecuteReader()

        'ciclo que llena el combobox 
        If variable2.Read = True Then

            TextBox2.Text = (variable2(0))

        End If
        conexion.Close()
    End Sub

    
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class