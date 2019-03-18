
Imports System.Data
Imports System.Data.SqlClient

Public Class Neo_clas

    Dim guardar As New SqlCommand
    Dim buscar As New SqlCommand
    Dim variable1 As SqlDataReader
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")

    Private Sub Neo_clas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.CharacterCasing = CharacterCasing.Upper

    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then

            MsgBox("Campo Clasificacion no puede estar vacio, reintente o cancele la operacion ", MsgBoxStyle.Critical)
        Else
            buscar.CommandType = CommandType.Text
            buscar.CommandText = ("select NOM_CLASIFICACION from CLASIFICACION WHERE NOM_CLASIFICACION  = '" & TextBox1.Text & "'")
            buscar.Connection = (conexion)
            conexion.Open()
            variable1 = buscar.ExecuteReader()


            If variable1.Read = False Then
                conexion.Close()

                guardar.Parameters.Clear()
                guardar.CommandType = CommandType.StoredProcedure
                guardar.CommandText = "agreg_cla"
                guardar.Parameters.Add("@cla", SqlDbType.NVarChar).Value = (TextBox1.Text)
                guardar.Connection = (conexion)
                conexion.Open()
                guardar.ExecuteNonQuery()
                conexion.Close()

                MsgBox("Registro Guardado con exito", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            Else
                MsgBox("Registro ya se encuentra en la base de datos", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            End If

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        Mantenedores.Show()

    End Sub
End Class