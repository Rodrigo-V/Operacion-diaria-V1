
Imports System.Data
Imports System.Data.SqlClient


Public Class Neo_empresa
    Dim guardar As New SqlCommand
    Dim variable1 As SqlDataReader
    Dim buscar As New SqlCommand
    Dim clasif As New SqlCommand
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        Mantenedores.Show()

    End Sub

    Private Sub Neo_empresa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.CharacterCasing = CharacterCasing.Upper


      

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then

            MsgBox("Campo empresa no puede estar vacio, reintente o cancele la operacion ", MsgBoxStyle.Critical)
            
        Else
            conexion.Close()
            buscar.CommandType = CommandType.Text
            buscar.CommandText = ("select NOM_EMPRESA from EMPRESA1 WHERE NOM_EMPRESA  = '" & TextBox1.Text & "'")
            buscar.Connection = (conexion)
            conexion.Open()
            variable1 = buscar.ExecuteReader()


            If variable1.Read = True Then
                MsgBox("Registro ya se encuentra en la base de datos", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            Else


                conexion.Close()

                guardar.Parameters.Clear()
                guardar.CommandType = CommandType.StoredProcedure
                guardar.CommandText = "agreg_empresa"
                guardar.Parameters.Add("@empresa", SqlDbType.NVarChar).Value = (TextBox1.Text)
                guardar.Connection = (conexion)
                conexion.Open()
                guardar.ExecuteNonQuery()
                conexion.Close()

                MsgBox("Registro Guardado con exito", MsgBoxStyle.Information)
                TextBox1.Clear()
                TextBox1.Focus()
            End If
        End If


    End Sub


End Class