
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop


Module Module1

    Dim G4dt As New DataTable
    Dim fechahora As String = DateTime.Now.ToString()
    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")
    Dim guardar As New SqlCommand
    Dim param As New SqlCommand
    Dim leer As SqlDataReader
    Dim variable As SqlDataReader
    Dim registrodt As New DataTable
    Dim comando As New SqlCommand
    Dim comando1 As New SqlCommand
    Dim comando2 As New SqlCommand
    Dim empleado As New SqlCommand
    Dim fecha As Char


    Public Sub armardatatable()

        'arma datatable
        registrodt.Columns.Clear()
        registrodt.Columns.Add("ID")
        registrodt.Columns.Add("FECHA")
        registrodt.Columns.Add("TIPO DIA")
        registrodt.Columns.Add("HORA")
        registrodt.Columns.Add("EMPRESA")
        registrodt.Columns.Add("SERVICIO")
        registrodt.Columns.Add("PPU")
        registrodt.Columns.Add("INICIO")
        registrodt.Columns.Add("TERMINO")
        registrodt.Columns.Add("CLASIFICACION")
        registrodt.Columns.Add("DETALLE")
        registrodt.Columns.Add("OBSERVACIONES")
    End Sub




    Public Sub llenargrilla()


        registrodt.Rows.Clear()
        comando.CommandText = "select * from REGISTRO  where FECHA= '" & REGISTRO.TextBox1.Text & "'ORDER BY HORA DESC"
        comando.Connection = (conexion)
        conexion.Open()
        leer = comando.ExecuteReader
        While leer.Read
            registrodt.Rows.Add(leer.Item("ID_REGISTRO"), leer.Item("FECHA").ToString.Remove(10, 8), leer.Item("TIPO_DIA"), leer.Item("HORA"), leer.Item("EMPRESA"), leer.Item("SERVICIO"), leer.Item("PATENTE"), leer.Item("HORA_INI"), leer.Item("HORA_TER"), leer.Item("CLASIFICACION"), leer.Item("DETALLE_CMB"), leer.Item("OBSERVACIONES"))
        End While
        leer.Close()
        conexion.Close()

        REGISTRO.DataGridView1.DataSource = registrodt


    End Sub

  

End Module
