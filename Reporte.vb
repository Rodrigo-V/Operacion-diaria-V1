
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class Reporte

    Dim conexion As New SqlConnection("Data Source=.\SQLEXPRESS;Initial Catalog=operacionLC;Integrated Security=True")
    Dim leer As SqlDataReader
    Dim registrodt As New DataTable
    Dim comando As New SqlCommand
    Dim dts As DataSet
    Dim adpt As SqlDataAdapter
    Dim variable51 As SqlDataReader
    Dim edi1 As New SqlCommand
   


    Private Sub Reporte_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        ProgressBar1.Visible = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 100
        ProgressBar1.Value = 0

        'muestra fecha en el form
        TextBox1.Text = DateValue(Now)
        TextBox1.Enabled = False

        'arma datatable
        registrodt.Columns.Add("ID")
        registrodt.Columns.Add("FECHA")
        registrodt.Columns.Add("TIPO DIA")
        registrodt.Columns.Add("HORA")
        registrodt.Columns.Add("EMPRESA")
        registrodt.Columns.Add("SERVICIO")
        registrodt.Columns.Add("PATENTE")
        registrodt.Columns.Add("INICIO")
        registrodt.Columns.Add("TERMINO")
        registrodt.Columns.Add("CLASIFICACION")
        registrodt.Columns.Add("DETALLE")
        registrodt.Columns.Add("OBSERVACIONES")


    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        REGISTRO.Show()
        Me.Close()


    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Me.DataGridView1.Enabled = False
        Me.DataGridView1.DataSource = Nothing
        registrodt.Rows.Clear()
        Me.DataGridView1.Enabled = True

        comando.CommandType = CommandType.Text
        comando.CommandText = "select * from REGISTRO WHERE FECHA BETWEEN  '" & Me.DateTimePicker1.Value.ToShortDateString & "' And '" & Me.DateTimePicker2.Value.ToShortDateString & "'ORDER BY FECHA DESC"
        comando.Connection = (conexion)
        conexion.Open()
        leer = comando.ExecuteReader
        While leer.Read
            registrodt.Rows.Add(leer.Item("ID_REGISTRO"), leer.Item("FECHA").ToString.Remove(10, 8), leer.Item("TIPO_DIA"), leer.Item("HORA"), leer.Item("EMPRESA"), leer.Item("SERVICIO"), leer.Item("PATENTE"), leer.Item("HORA_INI"), leer.Item("HORA_TER"), leer.Item("CLASIFICACION"), leer.Item("DETALLE_CMB"), leer.Item("OBSERVACIONES"))
        End While
        leer.Close()
        DataGridView1.DataSource = registrodt
        conexion.Close()
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(1).Width = 65
        DataGridView1.Columns(2).Width = 35
        DataGridView1.Columns(3).Width = 60
        DataGridView1.Columns(4).Width = 70
        DataGridView1.Columns(5).Width = 50
        DataGridView1.Columns(6).Width = 50
        DataGridView1.Columns(7).Width = 50
        DataGridView1.Columns(8).Width = 50
        DataGridView1.Columns(9).Width = 120
        DataGridView1.Columns(10).Width = 120
        DataGridView1.Columns(11).Width = 200


    End Sub
  
    Private Sub ExportarExcel()


        Try



            ProgressBar1.Value = 0
            ProgressBar1.Value += 10


            'Variables del programa
            Dim rutaPlantilla As String = "C:\plantillaxls\plantilla_novedades.xls"
            Dim rutaGuardado As String = "C:\Reportes\Novedades desde " + Me.DateTimePicker1.Value.ToShortDateString & " Hasta " + Me.DateTimePicker2.Value.ToShortDateString & ".xls"
            Dim xlApp As Excel.Application = New Excel.Application()
            Dim _libroExcel As Excel.Workbook = Nothing
            Dim _HojaExcel As Excel.Worksheet = Nothing
            Dim _Rango As Excel.Range = Nothing
            Dim misValue As Object = System.Reflection.Missing.Value

            ProgressBar1.Value += 10
            _libroExcel = xlApp.Workbooks.Open(rutaPlantilla, misValue, misValue, misValue _
                    , misValue, misValue, misValue, misValue _
                   , misValue, misValue, misValue, misValue _
                  , misValue, misValue, misValue)

            ProgressBar1.Value += 10
            _libroExcel = xlApp.ActiveWorkbook

            _HojaExcel = CType(_libroExcel.Worksheets.Item(1), Excel.Worksheet)
            _HojaExcel = _libroExcel.Worksheets(1)
            _HojaExcel.Columns("A").NumberFormat = "@"
            _HojaExcel.Columns("B").NumberFormat = "DD/MM/YY"
            Try
                ProgressBar1.Value += 10
                ' columnas y filas
                Dim ncol As Integer = DataGridView1.ColumnCount
                Dim nrow As Integer = DataGridView1.RowCount
                ProgressBar1.Value += 10
                For i As Integer = 1 To ncol
                    _HojaExcel.Cells.Item(8, i) = DataGridView1.Columns(i - 1).Name.Trim
                    _HojaExcel.Cells.Item(1, i).horizontalalignment = 6


                Next
                ProgressBar1.Value += 10
                For fila As Integer = 0 To nrow - 1
                    For col As Integer = 0 To ncol - 1
                        _HojaExcel.Cells.Item(fila + 9, col + 1) = DataGridView1.Rows(fila).Cells(col).Value
                    Next
                Next

                ProgressBar1.Value += 10

                _HojaExcel.Rows.Item(1).Font.Bold = 1
                _HojaExcel.Rows.Item(1).Autofit()

                'Guardamos el libro 
                _libroExcel.SaveAs(rutaGuardado, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
                ProgressBar1.Value += 10
                _libroExcel.Close(True, misValue, misValue)
                ProgressBar1.Value += 10
                xlApp.Quit()

                ProgressBar1.Value += 10

                MessageBox.Show("Datos Exportados" & rutaGuardado)

            Catch ex As System.Exception

                MessageBox.Show(ex.Message & "\n\n=======  Error al escribir el excel:  ======\n\n" & _
                   ex.StackTrace)

            Finally
            End Try
        Catch exl As System.Exception

            MessageBox.Show(exl.Message & _
             "\n\n=======   Error al abrir el archivo  ======\n\n" & _
             exl.StackTrace)

        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        ProgressBar1.Visible = True
        ExportarExcel()

    End Sub



    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MDInicio.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim res As String
        res = MsgBox("Va a modificar el registro seleccionado en la grilla, si continua se abrira una ventana para edicion", MsgBoxStyle.YesNo, "Eliminar registro")
        If res = vbYes And DataGridView1.RowCount > 1 Then
            Editar.Show()


            edi1.CommandType = CommandType.Text
            edi1.CommandText = "select * from REGISTRO where ID_REGISTRO= '" & DataGridView1.SelectedRows(0).Cells(0).Value.ToString & "'"
            edi1.Connection = (conexion)
            conexion.Open()
            variable51 = edi1.ExecuteReader()

            If variable51.Read = True Then

                Editar.TextBox1.Text = variable51.Item("ID_REGISTRO")
                Editar.TextBox2.Text = variable51.Item("FECHA")
                Editar.TextBox3.Text = variable51.Item("TIPO_DIA")
                Editar.DateTimePicker1.Text = variable51.Item("HORA").ToString
                Editar.TextBox5.Text = variable51.Item("EMPRESA")
                Editar.TextBox6.Text = variable51.Item("SERVICIO")
                Editar.TextBox11.Text = variable51.Item("PATENTE")
                Editar.DateTimePicker2.Text = variable51.Item("HORA_INI").ToString
                Editar.DateTimePicker3.Text = variable51.Item("HORA_TER").ToString
                Editar.TextBox8.Text = variable51.Item("CLASIFICACION")
                Editar.TextBox7.Text = variable51.Item("DETALLE_CMB")
                Editar.TextBox12.Text = variable51.Item("OBSERVACIONES")
                Me.Close()

            End If
        End If
    End Sub

End Class