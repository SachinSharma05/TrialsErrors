Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Net
Imports System.Web
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Collections.Generic

Module ModCS
    Public provider As String
    Public dataFile As String
    Public connString As String
    Public myConnection As OleDbConnection = New OleDbConnection
    Public ds As New DataSet
    Public dt As New DataTable
    Public da As New OleDbDataAdapter
    Public dr As OleDbDataReader
    Public rdr As OleDbDataReader = Nothing
    Public rdr1 As OleDbDataReader = Nothing
    Public rdr2 As OleDbDataReader = Nothing
    Public cmd As OleDbCommand

    Public focusedForeColor As Color = Color.Black
    Public focusedBackColor As Color = Color.BurlyWood

    Public Sub b_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        Dim b = DirectCast(sender, Button)
        Dim colors = DirectCast(b.Tag, Tuple(Of Color, Color))
        b.ForeColor = colors.Item1
        b.BackColor = colors.Item2
    End Sub

    Public Sub b_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
        Dim b = DirectCast(sender, Button)
        b.ForeColor = focusedForeColor
        b.BackColor = focusedBackColor
    End Sub

    Public Sub ExportExcel(ByVal obj As Object)
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = obj.RowCount
            colsTotal = obj.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = obj.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = obj.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub
End Module