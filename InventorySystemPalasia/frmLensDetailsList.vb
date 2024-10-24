Imports System.Data.OleDb

Public Class frmLensDetailsList

    Private Sub frmLensDetailsList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmLensDetailsList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Function GetValue(ByVal Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub frmLensDetailsList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd from SaleInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False

        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 150
        Me.DataGridView1.Columns(6).Width = 200
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            frmLensDetailsEntry.TextBox1.Text = DataGridView1.Item(0, i).Value.ToString
            frmLensDetailsEntry.TextBox2.Text = DataGridView1.Item(1, i).Value.ToString
            frmLensDetailsEntry.TextBox3.Text = DataGridView1.Item(2, i).Value.ToString
            frmLensDetailsEntry.TextBox28.Text = DataGridView1.Item(3, i).Value.ToString
            frmLensDetailsEntry.TextBox24.Text = DataGridView1.Item(4, i).Value.ToString
            frmLensDetailsEntry.TextBox4.Text = DataGridView1.Item(5, i).Value.ToString
            frmLensDetailsEntry.TextBox5.Text = DataGridView1.Item(6, i).Value.ToString
            frmLensDetailsEntry.DateTimePicker1.Value = DataGridView1.Item(7, i).Value.ToString
            frmLensDetailsEntry.TextBox6.Text = DataGridView1.Item(8, i).Value.ToString
            frmLensDetailsEntry.TextBox7.Text = DataGridView1.Item(9, i).Value.ToString
            frmLensDetailsEntry.TextBox8.Text = DataGridView1.Item(10, i).Value.ToString
            frmLensDetailsEntry.TextBox9.Text = DataGridView1.Item(11, i).Value.ToString
            frmLensDetailsEntry.TextBox10.Text = DataGridView1.Item(12, i).Value.ToString
            frmLensDetailsEntry.TextBox11.Text = DataGridView1.Item(13, i).Value.ToString
            frmLensDetailsEntry.TextBox12.Text = DataGridView1.Item(14, i).Value.ToString
            frmLensDetailsEntry.TextBox13.Text = DataGridView1.Item(15, i).Value.ToString
            frmLensDetailsEntry.TextBox14.Text = DataGridView1.Item(16, i).Value.ToString
            frmLensDetailsEntry.TextBox15.Text = DataGridView1.Item(17, i).Value.ToString
            frmLensDetailsEntry.TextBox16.Text = DataGridView1.Item(18, i).Value.ToString
            frmLensDetailsEntry.ComboBox1.Text = DataGridView1.Item(19, i).Value.ToString
            frmLensDetailsEntry.TextBox17.Text = DataGridView1.Item(20, i).Value.ToString
            frmLensDetailsEntry.TextBox25.Text = DataGridView1.Item(21, i).Value.ToString
            frmLensDetailsEntry.TextBox26.Text = DataGridView1.Item(22, i).Value.ToString
            frmLensDetailsEntry.TextBox27.Text = DataGridView1.Item(23, i).Value.ToString
            frmLensDetailsEntry.TextBox18.Text = DataGridView1.Item(24, i).Value.ToString
            frmLensDetailsEntry.TextBox20.Text = DataGridView1.Item(25, i).Value.ToString
            frmLensDetailsEntry.TextBox22.Text = DataGridView1.Item(26, i).Value.ToString
            frmLensDetailsEntry.TextBox23.Text = DataGridView1.Item(27, i).Value.ToString
            frmLensDetailsEntry.Show()
        Catch ex As Exception
            MsgBox("Row is Empty")
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd FROM SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd FROM SaleInvoice WHERE Mobile LIKE'%" &
        TextBox2.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd FROM SaleInvoice WHERE Phone LIKE'%" &
        TextBox3.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd FROM SaleInvoice WHERE REFBY LIKE'%" &
        TextBox4.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RefreshData()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim filepath As String = Application.StartupPath & "\Excel\Power_Details.xml"
        Dim _mFileStream As New IO.StreamWriter(filepath, False)
        Try
            _mFileStream.WriteLine("<?xml version=""1.0""?>")
            _mFileStream.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
            _mFileStream.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
            _mFileStream.WriteLine("    <ss:Styles>")
            _mFileStream.WriteLine("        <ss:Style ss:ID=""1"">")
            _mFileStream.WriteLine("           <ss:Font ss:Bold=""1""/>")
            _mFileStream.WriteLine("           <ss:FontName=""Courier New""/>")
            _mFileStream.WriteLine("        </ss:Style>")
            _mFileStream.WriteLine("    </ss:Styles>")
            _mFileStream.WriteLine("    <ss:Worksheet ss:Name=""Sheet1$"">")
            _mFileStream.WriteLine("        <ss:Table>")
            For x As Integer = 0 To DataGridView1.Columns.Count - 1
                _mFileStream.WriteLine("            <ss:Column ss:Width=""{0}""/>", DataGridView1.Columns.Item(x).Width)
            Next
            _mFileStream.WriteLine("            <ss:Row ss:StyleID=""1"">")
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                _mFileStream.WriteLine("                <ss:Cell>")
                _mFileStream.WriteLine(String.Format("                   <ss:Data ss:Type=""String"">{0}</ss:Data>", DataGridView1.Columns.Item(i).HeaderText))
                _mFileStream.WriteLine("</ss:Cell>")
            Next
            _mFileStream.WriteLine("            </ss:Row>")
            For intRow As Integer = 0 To DataGridView1.RowCount - 2
                _mFileStream.WriteLine(String.Format("            <ss:Row ss:Height =""{0}"">", DataGridView1.Rows(intRow).Height))
                For intCol As Integer = 0 To DataGridView1.Columns.Count - 1
                    _mFileStream.WriteLine("                <ss:Cell>")
                    _mFileStream.WriteLine(String.Format("                   <ss:Data ss:Type=""String"">{0}</ss:Data>", DataGridView1.Item(intCol, intRow).Value.ToString))
                    _mFileStream.WriteLine("                </ss:Cell>")
                Next
                _mFileStream.WriteLine("            </ss:Row>")
            Next
            _mFileStream.WriteLine("        </ss:Table>")
            _mFileStream.WriteLine("    </ss:Worksheet>")
            _mFileStream.WriteLine("</ss:Workbook>")
            _mFileStream.Close()
            _mFileStream.Dispose()
            _mFileStream = Nothing
            MessageBox.Show("Exported Successfully to" & Application.StartupPath & "\Excel\AO.xlsx")
        Catch ex As Exception
            _mFileStream.Close()
            _mFileStream.Dispose()
            _mFileStream = Nothing
            MessageBox.Show("Error While Exporting Data To Excel. Error : " & ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd from SaleInvoice Where BookingDate Between @d1 and @d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table
        myConnection.Close()

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(19).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False
    End Sub
End Class