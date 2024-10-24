Imports System.Data.OleDb

Public Class frmSpectacleSaleList

    Dim DBProvider As String
    Dim DBSource As String
    Dim con As New OleDbConnection
    Dim sql As String = "Select * from SaleInvoice"
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim i As Integer
    Dim len As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice Where BookingDate Between @d1 and @d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table
        myConnection.Close()
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub frmSpectacleSaleList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmSpectacleSaleList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmSpectacleSaleList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        LoadText()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        DBProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
        DBSource = "Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        con.ConnectionString = DBProvider & DBSource
        con.Open()
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "Billing")
        len = da.Fill(ds, "Billing")
    End Sub

    Private Sub Nav()
        TextBox6.Text = ds.Tables("Billing").Rows(i).Item(1).ToString
        TextBox2.Text = ds.Tables("Billing").Rows(i).Item(2).ToString
        TextBox3.Text = ds.Tables("Billing").Rows(i).Item(3).ToString
        TextBox4.Text = ds.Tables("Billing").Rows(i).Item(4).ToString
        TextBox10.Text = ds.Tables("Billing").Rows(i).Item(6).ToString
        TextBox9.Text = ds.Tables("Billing").Rows(i).Item(7).ToString
        TextBox8.Text = ds.Tables("Billing").Rows(i).Item(8).ToString
        TextBox7.Text = ds.Tables("Billing").Rows(i).Item(9).ToString
        TextBox34.Text = ds.Tables("Billing").Rows(i).Item(10).ToString
        TextBox35.Text = ds.Tables("Billing").Rows(i).Item(13).ToString
        TextBox30.Text = ds.Tables("Billing").Rows(i).Item(14).ToString
        TextBox37.Text = ds.Tables("Billing").Rows(i).Item(15).ToString
        TextBox38.Text = ds.Tables("Billing").Rows(i).Item(16).ToString
        TextBox26.Text = ds.Tables("Billing").Rows(i).Item(17).ToString
        TextBox27.Text = ds.Tables("Billing").Rows(i).Item(18).ToString
        TextBox28.Text = ds.Tables("Billing").Rows(i).Item(19).ToString
        TextBox29.Text = ds.Tables("Billing").Rows(i).Item(20).ToString
        TextBox31.Text = ds.Tables("Billing").Rows(i).Item(21).ToString
        TextBox33.Text = ds.Tables("Billing").Rows(i).Item(22).ToString
        TextBox32.Text = ds.Tables("Billing").Rows(i).Item(23).ToString
        TextBox12.Text = ds.Tables("Billing").Rows(i).Item(24).ToString
        TextBox11.Text = ds.Tables("Billing").Rows(i).Item(25).ToString
        TextBox17.Text = ds.Tables("Billing").Rows(i).Item(26).ToString
        TextBox16.Text = ds.Tables("Billing").Rows(i).Item(27).ToString
        TextBox15.Text = ds.Tables("Billing").Rows(i).Item(28).ToString
        TextBox14.Text = ds.Tables("Billing").Rows(i).Item(29).ToString
        TextBox13.Text = ds.Tables("Billing").Rows(i).Item(30).ToString
        TextBox18.Text = ds.Tables("Billing").Rows(i).Item(31).ToString
        TextBox19.Text = ds.Tables("Billing").Rows(i).Item(32).ToString
        TextBox20.Text = ds.Tables("Billing").Rows(i).Item(33).ToString
        TextBox21.Text = ds.Tables("Billing").Rows(i).Item(34).ToString
        TextBox22.Text = ds.Tables("Billing").Rows(i).Item(35).ToString
        TextBox23.Text = ds.Tables("Billing").Rows(i).Item(36).ToString
        TextBox24.Text = ds.Tables("Billing").Rows(i).Item(37).ToString
        TextBox25.Text = ds.Tables("Billing").Rows(i).Item(38).ToString
        TextBox39.Text = ds.Tables("Billing").Rows(i).Item(39).ToString
        TextBox40.Text = ds.Tables("Billing").Rows(i).Item(40).ToString
        TextBox41.Text = ds.Tables("Billing").Rows(i).Item(41).ToString
        con.Close()

        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT ProdName, Qty, Price, Discount, Total, Price*Qty As Gross, Gross*6/112 As CGST, Gross*6/112 As SGST FROM InvoiceProduct WHERE InvoiceProduct.Cust_ID LIKE'%" &
        TextBox9.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("InvoiceProduct")
        adapter.Fill(dt)
        Me.DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice WHERE BookedBy LIKE'%" &
        TextBox5.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text = "Payment By Cash" Then
                Dim sqlsearch As String
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                'Change the following to your access database location
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice WHERE Paymode LIKE'%" &
                ComboBox1.Text & "%'"
                Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
                Dim dt As New DataTable("SaleInvoice")
                adapter.Fill(dt)
                DataGridView1.DataSource = dt
                myConnection.Close()
            Else
                If ComboBox1.Text = "Payment By Card" Then
                    Dim sqlsearch As String
                    provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                    'Change the following to your access database location
                    dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                    connString = provider & dataFile
                    myConnection.ConnectionString = connString
                    myConnection.Open()
                    sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice WHERE Paymode LIKE'%" &
                    ComboBox1.Text & "%'"
                    Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
                    Dim dt As New DataTable("SaleInvoice")
                    adapter.Fill(dt)
                    DataGridView1.DataSource = dt
                    myConnection.Close()
                Else
                    If ComboBox1.Text = "Payment By Both" Then
                        Dim sqlsearch As String
                        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                        'Change the following to your access database location
                        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                        connString = provider & dataFile
                        myConnection.ConnectionString = connString
                        myConnection.Open()
                        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, BookedBy, Status, NetAmt, PaidAmt, DueAmt, Paymode from SaleInvoice WHERE Paymode LIKE'%" &
                        ComboBox1.Text & "%'"
                        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
                        Dim dt As New DataTable("SaleInvoice")
                        adapter.Fill(dt)
                        DataGridView1.DataSource = dt
                        myConnection.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Select Listed Options Only")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Clear()
        RefreshData()
    End Sub

    Sub Clear()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        TextBox1.Clear()
        TextBox5.Clear()
        ComboBox1.SelectedIndex = -1
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        i = 0
        Nav()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If (i = len - 1) Then
            MsgBox("This is the Last Record")
        Else
            i = i + 1
            Nav()
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If (i = 0) Then
            MsgBox("This is the First Record")
        Else
            i = i - 1
            Nav()
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        i = len - 1
        Nav()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If TextBox36.Text <> "" Then
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "SELECT * FROM SaleInvoice WHERE (Cust_Name = '" & TextBox36.Text & "')"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            dr = cmd.ExecuteReader
            While dr.Read()
                TextBox6.Focus()
                TextBox6.Text = dr("Cust_ID").ToString
                TextBox2.Text = dr("Cust_Name").ToString
                TextBox3.Text = dr("Mobile").ToString
                TextBox4.Text = dr("Phone").ToString
                TextBox10.Text = dr("BookingDate").ToString
                TextBox9.Text = dr("ReceiptNo").ToString
                TextBox8.Text = dr("DeliveryDate").ToString
                TextBox7.Text = dr("BookedBy").ToString
                TextBox34.Text = dr("DeliveryTime").ToString
                TextBox35.Text = dr("JobStatus").ToString
                TextBox30.Text = dr("GrossAmt").ToString
                TextBox37.Text = dr("CGST").ToString
                TextBox38.Text = dr("SGST").ToString
                TextBox26.Text = dr("ScmAmt").ToString
                TextBox27.Text = dr("NetAmt").ToString
                TextBox28.Text = dr("PaidAmt").ToString
                TextBox29.Text = dr("DueAmt").ToString
                TextBox31.Text = dr("Paymode").ToString
                TextBox33.Text = dr("RSPH").ToString
                TextBox32.Text = dr("RCYL").ToString
                TextBox12.Text = dr("RAXIS").ToString
                TextBox11.Text = dr("RVN").ToString
                TextBox17.Text = dr("RADD").ToString
                TextBox16.Text = dr("LSPH").ToString
                TextBox15.Text = dr("LCYL").ToString
                TextBox14.Text = dr("LAXIS").ToString
                TextBox13.Text = dr("LVN").ToString
                TextBox18.Text = dr("LADD").ToString
                TextBox19.Text = dr("PD").ToString
                TextBox20.Text = dr("REFBY").ToString
                TextBox21.Text = dr("LensType").ToString
                TextBox22.Text = dr("LensType1").ToString
                TextBox23.Text = dr("LensType2").ToString
                TextBox24.Text = dr("LensType3").ToString
                TextBox25.Text = dr("Remarks1").ToString
                TextBox39.Text = dr("Right").ToString
                TextBox40.Text = dr("Left").ToString
                TextBox41.Text = dr("RLAdd").ToString
            End While
            myConnection.Close()

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT ProdName, Qty, Price, Discount, Total, Price*Qty As Gross, Gross*6/112 As CGST, Gross*6/112 As SGST FROM InvoiceProduct WHERE InvoiceProduct.Cust_ID LIKE'%" &
            TextBox9.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("InvoiceProduct")
            adapter.Fill(dt)
            Me.DataGridView2.DataSource = dt
            myConnection.Close()
        Else
            MsgBox("Please enter Customer Name to Search")
        End If
    End Sub

    Sub LoadText()
        Try
            Dim con As New OleDbConnection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb"
            con.Open()
            Dim dt As New DataTable
            Dim ds As New DataSet
            ds.Tables.Add(dt)
            Dim da As New OleDbDataAdapter("SELECT Cust_Name FROM SaleInvoice", con)
            da.Fill(dt)
            Dim r As DataRow
            TextBox36.AutoCompleteCustomSource.Clear()
            For Each r In dt.Rows
                TextBox36.AutoCompleteCustomSource.Add(r.Item(0).ToString)
            Next
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Print()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New JobCard 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SaleInvoice.ReceiptNo, SaleInvoice.Cust_Name, SaleInvoice.Mobile, SaleInvoice.BookingDate, SaleInvoice.BookedBy, SaleInvoice.DeliveryDate, SaleInvoice.DeliveryTime, SaleInvoice.GrossAmt, SaleInvoice.REFBY, SaleInvoice.NetAmt, SaleInvoice.PaidAmt, SaleInvoice.DueAmt, SaleInvoice.ScmAmt, SaleInvoice.Paymode, SaleInvoice.RSPH, SaleInvoice.RCYL, SaleInvoice.RAXIS, SaleInvoice.RVN, SaleInvoice.RADD, SaleInvoice.LSPH, SaleInvoice.LCYL, SaleInvoice.LAXIS, SaleInvoice.LVN, SaleInvoice.LADD, SaleInvoice.PD, SaleInvoice.LensType, SaleInvoice.LensType1, SaleInvoice.LensType2, SaleInvoice.LensType3, SaleInvoice.Remarks1, SaleInvoice.Right, SaleInvoice.Left, SaleInvoice.RLAdd, SaleInvoice.PRGRight, SaleInvoice.PRGLeft, InvoiceProduct.ProdName, InvoiceProduct.Qty, InvoiceProduct.Price, InvoiceProduct.Discount, InvoiceProduct.Total FROM InvoiceProduct INNER JOIN SaleInvoice ON SaleInvoice.ReceiptNo=InvoiceProduct.Cust_ID Where SaleInvoice.ReceiptNo=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox9.Text)
            MyCommand1.CommandText = "SELECT * from SaleInvoice"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SaleInvoice")
            myDA.Fill(myDS, "InvoiceProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox9.Text)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Print1()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New InvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SaleInvoice.ReceiptNo, SaleInvoice.Cust_Name, SaleInvoice.Mobile, SaleInvoice.BookingDate, SaleInvoice.GrossAmt, SaleInvoice.NetAmt, SaleInvoice.PaidAmt, SaleInvoice.DueAmt, SaleInvoice.CGST, SaleInvoice.SGST, SaleInvoice.Paymode, SaleInvoice.ScmAmt, InvoiceProduct.ProdName, InvoiceProduct.Qty, InvoiceProduct.Price, InvoiceProduct.Discount, InvoiceProduct.Total FROM InvoiceProduct INNER JOIN SaleInvoice ON SaleInvoice.ReceiptNo=InvoiceProduct.Cust_ID Where SaleInvoice.ReceiptNo=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox9.Text)
            MyCommand1.CommandText = "SELECT * from SaleInvoice"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SaleInvoice")
            myDA.Fill(myDS, "InvoiceProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox9.Text)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Print()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Print1()
    End Sub
End Class