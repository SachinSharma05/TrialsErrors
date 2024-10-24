Imports System.Data.OleDb

Public Class frmContactLensListJobCardInvoice

    Dim DBProvider As String
    Dim DBSource As String
    Dim con As New OleDbConnection
    Dim sql As String = "Select * from CLSale"
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim i As Integer
    Dim len As Integer

    Private Sub frmContactLensListJobCardInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
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
        TextBox26.Text = ds.Tables("Billing").Rows(i).Item(1).ToString
        TextBox2.Text = ds.Tables("Billing").Rows(i).Item(2).ToString
        TextBox3.Text = ds.Tables("Billing").Rows(i).Item(3).ToString
        TextBox4.Text = ds.Tables("Billing").Rows(i).Item(4).ToString
        TextBox8.Text = ds.Tables("Billing").Rows(i).Item(5).ToString
        TextBox6.Text = ds.Tables("Billing").Rows(i).Item(6).ToString
        TextBox35.Text = ds.Tables("Billing").Rows(i).Item(7).ToString
        TextBox32.Text = ds.Tables("Billing").Rows(i).Item(8).ToString
        TextBox37.Text = ds.Tables("Billing").Rows(i).Item(9).ToString
        TextBox38.Text = ds.Tables("Billing").Rows(i).Item(10).ToString
        TextBox30.Text = ds.Tables("Billing").Rows(i).Item(11).ToString
        TextBox27.Text = ds.Tables("Billing").Rows(i).Item(12).ToString
        TextBox28.Text = ds.Tables("Billing").Rows(i).Item(13).ToString
        TextBox29.Text = ds.Tables("Billing").Rows(i).Item(14).ToString
        TextBox31.Text = ds.Tables("Billing").Rows(i).Item(15).ToString
        TextBox12.Text = ds.Tables("Billing").Rows(i).Item(17).ToString
        TextBox11.Text = ds.Tables("Billing").Rows(i).Item(18).ToString
        TextBox10.Text = ds.Tables("Billing").Rows(i).Item(19).ToString
        TextBox9.Text = ds.Tables("Billing").Rows(i).Item(20).ToString
        TextBox17.Text = ds.Tables("Billing").Rows(i).Item(21).ToString
        TextBox16.Text = ds.Tables("Billing").Rows(i).Item(22).ToString
        TextBox15.Text = ds.Tables("Billing").Rows(i).Item(23).ToString
        TextBox14.Text = ds.Tables("Billing").Rows(i).Item(24).ToString
        TextBox13.Text = ds.Tables("Billing").Rows(i).Item(25).ToString
        TextBox18.Text = ds.Tables("Billing").Rows(i).Item(26).ToString
        con.Close()

        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT ProdName, Price, Qty, Discount, Total FROM CLSaleProduct WHERE Cust_ID LIKE'%" &
        TextBox26.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("InvoiceProduct")
        adapter.Fill(dt)
        Me.DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 150
        Me.DataGridView1.Columns(8).Width = 150
        Me.DataGridView1.Columns(9).Width = 130
        Me.DataGridView1.Columns(10).Width = 130
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale Where BookingDate Between @d1 and @d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker1.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker2.Value.Date.AddDays(1)
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table
        myConnection.Close()
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale WHERE Cust_Name LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks FROM CLSale WHERE LensType ='" & ComboBox2.Text & "'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("Items")
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
                sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale WHERE Paymode LIKE'%" &
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
                    sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale WHERE Paymode LIKE'%" &
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
                        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, GrossAmt, NetAmt, PaidAmt, DueAmt, LensType, Paymode, Remarks from CLSale WHERE Paymode LIKE'%" &
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

    Sub Clear()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        TextBox1.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Clear()
        RefreshData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        i = 0
        Nav()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If (i = len - 1) Then
            MsgBox("This is the Last Record")
        Else
            i = i + 1
            Nav()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Print()
    End Sub

    Sub Print()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New CLJobCard 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select CLSale.Cust_ID, CLSale.Cust_Name, CLSale.Mobile, CLSale.BookingDate, CLSale.DeliveryDate, CLSale.GrossAmt, CLSale.ScmAmt, CLSale.REFBY, CLSale.NetAmt, CLSale.PaidAmt, CLSale.DueAmt, CLSale.Paymode, CLSale.RSPH, CLSale.RCYL, CLSale.RAXIS, CLSale.RVN, CLSale.LSPH, CLSale.LCYL, CLSale.LAXIS, CLSale.LVN, CLSale.LensType, CLSale.Remarks, CLSaleProduct.ProdName, CLSaleProduct.Qty, CLSaleProduct.Price, CLSaleProduct.Discount, CLSaleProduct.Total FROM CLSaleProduct INNER JOIN CLSale ON CLSale.Cust_ID=CLSaleProduct.Cust_ID Where CLSale.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox26.Text)
            MyCommand1.CommandText = "SELECT * from CLSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "CLSale")
            myDA.Fill(myDS, "CLSaleProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox26.Text)
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
            Dim rpt As New CLInvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select CLSale.Cust_ID, CLSale.Cust_Name, CLSale.Mobile, CLSale.BookingDate, CLSale.DeliveryDate, CLSale.REFBY, CLSale.GrossAmt, CLSale.CGST, CLSale.SGST, CLSale.ScmAmt, CLSale.NetAmt, CLSale.PaidAmt, CLSale.DueAmt, CLSale.Paymode, CLSale.RSPH, CLSale.RCYL, CLSale.RAXIS, CLSale.RVN, CLSale.LSPH, CLSale.LCYL, CLSale.LAXIS, CLSale.LVN, CLSale.LensType, CLSale.Remarks1, CLSaleProduct.ProdName, CLSaleProduct.Qty, CLSaleProduct.Price, CLSaleProduct.Discount, CLSaleProduct.Total FROM CLSaleProduct INNER JOIN CLSale ON CLSale.Cust_ID=CLSaleProduct.Cust_ID Where CLSale.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox26.Text)
            MyCommand1.CommandText = "SELECT * from CLSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "CLSale")
            myDA.Fill(myDS, "CLSaleProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox26.Text)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Print1()
    End Sub
End Class