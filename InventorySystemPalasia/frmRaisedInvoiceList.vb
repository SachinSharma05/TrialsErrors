Imports System.Data.OleDb

Public Class frmRaisedInvoiceList

    Private Sub frmRaisedInvoiceList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmRaisedInvoiceList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmRaisedInvoiceList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData12()
        RefreshData18()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub RefreshData12()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate, ReceiptNo, GrossAmt, CGST, SGST, ScmAmt, NetAmt, PaidAmt, DueAmt, Paymode from BilledInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub RefreshData18()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate, ReceiptNo, GrossAmt, CGST, SGST, ScmAmt, NetAmt, PaidAmt, DueAmt, Paymode from SunglassBilledInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ExportExcel(DataGridView2)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Clear()
        RefreshData12()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox2.Clear()
        RefreshData18()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * FROM SunglassBilledInvoice WHERE Cust_Name LIKE'%" &
        TextBox2.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
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
        sqlsearch = "SELECT * FROM BilledInvoice WHERE Cust_Name LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select * from BilledInvoice Where BookingDate >=@d1 And BookingDate <@d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date.AddDays(1)
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table.DefaultView
        myConnection.Close()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select * from SunglassBilledInvoice Where BookingDate >=@d1 And BookingDate <@d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker4.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker3.Value.Date.AddDays(1)
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView2.DataSource = table.DefaultView
        myConnection.Close()
    End Sub
End Class