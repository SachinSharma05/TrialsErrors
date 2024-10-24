Imports System.Data.OleDb

Public Class frmStockInList

    Private Sub frmStockInList_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmStockInList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmStockInList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now

        Dim Currstock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Currstock += row.Cells(4).Value
        Next
        Me.TextBox2.Text = Currstock

        Dim Added As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Added += row.Cells(5).Value
        Next
        Me.TextBox3.Text = Added
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
        Dim str As String = "select AddType, AddCategory, AddSubCat, AddProdName, AddCurrStock, StockAdded, AddDate, AddedBy from StockAdded"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select AddType, AddCategory, AddSubCat, AddProdName, AddCurrStock, StockAdded, AddDate, AddedBy from StockAdded Where AddDate Between @d1 and @d2", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "AddDate").Value = DateTimePicker1.Value.Date
        Command.Parameters.Add("@d2", OleDbType.Date, 30, "AddDate").Value = DateTimePicker2.Value
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table
        myConnection.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Clear()
        RefreshData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT AddType, AddCategory, AddSubCat, AddProdName, AddCurrStock, StockAdded, AddDate, AddedBy FROM StockAdded WHERE AddProdName LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Currstock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Currstock += row.Cells(4).Value
        Next
        Me.TextBox2.Text = Currstock

        Dim Added As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Added += row.Cells(5).Value
        Next
        Me.TextBox3.Text = Added
    End Sub
End Class