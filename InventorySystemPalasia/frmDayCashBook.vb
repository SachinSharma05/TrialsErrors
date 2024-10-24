Imports System.Data.OleDb

Public Class frmDayCashBook

    Private Sub frmDayCashBook_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmDayCashBook_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmDayCashBook_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        Try
            Dim Paid As Integer
            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                Paid += row.Cells(3).Value
            Next
            Me.TextBox2.Text = Paid
        Catch ex As Exception
            MsgBox("No record to calculate")
        End Try
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
        Dim str As String = "SELECT ReceiptName As Cust_Name, ReceiptNo As Cust_ID, ReceiptDate As Booking_Date, ReceiptBal As Amt_Received, ReceiptStatus As Mode_Of_Payment from PaymentVoucher WHERE ReceiptBal IS NOT NULL AND ReceiptBal<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView1.Columns(0).Width = 345
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 200
        Me.DataGridView1.Columns(3).Width = 200
        Me.DataGridView1.Columns(4).Width = 200
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If ComboBox1.Text <> "" Then
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim table As New DataTable
            Dim Command As New OleDbCommand("select ReceiptName As Cust_Name, ReceiptNo As Cust_ID, ReceiptDate As Booking_Date, ReceiptBal As Amt_Received, ReceiptStatus As Mode_Of_Payment from PaymentVoucher Where ReceiptDate >=@d1 And ReceiptDate <@d2 And ReceiptStatus=@d3 And ReceiptBal<>''", myConnection)
            Command.Parameters.Add("@d1", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker1.Value.Date
            Command.Parameters.Add("@d2", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker2.Value.Date.AddDays(1)
            Command.Parameters.AddWithValue("@d3", ComboBox1.Text)
            Dim adapter As New OleDbDataAdapter(Command)
            adapter.Fill(table)
            DataGridView1.DataSource = table.DefaultView
            myConnection.Close()
        Else
            If ComboBox1.Text = "" Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                'Change the following to your access database location
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim table As New DataTable
                Dim Command As New OleDbCommand("select ReceiptName As Cust_Name, ReceiptNo As Cust_ID, ReceiptDate As Booking_Date, ReceiptBal As Amt_Received, ReceiptStatus As Mode_Of_Payment from PaymentVoucher Where ReceiptDate >=@d1 And ReceiptDate <@d2 And ReceiptBal<>''", myConnection)
                Command.Parameters.Add("@d1", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker1.Value.Date
                Command.Parameters.Add("@d2", OleDbType.Date, 30, "ReceiptDate").Value = DateTimePicker2.Value.Date.AddDays(1)
                Dim adapter As New OleDbDataAdapter(Command)
                adapter.Fill(table)
                DataGridView1.DataSource = table.DefaultView
                myConnection.Close()
            End If
        End If

        Dim Paid As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Paid += row.Cells(3).Value
        Next
        Me.TextBox2.Text = Paid
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        TextBox1.Clear()
        ComboBox1.SelectedIndex = -1
        RefreshData()
        Dim Paid As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Paid += row.Cells(3).Value
        Next
        Me.TextBox2.Text = Paid
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT ReceiptName As Cust_Name, ReceiptNo As Cust_ID, ReceiptDate As Booking_Date, ReceiptBal As Amt_Received, ReceiptStatus As Mode_Of_Payment from PaymentVoucher WHERE ReceiptName LIKE'%" &
        TextBox1.Text & "%' AND ReceiptBal<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Dim Paid As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            Paid += row.Cells(3).Value
        Next
        Me.TextBox2.Text = Paid

        Me.DataGridView1.Columns(0).Width = 345
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 200
        Me.DataGridView1.Columns(3).Width = 200
        Me.DataGridView1.Columns(4).Width = 200
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub
End Class