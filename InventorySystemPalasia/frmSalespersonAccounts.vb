Imports System.Data.OleDb

Public Class frmSalespersonAccounts

    Private Sub frmSalespersonAccounts_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmSalespersonAccounts_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub frmSalespersonAccounts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        SunglassData()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "SELECT Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SaleInvoice where NetAmt<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        Spectacle()
    End Sub

    Private Sub SunglassData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SunglassSale"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using

        Sunglass()
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
            Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SaleInvoice Where BookingDate >=@d1 And BookingDate <@d2 And BookedBy=@d3 And NetAmt<>''", myConnection)
            Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
            Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date.AddDays(1)
            Command.Parameters.AddWithValue("@d3", ComboBox1.Text)
            Dim adapter As New OleDbDataAdapter(Command)
            adapter.Fill(table)
            DataGridView1.DataSource = table.DefaultView
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim table1 As New DataTable
            Dim Command1 As New OleDbCommand("select Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SunglassSale Where BookingDate >=@d1 And BookingDate <@d2 And BookedBy=@d3 And NetAmt<>''", myConnection)
            Command1.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
            Command1.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date.AddDays(1)
            Command1.Parameters.AddWithValue("@d3", ComboBox1.Text)
            Dim adapter1 As New OleDbDataAdapter(Command1)
            adapter1.Fill(table1)
            DataGridView2.DataSource = table1.DefaultView
            myConnection.Close()

            Sunglass()
            Spectacle()
        Else
            If ComboBox1.Text = "" Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                'Change the following to your access database location
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim table As New DataTable
                Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SaleInvoice Where BookingDate >=@d1 And BookingDate <@d2 And NetAmt<>''", myConnection)
                Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
                Command.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date.AddDays(1)
                Dim adapter As New OleDbDataAdapter(Command)
                adapter.Fill(table)
                DataGridView1.DataSource = table.DefaultView
                myConnection.Close()

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                'Change the following to your access database location
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim table1 As New DataTable
                Dim Command1 As New OleDbCommand("select Cust_ID, Cust_Name, BookingDate, BookedBy, NetAmt from SunglassSale Where BookingDate >=@d1 And BookingDate <@d2 And NetAmt<>''", myConnection)
                Command1.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
                Command1.Parameters.Add("@d2", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date.AddDays(1)
                Dim adapter1 As New OleDbDataAdapter(Command1)
                adapter1.Fill(table1)
                DataGridView2.DataSource = table1.DefaultView
                myConnection.Close()
            End If
            Sunglass()
            Spectacle()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        ComboBox1.SelectedIndex = -1
        RefreshData()
        SunglassData()
        Sunglass()
        Spectacle()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
        ExportExcel(DataGridView2)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Sub Sunglass()
        Try
            Dim Paid As Integer
            For Each row As DataGridViewRow In Me.DataGridView2.Rows
                Paid += row.Cells(4).Value
            Next
            Me.TextBox3.Text = Paid
        Catch ex As Exception
            MsgBox("No record to calculate")
        End Try
    End Sub

    Sub Spectacle()
        Try
            Dim Paid As Integer
            For Each row As DataGridViewRow In Me.DataGridView1.Rows
                Paid += row.Cells(4).Value
            Next
            Me.TextBox2.Text = Paid
        Catch ex As Exception
            MsgBox("No record to calculate")
        End Try
    End Sub
End Class