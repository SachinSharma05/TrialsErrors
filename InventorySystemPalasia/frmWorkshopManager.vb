Imports System.Data.OleDb
Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class frmWorkshopManager

    Dim st2 As String

    Private Sub frmWorkshopManager_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmWorkshopManager_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmWorkshopManager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DateTimePicker1.Value = Date.Now
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice"
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
        Me.DataGridView1.Columns(3).Width = 150
        Me.DataGridView1.Columns(6).Width = 150
        Me.DataGridView1.Columns(7).Width = 200
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If ComboBox2.Text = "" Then
                MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox2.Focus()
                Exit Sub
            End If
            If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" Then
                MessageBox.Show("Please select Customer Detail first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "update [SaleInvoice] set [JobStatus] = '" & ComboBox2.Text & "' Where [Cust_ID] = '" & TextBox2.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MsgBox("Job Status Updated Successfully", vbOKOnly)

            Dim okToPrint As MsgBoxResult = MsgBox("Do you want to send SMS", MsgBoxStyle.YesNoCancel)
            If okToPrint = MsgBoxResult.Yes Then
                SendSMS()
            ElseIf okToPrint = MsgBoxResult.No Then
                Clear()
                RefreshData()
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Clear()
        RefreshData()
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        Me.TextBox2.Text = DataGridView1.Item(0, i).Value.ToString
        Me.TextBox3.Text = DataGridView1.Item(1, i).Value.ToString
        Me.TextBox4.Text = DataGridView1.Item(2, i).Value.ToString
        Me.TextBox6.Text = DataGridView1.Item(3, i).Value.ToString
        Me.TextBox7.Text = DataGridView1.Item(4, i).Value.ToString
        Me.TextBox8.Text = DataGridView1.Item(5, i).Value.ToString
        Me.TextBox9.Text = DataGridView1.Item(6, i).Value.ToString
        Me.TextBox10.Text = DataGridView1.Item(7, i).Value.ToString
    End Sub

    Sub Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        ComboBox2.SelectedIndex = -1
        ComboBox1.SelectedIndex = -1
        TextBox1.Clear()
    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim dgv As DataGridView = Me.DataGridView1
        For i As Integer = 0 To dgv.Rows.Count - 1
            For ColNo As Integer = 0 To 4
                If dgv.Rows(i).Cells(7).Value.ToString = "ON PROCESS" Then
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Red
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                Else
                    If dgv.Rows(i).Cells(7).Value.ToString = "READY" Then
                        dgv.Rows(i).Cells(7).Style.BackColor = Color.Green
                        dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                    Else
                        If dgv.Rows(i).Cells(7).Value.ToString = "DELIVERED" Then
                            dgv.Rows(i).Cells(7).Style.BackColor = Color.Blue
                            dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                        End If
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice Where BookingDate =@d1", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker1.Value.Date
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table.DefaultView
        myConnection.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Text = "ON PROCESS" Then
                Dim sqlsearch As String
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                'Change the following to your access database location
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                sqlsearch = "SELECT  Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice WHERE JobStatus LIKE'%" &
                ComboBox1.Text & "%'"
                Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
                Dim dt As New DataTable("SaleInvoice")
                adapter.Fill(dt)
                DataGridView1.DataSource = dt
                myConnection.Close()
            Else
                If ComboBox1.Text = "READY" Then
                    Dim sqlsearch As String
                    provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                    'Change the following to your access database location
                    dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                    connString = provider & dataFile
                    myConnection.ConnectionString = connString
                    myConnection.Open()
                    sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice WHERE JobStatus LIKE'%" &
                    ComboBox1.Text & "%'"
                    Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
                    Dim dt As New DataTable("SaleInvoice")
                    adapter.Fill(dt)
                    DataGridView1.DataSource = dt
                    myConnection.Close()
                Else
                    If ComboBox1.Text = "DELIVERED" Then
                        Dim sqlsearch As String
                        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                        'Change the following to your access database location
                        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                        connString = provider & dataFile
                        myConnection.ConnectionString = connString
                        myConnection.Open()
                        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice WHERE JobStatus LIKE'%" &
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate, DeliveryDate, BookedBy, DeliveryTime, JobStatus from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox1.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Clear()
        ComboBox1.SelectedIndex = -1
        RefreshData()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Sub SendSMS()
        Try
            Dim url As String
            url = "http://alerts.valueleaf.com/api/v4/?api_key=A7ce7d9a7a5bcb5f1cfdc9e60b9095d8c&method=sms&message=" + Me.TextBox5.Text + "&to=" + Me.TextBox4.Text + "&sender=AOPTIC"
            Dim myReq As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            Dim myResp As HttpWebResponse = DirectCast(myReq.GetResponse(), HttpWebResponse)
            Dim respStreamReader As New System.IO.StreamReader(myResp.GetResponseStream())
            Dim responseString As String = respStreamReader.ReadToEnd()
            respStreamReader.Close()
            myResp.Close()
            MsgBox("Message Send Successfully")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "READY" Then
            TextBox5.Text = "Dear Customer, your order is ready for delivery, kindly come personally to get it checked. Thanks American Optics Palasia."
        Else
            If ComboBox2.Text = "DELIVERED" Then
                TextBox5.Text = "Your order is delivered, Thanks for your precious order, do visit again. Thanks American Optics Palasia."
            End If
        End If
    End Sub
End Class