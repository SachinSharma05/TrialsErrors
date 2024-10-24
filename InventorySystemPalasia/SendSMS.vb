Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class SendSMS

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Dim st2 As String
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox3.Text = "" Then
            MessageBox.Show("Please enter Mobile No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox3.Focus()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select Order Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MessageBox.Show("Please enter Message Body", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox4.Focus()
            Exit Sub
        End If
        Try
            Dim url As String
            url = "http://alerts.valueleaf.com/api/v4/?api_key=A7ce7d9a7a5bcb5f1cfdc9e60b9095d8c&method=sms&message=" + Me.TextBox4.Text + "&to=" + Me.TextBox3.Text + "&sender=AOPTIC"
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

    Private Sub SendSMS_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub SendSMS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub SendSMS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile from SaleInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView1.Columns(0).Width = 100
        Me.DataGridView1.Columns(1).Width = 250
        Me.DataGridView1.Columns(2).Width = 140
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "READY" Then
            TextBox4.Text = "Dear Customer, your order is ready for delivery, kindly come personally to get it checked. Thanks American Optical Co. Palasia."
        Else
            If ComboBox1.Text = "DELIVERED" Then
                TextBox4.Text = "Your order is delivered, Thanks for your precious order, do visit again. Thanks American Optical Co. Palasia."
            End If
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox5.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox6.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.TextBox1.Text = DataGridView1.Item(0, i).Value.ToString
            Me.TextBox2.Text = DataGridView1.Item(1, i).Value.ToString
            Me.TextBox3.Text = DataGridView1.Item(2, i).Value.ToString
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try
    End Sub
End Class