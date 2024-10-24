Imports System.Data.OleDb
Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class frmPayRemainingMain

    Private Sub frmPayRemainingMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmPayRemainingMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmPayRemainingMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        RefreshData1()
        RefreshData2()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        autoID()
        autoIDSg()
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        TextBox6.Text = TextBox8.Text
        Compute()
        If Val(TextBox8.Text) > Val(TextBox4.Text) Then
            MsgBox("Final Pay cannot be more than Due Amount")
            TextBox8.Clear()
            TextBox8.Focus()
        End If
    End Sub

    Sub Compute()
        Dim i As Double
        i = Val(TextBox4.Text) - Val(TextBox8.Text)
        i = Math.Round(i, 2)
        TextBox7.Text = i
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
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
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, ReceiptNo As Bill_No, GrossAmt As [Net Payment], PaidAmt As [Advance Paid], DueAmt As [Balance Amt], Paymode, CGST, SGST, ScmAmt, NetAmt from SaleInvoice WHERE Paymode<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(8).Width = 110
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
    End Sub

    Private Sub RefreshData1()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, Cust_ID As Bill_No, GrossAmt As [Gross Payment], PaidAmt As [Advance Paid], DueAmt As [BalanceAmt], Paymode, CGST, SGST, ScmAmt, NetAmt from CLSale WHERE Paymode<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
    End Sub

    Private Sub RefreshData2()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView3.DataSource = Nothing
        DataGridView3.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, Cust_ID As Bill_No, GrossAmt As [Gross Payment], PaidAmt As [Advance Paid], DueAmt As [BalanceAmt], Paymode, CGST, SGST, ScmAmt, NetAmt from SunglassSale WHERE Paymode<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView3.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView3.Columns(9).Visible = False
        Me.DataGridView3.Columns(10).Visible = False
        Me.DataGridView3.Columns(11).Visible = False
        Me.DataGridView3.Columns(12).Visible = False
    End Sub

    Private Sub TextBox10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox10.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, ReceiptNo As Bill_No, GrossAmt As [Net Payment], PaidAmt As [Advance Paid], DueAmt As [Balance Amt], Paymode from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox10.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
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
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, ReceiptNo As Bill_No, GrossAmt As [Net Payment], PaidAmt As [Advance Paid], DueAmt As [Balance Amt], Paymode from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox1.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If TextBox8.Text = "" Then
                MessageBox.Show("Please enter Due Amount to Pay", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox8.Focus()
                Exit Sub
            End If
            If ComboBox1.Text = "" Then
                MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox7.Focus()
                Exit Sub
            End If
            If TextBox4.Text = "" Or TextBox4.Text = 0 Then
                MessageBox.Show("Due Amount is already Paid or Customer not Selected", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox7.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct1 As String = "select Cust_ID from BilledInvoice where Cust_ID=@d1"
            Dim cmd10 As OleDbCommand = New OleDbCommand(ct1)
            cmd10.Parameters.AddWithValue("@d1", TextBox27.Text)
            cmd10.Connection = myConnection
            rdr = cmd10.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                autoID()
            End If
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "update [SaleInvoice] set [NetAmt] = '" & TextBox5.Text & "', [PaidAmt] = '" & TextBox6.Text & "', [DueAmt] = '" & TextBox7.Text & "', [JobStatus] = '" & ComboBox1.Text & "' Where [Cust_ID] = '" & TextBox2.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str3 As String
            str3 = "INSERT INTO PaymentVoucher ([ReceiptName], [ReceiptNo], [ReceiptDate], [ReceiptAmt], [ReceiptBal], [ReceiptDue], [ReceiptStatus]) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str3, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("ReceiptName", CType(TextBox3.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox2.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDate", CType(DateTimePicker1.Value.Date, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptAmt", CType(TextBox5.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptBal", CType(TextBox6.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDue", CType(TextBox7.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptStatus", CType(ComboBox2.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str4 As String
            str4 = "INSERT INTO BilledInvoice ([Cust_ID], [Cust_Name], [Mobile], [BookingDate], [ReceiptNo], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd5 As OleDbCommand = New OleDbCommand(str4, myConnection)
            cmd5.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox27.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox3.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox9.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd5.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox2.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox5.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("CGST", CType(TextBox21.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("SGST", CType(TextBox22.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox23.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox24.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("PaidAmt", CType(TextBox6.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox7.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Paymode", CType(ComboBox2.Text, String)))
            cmd5.ExecuteNonQuery()
            cmd5.Dispose()
            myConnection.Close()

            MsgBox("Payment Submitted Successfuly...!", MsgBoxStyle.OkOnly)
            RefreshData()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim okToDelete As MsgBoxResult = MsgBox("Do you want to print Invoice", MsgBoxStyle.YesNo)
        If okToDelete = MsgBoxResult.Yes Then
            Dim i As Integer
            i = Me.DataGridView1.CurrentRow.Index
            frmBilledInvoicePrint.TextBox1.Text = Me.TextBox27.Text
            frmBilledInvoicePrint.TextBox2.Text = Me.TextBox3.Text
            frmBilledInvoicePrint.TextBox3.Text = Me.TextBox9.Text
            frmBilledInvoicePrint.DateTimePicker1.Value = Me.TextBox25.Text
            frmBilledInvoicePrint.TextBox5.Text = Me.TextBox26.Text
            frmBilledInvoicePrint.TextBox25.Text = Me.TextBox5.Text
            frmBilledInvoicePrint.TextBox37.Text = Me.TextBox21.Text
            frmBilledInvoicePrint.TextBox38.Text = Me.TextBox22.Text
            frmBilledInvoicePrint.TextBox26.Text = Me.TextBox23.Text
            frmBilledInvoicePrint.TextBox27.Text = Me.TextBox24.Text
            frmBilledInvoicePrint.TextBox28.Text = Me.TextBox6.Text
            frmBilledInvoicePrint.TextBox29.Text = Me.TextBox7.Text
            frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox2.Text
            frmBilledInvoicePrint.Button1.Enabled = False
            frmBilledInvoicePrint.Button2.Enabled = False
            frmBilledInvoicePrint.Button3.Enabled = True
            frmBilledInvoicePrint.ShowDialog()
        Else
        End If
        Clear()
        autoID()
    End Sub

    Sub Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox9.Clear()
        ComboBox2.SelectedIndex = -1
        ComboBox1.SelectedIndex = -1
        TextBox8.Clear()
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        Me.TextBox2.Text = DataGridView1.Item(0, i).Value.ToString
        Me.TextBox3.Text = DataGridView1.Item(1, i).Value.ToString
        Me.TextBox9.Text = DataGridView1.Item(2, i).Value.ToString
        Me.TextBox5.Text = DataGridView1.Item(5, i).Value.ToString
        Me.TextBox6.Text = DataGridView1.Item(6, i).Value.ToString
        Me.TextBox7.Text = DataGridView1.Item(7, i).Value.ToString
        Me.TextBox4.Text = DataGridView1.Item(7, i).Value.ToString
        Me.ComboBox2.Text = DataGridView1.Item(8, i).Value.ToString
        Me.TextBox21.Text = DataGridView1.Item(9, i).Value.ToString
        Me.TextBox22.Text = DataGridView1.Item(10, i).Value.ToString
        Me.TextBox23.Text = DataGridView1.Item(11, i).Value.ToString
        Me.TextBox24.Text = DataGridView1.Item(12, i).Value.ToString
        Me.TextBox25.Text = DataGridView1.Item(3, i).Value.ToString
        Me.TextBox26.Text = DataGridView1.Item(4, i).Value.ToString
    End Sub

    Private Sub DataGridView1_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim dgv As DataGridView = Me.DataGridView1
        For i As Integer = 0 To dgv.Rows.Count - 1
            For ColNo As Integer = 0 To 4
                If dgv.Rows(i).Cells(7).Value = 0 Then
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Green
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                Else
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Red
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                End If
            Next
        Next
    End Sub

    Private Sub DataGridView2_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim dgv As DataGridView = Me.DataGridView2
        For i As Integer = 0 To dgv.Rows.Count - 1
            For ColNo As Integer = 0 To 4
                If dgv.Rows(i).Cells(7).Value = 0 Then
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Green
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                Else
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Red
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                End If
            Next
        Next
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If Me.TextBox4.Text = "0" Then
            ComboBox1.Enabled = False
            TextBox8.Enabled = False
            ComboBox2.Enabled = False
            Button2.Enabled = False
        Else
            ComboBox1.Enabled = True
            TextBox8.Enabled = True
            ComboBox2.Enabled = True
            Button2.Enabled = True
        End If
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, Cust_ID As Bill_No, GrossAmt As Total_Transaction, PaidAmt As Credit, DueAmt As Debit, Paymode from CLSale WHERE Cust_Name LIKE'%" &
        TextBox12.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, BookingDate As Bill_Date, Cust_ID As Bill_No, GrossAmt As Total_Transaction, PaidAmt As Credit, DueAmt As Debit, Paymode from CLSale WHERE Mobile LIKE'%" &
        TextBox13.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox12.Clear()
        TextBox13.Clear()
        RefreshData1()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ExportExcel(DataGridView2)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox15.TextChanged
        TextBox17.Text = TextBox15.Text
        Compute1()
        If Val(TextBox15.Text) > Val(TextBox14.Text) Then
            MsgBox("Final Pay cannot be more than Due Amount")
            TextBox15.Clear()
            TextBox15.Focus()
        End If
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView2.CurrentRow.Index
        Me.TextBox20.Text = DataGridView2.Item(0, i).Value.ToString
        Me.TextBox19.Text = DataGridView2.Item(1, i).Value.ToString
        Me.TextBox35.Text = DataGridView2.Item(2, i).Value.ToString
        Me.DateTimePicker2.Value = DataGridView2.Item(3, i).Value.ToString
        Me.TextBox30.Text = DataGridView2.Item(0, i).Value.ToString
        Me.TextBox18.Text = DataGridView2.Item(5, i).Value.ToString
        Me.TextBox17.Text = DataGridView2.Item(6, i).Value.ToString
        Me.TextBox16.Text = DataGridView2.Item(7, i).Value.ToString
        Me.TextBox14.Text = DataGridView2.Item(7, i).Value.ToString
        Me.ComboBox4.Text = DataGridView2.Item(8, i).Value.ToString
        Me.TextBox34.Text = DataGridView2.Item(9, i).Value.ToString
        Me.TextBox33.Text = DataGridView2.Item(10, i).Value.ToString
        Me.TextBox32.Text = DataGridView2.Item(11, i).Value.ToString
        Me.TextBox31.Text = DataGridView2.Item(12, i).Value.ToString
    End Sub

    Sub Compute1()
        Dim i As Double
        i = Val(TextBox14.Text) - Val(TextBox15.Text)
        i = Math.Round(i, 2)
        TextBox16.Text = i
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            If TextBox15.Text = "" Then
                MessageBox.Show("Please enter Due Amount to Pay", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox15.Focus()
                Exit Sub
            End If
            If ComboBox3.Text = "" Then
                MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox3.Focus()
                Exit Sub
            End If
            If TextBox14.Text = "" Or TextBox14.Text = 0 Then
                MessageBox.Show("Due Amount is already Paid or Customer not Selected", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox14.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct1 As String = "select Cust_ID from BilledInvoice where Cust_ID=@d1"
            Dim cmd10 As OleDbCommand = New OleDbCommand(ct1)
            cmd10.Parameters.AddWithValue("@d1", TextBox29.Text)
            cmd10.Connection = myConnection
            rdr = cmd10.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                autoID()
            End If
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "update [CLSale] set [NetAmt] = '" & TextBox18.Text & "', [PaidAmt] = '" & TextBox17.Text & "', [DueAmt] = '" & TextBox16.Text & "', [Status] = '" & ComboBox3.Text & "' Where [Cust_ID] = '" & TextBox20.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str3 As String
            str3 = "INSERT INTO PaymentVoucher ([ReceiptName], [ReceiptNo], [ReceiptDate], [ReceiptAmt], [ReceiptBal], [ReceiptDue], [ReceiptStatus]) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str3, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("ReceiptName", CType(TextBox19.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox20.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDate", CType(DateTimePicker2.Value.Date, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptAmt", CType(TextBox18.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptBal", CType(TextBox17.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDue", CType(TextBox16.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptStatus", CType(ComboBox4.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str4 As String
            str4 = "INSERT INTO BilledInvoice ([Cust_ID], [Cust_Name], [Mobile], [BookingDate], [ReceiptNo], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd5 As OleDbCommand = New OleDbCommand(str4, myConnection)
            cmd5.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox29.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox19.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox35.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker2.Value.Date, String)))
            cmd5.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox30.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox18.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("CGST", CType(TextBox34.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("SGST", CType(TextBox33.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox31.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox32.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("PaidAmt", CType(TextBox17.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox16.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Paymode", CType(ComboBox4.Text, String)))
            cmd5.ExecuteNonQuery()
            cmd5.Dispose()
            myConnection.Close()
            MsgBox("Payment Submitted Successfuly...!", MsgBoxStyle.OkOnly)
            RefreshData1()

            Dim okToDelete As MsgBoxResult = MsgBox("Do you want to print Invoice", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                Dim i As Integer
                i = Me.DataGridView1.CurrentRow.Index
                frmBilledInvoicePrint.TextBox1.Text = Me.TextBox29.Text
                frmBilledInvoicePrint.TextBox2.Text = Me.TextBox19.Text
                frmBilledInvoicePrint.TextBox3.Text = Me.TextBox35.Text
                frmBilledInvoicePrint.DateTimePicker1.Value = Me.DateTimePicker2.Value.Date
                frmBilledInvoicePrint.TextBox5.Text = Me.TextBox30.Text
                frmBilledInvoicePrint.TextBox25.Text = Me.TextBox18.Text
                frmBilledInvoicePrint.TextBox37.Text = Me.TextBox34.Text
                frmBilledInvoicePrint.TextBox38.Text = Me.TextBox33.Text
                frmBilledInvoicePrint.TextBox26.Text = Me.TextBox31.Text
                frmBilledInvoicePrint.TextBox27.Text = Me.TextBox32.Text
                frmBilledInvoicePrint.TextBox28.Text = Me.TextBox17.Text
                frmBilledInvoicePrint.TextBox29.Text = Me.TextBox16.Text
                frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox4.Text
                frmBilledInvoicePrint.Button1.Enabled = False
                frmBilledInvoicePrint.Button3.Enabled = False
                frmBilledInvoicePrint.Button2.Enabled = True
                frmBilledInvoicePrint.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Clear1()
    End Sub

    Sub Clear1()
        TextBox20.Clear()
        TextBox19.Clear()
        TextBox18.Clear()
        TextBox17.Clear()
        TextBox16.Clear()
        DateTimePicker2.Value = Date.Now
        ComboBox4.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        TextBox14.Clear()
        TextBox15.Clear()
    End Sub

    Private Function GenerateCode() As String
        Dim con As OleDbConnection
        Dim cs As String
        Dim cmd As OleDbCommand
        Dim rdr As OleDbDataReader
        cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        con = New OleDbConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM BilledInvoice ORDER BY ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function

    Sub autoID()
        Try
            TextBox27.Text = GenerateCode()
            TextBox27.Text = "INV-" + GenerateCode()
            TextBox29.Text = "INV-" + GenerateCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        frmBilledInvoicePrint.TextBox1.Text = Me.TextBox27.Text
        frmBilledInvoicePrint.TextBox2.Text = Me.TextBox3.Text
        frmBilledInvoicePrint.TextBox3.Text = Me.TextBox9.Text
        frmBilledInvoicePrint.DateTimePicker1.Value = Me.TextBox25.Text
        frmBilledInvoicePrint.TextBox5.Text = Me.TextBox26.Text
        frmBilledInvoicePrint.TextBox25.Text = Me.TextBox5.Text
        frmBilledInvoicePrint.TextBox37.Text = Me.TextBox21.Text
        frmBilledInvoicePrint.TextBox38.Text = Me.TextBox22.Text
        frmBilledInvoicePrint.TextBox26.Text = Me.TextBox23.Text
        frmBilledInvoicePrint.TextBox27.Text = Me.TextBox24.Text
        frmBilledInvoicePrint.TextBox28.Text = Me.TextBox6.Text
        frmBilledInvoicePrint.TextBox29.Text = Me.TextBox7.Text
        frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox2.Text
        frmBilledInvoicePrint.Button1.Enabled = False
        frmBilledInvoicePrint.Button2.Enabled = False
        frmBilledInvoicePrint.Button3.Enabled = True
        frmBilledInvoicePrint.ShowDialog()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        frmBilledInvoicePrint.TextBox1.Text = Me.TextBox29.Text
        frmBilledInvoicePrint.TextBox2.Text = Me.TextBox19.Text
        frmBilledInvoicePrint.TextBox3.Text = Me.TextBox35.Text
        frmBilledInvoicePrint.DateTimePicker1.Value = Me.DateTimePicker2.Value.Date
        frmBilledInvoicePrint.TextBox5.Text = Me.TextBox30.Text
        frmBilledInvoicePrint.TextBox25.Text = Me.TextBox18.Text
        frmBilledInvoicePrint.TextBox37.Text = Me.TextBox34.Text
        frmBilledInvoicePrint.TextBox38.Text = Me.TextBox33.Text
        frmBilledInvoicePrint.TextBox26.Text = Me.TextBox31.Text
        frmBilledInvoicePrint.TextBox27.Text = Me.TextBox32.Text
        frmBilledInvoicePrint.TextBox28.Text = Me.TextBox17.Text
        frmBilledInvoicePrint.TextBox29.Text = Me.TextBox16.Text
        frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox4.Text
        frmBilledInvoicePrint.Button1.Enabled = False
        frmBilledInvoicePrint.Button3.Enabled = False
        frmBilledInvoicePrint.Button2.Enabled = True
        frmBilledInvoicePrint.ShowDialog()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        RefreshData2()
        TextBox36.Clear()
        TextBox37.Clear()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        ExportExcel(DataGridView3)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub DataGridView3_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView3.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView3.CurrentRow.Index
        Me.TextBox46.Text = DataGridView3.Item(0, i).Value.ToString
        Me.TextBox45.Text = DataGridView3.Item(1, i).Value.ToString
        Me.TextBox39.Text = DataGridView3.Item(2, i).Value.ToString
        Me.DateTimePicker3.Value = DataGridView3.Item(3, i).Value.ToString
        Me.TextBox46.Text = DataGridView3.Item(4, i).Value.ToString
        Me.TextBox44.Text = DataGridView3.Item(5, i).Value.ToString
        Me.TextBox43.Text = DataGridView3.Item(6, i).Value.ToString
        Me.TextBox42.Text = DataGridView3.Item(7, i).Value.ToString
        Me.TextBox40.Text = DataGridView3.Item(7, i).Value.ToString
        Me.ComboBox6.Text = DataGridView3.Item(8, i).Value.ToString
        Me.TextBox53.Text = DataGridView3.Item(9, i).Value.ToString
        Me.TextBox52.Text = DataGridView3.Item(10, i).Value.ToString
        Me.TextBox50.Text = DataGridView3.Item(11, i).Value.ToString
        Me.TextBox51.Text = DataGridView3.Item(12, i).Value.ToString
    End Sub

    Private Sub DataGridView3_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView3.RowPostPaint
        Dim dgv As DataGridView = Me.DataGridView3
        For i As Integer = 0 To dgv.Rows.Count - 1
            For ColNo As Integer = 0 To 4
                If dgv.Rows(i).Cells(7).Value = 0 Then
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Green
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                Else
                    dgv.Rows(i).Cells(7).Style.BackColor = Color.Red
                    dgv.Rows(i).Cells(7).Style.ForeColor = Color.White
                End If
            Next
        Next
    End Sub

    Private Sub TextBox41_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox41.TextChanged
        TextBox43.Text = TextBox41.Text
        Compute4()
        If Val(TextBox41.Text) > Val(TextBox40.Text) Then
            MsgBox("Final Pay cannot be more than Due Amount")
            TextBox41.Clear()
            TextBox41.Focus()
        End If
    End Sub

    Sub Compute4()
        Dim i As Double
        i = Val(TextBox40.Text) - Val(TextBox41.Text)
        i = Math.Round(i, 2)
        TextBox42.Text = i
    End Sub

    Sub Clear2()
        Me.TextBox46.Clear()
        Me.TextBox45.Clear()
        Me.TextBox39.Clear()
        Me.DateTimePicker3.Value = Date.Now
        Me.TextBox46.Clear()
        Me.TextBox44.Clear()
        Me.TextBox53.Clear()
        Me.TextBox52.Clear()
        Me.TextBox51.Clear()
        Me.TextBox43.Clear()
        Me.TextBox42.Clear()
        Me.ComboBox6.SelectedIndex = -1
        Me.TextBox50.Clear()
        Me.TextBox49.Clear()
    End Sub

    Private Function GenerateCodeSg() As String
        Dim con As OleDbConnection
        Dim cs As String
        Dim cmd As OleDbCommand
        Dim rdr As OleDbDataReader
        cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        con = New OleDbConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM SunglassBilledInvoice ORDER BY ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function

    Sub autoIDSg()
        Try
            TextBox49.Text = GenerateCodeSg()
            TextBox49.Text = "INV-" + GenerateCodeSg()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            If TextBox41.Text = "" Then
                MessageBox.Show("Please enter Due Amount to Pay", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox41.Focus()
                Exit Sub
            End If
            If ComboBox5.Text = "" Then
                MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox5.Focus()
                Exit Sub
            End If
            If TextBox40.Text = "" Or TextBox40.Text = 0 Then
                MessageBox.Show("Due Amount is already Paid or Customer not Selected", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox41.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct1 As String = "select Cust_ID from SunglassBilledInvoice where Cust_ID=@d1"
            Dim cmd10 As OleDbCommand = New OleDbCommand(ct1)
            cmd10.Parameters.AddWithValue("@d1", TextBox49.Text)
            cmd10.Connection = myConnection
            rdr = cmd10.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                autoID()
            End If
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "update [SunglassSale] set [NetAmt] = '" & TextBox44.Text & "', [PaidAmt] = '" & TextBox43.Text & "', [DueAmt] = '" & TextBox42.Text & "' Where [Cust_ID] = '" & TextBox46.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str3 As String
            str3 = "INSERT INTO PaymentVoucher ([ReceiptName], [ReceiptNo], [ReceiptDate], [ReceiptAmt], [ReceiptBal], [ReceiptDue], [ReceiptStatus]) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str3, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("ReceiptName", CType(TextBox45.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox46.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDate", CType(DateTimePicker3.Value.Date, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptAmt", CType(TextBox44.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptBal", CType(TextBox43.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDue", CType(TextBox42.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptStatus", CType(ComboBox6.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str4 As String
            str4 = "INSERT INTO SunglassBilledInvoice ([Cust_ID], [Cust_Name], [Mobile], [BookingDate], [ReceiptNo], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd5 As OleDbCommand = New OleDbCommand(str4, myConnection)
            cmd5.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox49.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox45.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox39.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker3.Value.Date, String)))
            cmd5.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox46.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox44.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("CGST", CType(TextBox53.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("SGST", CType(TextBox52.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox50.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox51.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("PaidAmt", CType(TextBox43.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox42.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Paymode", CType(ComboBox6.Text, String)))
            cmd5.ExecuteNonQuery()
            cmd5.Dispose()
            myConnection.Close()

            MsgBox("Payment Submitted Successfuly...!", MsgBoxStyle.OkOnly)
            RefreshData()

            Dim okToDelete As MsgBoxResult = MsgBox("Do you want to print Invoice", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                Dim i As Integer
                i = Me.DataGridView1.CurrentRow.Index
                frmBilledInvoicePrint.TextBox1.Text = Me.TextBox49.Text
                frmBilledInvoicePrint.TextBox2.Text = Me.TextBox45.Text
                frmBilledInvoicePrint.TextBox3.Text = Me.TextBox39.Text
                frmBilledInvoicePrint.DateTimePicker1.Value = Me.DateTimePicker3.Value
                frmBilledInvoicePrint.TextBox5.Text = Me.TextBox46.Text
                frmBilledInvoicePrint.TextBox25.Text = Me.TextBox44.Text
                frmBilledInvoicePrint.TextBox37.Text = Me.TextBox53.Text
                frmBilledInvoicePrint.TextBox38.Text = Me.TextBox52.Text
                frmBilledInvoicePrint.TextBox26.Text = Me.TextBox50.Text
                frmBilledInvoicePrint.TextBox27.Text = Me.TextBox51.Text
                frmBilledInvoicePrint.TextBox28.Text = Me.TextBox43.Text
                frmBilledInvoicePrint.TextBox29.Text = Me.TextBox42.Text
                frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox6.Text
                frmBilledInvoicePrint.Button1.Enabled = True
                frmBilledInvoicePrint.Button2.Enabled = False
                frmBilledInvoicePrint.Button3.Enabled = False
                frmBilledInvoicePrint.ShowDialog()
            Else
            End If
            Clear2()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
        autoIDSg()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        frmBilledInvoicePrint.TextBox1.Text = Me.TextBox49.Text
        frmBilledInvoicePrint.TextBox2.Text = Me.TextBox45.Text
        frmBilledInvoicePrint.TextBox3.Text = Me.TextBox39.Text
        frmBilledInvoicePrint.DateTimePicker1.Value = Me.DateTimePicker3.Value
        frmBilledInvoicePrint.TextBox5.Text = Me.TextBox46.Text
        frmBilledInvoicePrint.TextBox25.Text = Me.TextBox44.Text
        frmBilledInvoicePrint.TextBox37.Text = Me.TextBox53.Text
        frmBilledInvoicePrint.TextBox38.Text = Me.TextBox52.Text
        frmBilledInvoicePrint.TextBox26.Text = Me.TextBox50.Text
        frmBilledInvoicePrint.TextBox27.Text = Me.TextBox51.Text
        frmBilledInvoicePrint.TextBox28.Text = Me.TextBox43.Text
        frmBilledInvoicePrint.TextBox29.Text = Me.TextBox42.Text
        frmBilledInvoicePrint.TextBox6.Text = Me.ComboBox6.Text
        frmBilledInvoicePrint.Button1.Enabled = True
        frmBilledInvoicePrint.Button2.Enabled = False
        frmBilledInvoicePrint.Button3.Enabled = False
        frmBilledInvoicePrint.ShowDialog()
    End Sub
End Class