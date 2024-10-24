Imports System.Data.OleDb

Public Class frmExchangeProducts

    Private Sub frmExchangeProducts_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmExchangeProducts_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmExchangeProducts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
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
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select * from SaleInvoice WHERE Paymode IS NOT NULL AND Paymode<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView2.Columns(0).Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox32.Clear()
        TextBox33.Clear()
        RefreshData()
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Try
            Dim i As Integer
            i = Me.DataGridView2.CurrentRow.Index
            Me.TextBox1.Text = Me.DataGridView2.Item(1, i).Value.ToString
            Me.TextBox2.Text = Me.DataGridView2.Item(2, i).Value.ToString
            Me.TextBox3.Text = Me.DataGridView2.Item(3, i).Value.ToString
            Me.TextBox4.Text = Me.DataGridView2.Item(4, i).Value.ToString
            Me.TextBox42.Text = Me.DataGridView2.Item(5, i).Value.ToString
            Me.DateTimePicker1.Value = Me.DataGridView2.Item(6, i).Value.ToString
            Me.TextBox7.Text = Me.DataGridView2.Item(7, i).Value.ToString
            Me.TextBox6.Text = Me.DataGridView2.Item(8, i).Value.ToString
            Me.TextBox5.Text = Me.DataGridView2.Item(9, i).Value.ToString
            Me.TextBox34.Text = Me.DataGridView2.Item(10, i).Value.ToString
            Me.TextBox41.Text = Me.DataGridView2.Item(11, i).Value.ToString
            Me.TextBox43.Text = Me.DataGridView2.Item(12, i).Value.ToString
            Me.TextBox35.Text = Me.DataGridView2.Item(13, i).Value.ToString
            Me.TextBox30.Text = Me.DataGridView2.Item(14, i).Value.ToString
            Me.TextBox37.Text = Me.DataGridView2.Item(15, i).Value.ToString
            Me.TextBox38.Text = Me.DataGridView2.Item(16, i).Value.ToString
            Me.TextBox26.Text = Me.DataGridView2.Item(17, i).Value.ToString
            Me.TextBox27.Text = Me.DataGridView2.Item(18, i).Value.ToString
            Me.TextBox28.Text = Me.DataGridView2.Item(19, i).Value.ToString
            Me.TextBox29.Text = Me.DataGridView2.Item(20, i).Value.ToString
            Me.TextBox31.Text = Me.DataGridView2.Item(21, i).Value.ToString
            Me.TextBox25.Text = Me.DataGridView2.Item(22, i).Value.ToString
            Me.TextBox12.Text = Me.DataGridView2.Item(23, i).Value.ToString
            Me.TextBox11.Text = Me.DataGridView2.Item(24, i).Value.ToString
            Me.TextBox10.Text = Me.DataGridView2.Item(25, i).Value.ToString
            Me.TextBox9.Text = Me.DataGridView2.Item(26, i).Value.ToString
            Me.TextBox17.Text = Me.DataGridView2.Item(27, i).Value.ToString
            Me.TextBox16.Text = Me.DataGridView2.Item(28, i).Value.ToString
            Me.TextBox15.Text = Me.DataGridView2.Item(29, i).Value.ToString
            Me.TextBox14.Text = Me.DataGridView2.Item(30, i).Value.ToString
            Me.TextBox13.Text = Me.DataGridView2.Item(31, i).Value.ToString
            Me.TextBox18.Text = Me.DataGridView2.Item(32, i).Value.ToString
            Me.TextBox19.Text = Me.DataGridView2.Item(33, i).Value.ToString
            Me.TextBox20.Text = Me.DataGridView2.Item(34, i).Value.ToString
            Me.TextBox21.Text = Me.DataGridView2.Item(35, i).Value.ToString
            Me.TextBox22.Text = Me.DataGridView2.Item(36, i).Value.ToString
            Me.TextBox23.Text = Me.DataGridView2.Item(37, i).Value.ToString
            Me.TextBox24.Text = Me.DataGridView2.Item(38, i).Value.ToString
            Me.TextBox36.Text = Me.DataGridView2.Item(40, i).Value.ToString
            Me.TextBox39.Text = Me.DataGridView2.Item(41, i).Value.ToString
            Me.TextBox40.Text = Me.DataGridView2.Item(42, i).Value.ToString

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT ProdName, Price, Qty, Discount, Total FROM InvoiceProduct WHERE Cust_ID LIKE'%" &
            TextBox7.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("InvoiceProduct")
            adapter.Fill(dt)
            Me.DataGridView1.DataSource = dt
            myConnection.Close()

            Dim sqlsearch1 As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch1 = "SELECT ReceiptName, ReceiptNo, ReceiptDate, ReceiptAmt, ReceiptBal, ReceiptDue, ReceiptStatus FROM PaymentVoucher WHERE ReceiptNo LIKE'%" &
            TextBox1.Text & "%'"
            Dim adapter1 As New OleDbDataAdapter(sqlsearch1, myConnection)
            Dim dt1 As New DataTable("InvoiceProduct")
            adapter1.Fill(dt1)
            Me.DataGridView3.DataSource = dt1
            myConnection.Close()
        Catch ex As Exception
            MsgBox("Row is empty")
        End Try
    End Sub

    Private Sub TextBox32_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox32.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "select * from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox32.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        Me.DataGridView2.Columns(0).Visible = False
    End Sub

    Private Sub TextBox33_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox33.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "select * from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox33.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        Me.DataGridView2.Columns(0).Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If TextBox2.Text = "" Then
                MessageBox.Show("Please Select Customer First", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox2.Focus()
                Exit Sub
            End If

            Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to delete the current record?", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str As String
                str = "Delete from SaleInvoice Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
                Try
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str1 As String
                str1 = "Delete from InvoiceProduct Where Cust_ID = '" & Me.TextBox7.Text & "'"
                Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
                Try
                    cmd1.ExecuteNonQuery()
                    cmd1.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str2 As String
                str2 = "Delete from PaymentVoucher Where ReceiptNo = '" & Me.TextBox1.Text & "'"
                Dim cmd2 As OleDbCommand = New OleDbCommand(str2, myConnection)
                Try
                    cmd2.ExecuteNonQuery()
                    cmd2.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str3 As String
                str3 = "Delete from CustomerTable Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd3 As OleDbCommand = New OleDbCommand(str3, myConnection)
                Try
                    cmd3.ExecuteNonQuery()
                    cmd3.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            ElseIf okToDelete = MsgBoxResult.No Then
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str4 As String
            str4 = "insert into SaleInvoice ([Cust_ID], [Cust_Name], [Mobile], [Phone], [Address], [BookingDate], [Status], [Age], [Remarks], [RSPH], [RCYL], [RAXIS], [RVN], [RADD], [LSPH], [LCYL], [LAXIS], [LVN], [LADD], [PD], [REFBY], [LensType], [LensType1], [LensType2], [LensType3], [Right], [Left], [RLAdd]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str4, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox1.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Phone", CType(TextBox4.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Address", CType(TextBox42.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd4.Parameters.Add(New OleDbParameter("Status", CType(TextBox41.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Age", CType(TextBox43.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Remarks", CType(TextBox25.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RSPH", CType(TextBox12.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RCYL", CType(TextBox11.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RAXIS", CType(TextBox10.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RVN", CType(TextBox9.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RADD", CType(TextBox17.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LSPH", CType(TextBox16.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LCYL", CType(TextBox15.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LAXIS", CType(TextBox14.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LVN", CType(TextBox13.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LADD", CType(TextBox18.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("PD", CType(TextBox19.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("REFBY", CType(TextBox20.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LensType", CType(TextBox21.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LensType1", CType(TextBox22.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LensType2", CType(TextBox23.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("LensType3", CType(TextBox24.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Right", CType(TextBox36.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("Left", CType(TextBox39.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("RLAdd", CType(TextBox40.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MsgBox("Transferred Successfuly...!", MsgBoxStyle.OkOnly)

        Clear()
        DataGridView1.DataSource = Nothing
        DataGridView3.DataSource = Nothing
    End Sub

    Sub Clear()
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.DateTimePicker1.Value = Date.Now
        Me.TextBox7.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox34.Text = ""
        Me.TextBox35.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox17.Text = ""
        Me.TextBox16.Text = ""
        Me.TextBox15.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox18.Text = ""
        Me.TextBox19.Text = ""
        Me.TextBox20.Text = ""
        Me.TextBox21.Text = ""
        Me.TextBox22.Text = ""
        Me.TextBox23.Text = ""
        Me.TextBox24.Text = ""
        Me.TextBox25.Text = ""
        Me.TextBox30.Text = ""
        Me.TextBox26.Text = ""
        Me.TextBox27.Text = ""
        Me.TextBox28.Text = ""
        Me.TextBox29.Text = ""
        Me.TextBox37.Text = ""
        Me.TextBox38.Text = ""
        Me.TextBox31.Text = ""
        Me.TextBox41.Text = ""
        Me.TextBox42.Text = ""
        Me.TextBox43.Text = ""
    End Sub
End Class