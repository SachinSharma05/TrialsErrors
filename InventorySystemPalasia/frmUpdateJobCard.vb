Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class frmUpdateJobCard

    Private Sub frmUpdateJobCard_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmUpdateJobCard_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmUpdateJobCard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        LoadCombo2()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        Clear1()
        RefreshData()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        autoID()
    End Sub

    Sub GridCopy()
        Dim sourceGrid As DataGridView = Me.DataGridView1
        Dim targetGrid As DataGridView = Me.DataGridView3
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)

                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next

                targetRows.Add(targetRow)
            End If
        Next

        targetGrid.Columns.Clear()

        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Sub LoadCombo()
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;")
        cn.Open()

        Dim cmd As New OleDbCommand("SELECT SName FROM Salesperson;", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox1.DisplayMember = "SName"
        ComboBox1.ValueMember = "SName"
        ComboBox1.DataSource = dt

        cn.Close()
    End Sub

    Sub LoadCombo2()
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;")
        cn.Open()

        Dim cmd As New OleDbCommand("SELECT Prod_Name FROM ItemMaster;", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox2.DisplayMember = "Prod_Name"
        ComboBox2.ValueMember = "Prod_Name"
        ComboBox2.DataSource = dt

        cn.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "SELECT Price, Stock, TaxCat FROM ItemMaster WHERE (Prod_Name = '" & ComboBox2.Text & "')"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        dr = cmd.ExecuteReader
        While dr.Read()
            TextBox6.Focus()
            TextBox7.Text = dr("Price").ToString
            TextBox9.Text = dr("Price").ToString
            TextBox32.Text = dr("Stock").ToString
            TextBox40.Text = dr("TaxCat").ToString
        End While
        myConnection.Close()
        TextBox50.Text = ComboBox2.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "UPDATE ItemMaster SET [Stock] = Stock - " & Val(TextBox6.Text) & " Where [Prod_Name] = '" & ComboBox2.Text & "'"
            Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            Try
                cmd1.ExecuteNonQuery()
                cmd1.Dispose()
                myConnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            If ComboBox2.Text = "" Then
                MessageBox.Show("Please select Product", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox2.Focus()
                Exit Sub
            End If
            If TextBox6.Text = "" Then
                MessageBox.Show("Please enter quantity", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox6.Focus()
                Exit Sub
            End If
            If TextBox6.Text = 0 Then
                MessageBox.Show("Quantity can not be zero", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox6.Focus()
                Exit Sub
            End If
            For i As Integer = 0 To DataGridView3.Rows.Count - 1
                DataGridView3.Rows.Add(ComboBox2.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, TextBox35.Text, TextBox36.Text)
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                TextBox25.Text = k
                Dim c As Double = 0
                c = TotalDiscount()
                c = Math.Round(c, 2)
                TextBox26.Text = c
                Dim x As Double = 0
                x = TotalPayment()
                x = Math.Round(x, 2)
                TextBox27.Text = x
                Dim g As Double = 0
                g = TaxCGST()
                g = Math.Round(g, 2)
                TextBox37.Text = g
                Dim s As Double = 0
                s = TaxSGST()
                s = Math.Round(s, 2)
                TextBox38.Text = s
                Compute1()
                Clear1()
                Exit Sub
            Next

            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                If r.Cells(0).Value = ComboBox2.Text Then
                    r.Cells(0).Value = ComboBox2.Text
                    r.Cells(3).Value = TextBox6.Text
                    r.Cells(4).Value = Val(r.Cells(4).Value) + Val(TextBox7.Text)
                    r.Cells(5).Value = Val(r.Cells(5).Value) + Val(TextBox8.Text)
                    r.Cells(6).Value = Val(r.Cells(6).Value) + Val(TextBox9.Text)
                    r.Cells(7).Value = Val(r.Cells(6).Value) + Val(TextBox10.Text)
                    r.Cells(8).Value = Val(r.Cells(8).Value) + Val(TextBox11.Text)
                    r.Cells(9).Value = Val(r.Cells(9).Value) + Val(TextBox17.Text)
                    r.Cells(10).Value = Val(r.Cells(10).Value) + Val(TextBox35.Text)
                    r.Cells(11).Value = Val(r.Cells(11).Value) + Val(TextBox36.Text)
                    Dim i As Double = 0
                    i = GrandTotal()
                    i = Math.Round(i, 2)
                    TextBox25.Text = i
                    Dim a As Double = 0
                    a = TotalDiscount()
                    a = Math.Round(a, 2)
                    TextBox26.Text = a
                    Dim q As Double
                    q = TotalPayment()
                    q = Math.Round(q, 2)
                    TextBox27.Text = q
                    Dim d As Double = 0
                    d = TaxCGST()
                    d = Math.Round(d, 2)
                    TextBox37.Text = d
                    Dim u As Double = 0
                    u = TaxSGST()
                    u = Math.Round(u, 2)
                    TextBox38.Text = u
                    Compute1()
                    Clear1()
                    Exit Sub
                End If
            Next

            DataGridView3.Rows.Add(ComboBox2.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, TextBox35.Text, TextBox36.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            TextBox25.Text = j
            Dim b As Double = 0
            b = TotalDiscount()
            b = Math.Round(b, 2)
            TextBox26.Text = b
            Dim z As Double
            z = TotalPayment()
            z = Math.Round(z, 2)
            TextBox27.Text = z
            Dim y As Double = 0
            y = TaxCGST()
            y = Math.Round(y, 2)
            TextBox37.Text = y
            Dim w As Double = 0
            w = TaxSGST()
            w = Math.Round(w, 2)
            TextBox38.Text = w
            Compute1()
            Clear1()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            For Each row As DataGridViewRow In DataGridView3.SelectedRows
                DataGridView3.Rows.Remove(row)
                If Not row.IsNewRow Then
                    myConnection = New OleDbConnection(provider)
                    myConnection.Open()
                    Dim cb4 As String = "update ItemMaster set Stock = Stock + (" & row.Cells(1).Value & ") where Prod_Name= Textbox50.text"
                    Dim cmd2 As New OleDbCommand
                    cmd2 = New OleDbCommand(cb4)
                    cmd2.Connection = myConnection
                    cmd2.Parameters.Add(New OleDbParameter("& Textbox50.text &", row.Cells(0).Value))
                    cmd2.ExecuteNonQuery()
                    myConnection.Close()
                End If
            Next

            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            TextBox25.Text = k
            Dim c As Double = 0
            c = TotalDiscount()
            c = Math.Round(c, 2)
            TextBox26.Text = c
            Dim x As Double
            x = TotalPayment()
            x = Math.Round(x, 2)
            TextBox27.Text = x
            Dim y As Double = 0
            y = TaxCGST()
            y = Math.Round(y, 2)
            TextBox37.Text = y
            Dim w As Double = 0
            w = TaxSGST()
            w = Math.Round(w, 2)
            TextBox38.Text = w
            Compute()
            Compute1()
            TextBox28.Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Compute()
        Dim num1, num2 As Double
        num1 = CDbl(Val(TextBox6.Text) * Val(TextBox7.Text))
        num1 = Math.Round(num1, 2)
        TextBox9.Text = num1
        TextBox10.Text = num1
        num2 = CDbl(Val(TextBox9.Text) - Val(TextBox8.Text))
        num2 = Math.Round(num2, 2)
        TextBox9.Text = num2
    End Sub

    Public Function TaxCGST() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(6).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function TaxSGST() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(7).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function GrandTotal() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(5).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function TotalPayment() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(4).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function TotalDiscount() As Double
        Dim Dis As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                Dis = Dis + r.Cells(3).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Dis
    End Function

    Sub Compute1()
        Dim i As Double = 0
        i = Val(TextBox25.Text) - Val(TextBox26.Text)
        i = Math.Round(i, 2)
        TextBox29.Text = i
    End Sub

    Sub Clear1()
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = ""
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox32.Clear()
        ComboBox2.Focus()
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
            MyCommand.Parameters.AddWithValue("@d1", TextBox5.Text)
            MyCommand1.CommandText = "SELECT * from SaleInvoice"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SaleInvoice")
            myDA.Fill(myDS, "InvoiceProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox5.Text)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            TextBox7.ReadOnly = True
        Else
            TextBox7.ReadOnly = False
        End If
        If TextBox8.Text = "" Then
            TextBox8.Text = "0"
        Else
        End If
        Compute()

        Dim i As Double = 0
        If TextBox40.Text = "Tax @ 12%" Then
            i = Val(TextBox9.Text) * 12 / 112
            i = i / 2
            i = Math.Round(i, 2)
            TextBox35.Text = i
            TextBox36.Text = i
        Else
            If TextBox40.Text = "Tax @ 18%" Then
                i = Val(TextBox9.Text) * 18 / 118
                i = i / 2
                i = Math.Round(i, 2)
                TextBox35.Text = i
                TextBox36.Text = i
            End If
        End If
    End Sub

    Sub Compute2()
        Dim i As Double = 0
        i = Val(TextBox27.Text) - Val(TextBox28.Text)
        i = Math.Round(i, 2)
        TextBox29.Text = i
    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged
        If Val(TextBox28.Text) > Val(TextBox27.Text) Then
            MsgBox("Advance cannot be more than Net Amount")
            TextBox28.Clear()
        End If
        Compute2()
    End Sub

    Private Sub TextBox28_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox28.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select * from SaleInvoice WHERE NetAmt<>''"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(5).Visible = False
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Try
            Me.DataGridView3.Rows.Clear()
            Me.Refresh()

            Dim i As Integer
            i = DataGridView2.CurrentRow.Index
            Me.TextBox1.Text = DataGridView2.Item(1, i).Value.ToString
            Me.TextBox2.Text = DataGridView2.Item(2, i).Value.ToString
            Me.TextBox3.Text = DataGridView2.Item(3, i).Value.ToString
            Me.TextBox4.Text = DataGridView2.Item(4, i).Value.ToString
            Me.DateTimePicker1.Value = DataGridView2.Item(6, i).Value.ToString
            Me.TextBox5.Text = DataGridView2.Item(7, i).Value.ToString
            Me.DateTimePicker2.Value = DataGridView2.Item(8, i).Value.ToString
            Me.ComboBox1.Text = DataGridView2.Item(9, i).Value.ToString
            Me.TextBox44.Text = DataGridView2.Item(10, i).Value.ToString
            Me.TextBox39.Text = DataGridView2.Item(11, i).Value.ToString
            Me.TextBox45.Text = DataGridView2.Item(12, i).Value.ToString
            Me.ComboBox3.Text = DataGridView2.Item(13, i).Value.ToString
            Me.TextBox25.Text = DataGridView2.Item(14, i).Value.ToString
            Me.TextBox37.Text = DataGridView2.Item(15, i).Value.ToString
            Me.TextBox38.Text = DataGridView2.Item(16, i).Value.ToString
            Me.TextBox26.Text = DataGridView2.Item(17, i).Value.ToString
            Me.TextBox27.Text = DataGridView2.Item(18, i).Value.ToString
            Me.TextBox28.Text = DataGridView2.Item(19, i).Value.ToString
            Me.TextBox29.Text = DataGridView2.Item(20, i).Value.ToString
            If DataGridView2.Item(21, i).Value.ToString = "Payment By Cash" Then
                Me.RadioButton1.Checked = True
            Else
                If DataGridView2.Item(21, i).Value.ToString = "Payment By Card" Then
                    Me.RadioButton2.Checked = True
                Else
                    If DataGridView2.Item(21, i).Value.ToString = "Payment By Both" Then
                        Me.RadioButton3.Checked = True
                    End If
                End If
            End If
            Me.TextBox41.Text = DataGridView2.Item(22, i).Value.ToString
            Me.TextBox11.Text = DataGridView2.Item(23, i).Value.ToString
            Me.TextBox12.Text = DataGridView2.Item(24, i).Value.ToString
            Me.TextBox13.Text = DataGridView2.Item(25, i).Value.ToString
            Me.TextBox14.Text = DataGridView2.Item(26, i).Value.ToString
            Me.TextBox15.Text = DataGridView2.Item(27, i).Value.ToString
            Me.TextBox16.Text = DataGridView2.Item(28, i).Value.ToString
            Me.TextBox17.Text = DataGridView2.Item(29, i).Value.ToString
            Me.TextBox18.Text = DataGridView2.Item(30, i).Value.ToString
            Me.TextBox19.Text = DataGridView2.Item(31, i).Value.ToString
            Me.TextBox20.Text = DataGridView2.Item(32, i).Value.ToString
            Me.TextBox21.Text = DataGridView2.Item(33, i).Value.ToString
            Me.TextBox22.Text = DataGridView2.Item(34, i).Value.ToString
            Me.TextBox23.Text = DataGridView2.Item(35, i).Value.ToString
            Me.TextBox46.Text = DataGridView2.Item(36, i).Value.ToString
            Me.TextBox47.Text = DataGridView2.Item(37, i).Value.ToString
            Me.TextBox48.Text = DataGridView2.Item(38, i).Value.ToString
            Me.TextBox24.Text = DataGridView2.Item(39, i).Value.ToString
            Me.TextBox43.Text = DataGridView2.Item(40, i).Value.ToString
            Me.TextBox42.Text = DataGridView2.Item(41, i).Value.ToString
            Me.TextBox34.Text = DataGridView2.Item(42, i).Value.ToString

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT ProdName, Qty, Price, Discount, Total, Price*Qty As Gross, Gross*6/112 As CGST, Gross*6/112 As SGST FROM InvoiceProduct WHERE InvoiceProduct.Cust_ID LIKE'%" &
            TextBox5.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("InvoiceProduct")
            adapter.Fill(dt)
            Me.DataGridView1.DataSource = dt
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "SELECT Cust_ID FROM BilledInvoice WHERE (ReceiptNo = '" & TextBox5.Text & "')"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            dr = cmd.ExecuteReader
            While dr.Read()
                TextBox51.Text = dr("Cust_ID").ToString
            End While
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try

        GridCopy()
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Compute()
        Dim i As Double = 0
        If TextBox40.Text = "Tax @ 12%" Then
            i = Val(TextBox9.Text) * 12 / 112
            i = i / 2
            i = Math.Round(i, 2)
            TextBox35.Text = i
            TextBox36.Text = i
        Else
            If TextBox40.Text = "Tax @ 18%" Then
                i = Val(TextBox9.Text) * 18 / 118
                i = i / 2
                i = Math.Round(i, 2)
                TextBox35.Text = i
                TextBox36.Text = i
            End If
        End If
    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged
        Compute4()
        Dim i As Double = 0
        i = Val(TextBox27.Text) * 12 / 112
        i = i / 2
        i = Math.Round(i, 2)
        TextBox37.Text = i
        TextBox38.Text = i
    End Sub

    Sub Compute4()
        Dim i As Double
        i = Val(TextBox25.Text) - Val(TextBox26.Text)
        i = Math.Round(i, 2)
        TextBox27.Text = i
        TextBox29.Text = i
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "select * from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox30.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox31_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox31.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "select * from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox31.Text & "%' AND Paymode<>''"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to update the current record?", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str6 As String
                str6 = "Delete from InvoiceProduct Where Cust_ID = '" & Me.TextBox5.Text & "'"
                Dim cmd7 As OleDbCommand = New OleDbCommand(str6, myConnection)
                Try
                    cmd7.ExecuteNonQuery()
                    cmd7.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str5 As String
                str5 = "Delete from PaymentVoucher Where ReceiptNo = '" & Me.TextBox1.Text & "'"
                Dim cmd6 As OleDbCommand = New OleDbCommand(str5, myConnection)
                Try
                    cmd6.ExecuteNonQuery()
                    cmd6.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str4 As String
                str4 = "Delete from CustomerTable Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd0 As OleDbCommand = New OleDbCommand(str4, myConnection)
                Try
                    cmd0.ExecuteNonQuery()
                    cmd0.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            ElseIf okToDelete = MsgBoxResult.No Then
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "Update SaleInvoice SET [Cust_ID]='" & TextBox1.Text & "', [Cust_Name]='" & TextBox2.Text & "', [Mobile]='" & TextBox3.Text & "', [Phone]='" & TextBox4.Text & "', [Address]='" & "" & "', [BookingDate]='" & DateTimePicker1.Value.Date & "', [ReceiptNo]='" & TextBox5.Text & "', [DeliveryDate]='" & DateTimePicker2.Value.Date & "', [BookedBy]='" & ComboBox1.Text & "', [DeliveryTime]='" & TextBox44.Text & "', [Status]='" & TextBox39.Text & "', [Age]='" & TextBox45.Text & "', [JobStatus]='" & ComboBox3.Text & "', [GrossAmt]='" & TextBox25.Text & "', [CGST]='" & TextBox37.Text & "', [SGST]='" & TextBox38.Text & "', [ScmAmt]='" & TextBox26.Text & "', [NetAmt]='" & TextBox27.Text & "', [PaidAmt]='" & TextBox28.Text & "', [DueAmt]='" & TextBox29.Text & "', [Paymode]='" & TextBox33.Text & "', [Remarks]='" & TextBox41.Text & "', [RSPH]='" & TextBox11.Text & "', [RCYL]='" & TextBox12.Text & "', [RAXIS]='" & TextBox13.Text & "', [RVN]='" & TextBox14.Text & "', [RADD]='" & TextBox15.Text & "', [LSPH]='" & TextBox16.Text & "', [LCYL]='" & TextBox17.Text & "', [LAXIS]='" & TextBox18.Text & "', [LVN]='" & TextBox19.Text & "', [LADD]='" & TextBox20.Text & "', [PD]='" & TextBox21.Text & "', [REFBY]='" & TextBox22.Text & "', [LensType]='" & TextBox23.Text & "', [LensType1]='" & TextBox46.Text & "', [LensType2]='" & TextBox47.Text & "', [LensType3]='" & TextBox48.Text & "', [Remarks1]='" & TextBox24.Text & "', [Right]='" & TextBox43.Text & "', [Left]='" & TextBox42.Text & "', [RLAdd]='" & TextBox34.Text & "' Where [Cust_ID]='" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "INSERT INTO InvoiceProduct ([Cust_ID], [Cust_Name], [Mobile], [InvDate], [ProdName], [Qty], [Price], [Discount], [Total]) VALUES ('" & TextBox5.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & DateTimePicker1.Value.Date & "', ?, ?, ?, ?, ?)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            For Each row As DataGridViewRow In DataGridView3.Rows
                If Not row.IsNewRow Then
                    cmd1.Parameters.Add(New OleDbParameter("ProdName", row.Cells(0).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Qty", row.Cells(1).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Price", row.Cells(2).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Discount", row.Cells(3).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Total", row.Cells(4).Value))
                    cmd1.ExecuteNonQuery()
                    cmd1.Parameters.Clear()
                End If
            Next
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str2 As String
            str2 = "INSERT INTO CustomerTable ([Cust_ID], [Cust_Name], [Address], [City], [ContactNo], [Remarks]) VALUES (?, ?, ?, ?, ?, ?)"
            Dim cmd3 As OleDbCommand = New OleDbCommand(str2, myConnection)
            cmd3.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox5.Text, String)))
            cmd3.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd3.Parameters.Add(New OleDbParameter("Address", CType("", String)))
            cmd3.Parameters.Add(New OleDbParameter("City", CType("INDORE", String)))
            cmd3.Parameters.Add(New OleDbParameter("ContactNo", CType(TextBox3.Text, String)))
            cmd3.Parameters.Add(New OleDbParameter("Remarks", CType("", String)))
            cmd3.ExecuteNonQuery()
            cmd3.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str3 As String
            str3 = "INSERT INTO PaymentVoucher ([ReceiptName], [ReceiptNo], [ReceiptDate], [ReceiptAmt], [ReceiptBal], [ReceiptDue], [ReceiptStatus]) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str3, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("ReceiptName", CType(TextBox2.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox1.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDate", CType(DateTimePicker1.Value.Date, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptAmt", CType(TextBox27.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptBal", CType(TextBox28.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDue", CType(TextBox29.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptStatus", CType(TextBox33.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MsgBox("Updated Successfuly...!", MsgBoxStyle.OkOnly)
        Dim okToSMS As MsgBoxResult = MsgBox("Press Yes to Send SMS?", MsgBoxStyle.YesNo)
        If okToSMS = MsgBoxResult.Yes Then
            SendSMS()
        ElseIf okToSMS = MsgBoxResult.No Then
        End If

        Dim okToPrint As MsgBoxResult = MsgBox("Press Yes for JOB-CARD?", MsgBoxStyle.YesNo)
        If okToPrint = MsgBoxResult.Yes Then
            Print()
        ElseIf okToPrint = MsgBoxResult.No Then
            Exit Sub
        End If
        Clear()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox33.Text = RadioButton1.Text
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            TextBox33.Text = RadioButton2.Text
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            TextBox33.Text = RadioButton3.Text
        End If
    End Sub

    Sub Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        TextBox5.Clear()
        TextBox44.Clear()
        ComboBox1.SelectedIndex = -1
        TextBox45.Clear()
        ComboBox3.SelectedIndex = -1
        TextBox39.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        TextBox18.Clear()
        TextBox19.Clear()
        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox46.Clear()
        TextBox47.Clear()
        TextBox48.Clear()
        TextBox24.Clear()
        TextBox43.Clear()
        TextBox42.Clear()
        TextBox34.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        DataGridView1.DataSource = Nothing
        DataGridView3.DataSource = Nothing
        DataGridView3.Rows.Clear()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "ON PROCESS" Then
            TextBox49.Text = "Dear Sir/Madam Thanks for Visiting American Optics Palasia, your Transaction No - " + Me.TextBox1.Text + " on date " + Me.DateTimePicker1.Value + " is on process now."
        Else
            If ComboBox3.Text = "READY" Then
                TextBox49.Text = "Dear Customer, your order is ready for delivery, kindly come personally to get it checked. Thanks American Optics Palasia."
            Else
                If ComboBox3.Text = "DELIVERED" Then
                    TextBox49.Text = "Your order is delivered, Thanks for your precious order, do visit again. Thanks American Optics Palasia."
                End If
            End If
        End If
    End Sub

    Sub SendSMS()
        Try
            Dim url As String
            url = "http://alerts.valueleaf.com/api/v4/?api_key=A7ce7d9a7a5bcb5f1cfdc9e60b9095d8c&method=sms&message=" + Me.TextBox49.Text + "&to=" + Me.TextBox3.Text + "&sender=AOPTIC"
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

    Private Sub DataGridView3_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView3.MouseClick
        Dim i As Integer
        i = DataGridView3.CurrentRow.Index
        Me.TextBox50.Text = DataGridView3.Item(0, i).Value.ToString
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
            TextBox51.Text = GenerateCode()
            TextBox51.Text = "INV-" + GenerateCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class