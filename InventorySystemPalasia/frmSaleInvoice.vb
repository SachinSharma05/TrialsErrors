Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class frmSaleInvoice

    Dim st2 As String

    Private Sub frmSaleInvoice_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmSaleInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Function GetValue(ByVal Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub frmSaleInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadCombo()
        LoadCombo2()
        ComboBox1.SelectedIndex = -1
        Me.ComboBox2.SelectedIndex = -1
        Me.ComboBox2.Text = ""
        TextBox8.Text = "0"
        TextBox35.Text = "0"
        TextBox36.Text = "0"
        auto()
        autoID()
        RefreshData()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        Clear()
        ClearText()
        TextBox2.Select()
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

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

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, Age, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD,  REFBY, LensType, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox30.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub TextBox31_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox31.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox31_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox31.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Phone, BookingDate, Age, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD,  REFBY, LensType, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox31.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        RefreshData()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        TextBox30.Clear()
        TextBox31.Clear()
    End Sub

    Sub CheckTextBox()
        If TextBox3.Text = "" And TextBox11.Text = "" And TextBox12.Text = "" And TextBox13.Text = "" And TextBox14.Text = "" And TextBox15.Text = "" And TextBox16.Text = "" And TextBox17.Text = "" And TextBox18.Text = "" And TextBox19.Text = "" And TextBox20.Text = "" And TextBox21.Text = "" And TextBox22.Text = "" And TextBox23.Text = "" Then
            Me.Button3.Enabled = True
            Me.Button4.Enabled = False
        Else
            Me.Button3.Enabled = False
            Me.Button4.Enabled = True
        End If
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
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
        TextBox24.Clear()
        TextBox46.Clear()
        TextBox47.Clear()
        TextBox48.Clear()
        TextBox43.Clear()
        TextBox42.Clear()
        TextBox49.Clear()
        TextBox50.Clear()
        TextBox34.Clear()
        TextBox11.Focus()
    End Sub

    Sub Clear()
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
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
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows.Add(ComboBox2.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, TextBox35.Text, TextBox36.Text)
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
                Clear()
                Exit Sub
            Next

            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(ComboBox2.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, TextBox35.Text, TextBox36.Text)
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
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
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

    Public Sub New()
        InitializeComponent()
        Me.GetAllControls(Me).OfType(Of Button)().ToList() _
          .ForEach(Sub(b)
                       b.Tag = Tuple.Create(b.ForeColor, b.BackColor)
                       AddHandler b.GotFocus, AddressOf b_GotFocus
                       AddHandler b.LostFocus, AddressOf b_LostFocus
                   End Sub)
    End Sub

    Public Function GetAllControls(ByVal control As Control) As IEnumerable(Of Control)
        Dim controls = control.Controls.Cast(Of Control)()
        Return controls.SelectMany(Function(ctrl) GetAllControls(ctrl)).Concat(controls)
    End Function

    Private Function GenerateID() As String
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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM SaleInvoice ORDER BY ID DESC", con)
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

    Sub auto()
        Try
            TextBox1.Text = GenerateID()
            TextBox1.Text = "SR-" + GenerateID()
            TextBox5.Text = TextBox1.Text
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
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
            TextBox52.Text = GenerateCode()
            TextBox52.Text = "INV-" + GenerateCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox45.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
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
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
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

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox8.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Sub Compute2()
        Dim i As Double = 0
        i = Val(TextBox27.Text) - Val(TextBox28.Text)
        i = Math.Round(i, 2)
        TextBox29.Text = i
    End Sub

    Private Sub TextBox28_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox28.KeyDown
        If e.KeyCode = Keys.Enter Then
            RadioButton1.Focus()
        End If
    End Sub

    Private Sub TextBox28_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox28.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged
        If Val(TextBox28.Text) > Val(TextBox27.Text) Then
            MsgBox("Advance cannot be more than Net Amount")
            TextBox28.Clear()
        End If
        Compute2()
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            DateTimePicker2.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Invoice Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub ComboBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
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
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Clear()
        ClearText()
        Me.DataGridView1.Rows.Clear()
        Button3.Enabled = True
        Button4.Enabled = True
        auto()
    End Sub

    Sub ClearText()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        ComboBox1.SelectedIndex = -1
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
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
        TextBox45.Clear()
        TextBox46.Clear()
        TextBox47.Clear()
        TextBox48.Clear()
        TextBox43.Clear()
        TextBox42.Clear()
        TextBox34.Clear()
        TextBox44.Clear()
        TextBox49.Clear()
        TextBox50.Clear()
        ComboBox3.SelectedIndex = -1
        TextBox2.Focus()
    End Sub

    Private Sub Button11_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Please enter Customer Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MessageBox.Show("Please enter Mobile No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox3.Focus()
            Exit Sub
        End If
        If TextBox44.Text = "" Then
            MessageBox.Show("Please enter Delivery Time", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox44.Focus()
            Exit Sub
        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
            MessageBox.Show("Please Select Paymode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox28.Focus()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select Salesperson Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox28.Text = "" Then
            MessageBox.Show("Please enter Advance Amt", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox28.Focus()
            Exit Sub
        End If
        If ComboBox3.Text = "" Then
            MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox28.Focus()
            Exit Sub
        End If

        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct As String = "select Cust_ID from SaleInvoice where Cust_ID=@d1"
            Dim cmd9 As OleDbCommand = New OleDbCommand(ct)
            cmd9.Parameters.AddWithValue("@d1", TextBox1.Text)
            cmd9.Connection = myConnection
            rdr = cmd9.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                auto()
            End If
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into SaleInvoice ([Cust_ID], [Cust_Name], [Mobile], [Phone], [Address], [BookingDate], [ReceiptNo], [DeliveryDate], [BookedBy], [DeliveryTime], [Status], [Age], [JobStatus], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode], [Remarks], [RSPH], [RCYL], [RAXIS], [RVN], [RADD], [LSPH], [LCYL], [LAXIS], [LVN], [LADD], [PD], [REFBY], [LensType], [LensType1], [LensType2], [LensType3], [Remarks1], [Right], [Left], [RLAdd], [PRGRight], [PRGLeft]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Phone", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Address", CType("", String)))
            cmd.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox5.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("DeliveryDate", CType(DateTimePicker2.Value.Date, String)))
            cmd.Parameters.Add(New OleDbParameter("BookedBy", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("DeliveryTime", CType(TextBox44.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Status", CType(TextBox39.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Age", CType(TextBox45.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("JobStatus", CType(ComboBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox25.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("CGST", CType(TextBox37.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("SGST", CType(TextBox38.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox26.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox27.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PaidAmt", CType(TextBox28.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox29.Text, String)))
            If RadioButton1.Checked = True Then
                cmd.Parameters.Add(New OleDbParameter("Paymode", CType(RadioButton1.Text, String)))
            End If
            If RadioButton2.Checked = True Then
                cmd.Parameters.Add(New OleDbParameter("Paymode", CType(RadioButton2.Text, String)))
            End If
            If RadioButton3.Checked = True Then
                cmd.Parameters.Add(New OleDbParameter("Paymode", CType(RadioButton3.Text, String)))
            End If
            cmd.Parameters.Add(New OleDbParameter("Remarks", CType(TextBox41.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RSPH", CType(TextBox11.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RCYL", CType(TextBox12.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RAXIS", CType(TextBox13.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RVN", CType(TextBox14.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RADD", CType(TextBox15.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LSPH", CType(TextBox16.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LCYL", CType(TextBox17.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LAXIS", CType(TextBox18.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LVN", CType(TextBox19.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LADD", CType(TextBox20.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PD", CType(TextBox21.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("REFBY", CType(TextBox22.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType", CType(TextBox23.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType1", CType(TextBox46.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType2", CType(TextBox47.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType3", CType(TextBox48.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Remarks1", CType(TextBox24.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Right", CType(TextBox43.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Left", CType(TextBox42.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RLAdd", CType(TextBox34.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PRGRight", CType(TextBox49.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PRGLeft", CType(TextBox50.Text, String)))
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
            For Each row As DataGridViewRow In DataGridView1.Rows
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

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    myConnection = New OleDbConnection(provider)
                    myConnection.Open()
                    Dim cb4 As String = "update ItemMaster set Stock = Stock - (" & row.Cells(1).Value & ") where Prod_Name= ComboBox2.Text"
                    Dim cmd2 As New OleDbCommand
                    cmd2 = New OleDbCommand(cb4)
                    cmd2.Connection = myConnection
                    cmd2.Parameters.Add(New OleDbParameter("& ComboBox2.Text &", row.Cells(0).Value))
                    cmd2.ExecuteNonQuery()
                    myConnection.Close()
                End If
            Next

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

            MsgBox("Sale Created Successfuly...!", MsgBoxStyle.OkOnly)
            TextBox2.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Clear()
        ClearText()
        Me.DataGridView1.Rows.Clear()

        Dim okToPrint As MsgBoxResult = MsgBox("Press Yes for JOB-CARD, Press No for INVOICE?", MsgBoxStyle.YesNoCancel)
        If okToPrint = MsgBoxResult.Yes Then
            Print()
        ElseIf okToPrint = MsgBoxResult.No Then
            Print1()
        End If
        auto()
        autoID()
    End Sub

    Public Property Checked As Boolean

    Private Sub DateTimePicker1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            DateTimePicker2.Focus()
        End If
    End Sub

    Private Sub DateTimePicker2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox44.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox35.Focus()
        End If
    End Sub

    Private Sub TextBox35_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox35.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox36.Focus()
        End If
    End Sub

    Private Sub TextBox36_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox36.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox12_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox13.Focus()
        End If
    End Sub

    Private Sub TextBox13_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox14.Focus()
        End If
    End Sub

    Private Sub TextBox14_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox15_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox16.Focus()
        End If
    End Sub

    Private Sub TextBox16_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox17.Focus()
        End If
    End Sub

    Private Sub TextBox17_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox18.Focus()
        End If
    End Sub

    Private Sub TextBox18_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox19.Focus()
        End If
    End Sub

    Private Sub TextBox19_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox19.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox20.Focus()
        End If
    End Sub

    Private Sub TextBox20_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox21.Focus()
        End If
    End Sub

    Private Sub TextBox21_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox21.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox22.Focus()
        End If
    End Sub

    Private Sub TextBox22_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox23.Focus()
        End If
    End Sub

    Private Sub TextBox23_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox46.Focus()
        End If
    End Sub

    Private Sub TextBox24_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox43.Focus()
        End If
    End Sub

    Sub Compute3()
        Dim i As Double = 0
        i = Val(TextBox27.Text) - Val(TextBox41.Text)
        i = Math.Round(i, 2)
        TextBox39.Text = i
        TextBox29.Text = i
    End Sub

    Private Sub TextBox41_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Compute3()
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

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            TextBox41.ReadOnly = False
            TextBox41.Focus()
        End If
        If RadioButton3.Checked = True Then
            TextBox33.Text = RadioButton3.Text
        End If
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, Phone, BookingDate, Age, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD,  REFBY, LensType, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView2.Columns(5).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(8).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
    End Sub

    Private Sub DataGridView2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseClick
        Dim i As Integer
        i = DataGridView2.CurrentRow.Index
        Me.TextBox53.Text = Me.DataGridView2.Item(1, i).Value.ToString
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Dim okToShow As MsgBoxResult = MsgBox("Press Yes for New Job Card, Press No to Update Job Card", MsgBoxStyle.YesNo)
        If okToShow = MsgBoxResult.Yes Then
            auto()
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "SELECT * FROM SaleInvoice WHERE (Cust_Name = '" & TextBox53.Text & "')"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            dr = cmd.ExecuteReader
            While dr.Read()
                TextBox2.Text = dr("Cust_Name").ToString
                TextBox3.Text = dr("Mobile").ToString
                TextBox4.Text = dr("Phone").ToString
                TextBox45.Text = dr("Age").ToString
                TextBox11.Text = dr("RSPH").ToString
                TextBox12.Text = dr("RCYL").ToString
                TextBox13.Text = dr("RAXIS").ToString
                TextBox14.Text = dr("RVN").ToString
                TextBox15.Text = dr("RADD").ToString
                TextBox16.Text = dr("LSPH").ToString
                TextBox17.Text = dr("LCYL").ToString
                TextBox18.Text = dr("LAXIS").ToString
                TextBox19.Text = dr("LVN").ToString
                TextBox20.Text = dr("LADD").ToString
                TextBox21.Text = dr("PD").ToString
                TextBox22.Text = dr("REFBY").ToString
                TextBox23.Text = dr("LensType").ToString
                TextBox46.Text = dr("LensType1").ToString
                TextBox47.Text = dr("LensType2").ToString
                TextBox48.Text = dr("LensType3").ToString
                TextBox24.Text = dr("Remarks1").ToString
                TextBox43.Text = dr("Right").ToString
                TextBox42.Text = dr("Left").ToString
                TextBox34.Text = dr("RLAdd").ToString
                Me.Button4.Enabled = False
                Me.Button3.Enabled = True
            End While
            myConnection.Close()
        Else
            If okToShow = MsgBoxResult.No Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
                dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
                connString = provider & dataFile
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str As String
                str = "SELECT * FROM SaleInvoice WHERE (Cust_Name = '" & TextBox53.Text & "')"
                Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
                dr = cmd.ExecuteReader
                While dr.Read()
                    TextBox1.Text = dr("Cust_ID").ToString
                    TextBox2.Text = dr("Cust_Name").ToString
                    TextBox3.Text = dr("Mobile").ToString
                    TextBox4.Text = dr("Phone").ToString
                    DateTimePicker1.Value = dr("BookingDate").ToString
                    TextBox5.Text = TextBox1.Text
                    TextBox45.Text = dr("Age").ToString
                    TextBox11.Text = dr("RSPH").ToString
                    TextBox12.Text = dr("RCYL").ToString
                    TextBox13.Text = dr("RAXIS").ToString
                    TextBox14.Text = dr("RVN").ToString
                    TextBox15.Text = dr("RADD").ToString
                    TextBox16.Text = dr("LSPH").ToString
                    TextBox17.Text = dr("LCYL").ToString
                    TextBox18.Text = dr("LAXIS").ToString
                    TextBox19.Text = dr("LVN").ToString
                    TextBox20.Text = dr("LADD").ToString
                    TextBox21.Text = dr("PD").ToString
                    TextBox22.Text = dr("REFBY").ToString
                    TextBox23.Text = dr("LensType").ToString
                    TextBox46.Text = dr("LensType1").ToString
                    TextBox47.Text = dr("LensType2").ToString
                    TextBox48.Text = dr("LensType3").ToString
                    TextBox24.Text = dr("Remarks1").ToString
                    TextBox43.Text = dr("Right").ToString
                    TextBox42.Text = dr("Left").ToString
                    TextBox34.Text = dr("RLAdd").ToString
                    Me.Button3.Enabled = False
                    Me.Button4.Enabled = True
                End While
                myConnection.Close()
            End If
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub TextBox27_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox27.TextChanged
        If Val(TextBox27.Text) = 0 Or TextBox27.Text = "" Then
            TextBox39.Text = "Regular"
        End If
        If Val(TextBox27.Text) <= 3000 Then
            TextBox39.Text = "Silver"
        Else
            If Val(TextBox27.Text) > 3001 And Val(TextBox27.Text) <= 8000 Then
                TextBox39.Text = "Gold"
            Else
                If Val(TextBox27.Text) > 8001 And Val(TextBox27.Text) <= 15000 Then
                    TextBox39.Text = "Platinum"
                End If
            End If
        End If
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

    Sub Print1()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New InvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SaleInvoice.ReceiptNo, SaleInvoice.Cust_Name, SaleInvoice.Mobile, SaleInvoice.GrossAmt, SaleInvoice.NetAmt, SaleInvoice.PaidAmt, SaleInvoice.DueAmt, SaleInvoice.Paymode, SaleInvoice.BookingDate, SaleInvoice.CGST, SaleInvoice.SGST, SaleInvoice.ScmAmt, InvoiceProduct.ProdName, InvoiceProduct.Qty, InvoiceProduct.Price, InvoiceProduct.Discount, InvoiceProduct.Total FROM InvoiceProduct INNER JOIN SaleInvoice ON SaleInvoice.ReceiptNo=InvoiceProduct.Cust_ID Where SaleInvoice.ReceiptNo=@d1"
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
            rpt.SetParameterValue("p3", TextBox52.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value.Date)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RadioButton1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RadioButton1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.Focus()
        End If
    End Sub

    Private Sub RadioButton2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RadioButton2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.Focus()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If TextBox2.Text = "" Then
                MessageBox.Show("Please enter Customer Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox2.Focus()
                Exit Sub
            End If
            If TextBox3.Text = "" Then
                MessageBox.Show("Please enter Mobile No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox3.Focus()
                Exit Sub
            End If
            If TextBox44.Text = "" Then
                MessageBox.Show("Please enter Delivery Time", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox44.Focus()
                Exit Sub
            End If
            If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Please Select Paymode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox5.Focus()
                Exit Sub
            End If
            If ComboBox1.Text = "" Then
                MessageBox.Show("Please select Salesperson Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox1.Focus()
                Exit Sub
            End If
            If TextBox28.Text = "" Then
                MessageBox.Show("Please enter Advance Amt", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox28.Focus()
                Exit Sub
            End If
            If TextBox44.Text = "" Then
                MessageBox.Show("Please enter Delivery Time with AM/PM", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox44.Focus()
                Exit Sub
            End If
            If ComboBox3.Text = "" Then
                MessageBox.Show("Please select Job Status", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox3.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "UPDATE SaleInvoice SET [Cust_Name] = '" & TextBox2.Text & "', [Mobile] = '" & TextBox3.Text & "', [Phone] = '" & TextBox4.Text & "', [ReceiptNo] = '" & TextBox5.Text & "', [DeliveryDate] = '" & DateTimePicker2.Value.Date & "', [BookedBy] = '" & ComboBox1.Text & "', [DeliveryTime] = '" & TextBox44.Text & "', [Status] = '" & TextBox39.Text & "', [Age] = '" & TextBox45.Text & "', [JobStatus] = '" & ComboBox3.Text & "', [GrossAmt] = '" & TextBox25.Text & "', [CGST] = '" & TextBox37.Text & "', [SGST] = '" & TextBox38.Text & "', [ScmAmt] = '" & TextBox26.Text & "', [NetAmt] = '" & TextBox27.Text & "', [PaidAmt] = '" & TextBox28.Text & "', [DueAmt] = '" & TextBox29.Text & "', [Paymode] = '" & TextBox33.Text & "', [Remarks] = '" & TextBox41.Text & "', [RSPH] = '" & TextBox11.Text & "', [RCYL] = '" & TextBox12.Text & "', [RAXIS] = '" & TextBox13.Text & "', [RVN] = '" & TextBox14.Text & "', [RADD] = '" & TextBox15.Text & "', [LSPH] = '" & TextBox16.Text & "', [LCYL] = '" & TextBox17.Text & "', [LAXIS] = '" & TextBox18.Text & "', [LVN] = '" & TextBox19.Text & "', [LADD] = '" & TextBox20.Text & "', [PD] = '" & TextBox21.Text & "', [REFBY] = '" & TextBox22.Text & "', [LensType] = '" & TextBox23.Text & "', [LensType1] = '" & TextBox46.Text & "', [LensType2] = '" & TextBox47.Text & "', [LensType3] = '" & TextBox48.Text & "' ,[Remarks1] = '" & TextBox24.Text & "', [Right] = '" & TextBox43.Text & "', [Left] = '" & TextBox42.Text & "', [RLAdd] = '" & TextBox34.Text & "', [PRGRight] = '" & TextBox49.Text & "', [PRGLeft] = '" & TextBox50.Text & "' Where [Cust_ID] = '" & TextBox1.Text & "'"
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
            For Each row As DataGridViewRow In DataGridView1.Rows
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

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    myConnection = New OleDbConnection(provider)
                    myConnection.Open()
                    Dim cb4 As String = "update ItemMaster set Stock = Stock - (" & row.Cells(1).Value & ") where Prod_Name= ComboBox2.Text"
                    Dim cmd2 As New OleDbCommand
                    cmd2 = New OleDbCommand(cb4)
                    cmd2.Connection = myConnection
                    cmd2.Parameters.Add(New OleDbParameter("& ComboBox2.Text &", row.Cells(0).Value))
                    cmd2.ExecuteNonQuery()
                    myConnection.Close()
                End If
            Next

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

            MsgBox("Sale Updated/Created Successfuly...!", MsgBoxStyle.OkOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim okToPrint As MsgBoxResult = MsgBox("Press Yes for JOB-CARD, Press No for INVOICE?", MsgBoxStyle.YesNoCancel)
        If okToPrint = MsgBoxResult.Yes Then
            Print()
        ElseIf okToPrint = MsgBoxResult.No Then
            Print1()
        End If
        Clear()
        ClearText()
        Me.DataGridView1.Rows.Clear()
        auto()
        autoID()
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

    Private Sub TextBox43_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox43.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox42.Focus()
        End If
    End Sub

    Private Sub TextBox42_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox42.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox34.Focus()
        End If
    End Sub

    Private Sub TextBox34_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox34.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.Focus()
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

    Private Sub TextBox44_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox44.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox44.Text = TextBox44.Text + "PM"
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox45_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox45.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub TextBox46_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox46.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox47.Focus()
        End If
    End Sub

    Private Sub TextBox47_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox47.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox48.Focus()
        End If
    End Sub

    Private Sub TextBox48_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox48.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox24.Focus()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "ON PROCESS" Then
            TextBox51.Text = "Dear Sir/Madam Thanks for Visiting American Optics Palasia, your Transaction No - " + Me.TextBox1.Text + " on date " + Me.DateTimePicker1.Value + " is on process now."
        Else
            If ComboBox3.Text = "READY" Then
                TextBox51.Text = "Dear Customer, your order is ready for delivery, kindly come personally to get it checked. Thanks American Optics Palasia."
            Else
                If ComboBox3.Text = "DELIVERED" Then
                    TextBox51.Text = "Your order is delivered, Thanks for your precious order, do visit again. Thanks American Optics Palasia."
                End If
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox3.Focus()
        End If
    End Sub
End Class