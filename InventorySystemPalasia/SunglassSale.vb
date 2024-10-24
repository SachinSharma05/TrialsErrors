Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class SunglassSale

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ClearText()
        Clear()
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

    Sub ClearText()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        ComboBox1.SelectedIndex = -1
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
        TextBox2.Focus()
    End Sub

    Private Sub SunglassSale_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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

    Private Sub SunglassSale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        LoadCombo2()
        SunglassData()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        Button4.Enabled = False
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        autoID()
        Clear()
        ClearText()
        auto()
    End Sub

    Private Sub SunglassData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select * from SunglassSale"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(1).Width = 80
        Me.DataGridView2.Columns(2).Width = 130
        Me.DataGridView2.Columns(4).Visible = False
        Me.DataGridView2.Columns(5).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(8).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(16).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
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

        Dim cmd As New OleDbCommand("SELECT Prod_Name FROM ItemMaster Where Prod_Type='Sunglasses';", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox2.DisplayMember = "Prod_Name"
        ComboBox2.ValueMember = "Prod_Name"
        ComboBox2.DataSource = dt

        cn.Close()
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

    Private Sub TextBox6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
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
        If TextBox40.Text = "Tax @ 18%" Then
            i = Val(TextBox9.Text) * 18 / 118
            i = i / 2
            i = Math.Round(i, 2)
            TextBox35.Text = i
            TextBox36.Text = i
        Else
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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM SunglassSale ORDER BY ID DESC", con)
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
            TextBox1.Text = "SG-" + GenerateID()
            TextBox5.Text = TextBox1.Text
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
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

    Sub Compute2()
        Dim i As Double = 0
        i = Val(TextBox27.Text) - Val(TextBox28.Text)
        i = Math.Round(i, 2)
        TextBox29.Text = i
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
        End While
        myConnection.Close()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
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

        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct As String = "select Cust_ID from SunglassSale where Cust_ID=@d1"
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
            str = "insert into SunglassSale ([Cust_ID], [Cust_Name], [Mobile], [Phone], [Address], [BookingDate], [ReceiptNo], [DeliveryDate], [BookedBy], [Status], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode], [Remarks]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
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
            cmd.Parameters.Add(New OleDbParameter("Status", CType(TextBox39.Text, String)))
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
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "INSERT INTO SunglassProduct ([Cust_ID], [Cust_Name], [Mobile], [InvDate], [ProdName], [Qty], [Price], [Discount], [Total]) VALUES ('" & TextBox5.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & DateTimePicker1.Value.Date & "', ?, ?, ?, ?, ?)"
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
            SendSMS()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Clear()
        ClearText()
        Me.DataGridView1.Rows.Clear()

        Dim okToPrint As MsgBoxResult = MsgBox("Press Yes for Invoice", MsgBoxStyle.YesNo)
        If okToPrint = MsgBoxResult.Yes Then
            Print1()
        ElseIf okToPrint = MsgBoxResult.No Then
        End If
        auto()
        autoID()
    End Sub

    Sub SendSMS()
        Try
            Dim url As String
            url = "http://alerts.valueleaf.com/api/v4/?api_key=A7ce7d9a7a5bcb5f1cfdc9e60b9095d8c&method=sms&message=" + Me.TextBox11.Text + "&to=" + Me.TextBox3.Text + "&sender=AOPTIC"
            Dim myReq As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
            Dim myResp As HttpWebResponse = DirectCast(myReq.GetResponse(), HttpWebResponse)
            Dim respStreamReader As New System.IO.StreamReader(myResp.GetResponseStream())
            Dim responseString As String = respStreamReader.ReadToEnd()
            respStreamReader.Close()
            myResp.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Compute()
        Dim i As Double = 0
        If TextBox40.Text = "Tax @ 18%" Then
            i = Val(TextBox9.Text) * 18 / 118
            i = i / 2
            i = Math.Round(i, 2)
            TextBox35.Text = i
            TextBox36.Text = i
        Else
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

    Sub Print1()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New SunglassBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SunglassSale.ReceiptNo, SunglassSale.Cust_Name, SunglassSale.Mobile, SunglassSale.GrossAmt, SunglassSale.NetAmt, SunglassSale.PaidAmt, SunglassSale.DueAmt, SunglassSale.Paymode, SunglassSale.BookingDate, SunglassSale.CGST, SunglassSale.SGST, SunglassSale.ScmAmt, SunglassProduct.ProdName, SunglassProduct.Qty, SunglassProduct.Price, SunglassProduct.Discount, SunglassProduct.Total FROM SunglassProduct INNER JOIN SunglassSale ON SunglassSale.ReceiptNo=SunglassProduct.Cust_ID Where SunglassSale.ReceiptNo=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox5.Text)
            MyCommand1.CommandText = "SELECT * from SunglassSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SunglassSale")
            myDA.Fill(myDS, "SunglassProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox5.Text)
            rpt.SetParameterValue("p3", TextBox13.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TextBox26_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox26.TextChanged
        Compute4()
        Dim i As Double = 0
        i = Val(TextBox27.Text) * 18 / 118
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

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged
        If Val(TextBox28.Text) > Val(TextBox27.Text) Then
            MsgBox("Advance cannot be more than Net Amount")
            TextBox28.Clear()
        End If
        Compute2()
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
            TextBox41.ReadOnly = False
            TextBox41.Focus()
        End If
        If RadioButton3.Checked = True Then
            TextBox33.Text = RadioButton3.Text
        End If
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from SunglassSale WHERE Cust_Name LIKE'%" &
        TextBox12.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(1).Width = 80
        Me.DataGridView2.Columns(2).Width = 130
        Me.DataGridView2.Columns(4).Visible = False
        Me.DataGridView2.Columns(5).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(8).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(16).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
    End Sub

    Private Sub DataGridView2_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView2.CurrentRow.Index
        Me.TextBox1.Text = DataGridView2.Item(1, i).Value.ToString
        Me.TextBox2.Text = DataGridView2.Item(2, i).Value.ToString
        Me.TextBox3.Text = DataGridView2.Item(3, i).Value.ToString
        Me.TextBox4.Text = DataGridView2.Item(4, i).Value.ToString
        Me.DateTimePicker1.Value = DataGridView2.Item(6, i).Value.ToString
        Me.TextBox5.Text = DataGridView2.Item(7, i).Value.ToString
        Me.DateTimePicker2.Value = DataGridView2.Item(8, i).Value.ToString
        Me.ComboBox1.Text = DataGridView2.Item(9, i).Value.ToString
        Me.TextBox39.Text = DataGridView2.Item(10, i).Value.ToString
        Me.TextBox25.Text = DataGridView2.Item(11, i).Value.ToString
        Me.TextBox37.Text = DataGridView2.Item(12, i).Value.ToString
        Me.TextBox38.Text = DataGridView2.Item(13, i).Value.ToString
        Me.TextBox26.Text = DataGridView2.Item(14, i).Value.ToString
        Me.TextBox27.Text = DataGridView2.Item(15, i).Value.ToString
        Me.TextBox28.Text = DataGridView2.Item(16, i).Value.ToString
        Me.TextBox29.Text = DataGridView2.Item(17, i).Value.ToString
        If Me.DataGridView2.Item(18, i).Value = "Payment By Cash" Then
            RadioButton1.Checked = True
        Else
            If Me.DataGridView2.Item(18, i).Value = "Payment By Card" Then
                RadioButton2.Checked = True
            Else
                If Me.DataGridView2.Item(18, i).Value = "Payment By Both" Then
                    RadioButton3.Checked = True
                End If
            End If
        End If

        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT ProdName, Qty, Price, Discount, Total, Price*Qty As Gross, Total*9/118 As CGST, Total*9/118 As SGST FROM SunglassProduct WHERE SunglassProduct.Cust_ID LIKE'%" &
        TextBox5.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("InvoiceProduct")
        adapter.Fill(dt)
        Me.DataGridView3.DataSource = dt
        myConnection.Close()

        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "SELECT Cust_ID FROM SunglassBilledInvoice WHERE (ReceiptNo = '" & TextBox5.Text & "')"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        dr = cmd.ExecuteReader
        While dr.Read()
            TextBox13.Text = dr("Cust_ID").ToString
        End While
        myConnection.Close()

        Button4.Enabled = True
        GridCopy()
    End Sub

    Sub GridCopy()
        Dim sourceGrid As DataGridView = Me.DataGridView3
        Dim targetGrid As DataGridView = Me.DataGridView1
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
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

        Try
            Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to update the current record?", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str6 As String
                str6 = "Delete from SunglassProduct Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd7 As OleDbCommand = New OleDbCommand(str6, myConnection)
                Try
                    cmd7.ExecuteNonQuery()
                    cmd7.Dispose()
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
            str = "UPDATE SunglassSale SET [Cust_Name] = '" & TextBox2.Text & "', [Mobile] = '" & TextBox3.Text & "', [Phone] = '" & TextBox4.Text & "', [Address] = '" & "" & "', [BookingDate] = '" & DateTimePicker1.Value & "', [ReceiptNo] = '" & TextBox5.Text & "', [DeliveryDate] = '" & DateTimePicker2.Value & "', [BookedBy] ='" & ComboBox1.Text & "', [Status] ='" & TextBox39.Text & "', [GrossAmt] ='" & TextBox25.Text & "', [CGST] = '" & TextBox37.Text & "', [SGST] = '" & TextBox38.Text & "', [ScmAmt] ='" & TextBox26.Text & "', [NetAmt] = '" & TextBox27.Text & "', [PaidAmt] = '" & TextBox28.Text & "', [DueAmt] = '" & TextBox29.Text & "', [Paymode] = '" & TextBox33.Text & "' Where [Cust_ID] = '" & TextBox1.Text & "'"
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
            str1 = "INSERT INTO SunglassProduct ([Cust_ID], [Cust_Name], [Mobile], [InvDate], [ProdName], [Qty], [Price], [Discount], [Total]) VALUES ('" & TextBox5.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & DateTimePicker1.Value.Date & "', ?, ?, ?, ?, ?)"
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
            SendSMS()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Clear()
        ClearText()
        auto()
        autoID()
        Button4.Enabled = False
    End Sub

    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            Dim i As Integer
            i = Me.DataGridView1.CurrentRow.Index
            Me.TextBox50.Text = DataGridView1.Item(0, i).Value.ToString
        Catch ex As Exception
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

    Sub autoID()
        Try
            TextBox13.Text = GenerateCode()
            TextBox13.Text = "INV-" + GenerateCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class