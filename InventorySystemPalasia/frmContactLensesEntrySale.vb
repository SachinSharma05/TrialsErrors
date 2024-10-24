Imports System.Data.OleDb
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine

Public Class frmContactLensesEntrySale

    Private Sub frmContactLensesEntrySale_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmContactLensesEntrySale_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Function GetValue(ByVal Value As Object) As String
        If Value IsNot Nothing Then Return Value.ToString() Else Return ""
    End Function

    Private Sub frmContactLensesEntrySale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo2()
        Me.ComboBox2.SelectedIndex = -1
        Me.ComboBox2.Text = ""
        TextBox8.Text = "0"
        TextBox35.Text = "0"
        TextBox36.Text = "0"
        auto()
        autoID()
        DateTimePicker1.Value = Date.Now
        DateTimePicker2.Value = Date.Now
        RefreshData()
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        Clear()
        ClearText()
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM CLSale ORDER BY ID DESC", con)
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
            TextBox1.Text = "CL-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
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

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
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
        ComboBox1.SelectedIndex = -1
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
    End Sub

    Sub LoadCombo2()
        Dim cn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;")
        cn.Open()

        Dim cmd As New OleDbCommand("SELECT Prod_Name FROM ItemMaster Where Prod_Type='CONTACT LENSES';", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox2.DisplayMember = "Prod_Name"
        ComboBox2.ValueMember = "Prod_Name"
        ComboBox2.DataSource = dt

        cn.Close()
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

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            TextBox7.ReadOnly = True
            TextBox8.ReadOnly = True
        Else
            TextBox7.ReadOnly = False
            TextBox8.ReadOnly = False
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

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Compute()
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
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
        If Me.TextBox8.Text = "" Then
            TextBox8.Text = "0"
        End If
    End Sub

    Private Sub TextBox28_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox28.TextChanged
        If Val(TextBox28.Text) > Val(TextBox27.Text) Then
            MsgBox("Advance cannot be more than Net Amount")
            TextBox28.Clear()
        End If
        Compute2()
    End Sub

    Sub Compute2()
        Dim i As Double = 0
        i = Val(TextBox27.Text) - Val(TextBox28.Text)
        i = Math.Round(i, 2)
        TextBox29.Text = i
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
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
        ComboBox1.SelectedIndex = -1
        TextBox24.Clear()
        TextBox11.Focus()
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Clear()
        ClearText()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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
            If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Please Select Paymode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                RadioButton1.Focus()
                Exit Sub
            End If
            If TextBox28.Text = "" Then
                MessageBox.Show("Please enter Advance Amt", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox28.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into CLSale ([Cust_ID], [Cust_Name], [Mobile], [Phone], [BookingDate], [DeliveryDate], [Status], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode], [Remarks], [RSPH], [RCYL], [RAXIS], [RVN], [RADD], [LSPH], [LCYL], [LAXIS], [LVN], [LADD], [PD], [REFBY], [LensType], [Remarks1]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Phone", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd.Parameters.Add(New OleDbParameter("DeliveryDate", CType(DateTimePicker2.Value.Date, String)))
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
            cmd.Parameters.Add(New OleDbParameter("LensType", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Remarks1", CType(TextBox24.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "INSERT INTO CLSaleProduct ([Cust_ID], [ProdName], [Qty], [Price], [Discount], [Total]) VALUES ('" & TextBox1.Text & "', ?, ?, ?, ?, ?)"
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
            Dim str3 As String
            str3 = "INSERT INTO PaymentVoucher ([ReceiptName], [ReceiptNo], [ReceiptDate], [ReceiptAmt], [ReceiptBal], [ReceiptDue], [ReceiptStatus]) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim cmd4 As OleDbCommand = New OleDbCommand(str3, myConnection)
            cmd4.Parameters.Add(New OleDbParameter("ReceiptName", CType(TextBox2.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox1.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDate", CType(DateTimePicker1.Value, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptAmt", CType(TextBox27.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptBal", CType(TextBox28.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptDue", CType(TextBox29.Text, String)))
            cmd4.Parameters.Add(New OleDbParameter("ReceiptStatus", CType(TextBox34.Text, String)))
            cmd4.ExecuteNonQuery()
            cmd4.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str4 As String
            str4 = "INSERT INTO BilledInvoice ([Cust_ID], [Cust_Name], [Mobile], [Phone], [BookingDate], [ReceiptNo], [BookedBy], [Status], [GrossAmt], [CGST], [SGST], [ScmAmt], [NetAmt], [PaidAmt], [DueAmt], [Paymode]) VALUES ('" & TextBox42.Text & "', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd5 As OleDbCommand = New OleDbCommand(str4, myConnection)
            cmd5.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Phone", CType(TextBox4.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd5.Parameters.Add(New OleDbParameter("ReceiptNo", CType(TextBox1.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("BookedBy", CType(TextBox22.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Status", CType(TextBox39.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox25.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("CGST", CType(TextBox37.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("SGST", CType(TextBox38.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox26.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox27.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("PaidAmt", CType(TextBox28.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox29.Text, String)))
            cmd5.Parameters.Add(New OleDbParameter("Paymode", CType(TextBox34.Text, String)))
            cmd5.ExecuteNonQuery()
            cmd5.Dispose()
            myConnection.Close()

            MsgBox("Sale Created Successfuly...!", MsgBoxStyle.OkOnly)
            Dim okToPrint As MsgBoxResult = MsgBox("Press Yes for JOB-CARD, Press No for INVOICE?", MsgBoxStyle.YesNoCancel)
            If okToPrint = MsgBoxResult.Yes Then
                Print()
            ElseIf okToPrint = MsgBoxResult.No Then
                Print1()
            End If
            RefreshData()
            Clear()
            ClearText()
            Me.DataGridView1.Rows.Clear()
            auto()
            autoID()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Clear()
        ClearText()
        Me.DataGridView1.Rows.Clear()
        auto()
        autoID()
    End Sub

    Sub Print()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New CLJobCard 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select CLSale.Cust_ID, CLSale.Cust_Name, CLSale.Mobile, CLSale.BookingDate, CLSale.DeliveryDate, CLSale.REFBY, CLSale.NetAmt, CLSale.PaidAmt, CLSale.DueAmt, CLSale.Paymode, CLSale.RSPH, CLSale.RCYL, CLSale.RAXIS, CLSale.RVN, CLSale.LSPH, CLSale.LCYL, CLSale.LAXIS, CLSale.LVN, CLSale.LensType, CLSale.Remarks1, CLSaleProduct.ProdName, CLSaleProduct.Qty, CLSaleProduct.Price, CLSaleProduct.Discount, CLSaleProduct.Total FROM CLSaleProduct INNER JOIN CLSale ON CLSale.Cust_ID=CLSaleProduct.Cust_ID Where CLSale.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox1.Text)
            MyCommand1.CommandText = "SELECT * from CLSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "CLSale")
            myDA.Fill(myDS, "CLSaleProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox1.Text)
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
            Dim rpt As New CLInvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select CLSale.Cust_ID, CLSale.Cust_Name, CLSale.Mobile, CLSale.BookingDate, CLSale.GrossAmt, CLSale.NetAmt, CLSale.PaidAmt, CLSale.DueAmt, CLSale.Paymode, CLSale.ScmAmt, CLSaleProduct.ProdName, CLSaleProduct.Qty, CLSaleProduct.Price, CLSaleProduct.Discount, CLSaleProduct.Total FROM CLSaleProduct INNER JOIN CLSale ON CLSale.Cust_ID=CLSaleProduct.Cust_ID Where CLSale.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox1.Text)
            MyCommand1.CommandText = "SELECT * from CLSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "CLSale")
            myDA.Fill(myDS, "CLSaleProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox1.Text)
            rpt.SetParameterValue("p3", TextBox42.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        RefreshData()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select * from CLSale"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(22).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
        Me.DataGridView2.Columns(25).Visible = False
        Me.DataGridView2.Columns(26).Visible = False
        Me.DataGridView2.Columns(27).Visible = False

        Me.DataGridView2.Columns(1).Width = 100
        Me.DataGridView2.Columns(2).Width = 250
        Me.DataGridView2.Columns(3).Width = 100
        Me.DataGridView2.Columns(4).Width = 100
        Me.DataGridView2.Columns(28).Width = 150
        Me.DataGridView2.Columns(29).Width = 200
        Me.DataGridView2.Columns(30).Width = 150
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from CLSale WHERE Cust_Name LIKE'%" &
        TextBox30.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(22).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
        Me.DataGridView2.Columns(25).Visible = False
        Me.DataGridView2.Columns(26).Visible = False
        Me.DataGridView2.Columns(27).Visible = False
    End Sub

    Private Sub TextBox31_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox31.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from CLSale WHERE Mobile LIKE'%" &
        TextBox31.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(22).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
        Me.DataGridView2.Columns(25).Visible = False
        Me.DataGridView2.Columns(26).Visible = False
        Me.DataGridView2.Columns(27).Visible = False
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from CLSale WHERE Phone LIKE'%" &
        TextBox5.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(22).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
        Me.DataGridView2.Columns(25).Visible = False
        Me.DataGridView2.Columns(26).Visible = False
        Me.DataGridView2.Columns(27).Visible = False
    End Sub

    Private Sub TextBox23_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox23.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from CLSale WHERE LensType LIKE'%" &
        TextBox23.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView2.DataSource = dt
        myConnection.Close()

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(6).Visible = False
        Me.DataGridView2.Columns(7).Visible = False
        Me.DataGridView2.Columns(9).Visible = False
        Me.DataGridView2.Columns(10).Visible = False
        Me.DataGridView2.Columns(11).Visible = False
        Me.DataGridView2.Columns(12).Visible = False
        Me.DataGridView2.Columns(13).Visible = False
        Me.DataGridView2.Columns(14).Visible = False
        Me.DataGridView2.Columns(15).Visible = False
        Me.DataGridView2.Columns(17).Visible = False
        Me.DataGridView2.Columns(18).Visible = False
        Me.DataGridView2.Columns(19).Visible = False
        Me.DataGridView2.Columns(20).Visible = False
        Me.DataGridView2.Columns(21).Visible = False
        Me.DataGridView2.Columns(22).Visible = False
        Me.DataGridView2.Columns(23).Visible = False
        Me.DataGridView2.Columns(24).Visible = False
        Me.DataGridView2.Columns(25).Visible = False
        Me.DataGridView2.Columns(26).Visible = False
        Me.DataGridView2.Columns(27).Visible = False
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            DateTimePicker2.Focus()
        End If
    End Sub

    Private Sub DateTimePicker2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox11.Focus()
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
            TextBox16.Focus()
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
            TextBox22.Focus()
        End If
    End Sub

    Private Sub TextBox20_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox22.Focus()
        End If
    End Sub

    Private Sub TextBox22_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox24.Focus()
        End If
    End Sub

    Private Sub TextBox24_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Dim i As Integer
            i = Me.DataGridView2.CurrentRow.Index
            Me.TextBox1.Text = Me.DataGridView2.Item(1, i).Value.ToString
            Me.TextBox2.Text = Me.DataGridView2.Item(2, i).Value.ToString
            Me.TextBox3.Text = Me.DataGridView2.Item(3, i).Value.ToString
            Me.TextBox4.Text = Me.DataGridView2.Item(4, i).Value.ToString
            Me.DateTimePicker1.Value = Me.DataGridView2.Item(5, i).Value.ToString
            Me.DateTimePicker1.Value = Me.DataGridView2.Item(6, i).Value.ToString
            Me.TextBox39.Text = Me.DataGridView2.Item(7, i).Value.ToString
            Me.TextBox25.Text = Me.DataGridView2.Item(8, i).Value.ToString
            Me.TextBox37.Text = Me.DataGridView2.Item(9, i).Value.ToString
            Me.TextBox38.Text = Me.DataGridView2.Item(10, i).Value.ToString
            Me.TextBox26.Text = Me.DataGridView2.Item(11, i).Value.ToString
            Me.TextBox27.Text = Me.DataGridView2.Item(12, i).Value.ToString
            Me.TextBox28.Text = Me.DataGridView2.Item(13, i).Value.ToString
            Me.TextBox29.Text = Me.DataGridView2.Item(14, i).Value.ToString
            Me.TextBox41.Text = Me.DataGridView2.Item(16, i).Value.ToString
            Me.TextBox11.Text = Me.DataGridView2.Item(17, i).Value.ToString
            Me.TextBox12.Text = Me.DataGridView2.Item(18, i).Value.ToString
            Me.TextBox13.Text = Me.DataGridView2.Item(19, i).Value.ToString
            Me.TextBox14.Text = Me.DataGridView2.Item(20, i).Value.ToString
            Me.TextBox15.Text = Me.DataGridView2.Item(21, i).Value.ToString
            Me.TextBox16.Text = Me.DataGridView2.Item(22, i).Value.ToString
            Me.TextBox17.Text = Me.DataGridView2.Item(23, i).Value.ToString
            Me.TextBox18.Text = Me.DataGridView2.Item(24, i).Value.ToString
            Me.TextBox19.Text = Me.DataGridView2.Item(25, i).Value.ToString
            Me.TextBox20.Text = Me.DataGridView2.Item(26, i).Value.ToString
            Me.TextBox22.Text = Me.DataGridView2.Item(28, i).Value.ToString
            Me.ComboBox1.Text = Me.DataGridView2.Item(29, i).Value.ToString
            Me.TextBox24.Text = Me.DataGridView2.Item(30, i).Value.ToString

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT ProdName, Price, Qty, Discount, Total FROM CLSaleProduct WHERE Cust_ID LIKE'%" &
            TextBox33.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("InvoiceProduct")
            adapter.Fill(dt)
            Me.DataGridView1.DataSource = dt
            myConnection.Close()
        Catch ex As Exception
            MsgBox("Row is empty")
        End Try
    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        Try
            Dim i As Integer
            i = Me.DataGridView2.CurrentRow.Index
            Me.TextBox33.Text = Me.DataGridView2.Item(1, i).Value
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox34.Text = RadioButton1.Text
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            TextBox34.Text = RadioButton2.Text
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            TextBox34.Text = RadioButton3.Text
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
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
            TextBox42.Text = GenerateCode()
            TextBox42.Text = "INV-" + GenerateCode()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class