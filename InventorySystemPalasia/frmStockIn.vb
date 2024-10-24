Imports System.Data.OleDb

Public Class frmStockIn

    Private Sub frmStockIn_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmStockIn_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmStockIn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        ComboBox1.SelectedIndex = -1
        auto()
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        DateTimePicker1.Value = Date.Now
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

        Dim cmd As New OleDbCommand("SELECT Type FROM ProductType;", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox1.DisplayMember = "Type"
        ComboBox1.ValueMember = "Type"
        ComboBox1.DataSource = dt

        cn.Close()
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct As String = "SELECT distinct RTRIM(CategoryName) FROM Category,ProductType where Category.CatType=ProductType.Type and Type=@d1"
            Dim cmd As New OleDbCommand(ct)
            cmd.Connection = myConnection
            cmd.Parameters.AddWithValue("@d1", ComboBox1.Text)
            rdr = cmd.ExecuteReader
            ComboBox2.Items.Clear()
            While rdr.Read
                ComboBox2.Items.Add(rdr(0))
            End While
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim conn As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        conn.ConnectionString = connString
        conn.Open()
        Dim ct1 As String = "SELECT distinct RTRIM(SubCatName) FROM SubCategory,Category Where SubCategory.CatType=Category.CategoryName and CategoryName=@d2"
        Dim cmd1 As New OleDbCommand(ct1)
        cmd1.Connection = conn
        cmd1.Parameters.AddWithValue("@d2", ComboBox2.Text)
        rdr1 = cmd1.ExecuteReader
        ComboBox3.Items.Clear()
        While rdr1.Read
            ComboBox3.Items.Add(rdr1(0))
        End While
        conn.Close()
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox3.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim conn As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        conn.ConnectionString = connString
        conn.Open()
        Dim ct1 As String = "SELECT distinct RTRIM(Prod_Name) FROM ItemMaster,SubCategory Where ItemMaster.SubCat_Type=SubCategory.SubCatName and SubCatName=@d3"
        Dim cmd2 As New OleDbCommand(ct1)
        cmd2.Connection = conn
        cmd2.Parameters.AddWithValue("@d3", ComboBox3.Text)
        rdr2 = cmd2.ExecuteReader
        ComboBox4.Items.Clear()
        While rdr2.Read
            ComboBox4.Items.Add(rdr2(0))
        End While
        conn.Close()
    End Sub

    Private Sub ComboBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox4.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "SELECT Price, Stock, Limit FROM ItemMaster WHERE (Prod_Name = '" & ComboBox4.Text & "')"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        dr = cmd.ExecuteReader
        While dr.Read()
            TextBox4.Focus()
            TextBox2.Text = dr("Price").ToString
            TextBox3.Text = dr("Stock").ToString
            TextBox6.Text = dr("Limit").ToString
        End While
        myConnection.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ComboBox1.Text = "" Then
                MessageBox.Show("Please Select Type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox1.Focus()
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MessageBox.Show("Please Select Category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox2.Focus()
                Exit Sub
            End If
            If ComboBox3.Text = "" Then
                MessageBox.Show("Please Select Sub-Category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox3.Focus()
                Exit Sub
            End If
            If ComboBox4.Text = "" Then
                MessageBox.Show("Please Select Product Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox4.Focus()
                Exit Sub
            End If
            If TextBox4.Text = "" Then
                MessageBox.Show("Please Enter Add Quantity", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox4.Focus()
                Exit Sub
            End If
            If TextBox5.Text = "" Then
                MessageBox.Show("Please Enter Person Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox5.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into StockAdded ([AddID], [AddType], [AddCategory], [AddSubCat], [AddProdName], [AddPrice], [AddCurrStock], [StockAdded], [AddDate], [AddedBy]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("AddID", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddType", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddCategory", CType(ComboBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddSubCat", CType(ComboBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddProdName", CType(ComboBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddPrice", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddCurrStock", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("StockAdded", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("AddDate", CType(DateTimePicker1.Value.Date, String)))
            cmd.Parameters.Add(New OleDbParameter("AddedBy", CType(TextBox5.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MsgBox("Stock Added Successfully", vbOKOnly)

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "UPDATE ItemMaster SET [Stock] = Stock + " & Val(TextBox4.Text) & " Where [Prod_Name] = '" & ComboBox4.Text & "'"
            Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            Try
                cmd1.ExecuteNonQuery()
                cmd1.Dispose()
                myConnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
            Me.Close()
        End Try
        Clear()
        auto()
        RefreshData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Clear()
    End Sub

    Sub Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = ""
        ComboBox3.SelectedIndex = -1
        ComboBox3.Text = ""
        ComboBox4.SelectedIndex = -1
        ComboBox4.Text = ""
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        DateTimePicker1.Value = Date.Now
        ComboBox1.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM StockAdded ORDER BY ID DESC", con)
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
            TextBox1.Text = "ADD-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock from ItemMaster"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock from ItemMaster WHERE Cat_Type LIKE'%" &
        TextBox7.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If Val(TextBox4.Text) > Val(TextBox6.Text) Then
            MsgBox("Add cannot be more than Stock Limit")
            TextBox4.Clear()
            TextBox4.Focus()
        End If
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Dim i As Integer
        i = Me.DataGridView1.CurrentRow.Index
        Me.ComboBox1.Text = Me.DataGridView1.Item(0, i).Value.ToString
        Me.ComboBox2.Text = Me.DataGridView1.Item(1, i).Value.ToString
        Me.ComboBox3.Text = Me.DataGridView1.Item(2, i).Value.ToString
        Me.ComboBox4.Text = Me.DataGridView1.Item(3, i).Value.ToString
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox7.Clear()
        TextBox8.Clear()
        RefreshData()
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock from ItemMaster WHERE Prod_Name LIKE'%" &
        TextBox8.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

End Class