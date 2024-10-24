Imports System.Data.OleDb

Public Class frmDamageProductEntry

    Private Sub frmDamageProductEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmDamageProductEntry_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmDamageProductEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = Date.Now
        LoadCombo()
        ComboBox1.SelectedIndex = -1
        TextBox2.Clear()
        auto()
        RefreshData()
        DataGridView2.Sort(DataGridView2.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows.Add(ComboBox1.Text, ComboBox2.Text, ComboBox3.Text, ComboBox4.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text)
                Exit Sub
            Next
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                If r.Cells(0).Value = ComboBox1.Text Then
                    r.Cells(0).Value = ComboBox1.Text
                    r.Cells(1).Value = ComboBox2.Text
                    r.Cells(2).Value = ComboBox3.Text
                    r.Cells(3).Value = ComboBox4.Text
                    r.Cells(4).Value = TextBox1.Text
                    r.Cells(5).Value = TextBox2.Text
                    r.Cells(6).Value = TextBox3.Text
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(ComboBox1.Text, ComboBox2.Text, ComboBox3.Text, ComboBox4.Text, TextBox1.Text, TextBox2.Text, TextBox3.Text)
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1
            ComboBox4.SelectedIndex = -1
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            ComboBox1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

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

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
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
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub Compute()
        Dim i As Double = 0
        i = Val(TextBox2.Text) * Val(TextBox1.Text)
        i = Math.Round(i, 2)
        TextBox3.Text = i
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Compute()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Focus()
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

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
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
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "SELECT Price FROM ItemMaster WHERE (Prod_Name = '" & ComboBox4.Text & "')"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        dr = cmd.ExecuteReader
        While dr.Read()
            TextBox1.Focus()
            TextBox2.Text = dr("Price").ToString
        End While
        myConnection.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "INSERT INTO DamageProducts ([Code], [DamDate], [Type], [Category], [SubCategory], [Prod_Name], [Qty], [Price], [Total]) VALUES ('" & TextBox4.Text & "', '" & DateTimePicker1.Value.Date & "', ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd1.Parameters.Add(New OleDbParameter("Type", row.Cells(0).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Category", row.Cells(1).Value))
                    cmd1.Parameters.Add(New OleDbParameter("SubCategory", row.Cells(2).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Prod_Name", row.Cells(3).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Qty", row.Cells(4).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Price", row.Cells(5).Value))
                    cmd1.Parameters.Add(New OleDbParameter("Total", row.Cells(6).Value))
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
                    Dim cb4 As String = "update ItemMaster set Stock = Stock - (" & row.Cells(4).Value & ") where Prod_Name= ComboBox4.Text"
                    Dim cmd2 As New OleDbCommand
                    cmd2 = New OleDbCommand(cb4)
                    cmd2.Connection = myConnection
                    cmd2.Parameters.Add(New OleDbParameter("& ComboBox4.Text &", row.Cells(3).Value))
                    cmd2.ExecuteNonQuery()
                    myConnection.Close()
                End If
            Next
            MsgBox("Damage Product Entered Successfuly...!", MsgBoxStyle.OkOnly)
            RefreshData()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView2.DataSource = Nothing
        DataGridView2.Refresh()
        Dim str As String = "select Type, Category, SubCategory, Prod_Name, Qty, Price, Total from DamageProducts"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView2.DataSource = newtable
                End Using
            End Using
        End Using
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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM DamageProducts ORDER BY ID DESC", con)
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
            TextBox4.Text = GenerateID()
            TextBox4.Text = "DAM-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
End Class