Imports System.Data.OleDb

Public Class frmPresentStock

    Private Sub frmPresentStock_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmPresentStock_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmPresentStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadCombo()
        ComboBox1.SelectedIndex = -1
        ComboBox1.Text = ""
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = ""
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
        ComboBox4.Text = ""
        RefreshData()
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

        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
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

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock FROM Itemmaster WHERE Prod_Type ='" & ComboBox1.Text & "'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("Items")
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
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

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock FROM Itemmaster WHERE Prod_Type='" & ComboBox1.Text & "' And " & "Cat_Type='" & ComboBox2.Text & "'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("Items")
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox3.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
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

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock FROM Itemmaster WHERE Prod_Type='" & ComboBox1.Text & "' And Cat_Type='" & ComboBox2.Text & "' And " & "SubCat_Type='" & ComboBox3.Text & "'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("Items")
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub ComboBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox4.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock FROM Itemmaster WHERE Prod_Name='" & ComboBox4.Text & "'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("Items")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim stock As Integer
        For Each row As DataGridViewRow In Me.DataGridView1.Rows
            stock += row.Cells(4).Value
        Next
        Me.TextBox1.Text = stock
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Stock FROM Itemmaster WHERE Prod_Name LIKE'%" &
        TextBox2.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub
End Class