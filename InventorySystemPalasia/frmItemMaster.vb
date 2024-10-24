Imports System.Data.OleDb

Public Class frmItemMaster

    Private Sub frmItemMaster_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmItemMaster_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmItemMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        ComboBox4.SelectedIndex = -1
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox3.Focus()
        End If
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox3.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        frmProductType.ShowDialog()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        frmCategoryMaster.ShowDialog()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        frmSubCategoryMaster.ShowDialog()
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

    Sub Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox3.SelectedIndex = -1
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        ComboBox4.SelectedIndex = -1
        TextBox1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ComboBox1.Text = "" Then
                MessageBox.Show("Please Select Product Type", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox1.Focus()
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MessageBox.Show("Please Select Category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox2.Focus()
                Exit Sub
            End If
            If TextBox1.Text = "" Then
                MessageBox.Show("Please Enter Product Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Focus()
                Exit Sub
            End If
            If TextBox2.Text = "" Then
                MessageBox.Show("Please Enter Price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox2.Focus()
                Exit Sub
            End If
            If TextBox3.Text = "" Then
                MessageBox.Show("Please Enter Stock-In-Hand", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox3.Focus()
                Exit Sub
            End If
            If TextBox4.Text = "" Then
                MessageBox.Show("Please Enter Stock Reorder Point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox4.Focus()
                Exit Sub
            End If
            If TextBox5.Text = "" Then
                MessageBox.Show("Please Enter Stock Limit", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox5.Focus()
                Exit Sub
            End If
            If ComboBox4.Text = "" Then
                MessageBox.Show("Please Select Tax Category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox4.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into Itemmaster ([Prod_Type], [Cat_Type], [SubCat_Type], [Prod_Name], [Price], [Stock], [Reorder], [Limit], [TaxCat]) values (?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("Prod_Type", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Cat_Type", CType(ComboBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("SubCat_Type", CType(ComboBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Prod_Name", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Price", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Stock", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Reorder", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Limit", CType(TextBox5.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("TaxCat", CType(ComboBox4.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MessageBox.Show("Product Added", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshData()
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox2.Text = ""
            ComboBox3.SelectedIndex = -1
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            ComboBox4.SelectedIndex = -1
            ComboBox1.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Prod_Type As Type, Cat_Type As Category, SubCat_Type As Sub_Category, Prod_Name, Price, Stock, Reorder, Limit, TaxCat from ItemMaster"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "UPDATE ItemMaster SET [Prod_Type] = '" & ComboBox1.Text & "', [Cat_Type] = '" & ComboBox2.Text & "', [SubCat_Type] = '" & ComboBox3.Text & "', [Prod_Name] = '" & TextBox1.Text & "', [Price] = '" & TextBox2.Text & "', [Stock] = '" & TextBox3.Text & "', [Reorder] = '" & TextBox4.Text & "', [Limit] = '" & TextBox5.Text & "', [TaxCat] = '" & ComboBox4.Text & "' Where [Prod_Name] = '" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MsgBox("Record Updated Successfuly...!", MsgBoxStyle.OkOnly)
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox2.Text = ""
            ComboBox3.SelectedIndex = -1
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            ComboBox4.SelectedIndex = -1
            ComboBox1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        RefreshData()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to delete the current record?", MsgBoxStyle.YesNo)
        If okToDelete = MsgBoxResult.Yes Then
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "Delete from ItemMaster Where Prod_Name = '" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
                ComboBox1.SelectedIndex = -1
                ComboBox2.SelectedIndex = -1
                ComboBox2.Text = ""
                ComboBox3.SelectedIndex = -1
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                TextBox5.Clear()
                ComboBox4.SelectedIndex = -1
                ComboBox1.Focus()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf okToDelete = MsgBoxResult.No Then
        End If
        RefreshData()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox2.Text = ""
        ComboBox3.SelectedIndex = -1
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        ComboBox4.SelectedIndex = -1
        ComboBox1.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        RefreshData()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Prod_Type, Cat_Type, SubCat_Type, Prod_Name, Price, Stock, Reorder, Limit, TaxCat FROM ItemMaster WHERE Prod_Name LIKE'%" &
        TextBox6.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()
    End Sub

    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.ComboBox1.Text = DataGridView1.Item(0, i).Value.ToString
            Me.ComboBox2.Text = DataGridView1.Item(1, i).Value.ToString
            Me.ComboBox3.Text = DataGridView1.Item(2, i).Value.ToString
            Me.TextBox1.Text = DataGridView1.Item(3, i).Value.ToString
            Me.TextBox2.Text = DataGridView1.Item(4, i).Value.ToString
            Me.TextBox3.Text = DataGridView1.Item(5, i).Value.ToString
            Me.TextBox4.Text = DataGridView1.Item(6, i).Value.ToString
            Me.TextBox5.Text = DataGridView1.Item(7, i).Value.ToString
            Me.ComboBox4.Text = DataGridView1.Item(8, i).Value.ToString
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try
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
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox4.Focus()
        End If
    End Sub

    Private Sub ComboBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub
End Class