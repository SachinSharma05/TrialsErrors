Imports System.Data.OleDb

Public Class frmSubCategoryMaster

    Private Sub frmSubCategoryMaster_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmSubCategoryMaster_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmSubCategoryMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        ComboBox1.Focus()
        RefreshData()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
        TextBox1.Clear()
        ComboBox1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ComboBox1.Text = "" Then
                MsgBox("Please Select Product Type")
                ComboBox1.Focus()
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MsgBox("Please Select Category Name")
                ComboBox2.Focus()
                Exit Sub
            End If
            If TextBox1.Text = "" Then
                MsgBox("Please Enter Sub-Category Name")
                TextBox1.Focus()
                Exit Sub
            End If
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into SubCategory ([ProdType], [CatType], [SubCatName]) values (?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("ProdType", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("CatType", CType(ComboBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("SubCatName", CType(TextBox1.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MessageBox.Show("Sub-Category Type Added", "Sub-Category Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshData()
            TextBox1.Clear()
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox2.Text = ""
            ComboBox1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
            Me.Close()
        End Try
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select ProdType, CatType, SubCatName from SubCategory"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to delete the selected category?", MsgBoxStyle.YesNo)
        If okToDelete = MsgBoxResult.Yes Then
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "Delete from SubCategory Where SubCatName = '" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
                TextBox1.Clear()
                ComboBox1.SelectedIndex = -1
                ComboBox2.SelectedIndex = -1
                ComboBox1.Focus()
                ComboBox1.SelectedIndex = -1
                ComboBox1.Focus()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf okToDelete = MsgBoxResult.No Then
        End If
        RefreshData()
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

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub ComboBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox2.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.ComboBox1.Text = DataGridView1.Item(0, i).Value.ToString
            Me.ComboBox2.Text = DataGridView1.Item(1, i).Value.ToString
            Me.TextBox1.Text = DataGridView1.Item(2, i).Value.ToString
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try
    End Sub
End Class