Imports System.Data.OleDb

Public Class frmProductType

    Private Sub frmProductType_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmProductType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmProductType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Select()
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

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Type from ProductType"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
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

    Private Sub DataGridView1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.TextBox1.Text = DataGridView1.Item(0, i).Value.ToString
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If TextBox1.Text = "" Then
                MsgBox("Please enter Product Type")
                TextBox1.Focus()
                Exit Sub
            End If
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into ProductType ([Type]) values (?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("Type", CType(TextBox1.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MessageBox.Show("Product Type Added", "Product Type Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            RefreshData()
            TextBox1.Clear()
            TextBox1.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
            Me.Close()
        End Try
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to delete the selected category?", MsgBoxStyle.YesNo)
        If okToDelete = MsgBoxResult.Yes Then
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "Delete from ProductType Where Type = '" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
                TextBox1.Clear()
                TextBox1.Focus()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        ElseIf okToDelete = MsgBoxResult.No Then
        End If
        RefreshData()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub
End Class