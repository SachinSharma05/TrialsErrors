Imports System.Data.OleDb

Public Class frmChangePassword

    Private Sub frmChangePassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmChangePassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadCombo()
        ComboBox1.SelectedIndex = -1
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

        Dim cmd As New OleDbCommand("SELECT UserType FROM AddUser;", cn)
        Dim dr = cmd.ExecuteReader()

        Dim dt As New DataTable()
        dt.Load(dr)
        dr.Close()

        ComboBox1.DisplayMember = "UserType"
        ComboBox1.ValueMember = "UserType"
        ComboBox1.DataSource = dt

        cn.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "UPDATE AddUser SET [Password] = '" & TextBox3.Text & "' Where [UserType] = '" & ComboBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MsgBox("User Password Updated Successfully...!", MsgBoxStyle.OkOnly)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class