Imports System.Data.OleDb

Public Class Login

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Jet OLEDB:Database Password=brutusbozo;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()

        'the query:
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM [AddUser] WHERE [UserType] = '" & TextBox1.Text & "' AND [password] = '" & TextBox2.Text & "'", myConnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        ' the following variable is hold true if user is found, and false if user is not found 
        Dim userFound As Boolean = False
        ' the following variables will hold the user first and last name if found.
        Dim UserType As String = ""

        'if found:
        While dr.Read
            userFound = True
            UserType = dr("UserType").ToString
        End While

        'checking the result
        If userFound = True Then
            MainMenu.Show()
            Me.Hide()
            If UserType = "Admin" Then
                MainMenu.Show()
                Me.Hide()
            End If
        Else
            MsgBox("Sorry, username or password not found", MsgBoxStyle.OkOnly, "Invalid Login")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
        myConnection.Close()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
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

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
