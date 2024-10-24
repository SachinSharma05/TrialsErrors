Imports System.Data.OleDb

Public Class frmMaxSellingProducts

    Private Sub frmMaxSellingProducts_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmMaxSellingProducts_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmMaxSellingProducts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        myConnection.Open()
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "Select ProdName As Model_No, SUM(Qty) As Qty FROM InvoiceProduct GROUP BY ProdName ORDER BY Sum(Qty) DESC"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
        myConnection.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        RefreshData()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ExportExcel(DataGridView1)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub
End Class