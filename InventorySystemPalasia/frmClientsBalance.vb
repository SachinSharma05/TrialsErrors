Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class frmClientsBalance

    Private Sub frmClientsBalance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmClientsBalance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        DateTimePicker1.Value = Date.Now
        TextBox1.select
        auto()
        Button4.Enabled = False
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            DateTimePicker1.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox3.Focus()
        End If
    End Sub

    Sub Clear()
        TextBox6.Clear()
        DateTimePicker1.Value = Date.Now
        ComboBox1.SelectedIndex = -1
        TextBox5.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox4.Clear()
        TextBox15.Clear()
        TextBox6.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If TextBox6.Text = "" Then
                MessageBox.Show("Please enter JobCard No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox6.Focus()
                Exit Sub
            End If
            If ComboBox1.Text = "" Then
                MessageBox.Show("Please select Booked By", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBox1.Focus()
                Exit Sub
            End If
            If TextBox5.Text = "" Then
                MessageBox.Show("Please enter Client Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox5.Focus()
                Exit Sub
            End If
            If TextBox7.Text = "" Then
                MessageBox.Show("Please enter GrossAmt", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox7.Focus()
                Exit Sub
            End If
            If TextBox8.Text = "" Then
                MessageBox.Show("Please enter Scheme Amt", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox8.Focus()
                Exit Sub
            End If
            If TextBox9.Text = "" Then
                MessageBox.Show("Please enter Net Amount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox9.Focus()
                Exit Sub
            End If
            For i As Integer = 0 To DataGridView3.Rows.Count - 1
                DataGridView3.Rows.Add(TextBox6.Text, DateTimePicker1.Value, ComboBox1.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox4.Text, TextBox15.Text)
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                TextBox10.Text = k
                Dim c As Double = 0
                c = TotalDiscount()
                c = Math.Round(c, 2)
                TextBox13.Text = c
                Dim x As Double = 0
                x = TotalPayment()
                x = Math.Round(x, 2)
                TextBox18.Text = x
                Dim h As Double = 0
                h = TotalAdvance()
                h = Math.Round(h, 2)
                TextBox20.Text = h
                Dim u As Double = 0
                u = TotalDue()
                u = Math.Round(u, 2)
                TextBox19.Text = u
                Clear()
                Exit Sub
            Next

            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                If r.Cells(0).Value = TextBox6.Text Then
                    r.Cells(0).Value = TextBox6.Text
                    r.Cells(3).Value = TextBox6.Text
                    r.Cells(4).Value = Val(r.Cells(4).Value) + Val(DateTimePicker1.Value)
                    r.Cells(5).Value = Val(r.Cells(5).Value) + Val(ComboBox1.Text)
                    r.Cells(6).Value = Val(r.Cells(6).Value) + Val(TextBox5.Text)
                    r.Cells(7).Value = Val(r.Cells(6).Value) + Val(TextBox7.Text)
                    r.Cells(8).Value = Val(r.Cells(8).Value) + Val(TextBox8.Text)
                    r.Cells(9).Value = Val(r.Cells(9).Value) + Val(TextBox9.Text)
                    r.Cells(10).Value = Val(r.Cells(10).Value) + Val(TextBox4.Text)
                    r.Cells(11).Value = Val(r.Cells(11).Value) + Val(TextBox15.Text)
                    Dim i As Double = 0
                    i = GrandTotal()
                    i = Math.Round(i, 2)
                    TextBox10.Text = i
                    Dim a As Double = 0
                    a = TotalDiscount()
                    a = Math.Round(a, 2)
                    TextBox13.Text = a
                    Dim q As Double
                    q = TotalPayment()
                    q = Math.Round(q, 2)
                    TextBox18.Text = q
                    Dim w As Double = 0
                    w = TotalAdvance()
                    w = Math.Round(w, 2)
                    TextBox20.Text = w
                    Dim p As Double = 0
                    p = TotalDue()
                    p = Math.Round(p, 2)
                    TextBox19.Text = p
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView3.Rows.Add(TextBox6.Text, DateTimePicker1.Value, ComboBox1.Text, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox4.Text, TextBox15.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            TextBox10.Text = j
            Dim b As Double = 0
            b = TotalDiscount()
            b = Math.Round(b, 2)
            TextBox13.Text = b
            Dim z As Double
            z = TotalPayment()
            z = Math.Round(z, 2)
            TextBox18.Text = z
            Dim g As Double = 0
            g = TotalAdvance()
            g = Math.Round(g, 2)
            TextBox20.Text = g
            Dim s As Double = 0
            s = TotalDue()
            s = Math.Round(s, 2)
            TextBox19.Text = s
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function GrandTotal() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(4).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function TotalPayment() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                sum = sum + r.Cells(6).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function

    Public Function TotalDiscount() As Double
        Dim Dis As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                Dis = Dis + r.Cells(5).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Dis
    End Function

    Public Function TotalAdvance() As Double
        Dim Adv As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                Adv = Adv + r.Cells(7).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Adv
    End Function

    Public Function TotalDue() As Double
        Dim Due As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                Due = Due + r.Cells(8).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Due
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            For Each row As DataGridViewRow In DataGridView3.SelectedRows
                DataGridView3.Rows.Remove(row)
            Next
            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            TextBox10.Text = k
            Dim c As Double = 0
            c = TotalDiscount()
            c = Math.Round(c, 2)
            TextBox13.Text = c
            Dim x As Double
            x = TotalPayment()
            x = Math.Round(x, 2)
            TextBox14.Text = x
            Dim g As Double = 0
            g = TotalAdvance()
            g = Math.Round(g, 2)
            TextBox20.Text = g
            Dim s As Double = 0
            s = TotalDue()
            s = Math.Round(s, 2)
            TextBox19.Text = s
            Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ClearAll()
        auto()
    End Sub

    Sub ClearAll()
        Clear()
        DataGridView3.Rows.Clear()
        DataGridView3.DataSource = Nothing
        TextBox10.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox12.Clear()
        TextBox11.Clear()
        TextBox4.Clear()
        TextBox15.Clear()
        TextBox20.Clear()
        TextBox19.Clear()
        TextBox18.Clear()
        TextBox22.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        TextBox21.Clear()
        ComboBox2.SelectedIndex = -1
        TextBox2.Focus()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Please enter Group Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MessageBox.Show("Please enter Group Contact No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox3.Focus()
            Exit Sub
        End If

        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim ct As String = "select Cust_ID from BalanceSale where Cust_ID=@d1"
        Dim cmd9 As OleDbCommand = New OleDbCommand(ct)
        cmd9.Parameters.AddWithValue("@d1", TextBox1.Text)
        cmd9.Connection = myConnection
        rdr = cmd9.ExecuteReader()
        If rdr.Read() Then
            MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            auto()
        End If
        myConnection.Close()

        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        str = "insert into BalanceSale ([Cust_ID], [Cust_Name], [Mobile], [Phone], [GrossAmt], [ScmAmt], [NetAmt], [AdvAmt], [DueAmt], [Advance1], [Adv1Date], [Advance2], [Adv2Date], [Advance3], [Adv3Date]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox1.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Phone", CType(ComboBox2.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("GrossAmt", CType(TextBox10.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("ScmAmt", CType(TextBox13.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("NetAmt", CType(TextBox18.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("AdvAmt", CType(TextBox20.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("DueAmt", CType(TextBox19.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Advance1", CType(TextBox14.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Adv1Date", CType(DateTimePicker2.Value.Date, String)))
        cmd.Parameters.Add(New OleDbParameter("Advance2", CType(TextBox16.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Adv2Date", CType(DateTimePicker3.Value.Date, String)))
        cmd.Parameters.Add(New OleDbParameter("Advance3", CType(TextBox17.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("Adv3Date", CType(DateTimePicker4.Value.Date, String)))
        cmd.ExecuteNonQuery()
        cmd.Dispose()
        myConnection.Close()

        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str1 As String
        str1 = "INSERT INTO BalanceSaleClients ([Cust_ID], [JobCardNo], [InvDate], [BookedBy], [ClientName], [GrossAmt], [ScmAmt], [NetAmt], [AdvAmt], [DueAmt]) VALUES ('" & TextBox1.Text & "', ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
        For Each row As DataGridViewRow In DataGridView3.Rows
            If Not row.IsNewRow Then
                cmd1.Parameters.Add(New OleDbParameter("JobCardNo", row.Cells(0).Value))
                cmd1.Parameters.Add(New OleDbParameter("InvDate", row.Cells(1).Value))
                cmd1.Parameters.Add(New OleDbParameter("BookedBy", row.Cells(2).Value))
                cmd1.Parameters.Add(New OleDbParameter("ClientName", row.Cells(3).Value))
                cmd1.Parameters.Add(New OleDbParameter("GrossAmt", row.Cells(4).Value))
                cmd1.Parameters.Add(New OleDbParameter("ScmAmt", row.Cells(5).Value))
                cmd1.Parameters.Add(New OleDbParameter("NetAmt", row.Cells(6).Value))
                cmd1.Parameters.Add(New OleDbParameter("AdvAmt", row.Cells(7).Value))
                cmd1.Parameters.Add(New OleDbParameter("DueAmt", row.Cells(8).Value))
                cmd1.ExecuteNonQuery()
                cmd1.Parameters.Clear()
            End If
        Next
        myConnection.Close()
        MsgBox("Group Created Successfuly...!", MsgBoxStyle.OkOnly)
        TextBox2.Focus()
        RefreshData()
        ClearAll()
        Button4.Enabled = False
        auto()
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select * from BalanceSale"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(4).HeaderText = "Approved By"
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM BalanceSale ORDER BY ID DESC", con)
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
            TextBox1.Text = "GRP-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from BalanceSale WHERE Cust_Name LIKE'%" &
        TextBox12.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT * from BalanceSale WHERE Phone LIKE'%" &
        TextBox11.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox2.Focus()
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox12.Clear()
        TextBox11.Clear()
        RefreshData()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to update the current record?", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str6 As String
                str6 = "Delete from BalanceSaleClients Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd7 As OleDbCommand = New OleDbCommand(str6, myConnection)
                Try
                    cmd7.ExecuteNonQuery()
                    cmd7.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            ElseIf okToDelete = MsgBoxResult.No Then
                Exit Sub
            End If

            If TextBox2.Text = "" Then
                MessageBox.Show("Please enter Group Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox2.Focus()
                Exit Sub
            End If
            If TextBox3.Text = "" Then
                MessageBox.Show("Please enter Group Contact No", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox3.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "UPDATE BalanceSale SET [Cust_Name] = '" & TextBox2.Text & "', [Mobile] = '" & TextBox3.Text & "', [Phone] = '" & ComboBox2.Text & "', [GrossAmt] = '" & TextBox10.Text & "', [ScmAmt] = '" & TextBox13.Text & "', [NetAmt] = '" & TextBox14.Text & "', [AdvAmt] = '" & TextBox20.Text & "', [DueAmt] ='" & TextBox19.Text & "', [Advance1] ='" & TextBox14.Text & "', [Adv1Date] ='" & DateTimePicker2.Value & "', [Advance2] = '" & TextBox16.Text & "', [Adv2Date] = '" & DateTimePicker3.Value & "', [Advance3] ='" & TextBox17.Text & "', [Adv3Date] = '" & DateTimePicker4.Value & "' Where [Cust_ID] = '" & TextBox1.Text & "'"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str1 As String
            str1 = "INSERT INTO BalanceSaleClients ([Cust_ID], [JobCardNo], [InvDate], [BookedBy], [ClientName], [GrossAmt], [ScmAmt], [NetAmt], [AdvAmt], [DueAmt]) VALUES ('" & TextBox1.Text & "', ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            For Each row As DataGridViewRow In DataGridView3.Rows
                If Not row.IsNewRow Then
                    cmd1.Parameters.Add(New OleDbParameter("JobCardNo", row.Cells(0).Value))
                    cmd1.Parameters.Add(New OleDbParameter("InvDate", row.Cells(1).Value))
                    cmd1.Parameters.Add(New OleDbParameter("BookedBy", row.Cells(2).Value))
                    cmd1.Parameters.Add(New OleDbParameter("ClientName", row.Cells(3).Value))
                    cmd1.Parameters.Add(New OleDbParameter("GrossAmt", row.Cells(4).Value))
                    cmd1.Parameters.Add(New OleDbParameter("ScmAmt", row.Cells(5).Value))
                    cmd1.Parameters.Add(New OleDbParameter("NetAmt", row.Cells(6).Value))
                    cmd1.Parameters.Add(New OleDbParameter("AdvAmt", row.Cells(7).Value))
                    cmd1.Parameters.Add(New OleDbParameter("DueAmt", row.Cells(8).Value))
                    cmd1.ExecuteNonQuery()
                    cmd1.Parameters.Clear()
                End If
            Next
            myConnection.Close()
            MsgBox("Group Updated Successfuly...!", MsgBoxStyle.OkOnly)
            TextBox2.Focus()
            RefreshData()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Button4.Enabled = False
        ClearAll()
        auto()
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Sub GridCopy()
        Dim sourceGrid As DataGridView = Me.DataGridView2
        Dim targetGrid As DataGridView = Me.DataGridView3
        Dim targetRows = New List(Of DataGridViewRow)

        For Each sourceRow As DataGridViewRow In sourceGrid.Rows
            If (Not sourceRow.IsNewRow) Then
                Dim targetRow = CType(sourceRow.Clone(), DataGridViewRow)

                For Each cell As DataGridViewCell In sourceRow.Cells
                    targetRow.Cells(cell.ColumnIndex).Value = cell.Value
                Next

                targetRows.Add(targetRow)
            End If
        Next

        targetGrid.Columns.Clear()

        For Each column As DataGridViewColumn In sourceGrid.Columns
            targetGrid.Columns.Add(CType(column.Clone(), DataGridViewColumn))
        Next

        targetGrid.Rows.AddRange(targetRows.ToArray())
    End Sub

    Private Sub ComboBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox9.Focus()
        End If
    End Sub

    Private Sub TextBox9_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox15_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Compute()
    End Sub

    Sub Compute()
        Dim num2 As Double
        num2 = CDbl(Val(TextBox7.Text) - Val(TextBox8.Text))
        num2 = Math.Round(num2, 2)
        TextBox9.Text = num2
    End Sub

    Sub Compute1()
        Dim num2 As Double
        num2 = CDbl(Val(TextBox9.Text) - Val(TextBox4.Text))
        num2 = Math.Round(num2, 2)
        TextBox15.Text = num2
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Compute1()
    End Sub

    Private Sub DataGridView1_MouseDoubleClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Try
            Me.DataGridView3.Rows.Clear()
            Me.Refresh()

            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.TextBox1.Text = DataGridView1.Item(1, i).Value.ToString
            Me.TextBox2.Text = DataGridView1.Item(2, i).Value.ToString
            Me.TextBox3.Text = DataGridView1.Item(3, i).Value.ToString
            Me.ComboBox2.Text = DataGridView1.Item(4, i).Value.ToString
            Me.TextBox10.Text = DataGridView1.Item(5, i).Value.ToString
            Me.TextBox13.Text = DataGridView1.Item(6, i).Value.ToString
            Me.TextBox18.Text = DataGridView1.Item(7, i).Value.ToString
            Me.TextBox20.Text = DataGridView1.Item(8, i).Value.ToString
            Me.TextBox19.Text = DataGridView1.Item(9, i).Value.ToString
            Me.TextBox14.Text = DataGridView1.Item(10, i).Value.ToString
            Me.DateTimePicker2.Value = DataGridView1.Item(11, i).Value.ToString
            Me.TextBox16.Text = DataGridView1.Item(12, i).Value.ToString
            Me.DateTimePicker3.Value = DataGridView1.Item(13, i).Value.ToString
            Me.TextBox17.Text = DataGridView1.Item(14, i).Value.ToString
            Me.DateTimePicker4.Value = DataGridView1.Item(15, i).Value.ToString

            Dim sqlsearch As String
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            'Change the following to your access database location
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            sqlsearch = "SELECT JobCardNo, InvDate, BookedBy, ClientName, GrossAmt, ScmAmt, NetAmt, AdvAmt, DueAmt FROM BalanceSaleClients WHERE BalanceSaleClients.Cust_ID LIKE'%" &
            TextBox1.Text & "%'"
            Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
            Dim dt As New DataTable("InvoiceProduct")
            adapter.Fill(dt)
            Me.DataGridView2.DataSource = dt
            DataGridView2.Columns(1).DefaultCellStyle.Format = "dd/MM/yyy"
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
            Exit Sub
        End Try

        Button4.Enabled = True
        GridCopy()
    End Sub

    Sub Compute2()
        Dim num1 As Double = 0
        num1 = CDbl(Val(TextBox14.Text) + Val(TextBox16.Text) + Val(TextBox17.Text))
        num1 = Math.Round(num1, 2)
        TextBox22.Text = num1
    End Sub

    Sub Compute3()
        Dim num1 As Double = 0
        num1 = CDbl(Val(TextBox22.Text) - Val(TextBox21.Text))
        num1 = Math.Round(num1, 2)
        TextBox23.Text = num1
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        Compute2()
        Compute3()
    End Sub

    Private Sub TextBox16_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox16.TextChanged
        Compute2()
        Compute3()
    End Sub

    Private Sub TextBox17_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox17.TextChanged
        Compute2()
        Compute3()
    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged
        TextBox21.Text = TextBox19.Text
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            Dim okToDelete As MsgBoxResult = MsgBox("Are you sure you want to delete the current record?", MsgBoxStyle.YesNo)
            If okToDelete = MsgBoxResult.Yes Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str6 As String
                str6 = "Delete from BalanceSaleClients Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd7 As OleDbCommand = New OleDbCommand(str6, myConnection)
                Try
                    cmd7.ExecuteNonQuery()
                    cmd7.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
                connString = provider
                myConnection.ConnectionString = connString
                myConnection.Open()
                Dim str5 As String
                str5 = "Delete from BalanceSale Where Cust_ID = '" & Me.TextBox1.Text & "'"
                Dim cmd6 As OleDbCommand = New OleDbCommand(str5, myConnection)
                Try
                    cmd6.ExecuteNonQuery()
                    cmd6.Dispose()
                    myConnection.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            ElseIf okToDelete = MsgBoxResult.No Then
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Row is Empty")
            Exit Sub
        End Try
        Button4.Enabled = False
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            If TextBox3.Text = "" Then
                MsgBox("Please enter no to send SMS")
            Else
                Dim url As String
                url = "http://alerts.valueleaf.com/api/v4/?api_key=A7ce7d9a7a5bcb5f1cfdc9e60b9095d8c&method=sms&message=" + Me.TextBox29.Text + "&to=" + Me.TextBox3.Text + "&sender=AOPTIC"
                Dim myReq As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
                Dim myResp As HttpWebResponse = DirectCast(myReq.GetResponse(), HttpWebResponse)
                Dim respStreamReader As New System.IO.StreamReader(myResp.GetResponseStream())
                Dim responseString As String = respStreamReader.ReadToEnd()
                respStreamReader.Close()
                myResp.Close()
                MsgBox("Message Send Successfully")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class