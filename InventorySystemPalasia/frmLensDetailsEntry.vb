Imports System.Data.OleDb

Public Class frmLensDetailsEntry

    Private Sub frmLensDetailsEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub frmLensDetailsEntry_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
            MainMenu.Panel1.Visible = True
        End If
    End Sub

    Private Sub frmLensDetailsEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshData()
        auto()
        DateTimePicker1.Value = Date.Now

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        ComboBox1.DropDownStyle = ComboBoxStyle.DropDown

        'Assume the ArrayList object is source collection  
        aList.Add("DIGITAL EYE TESTING")
        aList.Add("DOCTOR")
        aList.Add("OLD POWER")

        'ComboBox AutoComplete feature settings  
        For Each item As String In aList
            ComboBox1.AutoCompleteCustomSource.Add(item)
            ComboBox1.Items.Add(item)
        Next

        ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Sub Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        TextBox18.Clear()
        ComboBox1.SelectedIndex = -1
        ComboBox1.Text = ""
        DateTimePicker1.Value = Date.Now
        TextBox20.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox30.Clear()
        TextBox2.Focus()
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
            cmd = New OleDbCommand("SELECT TOP 1 ID FROM SaleInvoice ORDER BY ID DESC", con)
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
            TextBox1.Text = "SR-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If TextBox2.Text = "" Then
                MessageBox.Show("Please Enter Customer Name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox2.Focus()
                Exit Sub
            End If
            If TextBox3.Text = "" Then
                MessageBox.Show("Please Enter Mobile", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox3.Focus()
                Exit Sub
            End If

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
            dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;" ' Change it to your Access Database location
            connString = provider & dataFile
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim ct As String = "select Cust_ID from SaleInvoice where Cust_ID=@d1"
            Dim cmd9 As OleDbCommand = New OleDbCommand(ct)
            cmd9.Parameters.AddWithValue("@d1", TextBox1.Text)
            cmd9.Connection = myConnection
            rdr = cmd9.ExecuteReader()
            If rdr.Read() Then
                MessageBox.Show("ID already Saved", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                auto()
            End If
            myConnection.Close()

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim str As String
            str = "insert into SaleInvoice ([Cust_ID], [Cust_Name], [Mobile], [Status], [Age], [Phone], [Address], [RSPH], [RCYL], [RAXIS], [RVN], [RADD], [LSPH], [LCYL], [LAXIS], [LVN], [LADD], [PD], [REFBY], [LensType], [LensType1], [LensType2], [LensType3], [BookingDate], [Remarks], [Right], [Left], [RLAdd], [PRGRight], [PRGLeft]) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            cmd.Parameters.Add(New OleDbParameter("Cust_ID", CType(TextBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Cust_Name", CType(TextBox2.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Mobile", CType(TextBox3.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Status", CType(TextBox28.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Age", CType(TextBox24.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Phone", CType(TextBox4.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Address", CType(TextBox5.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RSPH", CType(TextBox6.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RCYL", CType(TextBox7.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RAXIS", CType(TextBox8.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RVN", CType(TextBox9.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RADD", CType(TextBox10.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LSPH", CType(TextBox11.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LCYL", CType(TextBox12.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LAXIS", CType(TextBox13.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LVN", CType(TextBox14.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LADD", CType(TextBox15.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PD", CType(TextBox16.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("REFBY", CType(ComboBox1.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType", CType(TextBox17.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType1", CType(TextBox25.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType2", CType(TextBox26.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LensType3", CType(TextBox27.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("BookingDate", CType(DateTimePicker1.Value.Date, String)))
            cmd.Parameters.Add(New OleDbParameter("Remarks", CType(TextBox18.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Right", CType(TextBox20.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("Left", CType(TextBox22.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("RLAdd", CType(TextBox23.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PRGRight", CType(TextBox29.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("PRGLeft", CType(TextBox30.Text, String)))
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            MessageBox.Show("Power Details Added", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim okToPrint As MsgBoxResult = MsgBox("Print Power Details?", MsgBoxStyle.YesNo)
        If okToPrint = MsgBoxResult.Yes Then
            Print()
        ElseIf okToPrint = MsgBoxResult.No Then
        End If
        RefreshData()
        Clear()
        auto()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub RefreshData()
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        'Change the following to your access database location
        connString = provider
        myConnection.ConnectionString = connString
        DataGridView1.DataSource = Nothing
        DataGridView1.Refresh()
        Dim str As String = "select Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice"
        Using cmd As New OleDb.OleDbCommand(str, myConnection)
            Using da As New OleDbDataAdapter(cmd)
                Using newtable As New DataTable
                    da.Fill(newtable)
                    DataGridView1.DataSource = newtable
                End Using
            End Using
        End Using

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False
        Me.DataGridView1.Columns(27).Visible = False
        Me.DataGridView1.Columns(28).Visible = False
        Me.DataGridView1.Columns(29).Visible = False

        Me.DataGridView1.Columns(0).Width = 70
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 80
        Me.DataGridView1.Columns(3).Width = 80
        Me.DataGridView1.Columns(4).Width = 60
        Me.DataGridView1.Columns(5).Width = 100
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Clear()
        auto()
    End Sub

    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDoubleClick
        Try
            Dim i As Integer
            i = DataGridView1.CurrentRow.Index
            Me.TextBox1.Text = DataGridView1.Item(0, i).Value.ToString
            Me.TextBox2.Text = DataGridView1.Item(1, i).Value.ToString
            Me.TextBox3.Text = DataGridView1.Item(2, i).Value.ToString
            Me.TextBox28.Text = DataGridView1.Item(3, i).Value.ToString
            Me.TextBox24.Text = DataGridView1.Item(4, i).Value.ToString
            Me.TextBox4.Text = DataGridView1.Item(5, i).Value.ToString
            Me.TextBox5.Text = DataGridView1.Item(6, i).Value.ToString
            Me.DateTimePicker1.Value = DataGridView1.Item(7, i).Value.ToString
            Me.TextBox6.Text = DataGridView1.Item(8, i).Value.ToString
            Me.TextBox7.Text = DataGridView1.Item(9, i).Value.ToString
            Me.TextBox8.Text = DataGridView1.Item(10, i).Value.ToString
            Me.TextBox9.Text = DataGridView1.Item(11, i).Value.ToString
            Me.TextBox10.Text = DataGridView1.Item(12, i).Value.ToString
            Me.TextBox11.Text = DataGridView1.Item(13, i).Value.ToString
            Me.TextBox12.Text = DataGridView1.Item(14, i).Value.ToString
            Me.TextBox13.Text = DataGridView1.Item(15, i).Value.ToString
            Me.TextBox14.Text = DataGridView1.Item(16, i).Value.ToString
            Me.TextBox15.Text = DataGridView1.Item(17, i).Value.ToString
            Me.TextBox16.Text = DataGridView1.Item(18, i).Value.ToString
            Me.ComboBox1.Text = DataGridView1.Item(19, i).Value.ToString
            Me.TextBox17.Text = DataGridView1.Item(20, i).Value.ToString
            Me.TextBox25.Text = DataGridView1.Item(21, i).Value.ToString
            Me.TextBox26.Text = DataGridView1.Item(22, i).Value.ToString
            Me.TextBox27.Text = DataGridView1.Item(23, i).Value.ToString
            Me.TextBox18.Text = DataGridView1.Item(24, i).Value.ToString
            Me.TextBox20.Text = DataGridView1.Item(25, i).Value.ToString
            Me.TextBox22.Text = DataGridView1.Item(26, i).Value.ToString
            Me.TextBox23.Text = DataGridView1.Item(27, i).Value.ToString
            Me.TextBox29.Text = DataGridView1.Item(28, i).Value.ToString
            Me.TextBox30.Text = DataGridView1.Item(29, i).Value.ToString
        Catch ex As Exception
            MsgBox("Row is Empty")
        End Try
    End Sub

    Private Sub TextBox19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox19.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice WHERE Cust_Name LIKE'%" &
        TextBox19.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False
        Me.DataGridView1.Columns(27).Visible = False
        Me.DataGridView1.Columns(28).Visible = False
        Me.DataGridView1.Columns(29).Visible = False

        Me.DataGridView1.Columns(0).Width = 70
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 80
        Me.DataGridView1.Columns(3).Width = 80
        Me.DataGridView1.Columns(4).Width = 60
        Me.DataGridView1.Columns(5).Width = 100
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox24.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            MessageBox.Show("Enter No. only", "Customer Master", MessageBoxButtons.OK, MessageBoxIcon.Information)
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox21_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox21.TextChanged
        Dim sqlsearch As String
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        sqlsearch = "SELECT Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice WHERE Mobile LIKE'%" &
        TextBox21.Text & "%'"
        Dim adapter As New OleDbDataAdapter(sqlsearch, myConnection)
        Dim dt As New DataTable("ItemMaster")
        adapter.Fill(dt)
        DataGridView1.DataSource = dt
        myConnection.Close()

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False
        Me.DataGridView1.Columns(27).Visible = False
        Me.DataGridView1.Columns(28).Visible = False
        Me.DataGridView1.Columns(29).Visible = False

        Me.DataGridView1.Columns(0).Width = 70
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 80
        Me.DataGridView1.Columns(3).Width = 80
        Me.DataGridView1.Columns(4).Width = 60
        Me.DataGridView1.Columns(5).Width = 100
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

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
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
            TextBox10.Focus()
        End If
    End Sub

    Private Sub TextBox10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox11.Focus()
        End If
    End Sub

    Private Sub TextBox11_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox12.Focus()
        End If
    End Sub

    Private Sub TextBox12_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox13.Focus()
        End If
    End Sub

    Private Sub TextBox13_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox13.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox14.Focus()
        End If
    End Sub

    Private Sub TextBox14_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox15.Focus()
        End If
    End Sub

    Private Sub TextBox15_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox15.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox16.Focus()
        End If
    End Sub

    Private Sub TextBox16_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox16.KeyDown
        If e.KeyCode = Keys.Enter Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub TextBox17_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox17.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox25.Focus()
        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox18.Focus()
        End If
    End Sub

    Private Sub TextBox18_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox18.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox29.Focus()
        End If
    End Sub

    Dim aList As ArrayList = New ArrayList

    Private Sub ComboBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If aList.Contains(ComboBox1.Text) = False Then
                ComboBox1.AutoCompleteCustomSource.Add(ComboBox1.Text)
                aList.Add(ComboBox1.Text)
                ComboBox1.Items.Add(ComboBox1.Text)
            End If
        End If
        If e.KeyCode = Keys.Enter Then
            TextBox17.Focus()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        MainMenu.Panel1.Visible = True
    End Sub

    Private Sub TextBox20_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox20.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox22.Focus()
        End If
    End Sub

    Private Sub TextBox22_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox22.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox23.Focus()
        End If
    End Sub

    Private Sub TextBox23_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox23.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.Focus()
        End If
    End Sub

    Private Sub TextBox24_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox24.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Print()
    End Sub

    Sub Print()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New PowerDetails  'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SaleInvoice.Cust_ID, SaleInvoice.Cust_Name, SaleInvoice.Mobile, SaleInvoice.Age, SaleInvoice.BookingDate, SaleInvoice.RSPH, SaleInvoice.RCYL, SaleInvoice.RAXIS, SaleInvoice.RVN, SaleInvoice.RADD, SaleInvoice.LSPH, SaleInvoice.LCYL, SaleInvoice.LAXIS, SaleInvoice.LVN, SaleInvoice.LADD, SaleInvoice.PD, SaleInvoice.Remarks FROM SaleInvoice Where SaleInvoice.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox1.Text)
            MyCommand1.CommandText = "SELECT * from SaleInvoice"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SaleInvoice")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox1.Text)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        'Change the following to your access database location
        dataFile = "|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim table As New DataTable
        Dim Command As New OleDbCommand("select Cust_ID, Cust_Name, Mobile, Status, Age, Phone, Address, BookingDate, RSPH, RCYL, RAXIS, RVN, RADD, LSPH, LCYL, LAXIS, LVN, LADD, PD, REFBY, LensType, LensType1, LensType2, LensType3, Remarks, Right, Left, RLAdd, PRGRight, PRGLeft from SaleInvoice Where BookingDate =@d1", myConnection)
        Command.Parameters.Add("@d1", OleDbType.Date, 30, "BookingDate").Value = DateTimePicker2.Value.Date
        Dim adapter As New OleDbDataAdapter(Command)
        adapter.Fill(table)
        DataGridView1.DataSource = table.DefaultView
        myConnection.Close()

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)

        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False
        Me.DataGridView1.Columns(20).Visible = False
        Me.DataGridView1.Columns(21).Visible = False
        Me.DataGridView1.Columns(22).Visible = False
        Me.DataGridView1.Columns(23).Visible = False
        Me.DataGridView1.Columns(24).Visible = False
        Me.DataGridView1.Columns(25).Visible = False
        Me.DataGridView1.Columns(26).Visible = False
        Me.DataGridView1.Columns(27).Visible = False
        Me.DataGridView1.Columns(28).Visible = False
        Me.DataGridView1.Columns(29).Visible = False

        Me.DataGridView1.Columns(0).Width = 70
        Me.DataGridView1.Columns(1).Width = 200
        Me.DataGridView1.Columns(2).Width = 80
        Me.DataGridView1.Columns(3).Width = 80
        Me.DataGridView1.Columns(4).Width = 60
        Me.DataGridView1.Columns(5).Width = 100
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox19.Clear()
        TextBox21.Clear()
        RefreshData()
    End Sub

    Private Sub TextBox25_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox25.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox26.Focus()
        End If
    End Sub

    Private Sub TextBox26_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox26.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox27.Focus()
        End If
    End Sub

    Private Sub TextBox27_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox27.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox18.Focus()
        End If
    End Sub

    Private Sub TextBox29_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox29.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox30.Focus()
        End If
    End Sub

    Private Sub TextBox30_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox30.KeyDown
        If e.KeyCode = Keys.Enter Then
            TextBox20.Focus()
        End If
    End Sub
End Class