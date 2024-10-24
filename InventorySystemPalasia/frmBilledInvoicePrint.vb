Imports System.Data.OleDb

Public Class frmBilledInvoicePrint

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Print1()
    End Sub

    Sub Print1()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New InvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SaleInvoice.ReceiptNo, SaleInvoice.Cust_Name, SaleInvoice.Mobile, SaleInvoice.GrossAmt, SaleInvoice.NetAmt, SaleInvoice.PaidAmt, SaleInvoice.DueAmt, SaleInvoice.Paymode, SaleInvoice.BookingDate, SaleInvoice.CGST, SaleInvoice.SGST, SaleInvoice.ScmAmt, InvoiceProduct.ProdName, InvoiceProduct.Qty, InvoiceProduct.Price, InvoiceProduct.Discount, InvoiceProduct.Total FROM InvoiceProduct INNER JOIN SaleInvoice ON SaleInvoice.ReceiptNo=InvoiceProduct.Cust_ID Where SaleInvoice.ReceiptNo=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox5.Text)
            MyCommand1.CommandText = "SELECT * from SaleInvoice"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SaleInvoice")
            myDA.Fill(myDS, "InvoiceProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox5.Text)
            rpt.SetParameterValue("p3", TextBox1.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value.Date)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Print()
    End Sub

    Sub Print()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New SunglassBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select SunglassSale.ReceiptNo, SunglassSale.Cust_Name, SunglassSale.Mobile, SunglassSale.GrossAmt, SunglassSale.NetAmt, SunglassSale.PaidAmt, SunglassSale.DueAmt, SunglassSale.Paymode, SunglassSale.BookingDate, SunglassSale.CGST, SunglassSale.SGST, SunglassSale.ScmAmt, SunglassProduct.ProdName, SunglassProduct.Qty, SunglassProduct.Price, SunglassProduct.Discount, SunglassProduct.Total FROM SunglassProduct INNER JOIN SunglassSale ON SunglassSale.ReceiptNo=SunglassProduct.Cust_ID Where SunglassSale.ReceiptNo=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox5.Text)
            MyCommand1.CommandText = "SELECT * from SunglassSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "SunglassSale")
            myDA.Fill(myDS, "SunglassProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox5.Text)
            rpt.SetParameterValue("p3", TextBox1.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Print2()
    End Sub

    Sub Print2()
        Try
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Billing.accdb;Jet OLEDB:Database Password=brutusbozo;"
            'Change the following to your access database location
            connString = provider
            myConnection.ConnectionString = connString
            myConnection.Open()
            Dim rpt As New CLInvoiceBill 'The report you created.
            Dim MyCommand, MyCommand1 As New OleDbCommand
            Dim myDA, myDA1 As New OleDbDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "Select CLSale.Cust_ID, CLSale.Cust_Name, CLSale.Mobile, CLSale.BookingDate, CLSale.GrossAmt, CLSale.NetAmt, CLSale.PaidAmt, CLSale.DueAmt, CLSale.Paymode, CLSale.ScmAmt, CLSaleProduct.ProdName, CLSaleProduct.Qty, CLSaleProduct.Price, CLSaleProduct.Discount, CLSaleProduct.Total FROM CLSaleProduct INNER JOIN CLSale ON CLSale.Cust_ID=CLSaleProduct.Cust_ID Where CLSale.Cust_ID=@d1"
            MyCommand.Parameters.AddWithValue("@d1", TextBox5.Text)
            MyCommand1.CommandText = "SELECT * from CLSale"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "CLSale")
            myDA.Fill(myDS, "CLSaleProduct")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", TextBox2.Text)
            rpt.SetParameterValue("p2", TextBox5.Text)
            rpt.SetParameterValue("p3", TextBox1.Text)
            rpt.SetParameterValue("p4", DateTimePicker1.Value)
            frmShowReports.CrystalReportViewer1.ReportSource = rpt
            frmShowReports.ShowDialog()
            myConnection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class