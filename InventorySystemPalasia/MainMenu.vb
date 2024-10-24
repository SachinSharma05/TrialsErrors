Imports System.Windows.Forms
Imports System.IO
Imports System.Data.OleDb

Public Class MainMenu

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Dim ChildForm As New frmProductType
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim ChildForm As New frmCategoryMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim ChildForm As New frmSubCategoryMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim ChildForm As New frmItemMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Dim ChildForm As New frmContactCreation
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub SalesmanMasterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesmanMasterToolStripMenuItem.Click
        Dim ChildForm As New frmCustomerType
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub CustomerLensEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerLensEntryToolStripMenuItem.Click
        Dim ChildForm As New frmLensDetailsEntry
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub LensDetailsListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LensDetailsListToolStripMenuItem.Click
        Dim ChildForm As New frmLensDetailsList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub ContactLensesSaleDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContactLensesSaleDetailsToolStripMenuItem.Click
        Dim ChildForm As New frmContactLensesEntrySale
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub ContactLensesDetailsListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContactLensesDetailsListToolStripMenuItem.Click
        Dim ChildForm As New frmContactLensListJobCardInvoice
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub SaleInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaleInvoiceToolStripMenuItem.Click
        Dim ChildForm As New frmSaleInvoice
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub SaleInvoiceListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaleInvoiceListToolStripMenuItem.Click
        Dim ChildForm As New frmSpectacleSaleList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub ExchangeOldOrderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExchangeOldOrderToolStripMenuItem.Click
        Dim ChildForm As New frmSpectacleSaleList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub MainMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
    End Sub

    Private Sub DamageProductEntryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DamageProductEntryToolStripMenuItem.Click
        Dim ChildForm As New frmDamageProductEntry
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PayRemainingSpectaclesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PayRemainingSpectaclesToolStripMenuItem.Click
        Dim ChildForm As New frmPayRemainingMain
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PayRemainingContactLensesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PayRemainingContactLensesToolStripMenuItem.Click
        Dim ChildForm As New frmPayRemainingMain
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub EnquiryByCustomerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnquiryByCustomerToolStripMenuItem.Click
        Dim ChildForm As New frmStockIn
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub DamageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DamageToolStripMenuItem.Click
        Dim ChildForm As New frmStockInList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub DamageItemRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DamageItemRecordToolStripMenuItem.Click
        Dim ChildForm As New frmStockOut
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub StockOutListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StockOutListToolStripMenuItem.Click
        Dim ChildForm As New frmStockOutList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PresentStockListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PresentStockListToolStripMenuItem.Click
        Dim ChildForm As New frmPresentStock
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub RequiredStockListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RequiredStockListToolStripMenuItem.Click
        Dim ChildForm As New frmRequiredStock
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub MaxSellingProductsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxSellingProductsToolStripMenuItem.Click
        Dim ChildForm As New frmMaxSellingProducts
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub BackupDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackupDatabaseToolStripMenuItem.Click
        Backup()
    End Sub

    Sub Backup()
        Try
            If File.Exists(Application.StartupPath & "\Backup\Billing.accdb") Then
                File.Delete(Application.StartupPath & "\Backup\Billing.accdb")
                File.Copy("Billing.accdb", Application.StartupPath & "\Backup\Billing.accdb", True)
                MsgBox("Backup Completed Successfully..!!")
            Else
                File.Copy("Billing.accdb", Application.StartupPath & "\Backup\Billing.accdb")
                MsgBox("Backup Completed Successfully..!!")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Me.Close()
        End Try
    End Sub

    Private Sub ChangeUserPasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeUserPasswordToolStripMenuItem.Click
        Dim ChildForm As New frmChangePassword
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub LogoutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem1.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    Dim st As String = "Successfully logged out"
                    Me.Hide()
                    Login.TextBox1.Clear()
                    Login.TextBox2.Clear()
                    Login.TextBox1.Focus()
                    Login.Show()
                Else
                    Dim st As String = "Successfully logged out"
                    Me.Hide()
                    Login.TextBox1.Clear()
                    Login.TextBox2.Clear()
                    Login.TextBox1.Focus()
                    Login.Show()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub WorkshopManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkshopManagerToolStripMenuItem.Click
        Dim ChildForm As New frmWorkshopManager
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub CLJobCardInvoiceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CLJobCardInvoiceToolStripMenuItem.Click
        Dim ChildForm As New frmContactLensListJobCardInvoice
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub ExchangeSaleReturnToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExchangeSaleReturnToolStripMenuItem.Click
        Dim ChildForm As New frmExchangeProducts
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub AccountsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountsToolStripMenuItem.Click
        Dim ChildForm As New frmDayCashBook
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    Dim st As String = "Successfully logged out"
                    Me.Hide()
                    Login.TextBox1.Clear()
                    Login.TextBox2.Clear()
                    Login.TextBox1.Focus()
                    Login.Show()
                Else
                    Dim st As String = "Successfully logged out"
                    Me.Hide()
                    Login.TextBox1.Clear()
                    Login.TextBox2.Clear()
                    Login.TextBox1.Focus()
                    Login.Show()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Dim ChildForm As New frmCategoryMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Dim ChildForm As New frmCategoryMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Dim ChildForm As New frmSubCategoryMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        Dim ChildForm As New frmItemMaster
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        Dim ChildForm As New frmContactCreation
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        Dim ChildForm As New frmLensDetailsEntry
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        Dim ChildForm As New frmContactLensesEntrySale
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New frmSaleInvoice
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        Dim ChildForm As New frmDayCashBook
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        Dim ChildForm As New frmWorkshopManager
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        Dim ChildForm As New frmDamageProductEntry
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        Dim ChildForm As New frmExchangeProducts
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        Dim ChildForm As New frmStockIn
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        Dim ChildForm As New frmStockOut
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox27.Click
        Dim ChildForm As New frmChangePassword
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        Dim ChildForm As New frmLensDetailsList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        Dim ChildForm As New frmContactLensListJobCardInvoice
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        Dim ChildForm As New frmSpectacleSaleList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox17.Click
        Dim ChildForm As New frmPayRemainingMain
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox21.Click
        Dim ChildForm As New frmStockInList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox22.Click
        Dim ChildForm As New frmStockOutList
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        Dim ChildForm As New frmPresentStock
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        Dim ChildForm As New frmRequiredStock
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox25.Click
        Dim ChildForm As New frmMaxSellingProducts
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox26.Click
        Backup()
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave
        Me.PictureBox1.BackColor = Color.White
        Me.Label1.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Me.PictureBox1.BackColor = Color.BurlyWood
        Me.Label1.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox2.MouseLeave
        Me.PictureBox2.BackColor = Color.White
        Me.Label2.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseMove
        Me.PictureBox2.BackColor = Color.BurlyWood
        Me.Label2.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox3.MouseLeave
        Me.PictureBox3.BackColor = Color.White
        Me.Label3.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox3.MouseMove
        Me.PictureBox3.BackColor = Color.BurlyWood
        Me.Label3.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox4.MouseLeave
        Me.PictureBox4.BackColor = Color.White
        Me.Label10.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox4.MouseMove
        Me.PictureBox4.BackColor = Color.BurlyWood
        Me.Label10.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox5.MouseLeave
        Me.PictureBox5.BackColor = Color.White
        Me.Label5.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox5_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox5.MouseMove
        Me.PictureBox5.BackColor = Color.BurlyWood
        Me.Label5.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox6.MouseLeave
        Me.PictureBox6.BackColor = Color.White
        Me.Label6.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox6_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox6.MouseMove
        Me.PictureBox6.BackColor = Color.BurlyWood
        Me.Label6.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox7.MouseLeave
        Me.PictureBox7.BackColor = Color.White
        Me.Label7.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox7_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox7.MouseMove
        Me.PictureBox7.BackColor = Color.BurlyWood
        Me.Label7.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox8.MouseLeave
        Me.PictureBox8.BackColor = Color.White
        Me.Label8.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox8_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox8.MouseMove
        Me.PictureBox8.BackColor = Color.BurlyWood
        Me.Label8.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox9_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox9.MouseLeave
        Me.PictureBox9.BackColor = Color.White
        Me.Label9.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox9_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox9.MouseMove
        Me.PictureBox9.BackColor = Color.BurlyWood
        Me.Label9.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox15_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox15.MouseLeave
        Me.PictureBox15.BackColor = Color.White
        Me.Label15.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox15_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox15.MouseMove
        Me.PictureBox15.BackColor = Color.BurlyWood
        Me.Label15.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox16_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox16.MouseLeave
        Me.PictureBox16.BackColor = Color.White
        Me.Label16.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox16_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox16.MouseMove
        Me.PictureBox16.BackColor = Color.BurlyWood
        Me.Label16.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox19_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox19.MouseLeave
        Me.PictureBox19.BackColor = Color.White
        Me.Label19.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox19_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox19.MouseMove
        Me.PictureBox19.BackColor = Color.BurlyWood
        Me.Label19.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox20_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox20.MouseLeave
        Me.PictureBox20.BackColor = Color.White
        Me.Label20.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox20_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox20.MouseMove
        Me.PictureBox20.BackColor = Color.BurlyWood
        Me.Label20.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox27_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox27.MouseLeave
        Me.PictureBox27.BackColor = Color.White
        Me.Label27.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox27_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox27.MouseMove
        Me.PictureBox27.BackColor = Color.BurlyWood
        Me.Label27.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox10_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox10.MouseLeave
        Me.PictureBox10.BackColor = Color.White
        Me.Label4.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox10_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox10.MouseMove
        Me.PictureBox10.BackColor = Color.BurlyWood
        Me.Label4.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox11_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox11.MouseLeave
        Me.PictureBox11.BackColor = Color.White
        Me.Label11.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox11_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox11.MouseMove
        Me.PictureBox11.BackColor = Color.BurlyWood
        Me.Label11.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox12_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox12.MouseLeave
        Me.PictureBox12.BackColor = Color.White
        Me.Label12.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox12_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox12.MouseMove
        Me.PictureBox12.BackColor = Color.BurlyWood
        Me.Label12.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox17_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox17.MouseLeave
        Me.PictureBox17.BackColor = Color.White
        Me.Label17.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox17_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox17.MouseMove
        Me.PictureBox17.BackColor = Color.BurlyWood
        Me.Label17.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox21_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox21.MouseLeave
        Me.PictureBox21.BackColor = Color.White
        Me.Label21.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox21_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox21.MouseMove
        Me.PictureBox21.BackColor = Color.BurlyWood
        Me.Label21.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox22_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox22.MouseLeave
        Me.PictureBox22.BackColor = Color.White
        Me.Label22.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox22_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox22.MouseMove
        Me.PictureBox22.BackColor = Color.BurlyWood
        Me.Label22.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox23_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox23.MouseLeave
        Me.PictureBox23.BackColor = Color.White
        Me.Label23.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox23_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox23.MouseMove
        Me.PictureBox23.BackColor = Color.BurlyWood
        Me.Label23.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox24_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox24.MouseLeave
        Me.PictureBox24.BackColor = Color.White
        Me.Label24.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox24_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox24.MouseMove
        Me.PictureBox24.BackColor = Color.BurlyWood
        Me.Label24.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox25_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox25.MouseLeave
        Me.PictureBox25.BackColor = Color.White
        Me.Label25.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox25_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox25.MouseMove
        Me.PictureBox25.BackColor = Color.BurlyWood
        Me.Label25.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox26_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox26.MouseLeave
        Me.PictureBox26.BackColor = Color.White
        Me.Label26.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox26_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox26.MouseMove
        Me.PictureBox26.BackColor = Color.BurlyWood
        Me.Label26.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox28_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox28.MouseLeave
        Me.PictureBox28.BackColor = Color.White
        Me.Label28.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox28_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox28.MouseMove
        Me.PictureBox28.BackColor = Color.BurlyWood
        Me.Label28.ForeColor = Color.Red
    End Sub

    Private Sub MainMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PreVentFlicker()
    End Sub

    Private Sub PreVentFlicker()
        With Me
            .SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            .SetStyle(ControlStyles.UserPaint, True)
            .SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            .UpdateStyles()
        End With
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or 33554432
            Return cp
        End Get
    End Property

    Private Sub PictureBox30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox30.Click
        Dim ChildForm As New frmUpdateJobCard
        ChildForm.MdiParent = Me
        ChildForm.Show()
        Me.Panel1.Visible = False
    End Sub

    Private Sub PictureBox30_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox30.MouseLeave
        Me.PictureBox30.BackColor = Color.White
        Me.Label30.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox30_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox30.MouseMove
        Me.PictureBox30.BackColor = Color.BurlyWood
        Me.Label30.ForeColor = Color.Red
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Panel1.Visible = False
        Dim ChildForm As New frmSaleInvoice
        ChildForm.MdiParent = Me
        ChildForm.Button3.Enabled = False
        ChildForm.Show()
    End Sub

    Private Sub PictureBox31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox31.Click
        Try
            Dim portfolioPath As String = My.Application.Info.DirectoryPath
            If MessageBox.Show("Restoring the database will erase any changes you have made since you last backup. Are you sure you want to do this?", _
                        "Confirm Delete", _
                        MessageBoxButtons.OKCancel, _
                        MessageBoxIcon.Question, _
                        MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.OK Then

                'Restore the database from a backup copy.
                FileCopy(Application.StartupPath & "\Backup\Billing.accdb", portfolioPath & "\Billing.accdb")
                MsgBox("Database Restoration Successful")
            End If
        Catch ex As Exception
            Dim MessageString As String = "Report this error to the system administrator: " & ControlChars.NewLine & ex.Message
            Dim TitleString As String = "Customer Details Data Load Failed"
            MessageBox.Show(MessageString, TitleString, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PictureBox31_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox31.MouseLeave
        Me.PictureBox31.BackColor = Color.White
        Me.Label31.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox31_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox31.MouseMove
        Me.PictureBox31.BackColor = Color.BurlyWood
        Me.Label31.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox29.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New SendSMS
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox29_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox29.MouseLeave
        Me.PictureBox29.BackColor = Color.White
        Me.Label29.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox29_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox29.MouseMove
        Me.PictureBox29.BackColor = Color.BurlyWood
        Me.Label29.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox33.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New frmRaisedInvoiceList
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox33_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox33.MouseLeave
        Me.PictureBox33.BackColor = Color.White
        Me.Label33.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox33_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox33.MouseMove
        Me.PictureBox33.BackColor = Color.BurlyWood
        Me.Label33.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New SunglassSale
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox13_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox13.MouseLeave
        Me.PictureBox13.BackColor = Color.White
        Me.Label13.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox13_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox13.MouseMove
        Me.PictureBox13.BackColor = Color.BurlyWood
        Me.Label13.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New frmSalespersonAccounts
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox18_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox18.MouseLeave
        Me.PictureBox18.BackColor = Color.White
        Me.Label18.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox18_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox18.MouseMove
        Me.PictureBox18.BackColor = Color.BurlyWood
        Me.Label18.ForeColor = Color.Red
    End Sub

    Private Sub PictureBox34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox34.Click
        Me.Panel1.Visible = False
        Dim ChildForm As New frmClientsBalance
        ChildForm.MdiParent = Me
        ChildForm.Show()
    End Sub

    Private Sub PictureBox34_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox34.MouseLeave
        Me.PictureBox34.BackColor = Color.White
        Me.Label34.ForeColor = Color.Black
    End Sub

    Private Sub PictureBox34_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox34.MouseMove
        Me.PictureBox34.BackColor = Color.BurlyWood
        Me.Label34.ForeColor = Color.Red
    End Sub
End Class
