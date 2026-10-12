Imports MySql.Data.MySqlClient

Public Class frmReceipt

    Public SaleID As Integer = 0
    Private dgvItems As New DataGridView()
    Private lblReceipt As New Label()
    Private lblDate As New Label()
    Private lblEmployee As New Label()
    Private lblCustomer As New Label()
    Private lblSubtotal As New Label()
    Private lblTotal As New Label()
    Private lblPaid As New Label()
    Private lblChange As New Label()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Receipt"
        Me.Size = New Size(420, 680)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        BuildUI()
        LoadReceipt()
    End Sub

    Private Sub BuildUI()
        Dim y As Integer = 15

        ' Store header
        Dim lblStore As New Label() With {
            .Text = "GHC PET MART",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(110, y)
        }
        Me.Controls.Add(lblStore) : y += 30

        Dim lblTag As New Label() With {
            .Text = "One Home for Fur, Feathers, and Fins.",
            .Font = New Font("Arial", 8, FontStyle.Italic),
            .ForeColor = Color.Gray,
            .AutoSize = True,
            .Location = New Point(65, y)
        }
        Me.Controls.Add(lblTag) : y += 30

        ' Divider
        AddDivider(y) : y += 15

        ' Receipt info
        AddInfoRow("Receipt No:", lblReceipt, y) : y += 25
        AddInfoRow("Date:", lblDate, y) : y += 25
        AddInfoRow("Employee:", lblEmployee, y) : y += 25
        AddInfoRow("Customer:", lblCustomer, y) : y += 25

        AddDivider(y) : y += 15

        ' Items grid
        Dim lblItemsHead As New Label() With {
            .Text = "ITEMS PURCHASED",
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, y)
        }
        Me.Controls.Add(lblItemsHead) : y += 22

        dgvItems.Location = New Point(15, y)
        dgvItems.Size = New Size(375, 170)
        dgvItems.ReadOnly = True
        dgvItems.AllowUserToAddRows = False
        dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvItems.BackgroundColor = Color.White
        dgvItems.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Controls.Add(dgvItems) : y += 180

        AddDivider(y) : y += 15

        ' Totals
        AddTotalRow("Subtotal:", lblSubtotal, y) : y += 28
        AddTotalRow("TOTAL:", lblTotal, y, True) : y += 35
        AddDivider(y) : y += 15
        AddTotalRow("Amount Paid:", lblPaid, y) : y += 28
        AddTotalRow("Change:", lblChange, y) : y += 35

        AddDivider(y) : y += 15

        Dim lblThank As New Label() With {
            .Text = "Thank you for shopping at GHC Pet Mart!",
            .Font = New Font("Arial", 8, FontStyle.Italic),
            .ForeColor = Color.Gray,
            .AutoSize = True,
            .Location = New Point(55, y)
        }
        Me.Controls.Add(lblThank) : y += 30

        ' Buttons
        Dim btnPrint As New Button() With {
            .Text = "PRINT",
            .Size = New Size(175, 40),
            .Location = New Point(15, y),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        AddHandler btnPrint.Click, AddressOf PrintReceipt
        Me.Controls.Add(btnPrint)

        Dim btnClose As New Button() With {
            .Text = "CLOSE",
            .Size = New Size(175, 40),
            .Location = New Point(210, y),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10)
        }
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)
    End Sub

    Private Sub AddDivider(y As Integer)
        Dim div As New Label() With {
            .BorderStyle = BorderStyle.Fixed3D,
            .Size = New Size(375, 2),
            .Location = New Point(15, y)
        }
        Me.Controls.Add(div)
    End Sub

    Private Sub AddInfoRow(caption As String, valLbl As Label, y As Integer)
        Dim lCap As New Label() With {
            .Text = caption,
            .Font = New Font("Arial", 9),
            .AutoSize = True,
            .Location = New Point(20, y)
        }
        Me.Controls.Add(lCap)
        valLbl.Font = New Font("Arial", 9, FontStyle.Bold)
        valLbl.AutoSize = True
        valLbl.Location = New Point(140, y)
        Me.Controls.Add(valLbl)
    End Sub

    Private Sub AddTotalRow(caption As String, valLbl As Label, y As Integer, Optional isBold As Boolean = False)
        Dim lCap As New Label() With {
            .Text = caption,
            .Font = New Font("Arial", If(isBold, 11, 9), If(isBold, FontStyle.Bold, FontStyle.Regular)),
            .AutoSize = True,
            .Location = New Point(20, y)
        }
        Me.Controls.Add(lCap)
        valLbl.Font = New Font("Arial", If(isBold, 13, 9), FontStyle.Bold)
        valLbl.ForeColor = If(isBold, Color.FromArgb(63, 81, 181), Color.Black)
        valLbl.AutoSize = True
        valLbl.Location = New Point(250, y)
        Me.Controls.Add(valLbl)
    End Sub

    Private Sub LoadReceipt()
        If SaleID = 0 Then Return
        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
            "SELECT 
                s.receipt_no,
                s.sale_date,
                u.full_name AS employee,
                COALESCE(CONCAT(c.first_name,' ',c.last_name),'Walk-in') AS customer,

                SUM(sd.quantity * sd.unit_price - sd.discount) AS subtotal,
                SUM(sd.quantity * sd.unit_price - sd.discount) AS total_amount,

                s.amount_paid,

(
    s.amount_paid -
    SUM(sd.quantity * sd.unit_price - sd.discount)
) AS change_amount

             FROM tbl_sales s
             LEFT JOIN tbl_users u ON s.employee_id = u.user_id
             LEFT JOIN tbl_customers c ON s.customer_id = c.customer_id
             JOIN tbl_sale_details sd ON s.sale_id = sd.sale_id

             WHERE s.sale_id = @id
             GROUP BY
s.sale_id,
s.receipt_no,
s.sale_date,
u.full_name,
c.first_name,
c.last_name,
s.amount_paid", conn)

            cmd.Parameters.AddWithValue("@id", SaleID)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            If dt.Rows.Count = 0 Then Return

            Dim r As DataRow = dt.Rows(0)

            lblReceipt.Text = r("receipt_no").ToString()
            lblDate.Text = Convert.ToDateTime(r("sale_date")).ToString("MM/dd/yyyy hh:mm tt")
            lblEmployee.Text = r("employee").ToString()
            lblCustomer.Text = r("customer").ToString()

            lblSubtotal.Text = "P " & Convert.ToDecimal(r("subtotal")).ToString("N2")
            lblTotal.Text = "P " & Convert.ToDecimal(r("total_amount")).ToString("N2")
            lblPaid.Text = "P " & Convert.ToDecimal(r("amount_paid")).ToString("N2")
            lblChange.Text = "P " & Convert.ToDecimal(r("change_amount")).ToString("N2")


            Dim cmd2 As New MySqlCommand(
            "SELECT 
                p.product_name AS Product,
                sd.quantity AS Qty,
                sd.unit_price AS Price,
                (sd.quantity * sd.unit_price - sd.discount) AS Subtotal
             FROM tbl_sale_details sd
             JOIN tbl_products p ON sd.product_id = p.product_id
             WHERE sd.sale_id = @id", conn)

            cmd2.Parameters.AddWithValue("@id", SaleID)

            Dim dt2 As New DataTable()
            Dim da2 As New MySqlDataAdapter(cmd2)
            da2.Fill(dt2)

            dgvItems.DataSource = dt2

        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub PrintReceipt(sender As Object, e As EventArgs)
        Dim pd As New System.Drawing.Printing.PrintDocument()
        AddHandler pd.PrintPage, AddressOf OnPrintPage
        Dim preview As New System.Windows.Forms.PrintPreviewDialog()
        preview.Document = pd
        preview.ShowDialog()
    End Sub

    Private Sub OnPrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs)
        Dim g As Graphics = e.Graphics
        Dim y As Single = 20
        Dim font As New Font("Courier New", 10)
        Dim fontBold As New Font("Courier New", 10, FontStyle.Bold)
        Dim fontLarge As New Font("Courier New", 12, FontStyle.Bold)
        Dim brush As Brush = Brushes.Black
        Dim lineH As Single = 20

        g.DrawString("GHC PET MART", fontLarge, brush, 80, y) : y += 25
        g.DrawString("One Home for Fur, Feathers, and Fins.", font, brush, 30, y) : y += 25
        g.DrawString(New String("-"c, 42), font, brush, 10, y) : y += 20
        g.DrawString("Receipt: " & lblReceipt.Text, font, brush, 10, y) : y += lineH
        g.DrawString("Date:    " & lblDate.Text, font, brush, 10, y) : y += lineH
        g.DrawString("Cashier: " & lblEmployee.Text, font, brush, 10, y) : y += lineH
        g.DrawString("Cust:    " & lblCustomer.Text, font, brush, 10, y) : y += lineH
        g.DrawString(New String("-"c, 42), font, brush, 10, y) : y += 20

        For Each row As DataGridViewRow In dgvItems.Rows
            If row.Cells("Product").Value Is Nothing Then Continue For

            Dim product As String = row.Cells("Product").Value.ToString()
            Dim qty As String = row.Cells("Qty").Value.ToString()
            Dim subtotal As Decimal = Convert.ToDecimal(row.Cells("Subtotal").Value)

            If product.Length > 20 Then
                product = product.Substring(0, 20)
            End If

            Dim line As String = String.Format("{0,-20} x{1}  P{2}",
        product,
        qty,
        subtotal.ToString("N2"))

            g.DrawString(line, font, brush, 10, y)
            y += lineH
        Next

        g.DrawString(New String("-"c, 42), font, brush, 10, y) : y += 20
        g.DrawString("TOTAL:      P " & lblTotal.Text.Replace("P ", ""), fontBold, brush, 10, y) : y += lineH
        g.DrawString("PAID:       P " & lblPaid.Text.Replace("P ", ""), font, brush, 10, y) : y += lineH
        g.DrawString("CHANGE:     P " & lblChange.Text.Replace("P ", ""), font, brush, 10, y) : y += lineH
        g.DrawString(New String("-"c, 42), font, brush, 10, y) : y += 20
        g.DrawString("Thank you for your purchase!", font, brush, 55, y)
    End Sub

End Class