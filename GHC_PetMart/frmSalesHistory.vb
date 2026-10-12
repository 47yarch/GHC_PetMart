Imports MySql.Data.MySqlClient

Public Class frmSalesHistory

    Private dgvSales As New DataGridView()
    Private dgvItems As New DataGridView()
    Private txtSearch As New TextBox()
    Private dtpFrom As New DateTimePicker()
    Private dtpTo As New DateTimePicker()
    Private lblTotal As New Label()
    Private btnEdit As New Button()
    Private btnDelete As New Button()
    Private btnReceipt As New Button()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Sales History"
        Me.Size = New Size(1100, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Sales History",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lS As New Label() With {.Text = "Search:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lS)
        txtSearch.Size = New Size(180, 28) : txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim lF As New Label() With {.Text = "From:", .AutoSize = True, .Location = New Point(260, 55)}
        Me.Controls.Add(lF)
        dtpFrom.Location = New Point(300, 52) : dtpFrom.Size = New Size(140, 28)
        dtpFrom.Value = DateTime.Now.AddDays(-30)
        Me.Controls.Add(dtpFrom)

        Dim lT As New Label() With {.Text = "To:", .AutoSize = True, .Location = New Point(455, 55)}
        Me.Controls.Add(lT)
        dtpTo.Location = New Point(475, 52) : dtpTo.Size = New Size(140, 28)
        Me.Controls.Add(dtpTo)

        Dim btnLoad As New Button() With {
            .Text = "Load", .Size = New Size(80, 28), .Location = New Point(630, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnLoad.Click, AddressOf LoadSales
        Me.Controls.Add(btnLoad)

        lblTotal.AutoSize = True
        lblTotal.Font = New Font("Arial", 10, FontStyle.Bold)
        lblTotal.Location = New Point(730, 57)
        Me.Controls.Add(lblTotal)

        btnEdit.Text = "Edit Transaction"
        btnEdit.Size = New Size(130, 35)
        btnEdit.Location = New Point(15, 600)
        btnEdit.BackColor = Color.FromArgb(33, 150, 243)
        btnEdit.ForeColor = Color.White
        btnEdit.FlatStyle = FlatStyle.Flat
        AddHandler btnEdit.Click, AddressOf EditTransaction
        Me.Controls.Add(btnEdit)

        btnDelete.Text = "Delete Transaction"
        btnDelete.Size = New Size(140, 35)
        btnDelete.Location = New Point(160, 600)
        btnDelete.BackColor = Color.Firebrick
        btnDelete.ForeColor = Color.White
        btnDelete.FlatStyle = FlatStyle.Flat
        AddHandler btnDelete.Click, AddressOf DeleteTransaction
        Me.Controls.Add(btnDelete)

        btnReceipt.Text = "Reprint Receipt"
        btnReceipt.Size = New Size(130, 35)
        btnReceipt.Location = New Point(315, 600)
        btnReceipt.BackColor = Color.SeaGreen
        btnReceipt.ForeColor = Color.White
        btnReceipt.FlatStyle = FlatStyle.Flat
        AddHandler btnReceipt.Click, AddressOf ReprintReceipt
        Me.Controls.Add(btnReceipt)



        Dim lbl1 As New Label() With {.Text = "Transactions:", .AutoSize = True, .Font = New Font("Arial", 9, FontStyle.Bold), .Location = New Point(15, 92)}
        Me.Controls.Add(lbl1)

        dgvSales.Location = New Point(15, 112)
        dgvSales.Size = New Size(1060, 230)
        dgvSales.ReadOnly = True
        dgvSales.AllowUserToAddRows = False
        dgvSales.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvSales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvSales.BackgroundColor = Color.White
        AddHandler dgvSales.SelectionChanged, AddressOf LoadSaleItems
        Me.Controls.Add(dgvSales)

        Dim lbl2 As New Label() With {.Text = "Items in Selected Sale:", .AutoSize = True, .Font = New Font("Arial", 9, FontStyle.Bold), .Location = New Point(15, 352)}
        Me.Controls.Add(lbl2)

        dgvItems.Location = New Point(15, 372)
        dgvItems.Size = New Size(1060, 220)
        dgvItems.ReadOnly = True
        dgvItems.AllowUserToAddRows = False
        dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvItems.BackgroundColor = Color.White
        Me.Controls.Add(dgvItems)

        LoadSales(Nothing, Nothing)
    End Sub

    Private Sub LoadSales(sender As Object, e As EventArgs)

        If Not OpenConnection() Then Return

        Try

            Dim search As String = "%" & txtSearch.Text.Trim() & "%"

            Dim cmd As New MySqlCommand(
"SELECT
    s.sale_id,
    s.receipt_no AS 'Receipt',
    s.sale_date AS 'Date',
    s.status AS 'Status',

    CONCAT(c.first_name,' ',c.last_name) AS 'Customer',
    u.full_name AS 'Employee',

    IFNULL(SUM((sd.quantity * sd.unit_price) - sd.discount),0) AS 'Total',

    s.amount_paid AS 'Paid',

    (
        s.amount_paid -
        IFNULL(SUM((sd.quantity * sd.unit_price) - sd.discount),0)
    ) AS 'Change'

FROM tbl_sales s

LEFT JOIN tbl_customers c
    ON s.customer_id = c.customer_id

LEFT JOIN tbl_users u
    ON s.employee_id = u.user_id

LEFT JOIN tbl_sale_details sd
    ON s.sale_id = sd.sale_id

WHERE
(
    s.receipt_no LIKE @s
    OR CONCAT(c.first_name,' ',c.last_name) LIKE @s
)
AND DATE(s.sale_date) BETWEEN @fd AND @td

GROUP BY
    s.sale_id,
    s.receipt_no,
    s.sale_date,
    s.status,
    c.first_name,
    c.last_name,
    u.full_name,
    s.amount_paid

ORDER BY s.sale_date DESC", conn)

            cmd.Parameters.AddWithValue("@s", search)
            cmd.Parameters.AddWithValue("@fd", dtpFrom.Value.ToString("yyyy-MM-dd"))
            cmd.Parameters.AddWithValue("@td", dtpTo.Value.ToString("yyyy-MM-dd"))

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)

            da.Fill(dt)

            dgvSales.DataSource = dt

            dgvSales.Columns("sale_id").Visible = False

            Dim total As Decimal = 0D

            For Each row As DataRow In dt.Rows

                If row("Status").ToString() = "Completed" Then

                    total += CDec(row("Total"))

                End If

            Next

            lblTotal.Text = "Total : ₱" & total.ToString("N2")

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub LoadSaleItems(sender As Object, e As EventArgs)
        If dgvSales.SelectedRows.Count = 0 Then Return
        Dim saleID As Integer = CInt(dgvSales.SelectedRows(0).Cells("sale_id").Value)

        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand(
"SELECT
    sd.product_id,
    p.product_name AS 'Product',
    sd.quantity AS 'Qty',
    sd.unit_price AS 'Unit Price',
    sd.discount AS 'Discount',
    (sd.quantity * sd.unit_price - sd.discount) AS 'Subtotal'

FROM tbl_sale_details sd

JOIN tbl_products p
ON sd.product_id = p.product_id

WHERE sd.sale_id = @id", conn)
            cmd.Parameters.AddWithValue("@id", saleID)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgvItems.DataSource = dt
            If dgvItems.Columns.Contains("product_id") Then
                dgvItems.Columns("product_id").Visible = False
            End If

        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub EditTransaction(sender As Object, e As EventArgs)

        If dgvSales.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a transaction.")
            Exit Sub
        End If

        If dgvSales.Columns.Contains("Status") Then
            If dgvSales.SelectedRows(0).Cells("Status").Value.ToString() = "Voided" Then
                MessageBox.Show("Voided transactions cannot be edited.")
                Exit Sub
            End If
        End If

        Dim frm As New frmEditTransaction()

        frm.SaleID = CInt(dgvSales.SelectedRows(0).Cells("sale_id").Value)

        If frm.ShowDialog() = DialogResult.OK Then

            LoadSales(Nothing, Nothing)
            LoadSaleItems(Nothing, Nothing)

        End If

    End Sub

    Private Sub DeleteTransaction(sender As Object, e As EventArgs)

        If dgvSales.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a transaction.")
            Exit Sub
        End If

        If MessageBox.Show(
        "Deleting this transaction will restore all product stocks." &
        vbCrLf & vbCrLf &
        "Continue?",
        "Delete Transaction",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning) = DialogResult.No Then
            Exit Sub
        End If

        Dim saleID As Integer =
        CInt(dgvSales.SelectedRows(0).Cells("sale_id").Value)

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try

            '==========================================
            ' Get all sold products
            '==========================================

            Dim cmdItems As New MySqlCommand(
        "SELECT product_id, quantity
         FROM tbl_sale_details
         WHERE sale_id=@id", conn, trans)

            cmdItems.Parameters.AddWithValue("@id", saleID)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmdItems)
            da.Fill(dt)

            For Each row As DataRow In dt.Rows

                Dim productID As Integer = CInt(row("product_id"))
                Dim qty As Integer = CInt(row("quantity"))

                '==========================================
                ' Restore Stock
                '==========================================

                Dim cmdStock As New MySqlCommand(
            "UPDATE tbl_products
             SET stock_quantity = stock_quantity + @qty,
status =
CASE
    WHEN stock_quantity + @qty <= 0
    THEN 'Sold Out'
    ELSE 'In Stock'
END
             WHERE product_id=@pid", conn, trans)

                cmdStock.Parameters.AddWithValue("@qty", qty)
                cmdStock.Parameters.AddWithValue("@pid", productID)

                cmdStock.ExecuteNonQuery()

                '==========================================
                ' Inventory Log (Return)
                '==========================================

                Dim cmdReturn As New MySqlCommand(
            "INSERT INTO tbl_inventory_log
            (
                product_id,
                action_type,
                quantity,
                ref_id,
                notes,
                logged_by
            )
            VALUES
            (
                @pid,
                'Return',
                @qty,
                @sid,
                'Transaction Deleted',
                @uid
            )", conn, trans)

                cmdReturn.Parameters.AddWithValue("@pid", productID)
                cmdReturn.Parameters.AddWithValue("@qty", qty)
                cmdReturn.Parameters.AddWithValue("@sid", saleID)
                cmdReturn.Parameters.AddWithValue("@uid", LoggedUserID)

                cmdReturn.ExecuteNonQuery()

            Next




            '==========================================
            ' Delete Sale Header
            '==========================================

            Dim cmdSale As New MySqlCommand(
"UPDATE tbl_sales
 SET status='Voided'
 WHERE sale_id=@id", conn, trans)

            cmdSale.Parameters.AddWithValue("@id", saleID)

            cmdSale.ExecuteNonQuery()

            trans.Commit()

            MessageBox.Show(
            "Transaction deleted successfully.",
            "Success",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

            LoadSales(Nothing, Nothing)
            dgvItems.DataSource = Nothing
            dgvItems.Rows.Clear()
            dgvItems.Refresh()

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(
            "Delete failed." &
            vbCrLf &
            ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub ReprintReceipt(sender As Object, e As EventArgs)

        If dgvSales.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a transaction.")
            Exit Sub
        End If

        Dim saleID As Integer =
            CInt(dgvSales.SelectedRows(0).Cells("sale_id").Value)

        Try

            Dim rpt As New frmReceipt()
            rpt.SaleID = saleID
            rpt.ShowDialog()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

End Class