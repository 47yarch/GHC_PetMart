Imports MySql.Data.MySqlClient

Public Class frmEditTransaction

    Public SaleID As Integer

    Private dgvItems As New DataGridView()

    Private btnSave As New Button()
    Private btnCancel As New Button()
    Private btnAddItem As New Button()
    Private btnRemoveItem As New Button()

    Private lblTitle As New Label()
    Private lblTotal As New Label()
    Private lblGrandTotal As New Label()
    Private lblAmountPaid As New Label()
    Private txtAmountPaid As New TextBox()
    Private lblChange As New Label()

    Private dtItems As New DataTable()

    Private decTotal As Decimal = 0D


    Private Sub frmEditTransaction_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Edit Transaction"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(950, 620)
        Me.BackColor = Color.WhiteSmoke

        BuildUI()

        SetupTable()

        LoadSale()

    End Sub

    Private Sub BuildUI()

        lblTitle.Text = "Edit Transaction"
        lblTitle.Font = New Font("Segoe UI", 15, FontStyle.Bold)
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(20, 15)

        Me.Controls.Add(lblTitle)

        dgvItems.Location = New Point(20, 55)
        dgvItems.Size = New Size(890, 360)

        dgvItems.AllowUserToAddRows = False
        dgvItems.AllowUserToDeleteRows = False
        dgvItems.AllowUserToResizeRows = False

        dgvItems.MultiSelect = False
        dgvItems.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgvItems.EditMode = DataGridViewEditMode.EditOnEnter

        dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvItems.BackgroundColor = Color.White

        AddHandler dgvItems.CellEndEdit, AddressOf dgvItems_CellEndEdit
        AddHandler dgvItems.CellValueChanged, AddressOf dgvItems_CellValueChanged
        AddHandler dgvItems.CurrentCellDirtyStateChanged, AddressOf dgvItems_CurrentCellDirtyStateChanged

        Me.Controls.Add(dgvItems)

        lblTotal.Text = "TOTAL : ₱0.00"
        lblTotal.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTotal.AutoSize = True
        lblTotal.Location = New Point(20, 435)

        Me.Controls.Add(lblTotal)

        lblGrandTotal.Text = "Grand Total : ₱0.00"
        lblGrandTotal.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblGrandTotal.AutoSize = True
        lblGrandTotal.Location = New Point(20, 465)

        Me.Controls.Add(lblGrandTotal)

        lblAmountPaid.Text = "Amount Paid"
        lblAmountPaid.AutoSize = True
        lblAmountPaid.Location = New Point(300, 470)

        Me.Controls.Add(lblAmountPaid)

        txtAmountPaid.Location = New Point(400, 465)
        txtAmountPaid.Size = New Size(140, 30)
        txtAmountPaid.Text = "0.00"

        AddHandler txtAmountPaid.TextChanged, AddressOf AmountPaidChanged

        Me.Controls.Add(txtAmountPaid)

        lblChange.Text = "Change : ₱0.00"
        lblChange.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        lblChange.AutoSize = True
        lblChange.Location = New Point(560, 470)

        Me.Controls.Add(lblChange)

        btnAddItem.Text = "Add Product"
        btnAddItem.Size = New Size(130, 40)
        btnAddItem.Location = New Point(20, 515)

        AddHandler btnAddItem.Click, AddressOf AddProduct

        Me.Controls.Add(btnAddItem)

        btnRemoveItem.Text = "Remove Item"
        btnRemoveItem.Size = New Size(130, 40)
        btnRemoveItem.Location = New Point(165, 515)

        AddHandler btnRemoveItem.Click, AddressOf RemoveProduct

        Me.Controls.Add(btnRemoveItem)

        btnSave.Text = "Save Changes"
        btnSave.Size = New Size(150, 40)
        btnSave.Location = New Point(580, 515)

        btnSave.BackColor = Color.RoyalBlue
        btnSave.ForeColor = Color.White
        btnSave.FlatStyle = FlatStyle.Flat

        AddHandler btnSave.Click, AddressOf btnSave_Click

        Me.Controls.Add(btnSave)

        btnCancel.Text = "Cancel"
        btnCancel.Size = New Size(120, 40)
        btnCancel.Location = New Point(750, 515)

        AddHandler btnCancel.Click,
                Sub()

                    Me.DialogResult = DialogResult.Cancel
                    Me.Close()

                End Sub

        Me.Controls.Add(btnCancel)

    End Sub

    Private Sub SetupTable()

        dtItems = New DataTable()

        dtItems.Columns.Add("product_id", GetType(Integer))
        dtItems.Columns.Add("Product", GetType(String))
        dtItems.Columns.Add("Price", GetType(Decimal))
        dtItems.Columns.Add("Qty", GetType(Integer))
        dtItems.Columns.Add("Discount", GetType(Decimal))
        dtItems.Columns.Add("Subtotal", GetType(Decimal))

        dgvItems.DataSource = dtItems

        dgvItems.Columns("product_id").Visible = False

        dgvItems.Columns("Product").ReadOnly = True
        dgvItems.Columns("Subtotal").ReadOnly = True

        dgvItems.Columns("Price").ReadOnly = True
        dgvItems.Columns("Qty").ReadOnly = False
        dgvItems.Columns("Discount").ReadOnly = False

        dgvItems.Columns("Price").DefaultCellStyle.Format = "N2"
        dgvItems.Columns("Discount").DefaultCellStyle.Format = "N2"
        dgvItems.Columns("Subtotal").DefaultCellStyle.Format = "N2"

        dgvItems.Columns("Qty").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleCenter

        dgvItems.Columns("Price").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight

        dgvItems.Columns("Discount").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight

        dgvItems.Columns("Subtotal").DefaultCellStyle.Alignment =
                DataGridViewContentAlignment.MiddleRight

    End Sub

    Private Sub LoadSale()

        dtItems.Rows.Clear()

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand(
    "SELECT
        sd.product_id,
        p.product_name,
        sd.unit_price,
        sd.quantity,
        IFNULL(sd.discount,0) AS discount
    FROM tbl_sale_details sd
    INNER JOIN tbl_products p
        ON sd.product_id = p.product_id
    WHERE sd.sale_id=@id
    ORDER BY p.product_name", conn)

            cmd.Parameters.AddWithValue("@id", SaleID)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            For Each r As DataRow In dt.Rows

                Dim qty As Integer = 1

                If Not IsDBNull(r("quantity")) Then
                    qty = CInt(r("quantity"))
                End If

                Dim price As Decimal = 0D

                If Not IsDBNull(r("unit_price")) Then
                    price = CDec(r("unit_price"))
                End If

                Dim disc As Decimal = 0D

                If Not IsDBNull(r("discount")) Then
                    disc = CDec(r("discount"))
                End If

                dtItems.Rows.Add(
                    CInt(r("product_id")),
                    r("product_name").ToString(),
                    price,
                    qty,
                    disc,
                    (qty * price) - disc)

            Next

            ComputeTotal()

            Dim cmdPaid As New MySqlCommand(
    "SELECT amount_paid
    FROM tbl_sales
    WHERE sale_id=@id", conn)

            cmdPaid.Parameters.AddWithValue("@id", SaleID)

            Dim obj = cmdPaid.ExecuteScalar()

            If obj IsNot Nothing AndAlso Not IsDBNull(obj) Then
                txtAmountPaid.Text = CDec(obj).ToString("N2")
            End If


        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try


    End Sub

    Private Sub ComputeTotal()

        decTotal = 0D

        For Each row As DataRow In dtItems.Rows

            Dim qty As Integer = 0
            Dim price As Decimal = 0D
            Dim discount As Decimal = 0D

            If Not IsDBNull(row("Qty")) Then
                qty = Convert.ToInt32(row("Qty"))
            End If

            If Not IsDBNull(row("Price")) Then
                price = Convert.ToDecimal(row("Price"))
            End If

            If Not IsDBNull(row("Discount")) Then
                discount = Convert.ToDecimal(row("Discount"))
            End If

            Dim subtotal As Decimal = (qty * price) - discount

            row("Subtotal") = subtotal

            decTotal += subtotal

        Next

        lblTotal.Text = "TOTAL : ₱" & decTotal.ToString("N2")
        lblGrandTotal.Text = "Grand Total : ₱" & decTotal.ToString("N2")

        dgvItems.Refresh()

    End Sub

    Private Sub dgvItems_CellEndEdit(
        sender As Object,
        e As DataGridViewCellEventArgs)

        ComputeTotal()

    End Sub

    Private Sub dgvItems_CellValueChanged(
        sender As Object,
        e As DataGridViewCellEventArgs)

        If e.RowIndex < 0 Then Exit Sub

        ComputeTotal()

    End Sub

    Private Sub dgvItems_CurrentCellDirtyStateChanged(
        sender As Object,
        e As EventArgs)

        If dgvItems.IsCurrentCellDirty Then
            dgvItems.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub

    Private Sub AddProduct(sender As Object, e As EventArgs)

        Dim frm As New frmProductPicker()

        If frm.ShowDialog() <> DialogResult.OK Then Exit Sub


        For Each r As DataRow In dtItems.Rows

            If CInt(r("product_id")) = frm.ProductID Then

                MessageBox.Show(
                    "This product is already in the transaction.",
                    "Duplicate Product",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

                Exit Sub

            End If

        Next

        Dim selectedPrice As Decimal

        If MessageBox.Show(
            "Use SALE PRICE?" &
            vbCrLf & vbCrLf &
            "YES = Sale Price" &
            vbCrLf &
            "NO = Regular Price",
            "Select Price",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.Yes Then

            selectedPrice = frm.SalePrice

        Else

            selectedPrice = frm.RegularPrice

        End If

        dtItems.Rows.Add(
            frm.ProductID,
            frm.ProductName,
            selectedPrice,
            1,
            0D,
            selectedPrice)

        ComputeTotal()

        dgvItems.ClearSelection()

        dgvItems.Rows(dgvItems.Rows.Count - 1).Selected = True

    End Sub

    Private Sub RemoveProduct(sender As Object, e As EventArgs)

        If dgvItems.SelectedCells.Count = 0 Then

            MessageBox.Show(
                "Please select a product first.",
                "Remove Product",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

            Exit Sub

        End If

        Dim rowIndex As Integer =
            dgvItems.SelectedCells(0).RowIndex

        If MessageBox.Show(
            "Remove the selected product?",
            "Confirm",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question) = DialogResult.No Then

            Exit Sub

        End If

        dgvItems.Rows.RemoveAt(rowIndex)

        ComputeTotal()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs)

        If dgvItems.Rows.Count = 0 Then
            MessageBox.Show("There are no items to save.")
            Exit Sub
        End If

        If Not OpenConnection() Then Exit Sub

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try


            Dim cmdOld As New MySqlCommand(
    "SELECT product_id, quantity
    FROM tbl_sale_details
    WHERE sale_id=@id", conn, trans)

            cmdOld.Parameters.AddWithValue("@id", SaleID)

            Dim dtOld As New DataTable()

            Dim daOld As New MySqlDataAdapter(cmdOld)
            daOld.Fill(dtOld)

            For Each r As DataRow In dtOld.Rows

                Dim cmdRestore As New MySqlCommand(
"UPDATE tbl_products
SET stock_quantity = stock_quantity + @qty
WHERE product_id=@pid", conn, trans)

                cmdRestore.Parameters.AddWithValue("@qty", CInt(r("quantity")))
                cmdRestore.Parameters.AddWithValue("@pid", CInt(r("product_id")))

                cmdRestore.ExecuteNonQuery()

                Dim cmdLog As New MySqlCommand(
    "INSERT INTO tbl_inventory_log
    (product_id,action_type,quantity,ref_id,notes,logged_by)
    VALUES
    (@pid,'Adjustment',@qty,@sale,'Restored during transaction edit',@user)",
    conn, trans)

                cmdLog.Parameters.AddWithValue("@pid", CInt(r("product_id")))
                cmdLog.Parameters.AddWithValue("@qty", CInt(r("quantity")))
                cmdLog.Parameters.AddWithValue("@sale", SaleID)
                cmdLog.Parameters.AddWithValue("@user", LoggedUserID)

                cmdLog.ExecuteNonQuery()

            Next


            Dim cmdDelete As New MySqlCommand(
    "DELETE FROM tbl_sale_details
    WHERE sale_id=@id", conn, trans)

            cmdDelete.Parameters.AddWithValue("@id", SaleID)

            cmdDelete.ExecuteNonQuery()

            Dim grandTotal As Decimal = 0D


            For Each row As DataGridViewRow In dgvItems.Rows

                If row.IsNewRow Then Continue For

                Dim productID As Integer = CInt(row.Cells("product_id").Value)
                Dim qty As Integer = CInt(row.Cells("Qty").Value)

                If qty <= 0 Then
                    Throw New Exception("Quantity must be greater than zero.")
                End If

                Dim cmdCheck As New MySqlCommand(
    "SELECT stock_quantity
    FROM tbl_products
    WHERE product_id=@id", conn, trans)

                cmdCheck.Parameters.AddWithValue("@id", productID)

                Dim available As Integer =
                    Convert.ToInt32(cmdCheck.ExecuteScalar())

                If qty > available Then

                    Dim productName As String =
                        row.Cells("Product").Value.ToString()

                    Throw New Exception(
                        "Insufficient stock for:" &
                        vbCrLf &
                        productName &
                        vbCrLf &
                        "Available: " & available &
                        vbCrLf &
                        "Requested: " & qty)

                End If

            Next


            For Each row As DataGridViewRow In dgvItems.Rows

                If row.IsNewRow Then Continue For

                Dim productID As Integer = CInt(row.Cells("product_id").Value)
                Dim qty As Integer = CInt(row.Cells("Qty").Value)
                Dim price As Decimal = CDec(row.Cells("Price").Value)
                Dim discount As Decimal = CDec(row.Cells("Discount").Value)

                Dim subtotal As Decimal =
                        (qty * price) - discount

                grandTotal += subtotal


                Dim cmdInsert As New MySqlCommand(
    "INSERT INTO tbl_sale_details
    (
        sale_id,
        product_id,
        quantity,
        unit_price,
        discount
    )
    VALUES
    (
        @sale,
        @product,
        @qty,
        @price,
        @discount
    )", conn, trans)

                cmdInsert.Parameters.AddWithValue("@sale", SaleID)
                cmdInsert.Parameters.AddWithValue("@product", productID)
                cmdInsert.Parameters.AddWithValue("@qty", qty)
                cmdInsert.Parameters.AddWithValue("@price", price)
                cmdInsert.Parameters.AddWithValue("@discount", discount)

                cmdInsert.ExecuteNonQuery()


                Dim cmdStock As New MySqlCommand(
    "UPDATE tbl_products
    SET stock_quantity = stock_quantity - @qty,
        status =
        CASE
            WHEN stock_quantity - @qty <= 0
                THEN 'Sold Out'
            ELSE 'In Stock'
        END
    WHERE product_id=@pid", conn, trans)

                cmdStock.Parameters.AddWithValue("@qty", qty)
                cmdStock.Parameters.AddWithValue("@pid", productID)

                cmdStock.ExecuteNonQuery()


                Dim cmdLogDeduct As New MySqlCommand(
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
        'Adjustment',
        @qty,
        @sale,
        'Deducted after transaction edit',
        @user
    )", conn, trans)

                cmdLogDeduct.Parameters.AddWithValue("@pid", productID)
                cmdLogDeduct.Parameters.AddWithValue("@qty", -qty)
                cmdLogDeduct.Parameters.AddWithValue("@sale", SaleID)
                cmdLogDeduct.Parameters.AddWithValue("@user", LoggedUserID)

                cmdLogDeduct.ExecuteNonQuery()

            Next

            Dim cmdSale As New MySqlCommand(
"UPDATE tbl_sales
SET amount_paid=@paid
WHERE sale_id=@id", conn, trans)

            cmdSale.Parameters.AddWithValue("@paid", CDec(txtAmountPaid.Text))
            cmdSale.Parameters.AddWithValue("@id", SaleID)

            cmdSale.ExecuteNonQuery()

            trans.Commit()

            MessageBox.Show(
                    "Transaction updated successfully.",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(
                    ex.Message,
                    "Transaction Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub AmountPaidChanged(sender As Object, e As EventArgs)

        Dim paid As Decimal = 0D

        Decimal.TryParse(txtAmountPaid.Text, paid)

        lblChange.Text =
                "Change : ₱" &
                (paid - decTotal).ToString("N2")

    End Sub

End Class