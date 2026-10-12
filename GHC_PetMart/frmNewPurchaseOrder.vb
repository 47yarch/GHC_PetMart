Imports MySql.Data.MySqlClient

Public Class frmNewPurchaseOrder
    Public Property PurchaseOrderID As Integer = 0
    Private EditMode As Boolean = False

    Private cboSupplier As New ComboBox()
    Private cboProducts As New ComboBox()

    Private txtQty As New TextBox()
    Private txtCost As New TextBox()

    Private dtExpected As New DateTimePicker()

    Private dgvItems As New DataGridView()

    Private lblGrandTotal As New Label()

    Private btnAdd As New Button()
    Private btnRemove As New Button()
    Private btnSave As New Button()
    Private btnClear As New Button()

    Private Sub InitForm() Handles MyBase.Load

        Me.Text = "New Purchase Order"
        Me.Size = New Size(1180, 720)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label With {
            .Text = "Create Purchase Order",
            .Font = New Font("Arial", 15, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)


        Dim lblSupplier As New Label With {
            .Text = "Supplier",
            .AutoSize = True,
            .Location = New Point(20, 60)
        }

        Me.Controls.Add(lblSupplier)

        cboSupplier.Location = New Point(20, 80)
        cboSupplier.Size = New Size(260, 30)
        cboSupplier.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Controls.Add(cboSupplier)

        Dim lblExpected As New Label With {
            .Text = "Expected Delivery",
            .AutoSize = True,
            .Location = New Point(305, 60)
        }

        Me.Controls.Add(lblExpected)

        dtExpected.Location = New Point(305, 80)
        dtExpected.Size = New Size(180, 30)
        dtExpected.Value = Date.Today.AddDays(3)

        Me.Controls.Add(dtExpected)


        Dim lblProduct As New Label With {
            .Text = "Product",
            .AutoSize = True,
            .Location = New Point(20, 135)
        }

        Me.Controls.Add(lblProduct)

        cboProducts.Location = New Point(20, 155)
        cboProducts.Size = New Size(320, 30)
        cboProducts.DropDownStyle = ComboBoxStyle.DropDownList

        Me.Controls.Add(cboProducts)

        Dim lblQty As New Label With {
            .Text = "Quantity",
            .AutoSize = True,
            .Location = New Point(360, 135)
        }

        Me.Controls.Add(lblQty)

        txtQty.Location = New Point(360, 155)
        txtQty.Size = New Size(90, 30)

        Me.Controls.Add(txtQty)


        Dim lblCost As New Label With {
            .Text = "Unit Cost",
            .AutoSize = True,
            .Location = New Point(470, 135)
        }

        Me.Controls.Add(lblCost)

        txtCost.Location = New Point(470, 155)
        txtCost.Size = New Size(120, 30)
        txtCost.ReadOnly = True
        txtCost.BackColor = Color.Gainsboro

        Me.Controls.Add(txtCost)

        btnAdd.Text = "Add Item"
        btnAdd.Size = New Size(120, 35)
        btnAdd.Location = New Point(610, 152)
        btnAdd.BackColor = Color.FromArgb(63, 81, 181)
        btnAdd.ForeColor = Color.White
        btnAdd.FlatStyle = FlatStyle.Flat

        Me.Controls.Add(btnAdd)

        btnRemove.Text = "Remove Selected"
        btnRemove.Size = New Size(145, 35)
        btnRemove.Location = New Point(740, 152)
        btnRemove.BackColor = Color.Firebrick
        btnRemove.ForeColor = Color.White
        btnRemove.FlatStyle = FlatStyle.Flat

        Me.Controls.Add(btnRemove)

        dgvItems.Location = New Point(20, 210)
        dgvItems.Size = New Size(1120, 360)

        dgvItems.AllowUserToAddRows = False
        dgvItems.AllowUserToDeleteRows = False

        dgvItems.SelectionMode =
            DataGridViewSelectionMode.FullRowSelect

        dgvItems.MultiSelect = False

        dgvItems.ReadOnly = True

        dgvItems.AutoSizeColumnsMode =
            DataGridViewAutoSizeColumnsMode.Fill

        dgvItems.BackgroundColor = Color.White

        dgvItems.Columns.Add("ProductID", "Product ID")
        dgvItems.Columns.Add("Product", "Product")
        dgvItems.Columns.Add("Qty", "Quantity")
        dgvItems.Columns.Add("Cost", "Unit Cost")
        dgvItems.Columns.Add("Subtotal", "Subtotal")

        dgvItems.Columns("ProductID").Visible = False

        Me.Controls.Add(dgvItems)

        Dim lblTotal As New Label With {
            .Text = "Grand Total",
            .Font = New Font("Arial", 11, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 600)
        }

        Me.Controls.Add(lblTotal)

        lblGrandTotal.Text = "₱0.00"

        lblGrandTotal.Font =
            New Font("Arial", 16, FontStyle.Bold)

        lblGrandTotal.ForeColor =
            Color.FromArgb(0, 150, 136)

        lblGrandTotal.AutoSize = True

        lblGrandTotal.Location =
            New Point(20, 625)

        Me.Controls.Add(lblGrandTotal)


        btnSave.Text = "Save Purchase Order"

        btnSave.Size = New Size(220, 45)

        btnSave.Location = New Point(920, 610)

        btnSave.BackColor =
            Color.FromArgb(76, 175, 80)

        btnSave.ForeColor = Color.White

        btnSave.FlatStyle = FlatStyle.Flat

        Me.Controls.Add(btnSave)

        btnClear.Text = "Clear"

        btnClear.Size = New Size(120, 45)

        btnClear.Location = New Point(790, 610)

        btnClear.BackColor =
            Color.DarkOrange

        btnClear.ForeColor = Color.White

        btnClear.FlatStyle = FlatStyle.Flat

        Me.Controls.Add(btnClear)


        AddHandler btnAdd.Click, AddressOf btnAdd_Click
        AddHandler btnRemove.Click, AddressOf btnRemove_Click
        AddHandler btnSave.Click, AddressOf btnSave_Click
        AddHandler btnClear.Click, AddressOf btnClear_Click

        LoadSuppliers()
        LoadProducts()

        RegisterEvents()

        If PurchaseOrderID > 0 Then

            EditMode = True
            Me.Text = "Edit Purchase Order"

            LoadPurchaseOrder()

        End If

        AddHandler cboProducts.SelectedIndexChanged, AddressOf cboProducts_SelectedIndexChanged
    End Sub

    Private Sub LoadSuppliers()

        If Not OpenConnection() Then Return

        Try

            Dim da As New MySqlDataAdapter(
                "SELECT supplier_id, supplier_name
             FROM tbl_suppliers
             ORDER BY supplier_name", conn)

            Dim dt As New DataTable()
            da.Fill(dt)

            cboSupplier.DataSource = dt
            cboSupplier.DisplayMember = "supplier_name"
            cboSupplier.ValueMember = "supplier_id"

        Finally
            CloseConnection()
        End Try

    End Sub

    Private Sub LoadProducts()

        If Not OpenConnection() Then Return

        Try

            Dim da As New MySqlDataAdapter(
                "SELECT product_id, product_name
             FROM tbl_products
             WHERE status <> 'Discontinued'
             ORDER BY product_name", conn)

            Dim dt As New DataTable()
            da.Fill(dt)

            cboProducts.DataSource = dt
            cboProducts.DisplayMember = "product_name"
            cboProducts.ValueMember = "product_id"

        Finally
            CloseConnection()
        End Try

    End Sub

    Private Sub LoadPurchaseOrder()

        If PurchaseOrderID <= 0 Then Return

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
        "SELECT supplier_id, expected_date
         FROM tbl_purchase_orders
         WHERE po_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", PurchaseOrderID)

            Dim dr As MySqlDataReader = cmd.ExecuteReader()

            If dr.Read() Then

                cboSupplier.SelectedValue = CInt(dr("supplier_id"))

                If Not IsDBNull(dr("expected_date")) Then
                    dtExpected.Value = CDate(dr("expected_date"))
                End If

            End If

            dr.Close()

            dgvItems.Rows.Clear()

            Dim cmdItems As New MySqlCommand(
        "SELECT
            d.product_id,
            p.product_name,
            d.quantity,
            d.unit_cost
         FROM tbl_po_details d
         INNER JOIN tbl_products p
             ON d.product_id = p.product_id
         WHERE d.po_id=@id", conn)

            cmdItems.Parameters.AddWithValue("@id", PurchaseOrderID)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmdItems)

            da.Fill(dt)

            For Each row As DataRow In dt.Rows

                Dim subtotal As Decimal =
                CDec(row("quantity")) *
                CDec(row("unit_cost"))

                dgvItems.Rows.Add(
                row("product_id"),
                row("product_name"),
                row("quantity"),
                Format(CDec(row("unit_cost")), "N2"),
                Format(subtotal, "N2")
            )

            Next

            ComputeGrandTotal()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub
    Private Sub btnAdd_Click(sender As Object, e As EventArgs)

        If cboProducts.SelectedIndex = -1 Then

            MessageBox.Show("Select a product.")
            Return

        End If

        If txtQty.Text.Trim = "" Then

            MessageBox.Show("Enter quantity.")
            txtQty.Focus()
            Return

        End If

        If txtCost.Text.Trim = "" Then

            MessageBox.Show("Enter unit cost.")
            txtCost.Focus()
            Return

        End If

        Dim qty As Integer

        If Not Integer.TryParse(txtQty.Text, qty) Then

            MessageBox.Show("Quantity must be numeric.")
            Return

        End If

        If qty <= 0 Then

            MessageBox.Show("Quantity must be greater than zero.")
            Return

        End If

        Dim cost As Decimal

        If Not Decimal.TryParse(txtCost.Text, cost) Then

            MessageBox.Show("Invalid unit cost.")
            Return

        End If

        If cost <= 0 Then

            MessageBox.Show("Unit cost must be greater than zero.")
            Return

        End If

        Dim productID As Integer =
            CInt(cboProducts.SelectedValue)


        For Each r As DataGridViewRow In dgvItems.Rows

            If CInt(r.Cells("ProductID").Value) = productID Then

                MessageBox.Show("This product is already in the purchase order.")
                Return

            End If

        Next

        Dim subtotal As Decimal = qty * cost

        dgvItems.Rows.Add(
            productID,
            cboProducts.Text,
            qty,
            cost.ToString("N2"),
            subtotal.ToString("N2")
        )

        ComputeGrandTotal()

        txtQty.Clear()
        txtCost.Clear()

        cboProducts.Focus()

    End Sub


    Private Sub btnRemove_Click(sender As Object, e As EventArgs)

        If dgvItems.SelectedRows.Count = 0 Then

            MessageBox.Show("Select an item to remove.")
            Return

        End If

        dgvItems.Rows.RemoveAt(
            dgvItems.SelectedRows(0).Index)

        ComputeGrandTotal()

    End Sub



    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        cboSupplier.SelectedIndex = 0

        dtExpected.Value = Date.Today.AddDays(3)

        txtQty.Clear()

        txtCost.Clear()

        dgvItems.Rows.Clear()

        ComputeGrandTotal()

    End Sub

    Private Sub ComputeGrandTotal()

        Dim total As Decimal = 0

        For Each r As DataGridViewRow In dgvItems.Rows

            total += Convert.ToDecimal(
                r.Cells("Subtotal").Value)

        Next

        lblGrandTotal.Text = "₱" &
            total.ToString("N2")

    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs)

        If cboSupplier.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier.")
            Return
        End If

        If dgvItems.Rows.Count = 0 Then
            MessageBox.Show("Add at least one product.")
            Return
        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try

            Dim poID As Integer


            If Not EditMode Then

                Dim cmdPO As New MySqlCommand(
                "INSERT INTO tbl_purchase_orders
            (
                supplier_id,
                created_by,
                order_date,
                expected_date,
                status
            )
            VALUES
            (
                @supplier,
                @createdby,
                NOW(),
                @expected,
                'Pending'
            );

            SELECT LAST_INSERT_ID();", conn, trans)

                cmdPO.Parameters.AddWithValue("@supplier", cboSupplier.SelectedValue)
                cmdPO.Parameters.AddWithValue("@createdby", LoggedUserID)
                cmdPO.Parameters.AddWithValue("@expected", dtExpected.Value.Date)

                poID = Convert.ToInt32(cmdPO.ExecuteScalar())

            Else

                poID = PurchaseOrderID

                Dim cmdUpdate As New MySqlCommand(
                "UPDATE tbl_purchase_orders
             SET
                supplier_id=@supplier,
                expected_date=@expected
             WHERE po_id=@id", conn, trans)

                cmdUpdate.Parameters.AddWithValue("@supplier", cboSupplier.SelectedValue)
                cmdUpdate.Parameters.AddWithValue("@expected", dtExpected.Value.Date)
                cmdUpdate.Parameters.AddWithValue("@id", poID)

                cmdUpdate.ExecuteNonQuery()


                Dim cmdDelete As New MySqlCommand(
                "DELETE FROM tbl_po_details
             WHERE po_id=@id", conn, trans)

                cmdDelete.Parameters.AddWithValue("@id", poID)

                cmdDelete.ExecuteNonQuery()

            End If


            For Each r As DataGridViewRow In dgvItems.Rows

                Dim cmdDetail As New MySqlCommand(
            "INSERT INTO tbl_po_details
            (
                po_id,
                product_id,
                quantity,
                unit_cost
            )
            VALUES
            (
                @po,
                @product,
                @qty,
                @cost
            )", conn, trans)

                cmdDetail.Parameters.AddWithValue("@po", poID)
                cmdDetail.Parameters.AddWithValue("@product", CInt(r.Cells("ProductID").Value))
                cmdDetail.Parameters.AddWithValue("@qty", CInt(r.Cells("Qty").Value))
                cmdDetail.Parameters.AddWithValue("@cost", Convert.ToDecimal(r.Cells("Cost").Value))

                cmdDetail.ExecuteNonQuery()

            Next

            trans.Commit()

            If EditMode Then

                MessageBox.Show(
                "Purchase Order updated successfully.",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

            Else

                MessageBox.Show(
                "Purchase Order saved successfully.",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(
            ex.Message,
            "Database Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)

        Finally

            CloseConnection()

        End Try

    End Sub
    Private Sub UpdatePurchaseOrder()

        If cboSupplier.SelectedIndex = -1 Then
            MessageBox.Show("Please select a supplier.")
            Return
        End If

        If dgvItems.Rows.Count = 0 Then
            MessageBox.Show("Add at least one product.")
            Return
        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try

            Dim cmd As New MySqlCommand(
"UPDATE tbl_purchase_orders
SET
supplier_id=@supplier,
expected_date=@expected
WHERE po_id=@id", conn, trans)

            cmd.Parameters.AddWithValue("@supplier", cboSupplier.SelectedValue)
            cmd.Parameters.AddWithValue("@expected",
                                    dtExpected.Value.ToString("yyyy-MM-dd"))
            cmd.Parameters.AddWithValue("@id", PurchaseOrderID)

            cmd.ExecuteNonQuery()

            Dim cmdDelete As New MySqlCommand(
"DELETE FROM tbl_po_details
WHERE po_id=@id", conn, trans)

            cmdDelete.Parameters.AddWithValue("@id",
                                          PurchaseOrderID)

            cmdDelete.ExecuteNonQuery()

            For Each r As DataGridViewRow In dgvItems.Rows

                Dim cmdInsert As New MySqlCommand(
"INSERT INTO tbl_po_details
(
po_id,
product_id,
quantity,
unit_cost
)
VALUES
(
@po,
@product,
@qty,
@cost
)", conn, trans)

                cmdInsert.Parameters.AddWithValue("@po",
                                              PurchaseOrderID)

                cmdInsert.Parameters.AddWithValue("@product",
                                              r.Cells("ProductID").Value)

                cmdInsert.Parameters.AddWithValue("@qty",
                                              r.Cells("Qty").Value)

                cmdInsert.Parameters.AddWithValue("@cost",
                                              r.Cells("Cost").Value)

                cmdInsert.ExecuteNonQuery()

            Next

            trans.Commit()

            MessageBox.Show(
            "Purchase Order updated successfully.")

            Me.Close()

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub cboProducts_SelectedIndexChanged(sender As Object, e As EventArgs)

        If cboProducts.SelectedIndex = -1 Then Exit Sub

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand(
        "SELECT unit_cost
         FROM tbl_products
         WHERE product_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", cboProducts.SelectedValue)

            Dim result = cmd.ExecuteScalar()

            If result IsNot Nothing Then
                txtCost.Text = Format(CDec(result), "0.00")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            CloseConnection()

        End Try

    End Sub

    Private Sub txtQty_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Not Char.IsControl(e.KeyChar) AndAlso
           Not Char.IsDigit(e.KeyChar) Then

            e.Handled = True

        End If

    End Sub

    Private Sub txtCost_KeyPress(sender As Object, e As KeyPressEventArgs)

        If Char.IsControl(e.KeyChar) Then Exit Sub

        If Char.IsDigit(e.KeyChar) Then Exit Sub

        If e.KeyChar = "."c AndAlso
           Not txtCost.Text.Contains(".") Then Exit Sub

        e.Handled = True

    End Sub


    Private Sub txtQty_KeyDown(sender As Object, e As KeyEventArgs)

        If e.KeyCode = Keys.Enter Then

            txtCost.Focus()

        End If

    End Sub

    Private Sub txtCost_KeyDown(sender As Object, e As KeyEventArgs)

        If e.KeyCode = Keys.Enter Then

            btnAdd.PerformClick()

        End If

    End Sub

    Private Sub frmNewPurchaseOrder_FormClosing(
    sender As Object,
    e As FormClosingEventArgs) Handles Me.FormClosing

        If dgvItems.Rows.Count = 0 Then Return

        Dim ans As DialogResult

        ans = MessageBox.Show(
        "Discard this Purchase Order?",
        "Confirm",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)

        If ans = DialogResult.No Then

            e.Cancel = True

        End If

    End Sub



    Private Sub RegisterEvents()


        AddHandler txtQty.KeyDown,
            AddressOf txtQty_KeyDown

        AddHandler txtCost.KeyDown,
            AddressOf txtCost_KeyDown

        AddHandler cboProducts.SelectedIndexChanged, AddressOf cboProducts_SelectedIndexChanged

    End Sub
End Class