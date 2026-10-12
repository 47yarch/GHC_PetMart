Imports MySql.Data.MySqlClient


Public Class frmPurchaseOrders

    Private dgvOrders As New DataGridView()
    Private dgvItems As New DataGridView()
    Private cboSupplier As New ComboBox()
    Private cboStatus As New ComboBox()


    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Purchase Orders"
        Me.Size = New Size(1100, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Purchase Order Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lSup As New Label() With {.Text = "Supplier:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lSup)
        cboSupplier.Location = New Point(75, 52) : cboSupplier.Size = New Size(200, 28)
        cboSupplier.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Controls.Add(cboSupplier)

        Dim lSt As New Label() With {.Text = "Status:", .AutoSize = True, .Location = New Point(290, 55)}
        Me.Controls.Add(lSt)
        cboStatus.Location = New Point(340, 52) : cboStatus.Size = New Size(140, 28)
        cboStatus.DropDownStyle = ComboBoxStyle.DropDownList
        cboStatus.Items.AddRange(New String() {"All", "Pending", "Received", "Cancelled"})
        cboStatus.SelectedIndex = 0
        Me.Controls.Add(cboStatus)

        Dim btnLoad As New Button() With {
            .Text = "Load", .Size = New Size(75, 28), .Location = New Point(495, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnLoad.Click, AddressOf LoadOrders
        Me.Controls.Add(btnLoad)


        Dim btnNew As New Button() With {
    .Text = "New PO",
    .Size = New Size(110, 32),
    .Location = New Point(15, 92),
    .BackColor = Color.FromArgb(63, 81, 181),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnNew.Click, AddressOf NewPurchaseOrder
        Me.Controls.Add(btnNew)


        Dim btnEdit As New Button() With {
    .Text = "Edit",
    .Size = New Size(90, 32),
    .Location = New Point(135, 92),
    .BackColor = Color.FromArgb(255, 152, 0),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnEdit.Click, AddressOf EditPurchaseOrder
        Me.Controls.Add(btnEdit)


        Dim btnDelete As New Button() With {
    .Text = "Delete",
    .Size = New Size(90, 32),
    .Location = New Point(235, 92),
    .BackColor = Color.Firebrick,
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnDelete.Click, AddressOf DeletePurchaseOrder
        Me.Controls.Add(btnDelete)


        Dim btnReceive As New Button() With {
    .Text = "Mark Received",
    .Size = New Size(130, 32),
    .Location = New Point(335, 92),
    .BackColor = Color.FromArgb(76, 175, 80),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnReceive.Click, AddressOf MarkReceived
        Me.Controls.Add(btnReceive)



        Dim btnCancel As New Button() With {
    .Text = "Cancel PO",
    .Size = New Size(120, 32),
    .Location = New Point(475, 92),
    .BackColor = Color.FromArgb(244, 67, 54),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnCancel.Click, AddressOf MarkCancelled
        Me.Controls.Add(btnCancel)

        dgvOrders.Location = New Point(15, 135)
        dgvOrders.Size = New Size(1060, 220)
        dgvOrders.ReadOnly = True
        dgvOrders.AllowUserToAddRows = False
        dgvOrders.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvOrders.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvOrders.BackgroundColor = Color.White
        AddHandler dgvOrders.SelectionChanged, AddressOf LoadOrderItems
        Me.Controls.Add(dgvOrders)

        Dim lbl2 As New Label() With {.Text = "Order Items:", .AutoSize = True, .Font = New Font("Arial", 9, FontStyle.Bold), .Location = New Point(15, 365)}
        Me.Controls.Add(lbl2)

        dgvItems.Location = New Point(15, 385)
        dgvItems.Size = New Size(1060, 215)
        dgvItems.ReadOnly = True
        dgvItems.AllowUserToAddRows = False
        dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvItems.BackgroundColor = Color.White
        Me.Controls.Add(dgvItems)

        LoadSupplierDropdown()
        LoadOrders(Nothing, Nothing)


    End Sub

    Private Sub LoadSupplierDropdown()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter("SELECT supplier_id, supplier_name FROM tbl_suppliers ORDER BY supplier_name", conn)
            Dim dt As New DataTable() : da.Fill(dt)
            cboSupplier.Items.Clear()
            cboSupplier.Items.Add("All Suppliers")
            For Each row As DataRow In dt.Rows
                cboSupplier.Items.Add(row("supplier_name").ToString())
            Next
            cboSupplier.SelectedIndex = 0
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadOrders(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim sql As String =
            "SELECT po.po_id, " &
            "s.supplier_name AS 'Supplier', " &
            "po.order_date AS 'Order Date', " &
            "po.expected_date AS 'Expected', " &
            "po.received_date AS 'Received', " &
            "po.status AS 'Status', " &
            "SUM(d.quantity * d.unit_cost) AS 'Total' " &
            "FROM tbl_purchase_orders po " &
            "JOIN tbl_suppliers s ON po.supplier_id = s.supplier_id " &
            "JOIN tbl_po_details d ON po.po_id = d.po_id " &
            "WHERE 1=1 "

            If cboStatus.Text <> "All" Then
                sql &= "AND po.status=@status "
            End If

            If cboSupplier.SelectedIndex > 0 Then
                sql &= "AND s.supplier_name=@supplier "
            End If

            sql &= "GROUP BY po.po_id " &
               "ORDER BY po.order_date DESC"

            Dim cmd As New MySqlCommand(sql, conn)

            If cboStatus.Text <> "All" Then
                cmd.Parameters.AddWithValue("@status", cboStatus.Text)
            End If

            If cboSupplier.SelectedIndex > 0 Then
                cmd.Parameters.AddWithValue("@supplier", cboSupplier.Text)
            End If

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            dgvOrders.DataSource = dt

            If dgvOrders.Columns.Contains("po_id") Then
                dgvOrders.Columns("po_id").Visible = False
            End If

        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadOrderItems(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then Return

        Dim poID As Integer =
        CInt(dgvOrders.SelectedRows(0).Cells("po_id").Value)

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
            "SELECT " &
            "p.product_name AS 'Product', " &
            "d.quantity AS 'Qty', " &
            "d.unit_cost AS 'Unit Cost', " &
            "(d.quantity * d.unit_cost) AS 'Subtotal' " &
            "FROM tbl_po_details d " &
            "JOIN tbl_products p ON d.product_id = p.product_id " &
            "WHERE d.po_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", poID)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            dgvItems.DataSource = dt

        Finally
            CloseConnection()
        End Try

    End Sub

    Private Sub NewPurchaseOrder(sender As Object, e As EventArgs)

        Dim f As New frmNewPurchaseOrder()

        f.ShowDialog()

        LoadOrders(Nothing, Nothing)

    End Sub


    Private Sub EditPurchaseOrder(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then

            MessageBox.Show("Select a purchase order.")

            Return

        End If

        Dim status As String =
        dgvOrders.SelectedRows(0).Cells("Status").Value.ToString()

        If status <> "Pending" Then

            MessageBox.Show(
            "Only Pending purchase orders can be edited.")

            Return

        End If

        Dim poID As Integer =
        CInt(dgvOrders.SelectedRows(0).Cells("po_id").Value)

        Dim f As New frmNewPurchaseOrder()


        f.PurchaseOrderID = poID

        f.ShowDialog()

        LoadOrders(Nothing, Nothing)

    End Sub

    Private Sub MarkReceived(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then
            MessageBox.Show("Select a purchase order.")
            Return
        End If

        Dim poID As Integer =
        CInt(dgvOrders.SelectedRows(0).Cells("po_id").Value)

        Dim status As String =
        dgvOrders.SelectedRows(0).Cells("Status").Value.ToString()

        If status <> "Pending" Then
            MessageBox.Show("Only Pending purchase orders can be received.")
            Return
        End If

        If MessageBox.Show(
        "Mark this Purchase Order as RECEIVED?",
        "Confirm",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question) <> DialogResult.Yes Then

            Return

        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try



            Dim cmdItems As New MySqlCommand(
        "SELECT product_id, quantity, unit_cost
         FROM tbl_po_details
         WHERE po_id=@id", conn, trans)

            cmdItems.Parameters.AddWithValue("@id", poID)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmdItems)

            da.Fill(dt)

            Dim totalExpense As Decimal = 0D

            For Each row As DataRow In dt.Rows

                Dim productID As Integer = CInt(row("product_id"))
                Dim qty As Integer = CInt(row("quantity"))
                Dim unitCost As Decimal = CDec(row("unit_cost"))
                totalExpense += qty * unitCost


                Dim cmdStock As New MySqlCommand(
    "UPDATE tbl_products
     SET stock_quantity = stock_quantity + @qty,
    unit_cost = @cost,
         status='In Stock'
     WHERE product_id=@pid", conn, trans)

                cmdStock.Parameters.AddWithValue("@qty", qty)
                cmdStock.Parameters.AddWithValue("@pid", productID)
                cmdStock.Parameters.AddWithValue("@cost", CDec(row("unit_cost")))

                cmdStock.ExecuteNonQuery()


                Dim cmdLog As New MySqlCommand(
    "INSERT INTO tbl_inventory_log
    (product_id,
     action_type,
     quantity,
     notes,
     logged_by)
     VALUES
    (@pid,
     'Restock',
     @qty,
     @note,
     @uid)", conn, trans)

                cmdLog.Parameters.AddWithValue("@pid", productID)
                cmdLog.Parameters.AddWithValue("@qty", qty)
                cmdLog.Parameters.AddWithValue("@note", "Purchase Order #" & poID & " Received")
                cmdLog.Parameters.AddWithValue("@uid", LoggedUserID)

                cmdLog.ExecuteNonQuery()

            Next



            Dim cmdPO As New MySqlCommand(
        "UPDATE tbl_purchase_orders
         SET
             status='Received',
             received_date=CURDATE()
         WHERE po_id=@id", conn, trans)

            cmdPO.Parameters.AddWithValue("@id", poID)

            cmdPO.ExecuteNonQuery()

            Dim cmdExpense As New MySqlCommand(
"INSERT INTO tbl_expenses
(
expense_date,
category,
description,
amount,
recorded_by
)
VALUES
(
CURDATE(),
'Supplies',
@desc,
@amount,
@user
)", conn, trans)

            cmdExpense.Parameters.AddWithValue(
    "@desc",
    "Purchase Order #" & poID & " received")

            cmdExpense.Parameters.AddWithValue(
    "@amount",
    totalExpense)

            cmdExpense.Parameters.AddWithValue(
    "@user",
    LoggedUserID)

            cmdExpense.ExecuteNonQuery()

            trans.Commit()

            MessageBox.Show(
            "Purchase Order successfully received.",
            "Success",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

        LoadOrders(Nothing, Nothing)

        dgvItems.DataSource = Nothing

    End Sub

    Private Sub MarkCancelled(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then
            MessageBox.Show("Select a purchase order.")
            Return
        End If

        Dim poID As Integer =
        CInt(dgvOrders.SelectedRows(0).Cells("po_id").Value)

        Dim status As String =
        dgvOrders.SelectedRows(0).Cells("Status").Value.ToString()

        If status = "Received" Then
            MessageBox.Show("Received purchase orders cannot be cancelled.")
            Return
        End If

        If status = "Cancelled" Then
            MessageBox.Show("This purchase order is already cancelled.")
            Return
        End If

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
        "UPDATE tbl_purchase_orders
         SET status='Cancelled'
         WHERE po_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", poID)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Purchase Order cancelled.")

        Finally

            CloseConnection()

        End Try

        LoadOrders(Nothing, Nothing)

    End Sub

    Private Sub DeletePurchaseOrder(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a purchase order.")
            Return
        End If

        Dim poID As Integer =
            CInt(dgvOrders.SelectedRows(0).Cells("po_id").Value)

        Dim status As String =
            dgvOrders.SelectedRows(0).Cells("Status").Value.ToString()

        If status <> "Pending" Then
            MessageBox.Show("Only Pending purchase orders can be deleted.")
            Return
        End If

        If MessageBox.Show(
            "Delete this Purchase Order permanently?",
            "Confirm Delete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning) <> DialogResult.Yes Then

            Return

        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try


            Dim cmd1 As New MySqlCommand(
                "DELETE FROM tbl_po_details
             WHERE po_id=@id", conn, trans)

            cmd1.Parameters.AddWithValue("@id", poID)

            cmd1.ExecuteNonQuery()



            Dim cmd2 As New MySqlCommand(
                "DELETE FROM tbl_purchase_orders
             WHERE po_id=@id", conn, trans)

            cmd2.Parameters.AddWithValue("@id", poID)

            cmd2.ExecuteNonQuery()

            trans.Commit()

            MessageBox.Show(
                "Purchase Order deleted successfully.",
                "Deleted",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

        dgvItems.DataSource = Nothing

        LoadOrders(Nothing, Nothing)

    End Sub



End Class