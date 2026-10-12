Imports MySql.Data.MySqlClient

Public Class frmInventory

    Private dgvAll As New DataGridView()
    Private dgvLow As New DataGridView()
    Private dgvLog As New DataGridView()
    Private txtAdjQty As New TextBox()
    Private cboAdjType As New ComboBox()
    Private txtAdjNote As New TextBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Inventory Management"
        Me.Size = New Size(1100, 660)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim tc As New TabControl() With {.Size = New Size(1070, 600), .Location = New Point(10, 10)}

        Dim t1 As New TabPage("All Stock")
        dgvAll.Dock = DockStyle.Fill
        dgvAll.ReadOnly = True
        dgvAll.AllowUserToAddRows = False
        dgvAll.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvAll.BackgroundColor = Color.White
        t1.Controls.Add(dgvAll)

        Dim t2 As New TabPage("Low Stock Alert")
        dgvLow.Dock = DockStyle.Fill
        dgvLow.ReadOnly = True
        dgvLow.AllowUserToAddRows = False
        dgvLow.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvLow.BackgroundColor = Color.White
        t2.Controls.Add(dgvLow)

        Dim t3 As New TabPage("Manual Adjustment")
        Dim pnl As New Panel() With {.Dock = DockStyle.Fill, .Padding = New Padding(20)}

        Dim lbl1 As New Label() With {.Text = "Select product from All Stock tab first, then adjust here.", .AutoSize = True, .Location = New Point(20, 20)}
        pnl.Controls.Add(lbl1)

        Dim l2 As New Label() With {.Text = "Adjustment Type:", .AutoSize = True, .Location = New Point(20, 60)}
        pnl.Controls.Add(l2)
        cboAdjType.Items.AddRange(New String() {"Add Stock", "Deduct Stock"})
        cboAdjType.SelectedIndex = 0
        cboAdjType.Location = New Point(20, 80) : cboAdjType.Size = New Size(200, 28)
        cboAdjType.DropDownStyle = ComboBoxStyle.DropDownList
        pnl.Controls.Add(cboAdjType)

        Dim l3 As New Label() With {.Text = "Quantity:", .AutoSize = True, .Location = New Point(20, 120)}
        pnl.Controls.Add(l3)
        txtAdjQty.Location = New Point(20, 140) : txtAdjQty.Size = New Size(200, 28)
        pnl.Controls.Add(txtAdjQty)

        Dim l4 As New Label() With {.Text = "Reason / Notes:", .AutoSize = True, .Location = New Point(20, 180)}
        pnl.Controls.Add(l4)
        txtAdjNote.Location = New Point(20, 200) : txtAdjNote.Size = New Size(400, 28)
        pnl.Controls.Add(txtAdjNote)

        Dim btnAdj As New Button() With {
            .Text = "APPLY ADJUSTMENT", .Size = New Size(200, 40), .Location = New Point(20, 250),
            .BackColor = Color.FromArgb(255, 152, 0), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdj.Click, AddressOf ApplyAdjustment
        pnl.Controls.Add(btnAdj)
        t3.Controls.Add(pnl)

        Dim t4 As New TabPage("Inventory Log")
        dgvLog.Dock = DockStyle.Fill
        dgvLog.ReadOnly = True
        dgvLog.AllowUserToAddRows = False
        dgvLog.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvLog.BackgroundColor = Color.White
        t4.Controls.Add(dgvLog)

        tc.TabPages.AddRange(New TabPage() {t1, t2, t3, t4})
        Me.Controls.Add(tc)

        LoadAllStock()
        LoadLowStock()
        LoadLog()
    End Sub

    Private Sub LoadAllStock()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter(
                "SELECT p.product_id, p.product_name AS 'Product', c.category_name AS 'Category', " &
                "p.stock_quantity AS 'Stock', p.reorder_level AS 'Min Level', p.status AS 'Status' " &
                "FROM tbl_products p JOIN tbl_categories c ON p.category_id=c.category_id " &
                "WHERE p.status <> 'Discontinued' ORDER BY p.product_name", conn)
            Dim dt As New DataTable() : da.Fill(dt)
            dgvAll.DataSource = dt
            dgvAll.Columns("product_id").Visible = False
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadLowStock()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter(
                "SELECT p.product_name AS 'Product', c.category_name AS 'Category', " &
                "p.stock_quantity AS 'Stock', p.reorder_level AS 'Min Level' " &
                "FROM tbl_products p JOIN tbl_categories c ON p.category_id=c.category_id " &
                "WHERE p.stock_quantity <= p.reorder_level AND p.status = 'In Stock'", conn)
            Dim dt As New DataTable() : da.Fill(dt)
            dgvLow.DataSource = dt
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadLog()

        If Not OpenConnection() Then Return

        Try

            Dim da As New MySqlDataAdapter(
"SELECT
    l.log_date AS 'Date',
    p.product_name AS 'Product',
    l.action_type AS 'Action',
    l.quantity AS 'Qty',
    u.full_name AS 'User',
    l.notes AS 'Notes'
FROM tbl_inventory_log l
INNER JOIN tbl_products p
    ON l.product_id = p.product_id
LEFT JOIN tbl_users u
    ON l.logged_by = u.user_id
ORDER BY l.log_date DESC
LIMIT 100", conn)

            Dim dt As New DataTable

            da.Fill(dt)

            dgvLog.DataSource = dt

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub ApplyAdjustment(sender As Object, e As EventArgs)

        If dgvAll.SelectedRows.Count = 0 Then
            MessageBox.Show("Go to All Stock tab and select a product first.")
            Return
        End If

        If txtAdjQty.Text.Trim() = "" OrElse Not Integer.TryParse(txtAdjQty.Text, Nothing) Then
            MessageBox.Show("Enter a valid quantity.")
            txtAdjQty.Focus()
            Return
        End If

        Dim productID As Integer =
        CInt(dgvAll.SelectedRows(0).Cells("product_id").Value)

        Dim productName As String =
        dgvAll.SelectedRows(0).Cells("Product").Value.ToString()

        Dim qty As Integer = CInt(txtAdjQty.Text)

        Dim action As String = cboAdjType.Text

        Dim notes As String = txtAdjNote.Text.Trim()

        If MessageBox.Show(
        "Apply this inventory adjustment for " &
        productName & "?",
        "Confirm Adjustment",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question) <> DialogResult.Yes Then

            Return

        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try

            Dim currentStock As Integer

            Dim cmdStock As New MySqlCommand(
            "SELECT stock_quantity
             FROM tbl_products
             WHERE product_id=@id", conn, trans)

            cmdStock.Parameters.AddWithValue("@id", productID)

            currentStock = Convert.ToInt32(cmdStock.ExecuteScalar())

            If action = "Deduct Stock" AndAlso qty > currentStock Then

                MessageBox.Show(
                "Cannot deduct more than the available stock." &
                vbCrLf &
                "Current Stock: " & currentStock)

                trans.Rollback()
                CloseConnection()
                Return

            End If

            Dim change As Integer

            If action = "Add Stock" Then
                change = qty
            Else
                change = -qty
            End If

            Dim cmdUpdate As New MySqlCommand(
            "UPDATE tbl_products
             SET stock_quantity = stock_quantity + @change,
                 status =
                 CASE
                     WHEN stock_quantity + @change <= 0
                     THEN 'Sold Out'
                     ELSE 'In Stock'
                 END,
                 updated_at = NOW()
             WHERE product_id=@id", conn, trans)

            cmdUpdate.Parameters.AddWithValue("@change", change)
            cmdUpdate.Parameters.AddWithValue("@id", productID)

            cmdUpdate.ExecuteNonQuery()


            Dim cmdLog As New MySqlCommand(
            "INSERT INTO tbl_inventory_log
            (
                product_id,
                action_type,
                quantity,
                notes,
                logged_by
            )
            VALUES
            (
                @pid,
                'Adjustment',
                @qty,
                @notes,
                @uid
            )", conn, trans)

            cmdLog.Parameters.AddWithValue("@pid", productID)
            cmdLog.Parameters.AddWithValue("@qty", change)
            cmdLog.Parameters.AddWithValue("@notes", notes)
            cmdLog.Parameters.AddWithValue("@uid", LoggedUserID)

            cmdLog.ExecuteNonQuery()

            trans.Commit()

            MessageBox.Show(
            "Inventory adjustment applied successfully.",
            "Success",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

            txtAdjQty.Clear()
            txtAdjNote.Clear()

            cboAdjType.SelectedIndex = 0

            LoadAllStock()
            LoadLowStock()
            LoadLog()

            dgvAll.ClearSelection()

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



End Class