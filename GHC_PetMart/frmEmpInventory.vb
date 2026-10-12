Imports MySql.Data.MySqlClient

Public Class frmEmpInventory

    Private dgv As New DataGridView()
    Private txtSearch As New TextBox()
    Private cboFilter As New ComboBox()
    Private lblLow As New Label()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Inventory - View Only"
        Me.Size = New Size(950, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Current Inventory (View Only)",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lS As New Label() With {.Text = "Search:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lS)
        txtSearch.Size = New Size(200, 28) : txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim btnS As New Button() With {
            .Text = "Search", .Size = New Size(80, 28), .Location = New Point(275, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnS.Click, AddressOf LoadStock
        Me.Controls.Add(btnS)

        Dim lF As New Label() With {.Text = "Category:", .AutoSize = True, .Location = New Point(375, 55)}
        Me.Controls.Add(lF)
        cboFilter.Location = New Point(435, 52) : cboFilter.Size = New Size(180, 28)
        cboFilter.DropDownStyle = ComboBoxStyle.DropDownList
        AddHandler cboFilter.SelectedIndexChanged, AddressOf LoadStock
        Me.Controls.Add(cboFilter)

        lblLow.AutoSize = True
        lblLow.ForeColor = Color.Red
        lblLow.Font = New Font("Arial", 9, FontStyle.Bold)
        lblLow.Location = New Point(640, 57)
        Me.Controls.Add(lblLow)

        dgv.Location = New Point(15, 95)
        dgv.Size = New Size(910, 440)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)

        LoadCategories()
        LoadStock(Nothing, Nothing)
    End Sub

    Private Sub LoadCategories()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter("SELECT category_name FROM tbl_categories ORDER BY category_name", conn)
            Dim dt As New DataTable() : da.Fill(dt)
            cboFilter.Items.Clear()
            cboFilter.Items.Add("All Categories")
            For Each row As DataRow In dt.Rows
                cboFilter.Items.Add(row("category_name").ToString())
            Next
            cboFilter.SelectedIndex = 0
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadStock(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim search As String = "%" & txtSearch.Text.Trim() & "%"
            Dim sql As String =
                "SELECT p.product_name AS 'Product', c.category_name AS 'Category', " &
                "p.regular_price AS 'Price', p.stock_quantity AS 'Stock', " &
                "p.reorder_level AS 'Min Level', p.status AS 'Status' " &
                "FROM tbl_products p JOIN tbl_categories c ON p.category_id=c.category_id " &
                "WHERE p.product_name LIKE @s AND p.status <> 'Discontinued'"
            If cboFilter.SelectedIndex > 0 Then
                sql &= " AND c.category_name = '" & cboFilter.Text & "'"
            End If
            sql &= " ORDER BY p.product_name"

            Dim cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@s", search)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgv.DataSource = dt

            For Each row As DataGridViewRow In dgv.Rows
                If CInt(row.Cells("Stock").Value) <= CInt(row.Cells("Min Level").Value) Then
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 238)
                End If
            Next

            Dim lowCmd As New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_products WHERE stock_quantity <= reorder_level AND status='In Stock'", conn)
            lblLow.Text = "Low Stock: " & lowCmd.ExecuteScalar().ToString() & " items"
        Finally
            CloseConnection()
        End Try
    End Sub

End Class