Imports MySql.Data.MySqlClient

Public Class frmProducts

    Private dgv As New DataGridView()
    Private txtSearch As New TextBox()
    Private cboFilter As New ComboBox()
    Private cboSort As New ComboBox()
    Private lblCount As New Label()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Products"
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Product Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lblS As New Label() With {.Text = "Search:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lblS)
        txtSearch.Size = New Size(200, 28)
        txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim btnSearch As New Button() With {
            .Text = "Search", .Size = New Size(75, 28), .Location = New Point(275, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSearch.Click, AddressOf LoadProducts
        Me.Controls.Add(btnSearch)

        Dim lblF As New Label() With {.Text = "Filter:", .AutoSize = True, .Location = New Point(370, 55)}
        Me.Controls.Add(lblF)
        cboFilter.Size = New Size(160, 28)
        cboFilter.Location = New Point(415, 52)
        cboFilter.DropDownStyle = ComboBoxStyle.DropDownList
        AddHandler cboFilter.SelectedIndexChanged, AddressOf LoadProducts
        Me.Controls.Add(cboFilter)

        Dim lblSo As New Label() With {.Text = "Sort:", .AutoSize = True, .Location = New Point(590, 55)}
        Me.Controls.Add(lblSo)
        cboSort.Size = New Size(160, 28)
        cboSort.Location = New Point(625, 52)
        cboSort.DropDownStyle = ComboBoxStyle.DropDownList
        cboSort.Items.AddRange(New String() {"Name A-Z", "Name Z-A", "Price Low-High", "Price High-Low", "Stock Low-High"})
        cboSort.SelectedIndex = 0
        AddHandler cboSort.SelectedIndexChanged, AddressOf LoadProducts
        Me.Controls.Add(cboSort)

        lblCount.AutoSize = True
        lblCount.Location = New Point(800, 57)
        lblCount.Font = New Font("Arial", 9, FontStyle.Bold)
        Me.Controls.Add(lblCount)

        Dim btnAdd As New Button() With {
            .Text = "Add Product", .Size = New Size(110, 32), .Location = New Point(15, 92),
            .BackColor = Color.FromArgb(76, 175, 80), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdd.Click, AddressOf AddProduct
        Me.Controls.Add(btnAdd)

        Dim btnEdit As New Button() With {
            .Text = "Edit", .Size = New Size(80, 32), .Location = New Point(135, 92),
            .BackColor = Color.FromArgb(255, 152, 0), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnEdit.Click, AddressOf EditProduct
        Me.Controls.Add(btnEdit)

        Dim btnDel As New Button() With {
            .Text = "Delete", .Size = New Size(100, 32), .Location = New Point(225, 92),
            .BackColor = Color.FromArgb(244, 67, 54), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDel.Click, AddressOf DeleteProduct
        Me.Controls.Add(btnDel)

        Dim btnRef As New Button() With {
            .Text = "Refresh", .Size = New Size(80, 32), .Location = New Point(335, 92),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnRef.Click, AddressOf LoadProducts
        Me.Controls.Add(btnRef)

        dgv.Location = New Point(15, 135)
        dgv.Size = New Size(955, 420)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)

        LoadCategories()
        LoadProducts(Nothing, Nothing)
    End Sub

    Private Sub LoadCategories()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter("SELECT category_id, category_name FROM tbl_categories ORDER BY category_name", conn)
            Dim dt As New DataTable()
            da.Fill(dt)
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

    Private Sub LoadProducts(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim search As String = "%" & txtSearch.Text.Trim() & "%"
            Dim catFilter As String = If(cboFilter.SelectedIndex <= 0, "", cboFilter.Text)

            Dim orderBy As String = "p.product_name ASC"
            Select Case cboSort.Text
                Case "Name Z-A" : orderBy = "p.product_name DESC"
                Case "Price Low-High" : orderBy = "p.regular_price ASC"
                Case "Price High-Low" : orderBy = "p.regular_price DESC"
                Case "Stock Low-High" : orderBy = "p.stock_quantity ASC"
            End Select

            Dim sql As String =
                "SELECT p.product_id, p.product_name AS 'Product Name', " &
                "c.category_name AS 'Category', " &
                "p.regular_price AS 'Regular Price', p.sale_price AS 'Sale Price', " &
                "p.stock_quantity AS 'Stock', p.status AS 'Status' " &
                "FROM tbl_products p JOIN tbl_categories c ON p.category_id = c.category_id " &
                "WHERE p.product_name LIKE @s AND p.status <> 'Discontinued'"

            If catFilter <> "" Then sql &= " AND c.category_name = @cat"
            sql &= " ORDER BY " & orderBy

            Dim cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@s", search)
            If catFilter <> "" Then cmd.Parameters.AddWithValue("@cat", catFilter)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgv.DataSource = dt
            dgv.Columns("product_id").Visible = False
            lblCount.Text = "Total: " & dt.Rows.Count & " products"
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub AddProduct(sender As Object, e As EventArgs)
        Dim f As New frmProductForm()
        f.IsEditMode = False
        f.ShowDialog()
        LoadProducts(Nothing, Nothing)
    End Sub

    Private Sub EditProduct(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a product to edit.")
            Return
        End If
        Dim f As New frmProductForm()
        f.IsEditMode = True
        f.EditProductID = CInt(dgv.SelectedRows(0).Cells("product_id").Value)
        f.ShowDialog()
        LoadProducts(Nothing, Nothing)
    End Sub

    Private Sub DeleteProduct(sender As Object, e As EventArgs)

        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a product.")
            Return
        End If

        Dim id As Integer = CInt(dgv.SelectedRows(0).Cells("product_id").Value)
        Dim name As String = dgv.SelectedRows(0).Cells("Product Name").Value.ToString()

        If MessageBox.Show(
        "Are you sure you want to permanently delete '" & name & "'?" &
        vbCrLf & vbCrLf &
        "This action cannot be undone.",
        "Delete Product",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning) = DialogResult.No Then Exit Sub

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
            "DELETE FROM tbl_products WHERE product_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", id)

            cmd.ExecuteNonQuery()

            MessageBox.Show(
            "Product deleted successfully.",
            "Deleted",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show(
            "Cannot delete this product." &
            vbCrLf &
            "It may already exist in orders or carts." &
            vbCrLf & vbCrLf &
            ex.Message)

        Finally

            CloseConnection()

        End Try

        LoadProducts(Nothing, Nothing)

    End Sub

End Class