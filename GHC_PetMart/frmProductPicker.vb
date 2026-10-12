Imports MySql.Data.MySqlClient

Public Class frmProductPicker

    Public ProductID As Integer = 0
    Public ProductName As String = ""
    Public RegularPrice As Decimal = 0D
    Public SalePrice As Decimal = 0D

    Private txtSearch As New TextBox()
    Private dgvProducts As New DataGridView()
    Private btnSelect As New Button()
    Private btnCancel As New Button()

    Private Sub frmProductPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Select Product"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Size = New Size(850, 520)
        Me.BackColor = Color.WhiteSmoke

        Dim lblSearch As New Label With {
            .Text = "Search:",
            .Location = New Point(15, 18),
            .AutoSize = True
        }

        Me.Controls.Add(lblSearch)

        txtSearch.Location = New Point(75, 15)
        txtSearch.Size = New Size(260, 28)
        AddHandler txtSearch.TextChanged, AddressOf SearchProducts
        Me.Controls.Add(txtSearch)

        dgvProducts.Location = New Point(15, 55)
        dgvProducts.Size = New Size(800, 370)
        dgvProducts.AllowUserToAddRows = False
        dgvProducts.AllowUserToDeleteRows = False
        dgvProducts.AllowUserToResizeRows = False
        dgvProducts.ReadOnly = True
        dgvProducts.MultiSelect = False
        dgvProducts.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvProducts.BackgroundColor = Color.White

        AddHandler dgvProducts.CellDoubleClick, AddressOf SelectProduct

        Me.Controls.Add(dgvProducts)

        btnSelect.Text = "Select"
        btnSelect.Size = New Size(120, 35)
        btnSelect.Location = New Point(560, 440)
        btnSelect.BackColor = Color.SeaGreen
        btnSelect.ForeColor = Color.White
        btnSelect.FlatStyle = FlatStyle.Flat
        AddHandler btnSelect.Click, AddressOf SelectProduct
        Me.Controls.Add(btnSelect)

        btnCancel.Text = "Cancel"
        btnCancel.Size = New Size(120, 35)
        btnCancel.Location = New Point(695, 440)

        AddHandler btnCancel.Click,
            Sub()
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            End Sub

        Me.Controls.Add(btnCancel)

        LoadProducts()

    End Sub

    Private Sub LoadProducts()

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
"SELECT
    product_id,
    product_name AS 'Product',
    regular_price AS 'Regular Price',
    sale_price AS 'Sale Price',
    stock_quantity AS 'Stock'
FROM tbl_products
WHERE stock_quantity > 0
ORDER BY product_name", conn)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)

            da.Fill(dt)

            dgvProducts.DataSource = dt

            dgvProducts.Columns("product_id").Visible = False

            dgvProducts.Columns("Regular Price").DefaultCellStyle.Format = "N2"
            dgvProducts.Columns("Sale Price").DefaultCellStyle.Format = "N2"

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub SearchProducts(sender As Object, e As EventArgs)

        If Not OpenConnection() Then Return

        Try

            Dim cmd As New MySqlCommand(
"SELECT
    product_id,
    product_name AS 'Product',
    regular_price AS 'Regular Price',
    sale_price AS 'Sale Price',
    stock_quantity AS 'Stock'
FROM tbl_products
WHERE stock_quantity > 0
AND product_name LIKE @search
ORDER BY product_name", conn)

            cmd.Parameters.AddWithValue("@search", "%" & txtSearch.Text.Trim() & "%")

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)

            da.Fill(dt)

            dgvProducts.DataSource = dt

            dgvProducts.Columns("product_id").Visible = False

            dgvProducts.Columns("Regular Price").DefaultCellStyle.Format = "N2"
            dgvProducts.Columns("Sale Price").DefaultCellStyle.Format = "N2"

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub SelectProduct(sender As Object, e As EventArgs)

        If dgvProducts.SelectedRows.Count = 0 Then

            MessageBox.Show("Please select a product.")

            Exit Sub

        End If

        ProductID = CInt(dgvProducts.SelectedRows(0).Cells("product_id").Value)

        ProductName = dgvProducts.SelectedRows(0).Cells("Product").Value.ToString()

        RegularPrice = CDec(dgvProducts.SelectedRows(0).Cells("Regular Price").Value)

        SalePrice = CDec(dgvProducts.SelectedRows(0).Cells("Sale Price").Value)

        Me.DialogResult = DialogResult.OK

        Me.Close()

    End Sub


End Class