Imports MySql.Data.MySqlClient

Public Class frmProductForm

    Public IsEditMode As Boolean = False
    Public EditProductID As Integer = 0

    Private txtName As New TextBox()
    Private cboCategory As New ComboBox()
    Private cboSupplier As New ComboBox()
    Private txtRegPrice As New TextBox()
    Private txtSalePrice As New TextBox()
    Private txtStock As New TextBox()
    Private txtReorder As New TextBox()
    Private cboStatus As New ComboBox()
    Private txtDesc As New RichTextBox()
    Private txtUnitCost As New TextBox()
    Private pbImage As New PictureBox()
    Private btnBrowse As New Button()
    Private txtImagePath As New TextBox()
    Private ofd As New OpenFileDialog()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = If(IsEditMode, "Edit Product", "Add Product")
        Me.Size = New Size(470, 820)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White
        Me.AutoScroll = True

        Dim y As Integer = 20
        Dim lblTitle As New Label() With {
            .Text = If(IsEditMode, "EDIT PRODUCT", "ADD PRODUCT"),
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(140, y)
        }
        Me.Controls.Add(lblTitle) : y += 45

        AddRow("Product Name", txtName, y) : y += 55
        AddCombo("Category", cboCategory, y) : y += 55
        AddCombo("Supplier", cboSupplier, y) : y += 55
        AddRow("Unit Cost", txtUnitCost, y) : y += 55
        AddRow("Regular Price", txtRegPrice, y) : y += 55
        AddRow("Sale Price", txtSalePrice, y) : y += 55
        AddRow("Stock Quantity", txtStock, y) : y += 55
        AddRow("Reorder Level", txtReorder, y) : y += 55

        Dim lblSt As New Label() With {.Text = "Status:", .AutoSize = True, .Location = New Point(25, y)}
        Me.Controls.Add(lblSt)
        cboStatus.Location = New Point(25, y + 20)
        cboStatus.Size = New Size(380, 28)
        cboStatus.DropDownStyle = ComboBoxStyle.DropDownList
        cboStatus.Items.AddRange(New String() {"In Stock", "Sold Out", "Discontinued"})
        cboStatus.SelectedIndex = 0
        Me.Controls.Add(cboStatus) : y += 55

        Dim lblDesc As New Label() With {.Text = "Description:", .AutoSize = True, .Location = New Point(25, y)}
        Me.Controls.Add(lblDesc)
        txtDesc.Location = New Point(25, y + 20)
        txtDesc.Size = New Size(380, 90)
        Me.Controls.Add(txtDesc) : y += 110

        Dim lblImg As New Label() With {
    .Text = "Product Image:",
    .AutoSize = True,
    .Location = New Point(25, y)
}
        Me.Controls.Add(lblImg)

        y += 25

        pbImage.Location = New Point(25, y)
        pbImage.Size = New Size(150, 150)
        pbImage.BorderStyle = BorderStyle.FixedSingle
        pbImage.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(pbImage)

        txtImagePath.Location = New Point(185, y)
        txtImagePath.Size = New Size(220, 28)
        txtImagePath.ReadOnly = True
        Me.Controls.Add(txtImagePath)

        btnBrowse.Text = "Browse..."
        btnBrowse.Location = New Point(185, y + 40)
        btnBrowse.Size = New Size(100, 35)
        AddHandler btnBrowse.Click, AddressOf BrowseImage
        Me.Controls.Add(btnBrowse)

        y += 170

        Dim btnSave As New Button() With {
            .Text = "SAVE", .Size = New Size(180, 40), .Location = New Point(25, y),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        AddHandler btnSave.Click, AddressOf SaveProduct
        Me.Controls.Add(btnSave)

        Dim btnCancel As New Button() With {
            .Text = "CANCEL", .Size = New Size(180, 40), .Location = New Point(220, y),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnCancel.Click, Sub(s, ev) Me.Close()
        Me.Controls.Add(btnCancel)

        LoadDropdowns()
        If IsEditMode Then LoadProductData()
    End Sub

    Private Sub AddRow(lbl As String, txt As TextBox, y As Integer)
        Dim l As New Label() With {.Text = lbl & ":", .AutoSize = True, .Location = New Point(25, y)}
        Me.Controls.Add(l)
        txt.Size = New Size(380, 28) : txt.Location = New Point(25, y + 20)
        Me.Controls.Add(txt)
    End Sub

    Private Sub AddCombo(lbl As String, cbo As ComboBox, y As Integer)
        Dim l As New Label() With {.Text = lbl & ":", .AutoSize = True, .Location = New Point(25, y)}
        Me.Controls.Add(l)
        cbo.Size = New Size(380, 28) : cbo.Location = New Point(25, y + 20)
        cbo.DropDownStyle = ComboBoxStyle.DropDownList
        Me.Controls.Add(cbo)
    End Sub

    Private Sub LoadDropdowns()
        If Not OpenConnection() Then Return
        Try
            Dim da1 As New MySqlDataAdapter("SELECT category_id, category_name FROM tbl_categories ORDER BY category_name", conn)
            Dim dt1 As New DataTable() : da1.Fill(dt1)
            cboCategory.DataSource = dt1
            cboCategory.DisplayMember = "category_name"
            cboCategory.ValueMember = "category_id"

            Dim da2 As New MySqlDataAdapter("SELECT supplier_id, supplier_name FROM tbl_suppliers ORDER BY supplier_name", conn)
            Dim dt2 As New DataTable() : da2.Fill(dt2)
            cboSupplier.DataSource = dt2
            cboSupplier.DisplayMember = "supplier_name"
            cboSupplier.ValueMember = "supplier_id"
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadProductData()
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand("SELECT * FROM tbl_products WHERE product_id=@id", conn)
            cmd.Parameters.AddWithValue("@id", EditProductID)
            Dim da As New MySqlDataAdapter(cmd)
            Dim dt As New DataTable() : da.Fill(dt)
            If dt.Rows.Count = 0 Then Return
            Dim row As DataRow = dt.Rows(0)
            txtName.Text = row("product_name").ToString()
            txtRegPrice.Text = row("regular_price").ToString()
            txtSalePrice.Text = row("sale_price").ToString()
            txtStock.Text = row("stock_quantity").ToString()
            txtUnitCost.Text = row("unit_cost").ToString()
            txtReorder.Text = row("reorder_level").ToString()
            cboStatus.Text = row("status").ToString()
            cboCategory.SelectedValue = row("category_id")
            txtDesc.Text = row("description").ToString()
            If Not IsDBNull(row("image_path")) Then

                txtImagePath.Text = row("image_path").ToString()

                Dim imgPath As String = IO.Path.Combine(
        Application.StartupPath,
        "Images",
        txtImagePath.Text)

                If IO.File.Exists(imgPath) Then

                    If pbImage.Image IsNot Nothing Then
                        pbImage.Image.Dispose()
                        pbImage.Image = Nothing
                    End If

                    Using bmp As New Bitmap(imgPath)
                        pbImage.Image = New Bitmap(bmp)
                    End Using

                End If

            End If
            If Not IsDBNull(row("supplier_id")) Then cboSupplier.SelectedValue = row("supplier_id")
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub SaveProduct(sender As Object, e As EventArgs)
        If txtName.Text.Trim() = "" _
OrElse txtUnitCost.Text.Trim() = "" _
OrElse txtRegPrice.Text.Trim() = "" Then

            MessageBox.Show("Product Name, Unit Cost, and Regular Price are required.")
            Return
        End If

        Try
            If Not OpenConnection() Then Return
            If IsEditMode Then
                Dim cmd As New MySqlCommand(
                    "UPDATE tbl_products SET product_name=@n, category_id=@c, supplier_id=@s, description=@d, " &
                    "unit_cost=@uc, regular_price=@rp, sale_price=@sp, stock_quantity=@sq, reorder_level=@rl, " &
                    "status=@st, image_path=@img, updated_at=NOW() WHERE product_id=@id", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@c", cboCategory.SelectedValue)
                cmd.Parameters.AddWithValue("@s", cboSupplier.SelectedValue)
                cmd.Parameters.AddWithValue("@d", txtDesc.Text.Trim())
                cmd.Parameters.AddWithValue("@uc", CDec(txtUnitCost.Text))
                cmd.Parameters.AddWithValue("@rp", CDec(txtRegPrice.Text))
                cmd.Parameters.AddWithValue("@sp", If(txtSalePrice.Text = "", 0, CDec(txtSalePrice.Text)))
                cmd.Parameters.AddWithValue("@sq", CInt(txtStock.Text))
                cmd.Parameters.AddWithValue("@rl", If(txtReorder.Text = "", 5, CInt(txtReorder.Text)))
                cmd.Parameters.AddWithValue("@st", cboStatus.Text)
                cmd.Parameters.AddWithValue("@img", txtImagePath.Text)
                cmd.Parameters.AddWithValue("@id", EditProductID)
                cmd.ExecuteNonQuery()
            Else
                Dim cmd As New MySqlCommand(
                    "INSERT INTO tbl_products(product_name, category_id, supplier_id, unit_cost, regular_price, " &
                    "sale_price, stock_quantity, reorder_level, status, description, image_path) " &
                    "VALUES(@n,@c,@s,@uc,@rp,@sp,@sq,@rl,@st,@d,@img)", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@c", cboCategory.SelectedValue)
                cmd.Parameters.AddWithValue("@s", cboSupplier.SelectedValue)
                cmd.Parameters.AddWithValue("@uc", CDec(txtUnitCost.Text))
                cmd.Parameters.AddWithValue("@d", txtDesc.Text.Trim())
                cmd.Parameters.AddWithValue("@rp", CDec(txtRegPrice.Text))
                cmd.Parameters.AddWithValue("@sp", If(txtSalePrice.Text = "", 0, CDec(txtSalePrice.Text)))
                cmd.Parameters.AddWithValue("@sq", CInt(txtStock.Text))
                cmd.Parameters.AddWithValue("@rl", If(txtReorder.Text = "", 5, CInt(txtReorder.Text)))
                cmd.Parameters.AddWithValue("@st", cboStatus.Text)
                cmd.Parameters.AddWithValue("@img", txtImagePath.Text)
                cmd.ExecuteNonQuery()
            End If
            CloseConnection()
            MessageBox.Show("Product saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            CloseConnection()
        End Try
    End Sub

    Private Sub BrowseImage(sender As Object, e As EventArgs)

        ofd.Title = "Select Product Image"
        ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"

        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub

        Dim imagesFolder As String = IO.Path.Combine(Application.StartupPath, "Images")

        If Not IO.Directory.Exists(imagesFolder) Then
            IO.Directory.CreateDirectory(imagesFolder)
        End If

        Dim fileName As String = IO.Path.GetFileName(ofd.FileName)
        Dim destFile As String = IO.Path.Combine(imagesFolder, fileName)


        If pbImage.Image IsNot Nothing Then
            pbImage.Image.Dispose()
            pbImage.Image = Nothing
        End If

        If String.Compare(ofd.FileName, destFile, True) <> 0 Then
            IO.File.Copy(ofd.FileName, destFile, True)
        End If

        Using bmp As New Bitmap(destFile)
            pbImage.Image = New Bitmap(bmp)
        End Using

        txtImagePath.Text = fileName

    End Sub
End Class