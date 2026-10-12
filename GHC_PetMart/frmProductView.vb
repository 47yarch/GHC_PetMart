Imports MySql.Data.MySqlClient

Public Class frmProductView

    Public ProductID As Integer = 0
    Public CustomerUserID As Integer = 0

    Private lblName As New Label()
    Private lblCategory As New Label()
    Private lblPrice As New Label()
    Private lblOrigPrice As New Label()
    Private lblStatus As New Label()
    Private lblStock As New Label()
    Private lblDesc As New RichTextBox()
    Private pbImage As New PictureBox()
    Private lblQtyNum As New Label()
    Private btnMinus As New Button()
    Private btnPlus As New Button()
    Private currentQty As Integer = 1
    Private maxStock As Integer = 0
    Private currentPrice As Decimal = 0

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Product Details"
        Me.Size = New Size(750, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        BuildUI()
        LoadProduct()
    End Sub

    Private Sub BuildUI()

        Dim pLeft As New Panel() With {
            .Location = New Point(15, 15),
            .Size = New Size(280, 460),
            .BackColor = Color.FromArgb(245, 245, 245),
            .BorderStyle = BorderStyle.FixedSingle
        }

        pbImage.Location = New Point(10, 10)
        pbImage.Size = New Size(258, 280)
        pbImage.SizeMode = PictureBoxSizeMode.Zoom
        pbImage.BackColor = Color.FromArgb(240, 240, 240)
        pLeft.Controls.Add(pbImage)

        Dim lblQtyTitle As New Label() With {
            .Text = "Quantity:",
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(10, 305)
        }
        pLeft.Controls.Add(lblQtyTitle)

        btnMinus.Text = "-"
        btnMinus.Size = New Size(40, 38)
        btnMinus.Location = New Point(10, 328)
        btnMinus.FlatStyle = FlatStyle.Flat
        btnMinus.Font = New Font("Arial", 14, FontStyle.Bold)
        btnMinus.BackColor = Color.FromArgb(239, 239, 239)
        AddHandler btnMinus.Click, AddressOf DecreaseQty
        pLeft.Controls.Add(btnMinus)

        lblQtyNum.Text = "1"
        lblQtyNum.Font = New Font("Arial", 16, FontStyle.Bold)
        lblQtyNum.Size = New Size(80, 38)
        lblQtyNum.Location = New Point(55, 328)
        lblQtyNum.TextAlign = ContentAlignment.MiddleCenter
        pLeft.Controls.Add(lblQtyNum)

        btnPlus.Text = "+"
        btnPlus.Size = New Size(40, 38)
        btnPlus.Location = New Point(140, 328)
        btnPlus.FlatStyle = FlatStyle.Flat
        btnPlus.Font = New Font("Arial", 14, FontStyle.Bold)
        btnPlus.BackColor = Color.FromArgb(239, 239, 239)
        AddHandler btnPlus.Click, AddressOf IncreaseQty
        pLeft.Controls.Add(btnPlus)

        lblStock.Text = ""
        lblStock.Font = New Font("Arial", 8)
        lblStock.ForeColor = Color.Gray
        lblStock.AutoSize = True
        lblStock.Location = New Point(10, 375)
        pLeft.Controls.Add(lblStock)

        Dim btnAdd As New Button() With {
            .Text = "ADD TO CART",
            .Size = New Size(258, 45),
            .Location = New Point(10, 400),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 11, FontStyle.Bold)
        }
        btnAdd.FlatAppearance.BorderSize = 0
        AddHandler btnAdd.Click, AddressOf AddToCart
        pLeft.Controls.Add(btnAdd)

        Me.Controls.Add(pLeft)

        Dim pRight As New Panel() With {
            .Location = New Point(310, 15),
            .Size = New Size(410, 460),
            .BackColor = Color.White
        }

        Dim y As Integer = 5

        lblCategory.Text = ""
        lblCategory.Font = New Font("Arial", 8)
        lblCategory.ForeColor = Color.White
        lblCategory.BackColor = Color.FromArgb(63, 81, 181)
        lblCategory.AutoSize = True
        lblCategory.Location = New Point(0, y)
        lblCategory.Padding = New Padding(6, 2, 6, 2)
        pRight.Controls.Add(lblCategory)
        y += 35

        lblName.Text = ""
        lblName.Font = New Font("Arial", 16, FontStyle.Bold)
        lblName.Size = New Size(405, 65)
        lblName.Location = New Point(0, y)
        pRight.Controls.Add(lblName)
        y += 72

        lblPrice.Text = ""
        lblPrice.Font = New Font("Arial", 18, FontStyle.Bold)
        lblPrice.ForeColor = Color.FromArgb(63, 81, 181)
        lblPrice.AutoSize = True
        lblPrice.Location = New Point(0, y)
        pRight.Controls.Add(lblPrice)

        lblOrigPrice.Text = ""
        lblOrigPrice.Font = New Font("Arial", 10, FontStyle.Strikeout)
        lblOrigPrice.ForeColor = Color.Gray
        lblOrigPrice.AutoSize = True
        lblOrigPrice.Location = New Point(0, y + 28)
        pRight.Controls.Add(lblOrigPrice)
        y += 75

        Dim div1 As New Label() With {
            .BorderStyle = BorderStyle.Fixed3D,
            .Size = New Size(405, 2),
            .Location = New Point(0, y)
        }
        pRight.Controls.Add(div1)
        y += 15

        lblStatus.Text = ""
        lblStatus.Font = New Font("Arial", 10, FontStyle.Bold)
        lblStatus.AutoSize = True
        lblStatus.Location = New Point(0, y)
        pRight.Controls.Add(lblStatus)
        y += 35

        Dim div2 As New Label() With {
            .BorderStyle = BorderStyle.Fixed3D,
            .Size = New Size(405, 2),
            .Location = New Point(0, y)
        }
        pRight.Controls.Add(div2)
        y += 15

        Dim lblDescTitle As New Label() With {
            .Text = "Description:",
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(0, y)
        }
        pRight.Controls.Add(lblDescTitle)
        y += 22

        lblDesc.Location = New Point(0, y)
        lblDesc.Size = New Size(405, 130)
        lblDesc.ReadOnly = True
        lblDesc.BorderStyle = BorderStyle.None
        lblDesc.BackColor = Color.White
        lblDesc.Font = New Font("Arial", 9)
        pRight.Controls.Add(lblDesc)

        Me.Controls.Add(pRight)


    End Sub

    Private Sub LoadProduct()
        If ProductID = 0 Then Return

        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand(
                    "SELECT p.product_name, c.category_name, " &
                    "p.regular_price, p.sale_price, " &
                    "p.stock_quantity, p.status, " &
                    "p.description, p.image_path " &
                    "FROM tbl_products p " &
                    "JOIN tbl_categories c ON p.category_id = c.category_id " &
                    "WHERE p.product_id = @id", con)
                cmd.Parameters.AddWithValue("@id", ProductID)

                Dim dt As New DataTable()
                Dim da As New MySqlDataAdapter(cmd)
                da.Fill(dt)

                If dt.Rows.Count = 0 Then Return
                Dim r As DataRow = dt.Rows(0)

                Dim regPrice As Decimal = Convert.ToDecimal(r("regular_price"))
                Dim salePrice As Decimal = Convert.ToDecimal(r("sale_price"))
                Dim status As String = r("status").ToString()
                maxStock = CInt(r("stock_quantity"))
                currentPrice = If(salePrice > 0, salePrice, regPrice)
                Dim isInStock As Boolean = (status = "In Stock" AndAlso maxStock > 0)

                Me.Text = r("product_name").ToString()
                lblName.Text = r("product_name").ToString()
                lblCategory.Text = "  " & r("category_name").ToString() & "  "
                lblPrice.Text = "P " & currentPrice.ToString("N2")
                lblDesc.Text = If(r("description").ToString() = "",
                                  "No description available.",
                                  r("description").ToString())

                If salePrice > 0 AndAlso salePrice < regPrice Then
                    lblOrigPrice.Text = "P " & regPrice.ToString("N2")
                End If

                lblStatus.Text = If(isInStock, "✔ In Stock", "✘ " & status)
                lblStatus.ForeColor = If(isInStock,
                    Color.FromArgb(56, 142, 60),
                    Color.FromArgb(198, 40, 40))

                lblStock.Text = maxStock & " unit(s) available"

                Try
                    Dim dbPath As String = r("image_path").ToString().Trim()

                    If dbPath <> "" Then

                        Dim imgPath As String

                        If dbPath.StartsWith("Images\", StringComparison.OrdinalIgnoreCase) Then
                            imgPath = IO.Path.Combine(Application.StartupPath, dbPath)
                        Else
                            imgPath = IO.Path.Combine(Application.StartupPath, "Images", dbPath)
                        End If

                        If IO.File.Exists(imgPath) Then
                            Using fs As New IO.FileStream(imgPath, IO.FileMode.Open, IO.FileAccess.Read)
                                pbImage.Image = Image.FromStream(fs)
                            End Using
                        End If

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                If Not isInStock Then
                    btnMinus.Enabled = False
                    btnPlus.Enabled = False
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading product: " & ex.Message)
        End Try
    End Sub

    Private Sub IncreaseQty(sender As Object, e As EventArgs)
        If currentQty < maxStock Then
            currentQty += 1
            lblQtyNum.Text = currentQty.ToString()
        Else
            MessageBox.Show("Only " & maxStock & " unit(s) available.",
                            "Stock Limit",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub DecreaseQty(sender As Object, e As EventArgs)
        If currentQty > 1 Then
            currentQty -= 1
            lblQtyNum.Text = currentQty.ToString()
        End If
    End Sub

    Private Sub AddToCart(sender As Object, e As EventArgs)
        If maxStock = 0 Then
            MessageBox.Show("This product is not available.") : Return
        End If

        Try
            Using con As MySqlConnection = GetConn()

                Dim chk As New MySqlCommand(
                    "SELECT IFNULL(quantity,0) FROM tbl_cart " &
                    "WHERE user_id = @uid AND product_id = @pid", con)
                chk.Parameters.AddWithValue("@uid", CustomerUserID)
                chk.Parameters.AddWithValue("@pid", ProductID)
                Dim existingQty As Integer = Convert.ToInt32(chk.ExecuteScalar())

                If existingQty + currentQty > maxStock Then
                    MessageBox.Show(
                        "Cannot add " & currentQty & " more." & vbCrLf &
                        "Already " & existingQty & " in cart." & vbCrLf &
                        "Max stock: " & maxStock,
                        "Stock Limit",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
                    Return
                End If

                If existingQty > 0 Then
                    Dim upd As New MySqlCommand(
                        "UPDATE tbl_cart SET quantity = quantity + @qty " &
                        "WHERE user_id = @uid AND product_id = @pid", con)
                    upd.Parameters.AddWithValue("@qty", currentQty)
                    upd.Parameters.AddWithValue("@uid", CustomerUserID)
                    upd.Parameters.AddWithValue("@pid", ProductID)
                    upd.ExecuteNonQuery()
                Else
                    Dim ins As New MySqlCommand(
                        "INSERT INTO tbl_cart(user_id, product_id, quantity) " &
                        "VALUES (@uid, @pid, @qty)", con)
                    ins.Parameters.AddWithValue("@uid", CustomerUserID)
                    ins.Parameters.AddWithValue("@pid", ProductID)
                    ins.Parameters.AddWithValue("@qty", currentQty)
                    ins.ExecuteNonQuery()
                End If
            End Using

            MessageBox.Show(
                currentQty & "x " & lblName.Text & " added to cart!" & vbCrLf &
                "Subtotal: P " & (currentPrice * currentQty).ToString("N2"),
                "Added to Cart",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)


            currentQty = 1
            lblQtyNum.Text = "1"

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

End Class