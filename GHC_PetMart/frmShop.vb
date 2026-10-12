Imports MySql.Data.MySqlClient

Public Class frmShop

    Private flpProducts As New FlowLayoutPanel()
    Private txtSearch As New TextBox()
    Private cboFilter As New ComboBox()
    Private cboSort As New ComboBox()
    Private lblCount As New Label()
    Private lblCartCount As New Label()

    Private pnlHeader As New Panel()
    Private pnlFilter As New Panel()
    Private pnlScroll As New Panel()
    Private pnlFooter As New Panel()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "GHC Pet Mart - Shop"
        Me.Size = New Size(1150, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.WhiteSmoke

        AddHandler Me.Resize, AddressOf OnResize

        BuildHeader()
        BuildFilterBar()
        BuildProductArea()
        LoadCategories()
        LoadProducts()
        RefreshCartCount()
        BuildFooter()
    End Sub

    Private Sub OnResize(sender As Object, e As EventArgs)

        pnlHeader.Width = Me.ClientSize.Width

        pnlFilter.Width = Me.ClientSize.Width

        pnlFooter.Width = Me.ClientSize.Width
        pnlFooter.Top = Me.ClientSize.Height - pnlFooter.Height

        pnlScroll.Width = Me.ClientSize.Width
        pnlScroll.Height =
        Me.ClientSize.Height -
        pnlHeader.Height -
        pnlFilter.Height -
        pnlFooter.Height

    End Sub

    Private Sub BuildHeader()
        pnlHeader.Size = New Size(Me.ClientSize.Width, 65)
        pnlHeader.Location = New Point(0, 0)
        pnlHeader.BackColor = Color.FromArgb(40, 40, 80)
        pnlHeader.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Dim lblLogo As New Label() With {
            .Text = "GHC PET MART",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(20, 18)
        }
        pnlHeader.Controls.Add(lblLogo)

        Dim lblWelcome As New Label() With {
            .Text = "Hello, " & LoggedFullName,
            .Font = New Font("Arial", 9),
            .ForeColor = Color.LightGray,
            .AutoSize = True,
            .Location = New Point(300, 24)
        }
        pnlHeader.Controls.Add(lblWelcome)



        Dim btnCart As New Button() With {
            .Text = "",
            .Size = New Size(130, 38),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        lblCartCount = New Label() With {
            .Text = "Cart (0)",
            .AutoSize = True,
            .ForeColor = Color.White,
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .Location = New Point(10, 10)
        }
        btnCart.Controls.Add(lblCartCount)
        AddHandler btnCart.Click, Sub(s, e)
                                      Dim f As New frmCart()
                                      f.CustomerUserID = LoggedUserID
                                      f.ShowDialog()
                                      RefreshCartCount()
                                  End Sub

        Dim btnOrders As New Button() With {
            .Text = "My Orders",
            .Size = New Size(110, 38),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        AddHandler btnOrders.Click, Sub(s, e)
                                        Dim f As New frmOrderHistory()
                                        f.CustomerUserID = LoggedUserID
                                        f.ShowDialog()
                                    End Sub

        Dim btnProfile As New Button() With {
    .Text = "My Profile",
    .Size = New Size(110, 38),
    .BackColor = Color.FromArgb(0, 150, 136),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat,
    .Anchor = AnchorStyles.Top Or AnchorStyles.Right
}

        AddHandler btnProfile.Click,
            Sub(s, e)
                Dim f As New frmMyProfile()
                f.ShowDialog()
            End Sub

        Dim btnLogout As New Button() With {
            .Text = "Logout",
            .Size = New Size(85, 38),
            .BackColor = Color.FromArgb(244, 67, 54),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right
        }
        AddHandler btnLogout.Click, Sub(s, e)
                                        Dim f As New frmLogin()
                                        f.Show()
                                        Me.Close()
                                    End Sub

        AddHandler pnlHeader.Resize, Sub(s, e)

                                         btnLogout.Location = New Point(pnlHeader.Width - 100, 13)

                                         btnProfile.Location = New Point(pnlHeader.Width - 225, 13)

                                         btnOrders.Location = New Point(pnlHeader.Width - 350, 13)

                                         btnCart.Location = New Point(pnlHeader.Width - 495, 13)

                                     End Sub

        pnlHeader.Controls.Add(btnCart)
        pnlHeader.Controls.Add(btnOrders)
        pnlHeader.Controls.Add(btnProfile)
        pnlHeader.Controls.Add(btnLogout)
        Me.Controls.Add(pnlHeader)
    End Sub

    Private Sub BuildFilterBar()
        pnlFilter.Size = New Size(Me.ClientSize.Width, 50)
        pnlFilter.Location = New Point(0, pnlHeader.Height)
        pnlFilter.BackColor = Color.White
        pnlFilter.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Dim lS As New Label() With {
            .Text = "Search:",
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        pnlFilter.Controls.Add(lS)

        txtSearch.Size = New Size(220, 28)
        txtSearch.Location = New Point(68, 12)
        pnlFilter.Controls.Add(txtSearch)

        Dim btnS As New Button() With {
            .Text = "Search",
            .Size = New Size(80, 28),
            .Location = New Point(298, 12),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnS.Click, Sub(s, e) LoadProducts()
        pnlFilter.Controls.Add(btnS)

        Dim lF As New Label() With {
            .Text = "Filter by:",
            .AutoSize = True,
            .Location = New Point(400, 15)
        }
        pnlFilter.Controls.Add(lF)

        cboFilter.Location = New Point(465, 12)
        cboFilter.Size = New Size(180, 28)
        cboFilter.DropDownStyle = ComboBoxStyle.DropDownList
        AddHandler cboFilter.SelectedIndexChanged, Sub(s, e) LoadProducts()
        pnlFilter.Controls.Add(cboFilter)

        Dim lSo As New Label() With {
            .Text = "Sort by:",
            .AutoSize = True,
            .Location = New Point(660, 15)
        }
        pnlFilter.Controls.Add(lSo)

        cboSort.Location = New Point(715, 12)
        cboSort.Size = New Size(180, 28)
        cboSort.DropDownStyle = ComboBoxStyle.DropDownList
        cboSort.Items.AddRange(New String() {"Name A-Z", "Price Low-High", "Price High-Low"})
        cboSort.SelectedIndex = 0
        AddHandler cboSort.SelectedIndexChanged, Sub(s, e) LoadProducts()
        pnlFilter.Controls.Add(cboSort)

        lblCount.AutoSize = True
        lblCount.Font = New Font("Arial", 9, FontStyle.Bold)
        lblCount.Location = New Point(915, 17)
        pnlFilter.Controls.Add(lblCount)

        Me.Controls.Add(pnlFilter)
    End Sub

    Private Sub BuildProductArea()

        pnlScroll.Location = New Point(0, pnlHeader.Height + pnlFilter.Height)
        pnlScroll.Size = New Size(
Me.ClientSize.Width,
Me.ClientSize.Height -
pnlHeader.Height -
pnlFilter.Height -
130)

        pnlScroll.Anchor = AnchorStyles.Top Or
                       AnchorStyles.Bottom Or
                       AnchorStyles.Left Or
                       AnchorStyles.Right

        pnlScroll.BackColor = Color.WhiteSmoke

        pnlScroll.AutoScroll = False

        flpProducts.Location = New Point(0, 0)
        flpProducts.Dock = DockStyle.Fill

        flpProducts.AutoScroll = True
        flpProducts.WrapContents = True
        flpProducts.FlowDirection = FlowDirection.LeftToRight

        flpProducts.Margin = New Padding(10)
        flpProducts.Padding = New Padding(10)

        pnlScroll.Controls.Add(flpProducts)
        Me.Controls.Add(pnlScroll)

    End Sub


    Private Sub LoadCategories()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter(
                "SELECT category_name FROM tbl_categories ORDER BY category_name", conn)
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

    Public Sub LoadProducts()
        flpProducts.Controls.Clear()

        If Not OpenConnection() Then Return
        Try
            Dim search As String = "%" & txtSearch.Text.Trim() & "%"
            Dim orderBy As String = "p.product_name ASC"
            Select Case cboSort.Text
                Case "Price Low-High" : orderBy = "p.regular_price ASC"
                Case "Price High-Low" : orderBy = "p.regular_price DESC"
            End Select

            Dim sql As String =
                "SELECT p.product_id, p.product_name, c.category_name, " &
                "p.regular_price, p.sale_price, p.stock_quantity, p.status, p.image_path " &
                "FROM tbl_products p " &
                "JOIN tbl_categories c ON p.category_id = c.category_id " &
                "WHERE p.product_name LIKE @s AND p.status <> 'Discontinued'"

            If cboFilter.SelectedIndex > 0 Then
                sql &= " AND c.category_name = '" & cboFilter.Text.Replace("'", "''") & "'"
            End If
            sql &= " ORDER BY " & orderBy

            Dim cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@s", search)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            lblCount.Text = "No. of Products: " & dt.Rows.Count
            flpProducts.SuspendLayout()

            For Each row As DataRow In dt.Rows
                flpProducts.Controls.Add(CreateProductCard(row))
            Next

            flpProducts.ResumeLayout()
            flpProducts.PerformLayout()

        Finally
            CloseConnection()
        End Try
    End Sub

    Private Function CreateProductCard(row As DataRow) As Panel
        Dim pid As Integer = CInt(row("product_id"))
        Dim pname As String = row("product_name").ToString()
        Dim category As String = row("category_name").ToString()
        Dim regPrice As Decimal = Convert.ToDecimal(row("regular_price"))
        Dim salePrice As Decimal = Convert.ToDecimal(row("sale_price"))
        Dim stock As Integer = CInt(row("stock_quantity"))
        Dim status As String = row("status").ToString()
        Dim isInStock As Boolean = (status = "In Stock" AndAlso stock > 0)
        Dim displayPrice As Decimal = If(salePrice > 0, salePrice, regPrice)

        Dim card As New Panel() With {
            .Size = New Size(220, 355),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Margin = New Padding(8),
            .Tag = pid,
            .Cursor = Cursors.Hand
        }

        Dim pic As New PictureBox() With {
    .Size = New Size(200, 140),
    .Location = New Point(10, 10),
    .SizeMode = PictureBoxSizeMode.Zoom,
    .BackColor = Color.White,
    .BorderStyle = BorderStyle.FixedSingle,
    .Cursor = Cursors.Hand,
    .Tag = pid
}
        Try
            Dim dbPath As String = row("image_path").ToString().Trim()

            If dbPath <> "" Then

                Dim imgPath As String

                If dbPath.StartsWith("Images\", StringComparison.OrdinalIgnoreCase) Then
                    imgPath = IO.Path.Combine(Application.StartupPath, dbPath)
                Else
                    imgPath = IO.Path.Combine(Application.StartupPath, "Images", dbPath)
                End If

                If IO.File.Exists(imgPath) Then
                    Using fs As New IO.FileStream(imgPath, IO.FileMode.Open, IO.FileAccess.Read)
                        pic.Image = Image.FromStream(fs)
                    End Using
                End If

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        AddHandler pic.Click, AddressOf OpenProductView
        card.Controls.Add(pic)

        Dim lblCat As New Label() With {
            .Text = category,
            .Font = New Font("Arial", 7),
            .ForeColor = Color.White,
            .BackColor = Color.FromArgb(63, 81, 181),
            .AutoSize = True,
            .Location = New Point(10, 155),
            .Padding = New Padding(3, 1, 3, 1)
        }
        card.Controls.Add(lblCat)

        Dim lblName As New Label() With {
            .Text = pname,
            .Font = New Font("Arial", 8, FontStyle.Bold),
            .Size = New Size(200, 38),
            .Location = New Point(10, 175),
            .TextAlign = ContentAlignment.TopLeft,
            .Cursor = Cursors.Hand,
            .Tag = pid
        }
        AddHandler lblName.Click, AddressOf OpenProductView
        card.Controls.Add(lblName)

        Dim lblPrice As New Label() With {
            .Text = "P " & displayPrice.ToString("N2"),
            .Font = New Font("Arial", 12, FontStyle.Bold),
            .ForeColor = Color.FromArgb(63, 81, 181),
            .AutoSize = True,
            .Location = New Point(10, 215)
        }
        card.Controls.Add(lblPrice)

        If salePrice > 0 AndAlso salePrice < regPrice Then
            Dim lblOrig As New Label() With {
                .Text = "P " & regPrice.ToString("N2"),
                .Font = New Font("Arial", 7, FontStyle.Strikeout),
                .ForeColor = Color.Gray,
                .AutoSize = True,
                .Location = New Point(130, 225)
            }
            card.Controls.Add(lblOrig)
        End If

        Dim lblStatus As New Label() With {
            .Text = If(isInStock, "In Stock (" & stock & ")", status),
            .Font = New Font("Arial", 7, FontStyle.Bold),
            .ForeColor = If(isInStock,
                Color.FromArgb(56, 142, 60),
                Color.FromArgb(198, 40, 40)),
            .AutoSize = True,
            .Location = New Point(10, 238)
        }
        card.Controls.Add(lblStatus)

        Dim lblQtyText As New Label() With {
            .Text = "Qty:",
            .Font = New Font("Arial", 8, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(10, 262)
        }
        card.Controls.Add(lblQtyText)

        Dim btnMinus As New Button() With {
            .Text = "-",
            .Size = New Size(28, 26),
            .Location = New Point(45, 259),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 11, FontStyle.Bold),
            .BackColor = Color.FromArgb(239, 239, 239),
            .Enabled = isInStock
        }
        card.Controls.Add(btnMinus)

        Dim txtQty As New TextBox() With {
            .Text = "1",
            .Size = New Size(38, 26),
            .Location = New Point(76, 259),
            .TextAlign = HorizontalAlignment.Center,
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .ReadOnly = True,
            .BackColor = Color.White,
            .Enabled = isInStock
        }
        card.Controls.Add(txtQty)

        Dim btnPlus As New Button() With {
            .Text = "+",
            .Size = New Size(28, 26),
            .Location = New Point(117, 259),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 11, FontStyle.Bold),
            .BackColor = Color.FromArgb(239, 239, 239),
            .Enabled = isInStock
        }
        card.Controls.Add(btnPlus)

        Dim lblMax As New Label() With {
            .Text = "/ " & stock,
            .Font = New Font("Arial", 7),
            .ForeColor = Color.Gray,
            .AutoSize = True,
            .Location = New Point(150, 264)
        }
        card.Controls.Add(lblMax)

        AddHandler btnMinus.Click, Sub(s, e)
                                       Dim v As Integer = CInt(txtQty.Text)
                                       If v > 1 Then txtQty.Text = (v - 1).ToString()
                                   End Sub

        AddHandler btnPlus.Click, Sub(s, e)
                                      Dim v As Integer = CInt(txtQty.Text)
                                      If v < stock Then
                                          txtQty.Text = (v + 1).ToString()
                                      Else
                                          MessageBox.Show("Only " & stock & " unit(s) available.")
                                      End If
                                  End Sub

        Dim btnView As New Button() With {
            .Text = "View Details",
            .Size = New Size(198, 26),
            .Location = New Point(11, 290),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 8),
            .BackColor = Color.FromArgb(230, 230, 230),
            .ForeColor = Color.FromArgb(50, 50, 50),
            .Tag = pid
        }
        AddHandler btnView.Click, AddressOf OpenProductView
        card.Controls.Add(btnView)

        Dim btnAdd As New Button() With {
            .Text = If(isInStock, "Add to Cart", "Sold Out"),
            .Size = New Size(198, 32),
            .Location = New Point(11, 320),
            .BackColor = If(isInStock,
                Color.FromArgb(76, 175, 80),
                Color.FromArgb(200, 200, 200)),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .Enabled = isInStock,
            .Tag = pid
        }
        btnAdd.FlatAppearance.BorderSize = 0
        AddHandler btnAdd.Click, Sub(s, e)
                                     AddToCartWithQty(pid, pname, CInt(txtQty.Text), stock)
                                     txtQty.Text = "1"
                                 End Sub
        card.Controls.Add(btnAdd)

        If Not isInStock Then
            card.BackColor = Color.FromArgb(250, 250, 250)
        End If

        Return card
    End Function

    Private Sub AddToCartWithQty(productID As Integer, productName As String,
                              qty As Integer, stock As Integer)
        If Not OpenConnection() Then Return
        Try
            Dim chkQty As New MySqlCommand(
            "SELECT IFNULL(quantity,0) FROM tbl_cart " &
            "WHERE user_id=@uid AND product_id=@pid", conn)
            chkQty.Parameters.AddWithValue("@uid", LoggedUserID)
            chkQty.Parameters.AddWithValue("@pid", productID)
            Dim existingQty As Integer = Convert.ToInt32(chkQty.ExecuteScalar())

            If existingQty + qty > stock Then
                MessageBox.Show(
                "Cannot add " & qty & " more." & vbCrLf &
                "Already " & existingQty & " in cart." & vbCrLf &
                "Available stock: " & stock,
                "Stock Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                CloseConnection()
                Return
            End If

            If existingQty > 0 Then
                Dim upd As New MySqlCommand(
                "UPDATE tbl_cart SET quantity = quantity + @qty " &
                "WHERE user_id=@uid AND product_id=@pid", conn)
                upd.Parameters.AddWithValue("@qty", qty)
                upd.Parameters.AddWithValue("@uid", LoggedUserID)
                upd.Parameters.AddWithValue("@pid", productID)
                upd.ExecuteNonQuery()
            Else
                Dim ins As New MySqlCommand(
                "INSERT INTO tbl_cart(user_id, product_id, quantity) " &
                "VALUES(@uid, @pid, @qty)", conn)
                ins.Parameters.AddWithValue("@uid", LoggedUserID)
                ins.Parameters.AddWithValue("@pid", productID)
                ins.Parameters.AddWithValue("@qty", qty)
                ins.ExecuteNonQuery()
            End If

            CloseConnection()
            RefreshCartCount()
            MessageBox.Show(qty & "x " & productName & " added to cart!",
                        "Added to Cart",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            CloseConnection()
        End Try
    End Sub

    Private Sub AddToCartClick(sender As Object, e As EventArgs)
        Dim btn As Button = DirectCast(sender, Button)
        Dim productID As Integer = CInt(btn.Tag)

        If Not OpenConnection() Then Return
        Try
            Dim chk As New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_cart WHERE user_id=@uid AND product_id=@pid", conn)
            chk.Parameters.AddWithValue("@uid", LoggedUserID)
            chk.Parameters.AddWithValue("@pid", productID)

            If Convert.ToInt32(chk.ExecuteScalar()) > 0 Then
                Dim upd As New MySqlCommand(
                    "UPDATE tbl_cart SET quantity=quantity+1 " &
                    "WHERE user_id=@uid AND product_id=@pid", conn)
                upd.Parameters.AddWithValue("@uid", LoggedUserID)
                upd.Parameters.AddWithValue("@pid", productID)
                upd.ExecuteNonQuery()
            Else
                Dim ins As New MySqlCommand(
                    "INSERT INTO tbl_cart(user_id, product_id, quantity) VALUES(@uid,@pid,1)", conn)
                ins.Parameters.AddWithValue("@uid", LoggedUserID)
                ins.Parameters.AddWithValue("@pid", productID)
                ins.ExecuteNonQuery()
            End If

            CloseConnection()
            RefreshCartCount()
            MessageBox.Show("Added to cart!", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            CloseConnection()
        End Try
    End Sub

    Private Sub RefreshCartCount()
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand(
                "SELECT IFNULL(SUM(quantity),0) FROM tbl_cart WHERE user_id=@uid", conn)
            cmd.Parameters.AddWithValue("@uid", LoggedUserID)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            lblCartCount.Text = "Cart (" & count & ")"
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub OpenProductView(sender As Object, e As EventArgs)
        Dim ctrl As Control = DirectCast(sender, Control)
        Dim pid As Integer = CInt(ctrl.Tag)

        Dim f As New frmProductView()
        f.ProductID = pid
        f.CustomerUserID = LoggedUserID
        f.ShowDialog()


        RefreshCartCount()
        LoadProducts()
    End Sub

    Public Sub RefreshShop()

        LoadProducts()

        RefreshCartCount()

    End Sub

    Private Sub BuildFooter()

        pnlFooter.Height = 130
        pnlFooter.Dock = DockStyle.Bottom
        pnlFooter.BackColor = Color.FromArgb(35, 35, 35)


        Dim picLogo As New PictureBox()

        picLogo.Size = New Size(70, 70)
        picLogo.Location = New Point(25, 28)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom

        Try

            Dim logoPath As String =
                IO.Path.Combine(
                Application.StartupPath,
                "Images",
                "logo.png")

            If IO.File.Exists(logoPath) Then

                picLogo.Image = Image.FromFile(logoPath)

            End If

        Catch
        End Try

        pnlFooter.Controls.Add(picLogo)



        Dim lblTitle As New Label()

        lblTitle.Text = "GHC PET MART"
        lblTitle.Font = New Font("Arial", 14, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(110, 18)

        pnlFooter.Controls.Add(lblTitle)


        Dim lblTag As New Label()

        lblTag.Text = "One Home for Fur, Feathers, and Fins."
        lblTag.Font = New Font("Arial", 9, FontStyle.Italic)
        lblTag.ForeColor = Color.Gainsboro
        lblTag.AutoSize = True
        lblTag.Location = New Point(110, 43)

        pnlFooter.Controls.Add(lblTag)



        Dim lblAddress As New Label()

        lblAddress.Text =
    "Address: Km. 14, East Service Road, South Super Highway," &
    vbCrLf &
    "Western Bicutan, Taguig, Metro Manila"

        lblAddress.ForeColor = Color.WhiteSmoke
        lblAddress.AutoSize = True
        lblAddress.Location = New Point(110, 66)

        pnlFooter.Controls.Add(lblAddress)


        Dim lblPhone As New Label()

        lblPhone.Text =
    "Contact: 091234556789"

        lblPhone.ForeColor = Color.WhiteSmoke
        lblPhone.AutoSize = True
        lblPhone.Location = New Point(520, 20)

        pnlFooter.Controls.Add(lblPhone)


        Dim lnkFacebook As New LinkLabel()

        lnkFacebook.Text = "Facebook"

        lnkFacebook.LinkColor = Color.DeepSkyBlue

        lnkFacebook.AutoSize = True

        lnkFacebook.Location = New Point(520, 50)

        AddHandler lnkFacebook.Click,
            Sub()

                Process.Start(
                "https://facebook.com")

            End Sub

        pnlFooter.Controls.Add(lnkFacebook)



        Dim lnkEmail As New LinkLabel()

        lnkEmail.Text = "ghcpetmart@gmail.com"

        lnkEmail.LinkColor = Color.DeepSkyBlue

        lnkEmail.AutoSize = True

        lnkEmail.Location = New Point(520, 75)

        AddHandler lnkEmail.Click,
            Sub()

                Process.Start(
                "mailto:ghcpetmart@gmail.com")

            End Sub

        pnlFooter.Controls.Add(lnkEmail)



        Dim lnkAbout As New LinkLabel()

        lnkAbout.Text = "About Us"

        lnkAbout.Font = New Font("Arial", 10, FontStyle.Bold)

        lnkAbout.LinkColor = Color.Gold

        lnkAbout.AutoSize = True

        lnkAbout.Location = New Point(900, 50)

        AddHandler lnkAbout.Click,
            Sub()

                Dim f As New frmAboutUs()

                f.ShowDialog()

            End Sub

        pnlFooter.Controls.Add(lnkAbout)

        Me.Controls.Add(pnlFooter)

    End Sub
End Class