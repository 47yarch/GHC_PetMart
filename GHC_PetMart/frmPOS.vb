Imports MySql.Data.MySqlClient

Public Class frmPOS

    Private txtSearch As New TextBox()
    Private dgvSearch As New DataGridView()
    Private dgvCart As New DataGridView()
    Private txtQty As New TextBox()
    Private cboCustomer As New ComboBox()
    Private lblSubtotal As New Label()
    Private lblTotal As New Label()
    Private txtPaid As New TextBox()
    Private lblChange As New Label()
    Private dtCart As New DataTable()
    Private decTotal As Decimal = 0

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Point of Sale"
        Me.Size = New Size(1150, 680)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        SetupCartTable()
        BuildUI()
        LoadCustomers()
    End Sub

    Private Sub SetupCartTable()
        dtCart.Columns.Add("product_id", GetType(Integer))
        dtCart.Columns.Add("Product Name", GetType(String))
        dtCart.Columns.Add("Price", GetType(Decimal))
        dtCart.Columns.Add("Qty", GetType(Integer))
        dtCart.Columns.Add("Subtotal", GetType(Decimal))
    End Sub

    Private Sub BuildUI()

        Dim pLeft As New Panel() With {
            .Location = New Point(10, 10),
            .Size = New Size(550, 630),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }

        Dim lblS As New Label() With {
            .Text = "Search Product:",
            .AutoSize = True,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .Location = New Point(10, 10)
        }
        pLeft.Controls.Add(lblS)

        txtSearch.Location = New Point(10, 30)
        txtSearch.Size = New Size(380, 28)
        pLeft.Controls.Add(txtSearch)

        Dim btnSearch As New Button() With {
            .Text = "Search",
            .Size = New Size(90, 28),
            .Location = New Point(400, 30),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSearch.Click, AddressOf SearchProduct
        pLeft.Controls.Add(btnSearch)

        dgvSearch.Location = New Point(10, 68)
        dgvSearch.Size = New Size(525, 200)
        dgvSearch.ReadOnly = True
        dgvSearch.AllowUserToAddRows = False
        dgvSearch.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvSearch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvSearch.BackgroundColor = Color.WhiteSmoke
        pLeft.Controls.Add(dgvSearch)

        Dim lblQ As New Label() With {
            .Text = "Quantity:",
            .AutoSize = True,
            .Location = New Point(10, 280)
        }
        pLeft.Controls.Add(lblQ)

        txtQty.Location = New Point(10, 300)
        txtQty.Size = New Size(100, 28)
        txtQty.Text = "1"
        pLeft.Controls.Add(txtQty)

        Dim btnAdd As New Button() With {
            .Text = "ADD TO CART",
            .Size = New Size(200, 35),
            .Location = New Point(120, 298),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        AddHandler btnAdd.Click, AddressOf AddToCart
        pLeft.Controls.Add(btnAdd)

        Dim lblC As New Label() With {
            .Text = "Customer (optional):",
            .AutoSize = True,
            .Location = New Point(10, 350)
        }
        pLeft.Controls.Add(lblC)

        cboCustomer.Location = New Point(10, 370)
        cboCustomer.Size = New Size(525, 28)
        cboCustomer.DropDownStyle = ComboBoxStyle.DropDownList
        pLeft.Controls.Add(cboCustomer)

        Dim lblCart As New Label() With {
            .Text = "CART:",
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(10, 415)
        }
        pLeft.Controls.Add(lblCart)

        dgvCart.Location = New Point(10, 438)
        dgvCart.Size = New Size(525, 150)
        dgvCart.ReadOnly = True
        dgvCart.AllowUserToAddRows = False
        dgvCart.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvCart.BackgroundColor = Color.WhiteSmoke
        dgvCart.DataSource = dtCart
        pLeft.Controls.Add(dgvCart)

        Dim btnRemove As New Button() With {
            .Text = "Remove Item",
            .Size = New Size(130, 30),
            .Location = New Point(10, 595),
            .BackColor = Color.FromArgb(244, 67, 54),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnRemove.Click, AddressOf RemoveItem
        pLeft.Controls.Add(btnRemove)

        Dim btnClear As New Button() With {
            .Text = "Clear Cart",
            .Size = New Size(130, 30),
            .Location = New Point(150, 595),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnClear.Click, AddressOf ClearCart
        pLeft.Controls.Add(btnClear)

        Me.Controls.Add(pLeft)


        Dim pRight As New Panel() With {
            .Location = New Point(575, 10),
            .Size = New Size(555, 630),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle
        }

        Dim lblHead As New Label() With {
            .Text = "PAYMENT",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(180, 15)
        }
        pRight.Controls.Add(lblHead)

        Dim lblST As New Label() With {
            .Text = "Subtotal:",
            .Font = New Font("Arial", 11),
            .AutoSize = True,
            .Location = New Point(20, 70)
        }
        pRight.Controls.Add(lblST)

        lblSubtotal.Text = "P 0.00"
        lblSubtotal.Font = New Font("Arial", 14, FontStyle.Bold)
        lblSubtotal.AutoSize = True
        lblSubtotal.Location = New Point(300, 68)
        pRight.Controls.Add(lblSubtotal)


        Dim div As New Label() With {
            .BorderStyle = BorderStyle.Fixed3D,
            .Size = New Size(510, 2),
            .Location = New Point(20, 110)
        }
        pRight.Controls.Add(div)


        Dim lblTL As New Label() With {
            .Text = "TOTAL:",
            .Font = New Font("Arial", 13, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 125)
        }
        pRight.Controls.Add(lblTL)

        lblTotal.Text = "P 0.00"
        lblTotal.Font = New Font("Arial", 18, FontStyle.Bold)
        lblTotal.ForeColor = Color.FromArgb(63, 81, 181)
        lblTotal.AutoSize = True
        lblTotal.Location = New Point(260, 120)
        pRight.Controls.Add(lblTotal)


        Dim lblPaid As New Label() With {
            .Text = "Amount Paid (Cash):",
            .Font = New Font("Arial", 11),
            .AutoSize = True,
            .Location = New Point(20, 180)
        }
        pRight.Controls.Add(lblPaid)

        txtPaid.Location = New Point(20, 205)
        txtPaid.Size = New Size(510, 38)
        txtPaid.Font = New Font("Arial", 14)
        AddHandler txtPaid.TextChanged, AddressOf CalcChange
        pRight.Controls.Add(txtPaid)


        Dim lblCh As New Label() With {
            .Text = "Change:",
            .Font = New Font("Arial", 11),
            .AutoSize = True,
            .Location = New Point(20, 260)
        }
        pRight.Controls.Add(lblCh)

        lblChange.Text = "P 0.00"
        lblChange.Font = New Font("Arial", 18, FontStyle.Bold)
        lblChange.ForeColor = Color.FromArgb(76, 175, 80)
        lblChange.AutoSize = True
        lblChange.Location = New Point(20, 285)
        pRight.Controls.Add(lblChange)

        Dim btnCheckout As New Button() With {
            .Text = "CHECKOUT",
            .Size = New Size(510, 65),
            .Location = New Point(20, 360),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 16, FontStyle.Bold)
        }
        AddHandler btnCheckout.Click, AddressOf ProcessCheckout
        pRight.Controls.Add(btnCheckout)

        Dim btnNew As New Button() With {
            .Text = "NEW TRANSACTION",
            .Size = New Size(510, 45),
            .Location = New Point(20, 440),
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 11)
        }
        AddHandler btnNew.Click, AddressOf ClearCart
        pRight.Controls.Add(btnNew)

        Me.Controls.Add(pRight)
    End Sub

    Private Sub LoadCustomers()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter(
                "SELECT customer_id, CONCAT(first_name,' ',last_name) AS fullname FROM tbl_customers ORDER BY first_name", conn)
            Dim dt As New DataTable()
            da.Fill(dt)
            cboCustomer.Items.Clear()
            cboCustomer.Items.Add("Walk-in Customer")
            For Each row As DataRow In dt.Rows
                cboCustomer.Items.Add(row("fullname").ToString())
            Next
            cboCustomer.SelectedIndex = 0
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub SearchProduct(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand(
                "SELECT p.product_id, p.product_name AS 'Product', " &
                "c.category_name AS 'Category', " &
                "p.regular_price AS 'Price', p.stock_quantity AS 'Stock' " &
                "FROM tbl_products p JOIN tbl_categories c ON p.category_id=c.category_id " &
                "WHERE p.product_name LIKE @s AND p.status='In Stock' AND p.stock_quantity > 0 " &
                "ORDER BY p.product_name", conn)
            cmd.Parameters.AddWithValue("@s", "%" & txtSearch.Text.Trim() & "%")
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgvSearch.DataSource = dt
            If dgvSearch.Columns.Contains("product_id") Then
                dgvSearch.Columns("product_id").Visible = False
            End If
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub AddToCart(sender As Object, e As EventArgs)
        If dgvSearch.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a product first.")
            Return
        End If

        If Not IsNumeric(txtQty.Text) OrElse CInt(txtQty.Text) <= 0 Then
            MessageBox.Show("Please enter a valid quantity.")
            Return
        End If

        Dim productID As Integer = CInt(dgvSearch.SelectedRows(0).Cells("product_id").Value)
        Dim productName As String = dgvSearch.SelectedRows(0).Cells("Product").Value.ToString()
        Dim price As Decimal = Convert.ToDecimal(dgvSearch.SelectedRows(0).Cells("Price").Value)
        Dim stock As Integer = CInt(dgvSearch.SelectedRows(0).Cells("Stock").Value)
        Dim qty As Integer = CInt(txtQty.Text)

        If qty > stock Then
            MessageBox.Show("Not enough stock. Available: " & stock)
            Return
        End If


        For Each row As DataRow In dtCart.Rows
            If CInt(row("product_id")) = productID Then
                Dim newQty As Integer = CInt(row("Qty")) + qty
                If newQty > stock Then
                    MessageBox.Show("Not enough stock for this quantity.")
                    Return
                End If
                row("Qty") = newQty
                row("Subtotal") = CDec(row("Price")) * newQty
                UpdateTotals()
                Return
            End If
        Next

        dtCart.Rows.Add(productID, productName, price, qty, price * qty)
        UpdateTotals()
    End Sub

    Private Sub RemoveItem(sender As Object, e As EventArgs)
        If dgvCart.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an item to remove.")
            Return
        End If
        Dim idx As Integer = dgvCart.SelectedRows(0).Index
        dtCart.Rows(idx).Delete()
        UpdateTotals()
    End Sub

    Private Sub ClearCart(sender As Object, e As EventArgs)
        dtCart.Rows.Clear()
        txtPaid.Clear()
        lblSubtotal.Text = "P 0.00"
        lblTotal.Text = "P 0.00"
        lblChange.Text = "P 0.00"
        decTotal = 0
    End Sub

    Private Sub UpdateTotals()
        decTotal = 0
        For Each row As DataRow In dtCart.Rows
            decTotal += Convert.ToDecimal(row("Subtotal"))
        Next
        lblSubtotal.Text = "P " & decTotal.ToString("N2")
        lblTotal.Text = "P " & decTotal.ToString("N2")
        CalcChange(Nothing, Nothing)
    End Sub

    Private Sub CalcChange(sender As Object, e As EventArgs)
        If IsNumeric(txtPaid.Text) Then
            Dim change As Decimal = CDec(txtPaid.Text) - decTotal
            lblChange.Text = "P " & change.ToString("N2")
            lblChange.ForeColor = If(change >= 0, Color.FromArgb(76, 175, 80), Color.Red)
        Else
            lblChange.Text = "P 0.00"
        End If
    End Sub

    Private Sub ProcessCheckout(sender As Object, e As EventArgs)

        If dtCart.Rows.Count = 0 Then
            MessageBox.Show("Cart is empty.")
            Return
        End If

        If Not IsNumeric(txtPaid.Text) Then
            MessageBox.Show("Please enter the amount paid.")
            Return
        End If

        Dim paid As Decimal = CDec(txtPaid.Text)

        If paid < decTotal Then
            MessageBox.Show("Amount paid is less than the total.")
            Return
        End If

        Dim change As Decimal = paid - decTotal

        Dim receiptNo As String =
        "RCPT-" & DateTime.Now.ToString("yyyyMMddHHmmss")



        Dim custID As Object = DBNull.Value

        If cboCustomer.SelectedIndex > 0 Then

            If Not OpenConnection() Then Return

            Dim cmdCust As New MySqlCommand(
            "SELECT customer_id
             FROM tbl_customers
             ORDER BY first_name
             LIMIT " & cboCustomer.SelectedIndex & ",1", conn)

            custID = cmdCust.ExecuteScalar()

            CloseConnection()

        End If

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try



            Dim cmdSale As New MySqlCommand(
            "INSERT INTO tbl_sales
            (
                customer_id,
                employee_id,
                amount_paid,
                receipt_no
            )
            VALUES
            (
                @cid,
                @eid,
                @paid,
                @rcpt
            )", conn, trans)

            cmdSale.Parameters.AddWithValue("@cid", custID)
            cmdSale.Parameters.AddWithValue("@eid", LoggedUserID)
            cmdSale.Parameters.AddWithValue("@paid", paid)
            cmdSale.Parameters.AddWithValue("@rcpt", receiptNo)

            cmdSale.ExecuteNonQuery()

            Dim saleID As Integer =
            Convert.ToInt32(New MySqlCommand(
                "SELECT LAST_INSERT_ID()", conn, trans).ExecuteScalar())


            For Each row As DataRow In dtCart.Rows

                Dim productID As Integer = CInt(row("product_id"))
                Dim qty As Integer = CInt(row("Qty"))
                Dim price As Decimal = CDec(row("Price"))

                Dim cmdDetail As New MySqlCommand(
                "INSERT INTO tbl_sale_details
                (
                    sale_id,
                    product_id,
                    quantity,
                    unit_price
                )
                VALUES
                (
                    @sid,
                    @pid,
                    @qty,
                    @price
                )", conn, trans)

                cmdDetail.Parameters.AddWithValue("@sid", saleID)
                cmdDetail.Parameters.AddWithValue("@pid", productID)
                cmdDetail.Parameters.AddWithValue("@qty", qty)
                cmdDetail.Parameters.AddWithValue("@price", price)

                cmdDetail.ExecuteNonQuery()


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


                Dim cmdLog As New MySqlCommand(
                "INSERT INTO tbl_inventory_log
                (
                    product_id,
                    action_type,
                    quantity,
                    ref_id,
                    logged_by
                )
                VALUES
                (
                    @pid,
                    'Sale',
                    @qty,
                    @sid,
                    @uid
                )", conn, trans)

                cmdLog.Parameters.AddWithValue("@pid", productID)
                cmdLog.Parameters.AddWithValue("@qty", -qty)
                cmdLog.Parameters.AddWithValue("@sid", saleID)
                cmdLog.Parameters.AddWithValue("@uid", LoggedUserID)

                cmdLog.ExecuteNonQuery()

            Next

            trans.Commit()

            CloseConnection()

            Dim receipt As New frmReceipt()
            receipt.SaleID = saleID
            receipt.ShowDialog()

            ClearCart(Nothing, Nothing)
            LoadCustomers()

        Catch ex As Exception

            trans.Rollback()

            CloseConnection()

            MessageBox.Show("Checkout Error: " & ex.Message)

        End Try

    End Sub

End Class