Imports MySql.Data.MySqlClient

Public Class frmCart

    Public CustomerUserID As Integer = 0

    Private dgvCart As New DataGridView()
    Private lblSubtotal As New Label()
    Private lblTotal As New Label()
    Private txtPaid As New TextBox()
    Private lblChange As New Label()
    Private lblItemCount As New Label()
    Private decTotal As Decimal = 0

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "My Cart"
        Me.Size = New Size(1000, 600)
        Me.MinimumSize = New Size(900, 550)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        BuildUI()
        LoadCart()
    End Sub

    Private Sub BuildUI()

        Dim lblTitle As New Label() With {
                    .Text = "My Shopping Cart",
                    .Font = New Font("Arial", 14, FontStyle.Bold),
                    .AutoSize = True,
                    .Location = New Point(15, 15)
                }
        Me.Controls.Add(lblTitle)

        lblItemCount.AutoSize = True
        lblItemCount.Font = New Font("Arial", 9)
        lblItemCount.ForeColor = Color.Gray
        lblItemCount.Location = New Point(15, 42)
        Me.Controls.Add(lblItemCount)

        Dim pLeft As New Panel() With {
                    .Location = New Point(10, 60),
                    .BackColor = Color.White,
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or
                              AnchorStyles.Left Or AnchorStyles.Right
                }

        dgvCart.Location = New Point(5, 5)
        dgvCart.ReadOnly = True
        dgvCart.AllowUserToAddRows = False
        dgvCart.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvCart.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvCart.BackgroundColor = Color.White
        dgvCart.BorderStyle = BorderStyle.None
        dgvCart.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or
                                 AnchorStyles.Left Or AnchorStyles.Right
        pLeft.Controls.Add(dgvCart)

        Dim btnIncrease As New Button() With {
                    .Text = "+ Qty",
                    .Size = New Size(90, 35),
                    .BackColor = Color.FromArgb(76, 175, 80),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Arial", 9, FontStyle.Bold),
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                }
        AddHandler btnIncrease.Click, AddressOf IncreaseQty
        pLeft.Controls.Add(btnIncrease)

        Dim btnDecrease As New Button() With {
                    .Text = "- Qty",
                    .Size = New Size(90, 35),
                    .BackColor = Color.FromArgb(255, 152, 0),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Arial", 9, FontStyle.Bold),
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                }
        AddHandler btnDecrease.Click, AddressOf DecreaseQty
        pLeft.Controls.Add(btnDecrease)

        Dim btnRemove As New Button() With {
                    .Text = "Remove",
                    .Size = New Size(90, 35),
                    .BackColor = Color.FromArgb(244, 67, 54),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                }
        AddHandler btnRemove.Click, AddressOf RemoveItem
        pLeft.Controls.Add(btnRemove)

        Dim btnClear As New Button() With {
                    .Text = "Clear All",
                    .Size = New Size(90, 35),
                    .FlatStyle = FlatStyle.Flat,
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                }
        AddHandler btnClear.Click, AddressOf ClearCart
        pLeft.Controls.Add(btnClear)

        Dim btnBack As New Button() With {
                    .Text = "Continue Shopping",
                    .Size = New Size(145, 35),
                    .FlatStyle = FlatStyle.Flat,
                    .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
                }
        AddHandler btnBack.Click, Sub(s, e) Me.Close()
        pLeft.Controls.Add(btnBack)

        Me.Controls.Add(pLeft)

        Dim pRight As New Panel() With {
                    .Size = New Size(310, 490),
                    .BackColor = Color.White,
                    .BorderStyle = BorderStyle.FixedSingle,
                    .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Right
                }

        Dim y As Integer = 15

        Dim lblPayHead As New Label() With {
                    .Text = "ORDER SUMMARY",
                    .Font = New Font("Arial", 12, FontStyle.Bold),
                    .AutoSize = True,
                    .Location = New Point(70, y)
                }
        pRight.Controls.Add(lblPayHead) : y += 40

        pRight.Controls.Add(MakeLine(y)) : y += 15

        Dim lSub As New Label() With {
                    .Text = "Subtotal:",
                    .Font = New Font("Arial", 10),
                    .AutoSize = True,
                    .Location = New Point(15, y)
                }
        pRight.Controls.Add(lSub)

        lblSubtotal.Text = "P 0.00"
        lblSubtotal.Font = New Font("Arial", 10, FontStyle.Bold)
        lblSubtotal.AutoSize = True
        lblSubtotal.Location = New Point(185, y)
        pRight.Controls.Add(lblSubtotal) : y += 35

        pRight.Controls.Add(MakeLine(y)) : y += 15


        Dim lTot As New Label() With {
                    .Text = "TOTAL:",
                    .Font = New Font("Arial", 13, FontStyle.Bold),
                    .AutoSize = True,
                    .Location = New Point(15, y)
                }
        pRight.Controls.Add(lTot)

        lblTotal.Text = "P 0.00"
        lblTotal.Font = New Font("Arial", 15, FontStyle.Bold)
        lblTotal.ForeColor = Color.FromArgb(63, 81, 181)
        lblTotal.AutoSize = True
        lblTotal.Location = New Point(150, y - 2)
        pRight.Controls.Add(lblTotal) : y += 50

        pRight.Controls.Add(MakeLine(y)) : y += 20

        Dim lPaid As New Label() With {
                    .Text = "Amount Paid (Cash):",
                    .Font = New Font("Arial", 9),
                    .AutoSize = True,
                    .Location = New Point(15, y)
                }
        pRight.Controls.Add(lPaid) : y += 22

        txtPaid.Location = New Point(15, y)
        txtPaid.Size = New Size(275, 35)
        txtPaid.Font = New Font("Arial", 14)
        AddHandler txtPaid.TextChanged, AddressOf CalcChange
        pRight.Controls.Add(txtPaid) : y += 48

        Dim lChg As New Label() With {
                    .Text = "Change:",
                    .Font = New Font("Arial", 9),
                    .AutoSize = True,
                    .Location = New Point(15, y)
                }
        pRight.Controls.Add(lChg)

        lblChange.Text = "P 0.00"
        lblChange.Font = New Font("Arial", 14, FontStyle.Bold)
        lblChange.ForeColor = Color.FromArgb(76, 175, 80)
        lblChange.AutoSize = True
        lblChange.Location = New Point(155, y - 2)
        pRight.Controls.Add(lblChange) : y += 55

        Dim btnCheckout As New Button() With {
                    .Text = "CHECKOUT",
                    .Size = New Size(275, 55),
                    .Location = New Point(15, y),
                    .BackColor = Color.FromArgb(63, 81, 181),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat,
                    .Font = New Font("Arial", 14, FontStyle.Bold)
                }
        AddHandler btnCheckout.Click, AddressOf ProcessCheckout
        pRight.Controls.Add(btnCheckout)

        Me.Controls.Add(pRight)

        LayoutPanels(pLeft, pRight, btnIncrease, btnDecrease, btnRemove, btnClear, btnBack)
        AddHandler Me.Resize, Sub(s, e)
                                  LayoutPanels(pLeft, pRight, btnIncrease, btnDecrease, btnRemove, btnClear, btnBack)
                              End Sub
    End Sub

    Private Function MakeLine(y As Integer) As Label
        Return New Label() With {
                    .BorderStyle = BorderStyle.Fixed3D,
                    .Size = New Size(275, 2),
                    .Location = New Point(15, y)
                }
    End Function

    Private Sub LayoutPanels(pLeft As Panel, pRight As Panel,
                                      btnI As Button, btnD As Button,
                                      btnR As Button, btnC As Button, btnB As Button)
        Dim margin As Integer = 10
        Dim rightW As Integer = 320
        Dim topY As Integer = 60
        Dim availH As Integer = Me.ClientSize.Height - topY - margin
        Dim availW As Integer = Me.ClientSize.Width - (margin * 3) - rightW

        pLeft.Location = New Point(margin, topY)
        pLeft.Size = New Size(availW, availH)

        dgvCart.Size = New Size(availW - 12, availH - 55)

        Dim btnY As Integer = availH - 45
        btnI.Location = New Point(5, btnY)
        btnD.Location = New Point(100, btnY)
        btnR.Location = New Point(195, btnY)
        btnC.Location = New Point(290, btnY)
        btnB.Location = New Point(385, btnY)

        pRight.Location = New Point(Me.ClientSize.Width - rightW - margin, topY)
        pRight.Height = availH
    End Sub

    Private Sub LoadCart()
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand(
                        "SELECT p.product_id, p.product_name AS 'Product', " &
                        "p.regular_price AS 'Price', c.quantity AS 'Qty', " &
                        "(p.regular_price * c.quantity) AS 'Subtotal', " &
                        "p.stock_quantity AS 'Stock' " &
                        "FROM tbl_cart c " &
                        "JOIN tbl_products p ON c.product_id = p.product_id " &
                        "WHERE c.user_id = @uid " &
                        "ORDER BY p.product_name", conn)
            cmd.Parameters.AddWithValue("@uid", CustomerUserID)

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            dgvCart.DataSource = dt

            For Each col As String In {"product_id", "Stock"}
                If dgvCart.Columns.Contains(col) Then
                    dgvCart.Columns(col).Visible = False
                End If
            Next

            decTotal = 0
            For Each row As DataRow In dt.Rows
                decTotal += Convert.ToDecimal(row("Subtotal"))
            Next

            lblSubtotal.Text = "P " & decTotal.ToString("N2")
            lblTotal.Text = "P " & decTotal.ToString("N2")
            lblItemCount.Text = dt.Rows.Count & " item(s) in cart"
            CalcChange(Nothing, Nothing)

        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub IncreaseQty(sender As Object, e As EventArgs)

        If dgvCart.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an item first.")
            Return
        End If

        Dim productID As Integer = CInt(dgvCart.SelectedRows(0).Cells("product_id").Value)
        Dim currentQty As Integer = CInt(dgvCart.SelectedRows(0).Cells("Qty").Value)
        Dim stock As Integer = CInt(dgvCart.SelectedRows(0).Cells("Stock").Value)

        If currentQty >= stock Then
            MessageBox.Show("Maximum stock reached.")
            Return
        End If

        If Not OpenConnection() Then Return

        Dim cmd As New MySqlCommand(
        "UPDATE tbl_cart
         SET quantity = quantity + 1
         WHERE user_id=@uid
         AND product_id=@pid", conn)

        cmd.Parameters.AddWithValue("@uid", CustomerUserID)
        cmd.Parameters.AddWithValue("@pid", productID)

        cmd.ExecuteNonQuery()

        CloseConnection()

        LoadCart()

    End Sub

    Private Sub DecreaseQty(sender As Object, e As EventArgs)

        If dgvCart.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an item first.")
            Return
        End If

        Dim productID As Integer = CInt(dgvCart.SelectedRows(0).Cells("product_id").Value)
        Dim qty As Integer = CInt(dgvCart.SelectedRows(0).Cells("Qty").Value)

        If qty <= 1 Then
            RemoveItem(Nothing, Nothing)
            Return
        End If

        If Not OpenConnection() Then Return

        Dim cmd As New MySqlCommand(
        "UPDATE tbl_cart
         SET quantity = quantity - 1
         WHERE user_id=@uid
         AND product_id=@pid", conn)

        cmd.Parameters.AddWithValue("@uid", CustomerUserID)
        cmd.Parameters.AddWithValue("@pid", productID)

        cmd.ExecuteNonQuery()

        CloseConnection()

        LoadCart()

    End Sub

    Private Sub RemoveItem(sender As Object, e As EventArgs)

        If dgvCart.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an item.")
            Return
        End If

        Dim productID As Integer = CInt(dgvCart.SelectedRows(0).Cells("product_id").Value)

        If Not OpenConnection() Then Return

        Dim cmd As New MySqlCommand(
        "DELETE FROM tbl_cart
         WHERE user_id=@uid
         AND product_id=@pid", conn)

        cmd.Parameters.AddWithValue("@uid", CustomerUserID)
        cmd.Parameters.AddWithValue("@pid", productID)

        cmd.ExecuteNonQuery()

        CloseConnection()

        LoadCart()

    End Sub

    Private Sub ClearCart(sender As Object, e As EventArgs)
        If MessageBox.Show("Remove ALL items from cart?",
                                   "Clear Cart", MessageBoxButtons.YesNo,
                                   MessageBoxIcon.Warning) = DialogResult.No Then Return

        If Not OpenConnection() Then Return
        Dim cmd As New MySqlCommand("DELETE FROM tbl_cart WHERE user_id = @uid", conn)
        cmd.Parameters.AddWithValue("@uid", CustomerUserID)
        cmd.ExecuteNonQuery()
        CloseConnection()
        txtPaid.Clear()
        decTotal = 0
        LoadCart()
    End Sub

    Private Sub CalcChange(sender As Object, e As EventArgs)
        If IsNumeric(txtPaid.Text) Then
            Dim change As Decimal = CDec(txtPaid.Text) - decTotal
            lblChange.Text = "P " & change.ToString("N2")
            lblChange.ForeColor = If(change >= 0,
                        Color.FromArgb(76, 175, 80), Color.Red)
        Else
            lblChange.Text = "P 0.00"
            lblChange.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub ProcessCheckout(sender As Object, e As EventArgs)

        If dgvCart.Rows.Count = 0 Then
            MessageBox.Show("Your cart is empty.")
            Return
        End If

        If Not IsNumeric(txtPaid.Text) Then
            MessageBox.Show("Please enter the amount paid.")
            Return
        End If

        Dim paid As Decimal = CDec(txtPaid.Text)

        If paid < decTotal Then
            MessageBox.Show("Insufficient payment.")
            Return
        End If

        Dim change As Decimal = paid - decTotal
        Dim receiptNo As String = "RCPT-" & DateTime.Now.ToString("yyyyMMddHHmmss")

        If Not OpenConnection() Then Return

        Dim trans As MySqlTransaction = conn.BeginTransaction()

        Try

            Dim custID As Object = DBNull.Value

            Dim cmdCust As New MySqlCommand(
            "SELECT customer_id
             FROM tbl_customers
             WHERE user_id=@uid", conn, trans)

            cmdCust.Parameters.AddWithValue("@uid", CustomerUserID)

            Dim custObj As Object = cmdCust.ExecuteScalar()

            If custObj IsNot Nothing AndAlso Not IsDBNull(custObj) Then
                custID = CInt(custObj)
            End If

            Dim cmdEmp As New MySqlCommand(
            "SELECT user_id
             FROM tbl_users
             WHERE role='Employee'
             LIMIT 1", conn, trans)

            Dim empObj As Object = cmdEmp.ExecuteScalar()

            If empObj Is Nothing OrElse IsDBNull(empObj) Then
                Throw New Exception("No employee account found.")
            End If

            Dim empID As Integer = CInt(empObj)

            Dim cmdSale As New MySqlCommand(
            "INSERT INTO tbl_sales
            (customer_id,employee_id,amount_paid,receipt_no)
            VALUES
            (@cid,@eid,@paid,@rcpt)", conn, trans)

            cmdSale.Parameters.AddWithValue("@cid", custID)
            cmdSale.Parameters.AddWithValue("@eid", empID)
            cmdSale.Parameters.AddWithValue("@paid", paid)
            cmdSale.Parameters.AddWithValue("@rcpt", receiptNo)

            cmdSale.ExecuteNonQuery()

            Dim saleID As Integer =
            Convert.ToInt32(New MySqlCommand(
            "SELECT LAST_INSERT_ID()", conn, trans).ExecuteScalar())

            For Each row As DataGridViewRow In dgvCart.Rows

                If row.IsNewRow Then Continue For

                Dim productID As Integer =
                CInt(row.Cells("product_id").Value)

                Dim qty As Integer =
                CInt(row.Cells("Qty").Value)

                Dim price As Decimal =
                CDec(row.Cells("Price").Value)

                Dim cmdDet As New MySqlCommand(
                "INSERT INTO tbl_sale_details
                (sale_id,product_id,quantity,unit_price)
                VALUES
                (@sid,@pid,@qty,@price)", conn, trans)

                cmdDet.Parameters.AddWithValue("@sid", saleID)
                cmdDet.Parameters.AddWithValue("@pid", productID)
                cmdDet.Parameters.AddWithValue("@qty", qty)
                cmdDet.Parameters.AddWithValue("@price", price)

                cmdDet.ExecuteNonQuery()

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
                (product_id,action_type,quantity,ref_id,logged_by)
                VALUES
                (@pid,'Sale',@qty,@sid,@uid)", conn, trans)

                cmdLog.Parameters.AddWithValue("@pid", productID)
                cmdLog.Parameters.AddWithValue("@qty", -qty)
                cmdLog.Parameters.AddWithValue("@sid", saleID)
                cmdLog.Parameters.AddWithValue("@uid", CustomerUserID)

                cmdLog.ExecuteNonQuery()

            Next

            Dim cmdClear As New MySqlCommand(
            "DELETE FROM tbl_cart
             WHERE user_id=@uid", conn, trans)

            cmdClear.Parameters.AddWithValue("@uid", CustomerUserID)

            cmdClear.ExecuteNonQuery()

            trans.Commit()

            CloseConnection()

            ReceiptGenerator.GenerateReceipt(saleID)

            MessageBox.Show(
"Order placed successfully!" &
vbCrLf & vbCrLf &
"Receipt No : " & receiptNo,
"Checkout Successful",
MessageBoxButtons.OK,
MessageBoxIcon.Information)

            LoadCart()
            txtPaid.Clear()
            If Application.OpenForms().OfType(Of frmShop).Any Then
                Application.OpenForms().OfType(Of frmShop).First.RefreshShop()
            End If

        Catch ex As Exception

            Try
                trans.Rollback()
            Catch
            End Try

            CloseConnection()

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

End Class