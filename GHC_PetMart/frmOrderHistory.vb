Imports MySql.Data.MySqlClient

Public Class frmOrderHistory

    Public CustomerUserID As Integer = 0

    Private dgvOrders As New DataGridView()
    Private dgvItems As New DataGridView()
    Private lblTotalSpent As New Label()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "My Order History"
        Me.Size = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "My Order History",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        lblTotalSpent.AutoSize = True
        lblTotalSpent.Font = New Font("Arial", 10, FontStyle.Bold)
        lblTotalSpent.ForeColor = Color.FromArgb(63, 81, 181)
        lblTotalSpent.Location = New Point(600, 20)
        Me.Controls.Add(lblTotalSpent)

        Dim lbl1 As New Label() With {
            .Text = "Orders (click to see items):",
            .AutoSize = True,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .Location = New Point(15, 55)
        }
        Me.Controls.Add(lbl1)

        dgvOrders.Location = New Point(15, 75)
        dgvOrders.Size = New Size(860, 220)
        dgvOrders.ReadOnly = True
        dgvOrders.AllowUserToAddRows = False
        dgvOrders.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvOrders.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvOrders.BackgroundColor = Color.White
        AddHandler dgvOrders.SelectionChanged, AddressOf LoadOrderItems
        Me.Controls.Add(dgvOrders)

        Dim lbl2 As New Label() With {
            .Text = "Items in Selected Order:",
            .AutoSize = True,
            .Font = New Font("Arial", 9, FontStyle.Bold),
            .Location = New Point(15, 308)
        }
        Me.Controls.Add(lbl2)

        dgvItems.Location = New Point(15, 328)
        dgvItems.Size = New Size(860, 210)
        dgvItems.ReadOnly = True
        dgvItems.AllowUserToAddRows = False
        dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvItems.BackgroundColor = Color.White
        Me.Controls.Add(dgvItems)

        Dim btnBack As New Button() With {
            .Text = "Back to Shop",
            .Size = New Size(150, 35),
            .Location = New Point(15, 550),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnBack.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnBack)

        LoadOrders()
    End Sub

    Private Sub LoadOrders()
        If Not OpenConnection() Then Return
        Try

            Dim cmdCust As New MySqlCommand(
                "SELECT customer_id FROM tbl_customers WHERE user_id=@uid", conn)
            cmdCust.Parameters.AddWithValue("@uid", CustomerUserID)
            Dim custObj As Object = cmdCust.ExecuteScalar()

            If custObj Is Nothing Then
                lblTotalSpent.Text = "No order history found."
                CloseConnection()
                Return
            End If

            Dim custID As Integer = CInt(custObj)

            Dim cmd As New MySqlCommand(
"SELECT
    s.sale_id,
    s.receipt_no AS 'Receipt No',
    s.sale_date AS 'Date',

    SUM(sd.quantity * sd.unit_price - sd.discount) AS 'Total',

    s.amount_paid AS 'Paid',

    (s.amount_paid -
        SUM(sd.quantity * sd.unit_price - sd.discount)
    ) AS 'Change'

FROM tbl_sales s

INNER JOIN tbl_sale_details sd
ON s.sale_id = sd.sale_id

WHERE s.customer_id=@cid

GROUP BY
    s.sale_id,
    s.receipt_no,
    s.sale_date,
    s.amount_paid

ORDER BY s.sale_date DESC", conn)
            cmd.Parameters.AddWithValue("@cid", custID)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgvOrders.DataSource = dt
            If dgvOrders.Columns.Contains("sale_id") Then
                dgvOrders.Columns("sale_id").Visible = False
            End If

            Dim cmdTotal As New MySqlCommand(
"SELECT IFNULL(
    SUM(sd.quantity * sd.unit_price - sd.discount),
0)

FROM tbl_sales s

INNER JOIN tbl_sale_details sd
ON s.sale_id = sd.sale_id

WHERE s.customer_id=@cid", conn)
            cmdTotal.Parameters.AddWithValue("@cid", custID)
            Dim totalSpent As Decimal = Convert.ToDecimal(cmdTotal.ExecuteScalar())
            lblTotalSpent.Text = "Total Spent: P " & totalSpent.ToString("N2") &
                                 "  |  Orders: " & dt.Rows.Count
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub LoadOrderItems(sender As Object, e As EventArgs)

        If dgvOrders.SelectedRows.Count = 0 Then Return

        Try

            Using con As MySqlConnection = GetConn()

                Dim saleID As Integer =
                CInt(dgvOrders.SelectedRows(0).Cells("sale_id").Value)

                Dim cmd As New MySqlCommand(
                "SELECT
                    p.product_name AS 'Product',
                    sd.quantity AS 'Qty',
                    sd.unit_price AS 'Price',
                    (sd.quantity * sd.unit_price - sd.discount) AS 'Subtotal'
                 FROM tbl_sale_details sd
                 INNER JOIN tbl_products p
                    ON sd.product_id = p.product_id
                 WHERE sd.sale_id=@id", con)

                cmd.Parameters.AddWithValue("@id", saleID)

                Dim dt As New DataTable()

                Dim da As New MySqlDataAdapter(cmd)
                da.Fill(dt)

                dgvItems.DataSource = dt

            End Using

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try

    End Sub

End Class