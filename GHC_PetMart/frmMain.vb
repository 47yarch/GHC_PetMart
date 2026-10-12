Imports MySql.Data.MySqlClient

Public Class frmMain

    Private lblTotalProd As New Label()
    Private lblTodaySales As New Label()
    Private lblTotalCust As New Label()
    Private lblLowCount As New Label()
    Private dgvLowStock As New DataGridView()
    Private dgvRecentSales As New DataGridView()

    Private Sub InitMain() Handles MyBase.Load
        Me.Text = "GHC Pet Mart - Admin Dashboard"
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.WhiteSmoke

        BuildMenu()
        BuildUI()
        LoadData()
    End Sub

    Private Sub BuildMenu()
        Dim ms As New MenuStrip()

        Dim menuItems As String() = {
    "Dashboard", "Products", "Inventory",
    "Customers", "Suppliers", "Employees",
    "Purchase Orders", "Sales History", "Expenses", "Reports",
    "Settings", "Logout"
}

        For Each item As String In menuItems
            Dim mi As New ToolStripMenuItem(item)
            AddHandler mi.Click, AddressOf OnMenuClick
            ms.Items.Add(mi)
        Next

        Me.Controls.Add(ms)
        Me.MainMenuStrip = ms
    End Sub

    Private Sub OnMenuClick(sender As Object, e As EventArgs)
        Dim mi As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)

        Select Case mi.Text
            Case "Products"
                Dim f As New frmProducts() : f.Show()
            Case "Inventory"
                Dim f As New frmInventory() : f.Show()
            Case "Customers"
                Dim f As New frmCustomers() : f.Show()
            Case "Employees"
                Dim f As New frmAdminEmployee() : f.ShowDialog()
            Case "Suppliers"
                Dim f As New frmSuppliers() : f.Show()
            Case "Purchase Orders"
                Dim f As New frmPurchaseOrders() : f.Show()
            Case "Sales History"
                Dim f As New frmSalesHistory() : f.Show()
            Case "Reports"
                Dim f As New frmReports() : f.Show()
            Case "Expenses"
                Dim f As New frmExpenses() : f.Show()
            Case "Settings"
                Dim f As New frmSettings() : f.Show()
            Case "Logout"
                Dim confirm As DialogResult = MessageBox.Show(
                    "Are you sure you want to logout?", "Logout",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If confirm = DialogResult.Yes Then
                    LoggedUserID = 0
                    LoggedFullName = ""
                    LoggedRole = ""
                    Dim f As New frmLogin()
                    f.Show()
                    Me.Close()
                End If
        End Select
    End Sub

    Private Sub BuildUI()
        Dim lblWelcome As New Label() With {
            .Text = "Welcome, " & LoggedFullName,
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 40)
        }
        Me.Controls.Add(lblWelcome)

        Dim lblDate As New Label() With {
            .Text = DateTime.Now.ToString("dddd, MMMM dd yyyy"),
            .Font = New Font("Arial", 9),
            .ForeColor = Color.Gray,
            .AutoSize = True,
            .Location = New Point(20, 68)
        }
        Me.Controls.Add(lblDate)

        AddCard("Total Products", lblTotalProd, 20, 100, Color.FromArgb(63, 81, 181))
        AddCard("Today's Sales", lblTodaySales, 210, 100, Color.FromArgb(0, 150, 136))
        AddCard("Customers", lblTotalCust, 400, 100, Color.FromArgb(255, 152, 0))
        AddCard("Low Stock", lblLowCount, 590, 100, Color.FromArgb(244, 67, 54))

        Dim lbl1 As New Label() With {
            .Text = "Low Stock Items",
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 225)
        }
        Me.Controls.Add(lbl1)

        dgvLowStock.Location = New Point(20, 248)
        dgvLowStock.Size = New Size(550, 200)
        dgvLowStock.ReadOnly = True
        dgvLowStock.AllowUserToAddRows = False
        dgvLowStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.Controls.Add(dgvLowStock)

        Dim lbl2 As New Label() With {
            .Text = "Recent Sales",
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(600, 225)
        }
        Me.Controls.Add(lbl2)

        dgvRecentSales.Location = New Point(600, 248)
        dgvRecentSales.Size = New Size(550, 200)
        dgvRecentSales.ReadOnly = True
        dgvRecentSales.AllowUserToAddRows = False
        dgvRecentSales.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.Controls.Add(dgvRecentSales)
    End Sub

    Private Sub AddCard(title As String, numLbl As Label, x As Integer, y As Integer, col As Color)
        Dim pnl As New Panel() With {
            .Size = New Size(175, 95),
            .Location = New Point(x, y),
            .BackColor = col
        }

        numLbl.Text = "..."
        numLbl.ForeColor = Color.White
        numLbl.Font = New Font("Arial", 20, FontStyle.Bold)
        numLbl.Location = New Point(10, 10)
        numLbl.AutoSize = True
        pnl.Controls.Add(numLbl)

        Dim lbl As New Label() With {
            .Text = title,
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(10, 58)
        }
        pnl.Controls.Add(lbl)

        Me.Controls.Add(pnl)
    End Sub

    Private Sub LoadData()
        If Not OpenConnection() Then Return
        Try
            lblTotalProd.Text = New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_products WHERE status <> 'Discontinued'",
                conn).ExecuteScalar().ToString()

            lblTodaySales.Text = "P" & Convert.ToDecimal(New MySqlCommand("
    SELECT IFNULL(SUM((SELECT SUM((quantity * unit_price) - discount)
    FROM tbl_sale_details sd
    WHERE sd.sale_id = s.sale_id)),0)
    FROM tbl_sales s
    WHERE DATE(s.sale_date)=CURDATE()", conn).ExecuteScalar()).ToString("N2")

            lblTotalCust.Text = New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_users WHERE role = 'Customer'",
                conn).ExecuteScalar().ToString()

            lblLowCount.Text = New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_products " &
                "WHERE stock_quantity <= reorder_level AND status = 'In Stock'",
                conn).ExecuteScalar().ToString()

            Dim da1 As New MySqlDataAdapter(
                "SELECT product_name AS 'Product', stock_quantity AS 'Stock', " &
                "reorder_level AS 'Min Level' FROM tbl_products " &
                "WHERE stock_quantity <= reorder_level AND status = 'In Stock'", conn)
            Dim dt1 As New DataTable()
            da1.Fill(dt1)
            dgvLowStock.DataSource = dt1

            Dim da2 As New MySqlDataAdapter("
SELECT
s.receipt_no AS 'Receipt',
s.sale_date AS 'Date',
IFNULL(SUM((sd.quantity * sd.unit_price)-sd.discount),0) AS 'Total'
FROM tbl_sales s
LEFT JOIN tbl_sale_details sd
ON s.sale_id=sd.sale_id
GROUP BY s.sale_id
ORDER BY s.sale_date DESC
LIMIT 5", conn)
            Dim dt2 As New DataTable()
            da2.Fill(dt2)
            dgvRecentSales.DataSource = dt2

        Catch ex As Exception
            MessageBox.Show("Error loading dashboard: " & ex.Message)
        Finally
            CloseConnection()
        End Try
    End Sub

End Class