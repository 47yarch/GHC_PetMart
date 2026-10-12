Imports MySql.Data.MySqlClient

Public Class frmEmployee

    Private lblTodaySales As New Label()
    Private lblTodayTx As New Label()
    Private dgvRecent As New DataGridView()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "GHC Pet Mart - Employee"
        Me.Size = New Size(900, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        BuildMenu()
        BuildUI()
        LoadData()
    End Sub

    Private Sub BuildMenu()
        Dim ms As New MenuStrip()
        Dim items() As String = {"Dashboard", "Point of Sale", "View Inventory", "Logout"}
        For Each item As String In items
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
            Case "Point of Sale"
                Dim f As New frmPOS()
                f.Show()
            Case "View Inventory"
                Dim f As New frmEmpInventory()
                f.Show()
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

        AddCard("Today's Sales (P)", lblTodaySales, 20, 105, Color.FromArgb(0, 150, 136))
        AddCard("Transactions Today", lblTodayTx, 220, 105, Color.FromArgb(63, 81, 181))

        Dim btnPOS As New Button() With {
            .Text = "OPEN POINT OF SALE",
            .Size = New Size(250, 60),
            .Location = New Point(420, 105),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 12, FontStyle.Bold)
        }
        AddHandler btnPOS.Click, Sub(s, e)
                                     Dim f As New frmPOS()
                                     f.Show()
                                 End Sub
        Me.Controls.Add(btnPOS)

        Dim lblRec As New Label() With {
            .Text = "My Recent Sales",
            .Font = New Font("Arial", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 220)
        }
        Me.Controls.Add(lblRec)

        dgvRecent.Location = New Point(20, 245)
        dgvRecent.Size = New Size(840, 300)
        dgvRecent.ReadOnly = True
        dgvRecent.AllowUserToAddRows = False
        dgvRecent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvRecent.BackgroundColor = Color.White
        Me.Controls.Add(dgvRecent)
    End Sub

    Private Sub AddCard(title As String, numLbl As Label, x As Integer, y As Integer, col As Color)
        Dim pnl As New Panel() With {
            .Size = New Size(185, 90),
            .Location = New Point(x, y),
            .BackColor = col
        }
        numLbl.Text = "..."
        numLbl.ForeColor = Color.White
        numLbl.Font = New Font("Arial", 18, FontStyle.Bold)
        numLbl.Location = New Point(10, 10)
        numLbl.AutoSize = True
        pnl.Controls.Add(numLbl)

        Dim lbl As New Label() With {
            .Text = title,
            .ForeColor = Color.White,
            .AutoSize = True,
            .Location = New Point(10, 55)
        }
        pnl.Controls.Add(lbl)
        Me.Controls.Add(pnl)
    End Sub

    Private Sub LoadData()
        If Not OpenConnection() Then Return

        Try
            Dim cmdSales As New MySqlCommand(
"SELECT IFNULL(SUM(sd.quantity*sd.unit_price-sd.discount),0) " &
"FROM tbl_sales s " &
"JOIN tbl_sale_details sd ON s.sale_id = sd.sale_id " &
"WHERE s.employee_id=@uid " &
"AND DATE(s.sale_date)=CURDATE()", conn)

            cmdSales.Parameters.AddWithValue("@uid", LoggedUserID)

            lblTodaySales.Text = "P" &
            Convert.ToDecimal(cmdSales.ExecuteScalar()).ToString("N2")

            Dim cmdTx As New MySqlCommand(
            "SELECT COUNT(*) FROM tbl_sales " &
            "WHERE employee_id=@uid AND DATE(sale_date)=CURDATE()", conn)
            cmdTx.Parameters.AddWithValue("@uid", LoggedUserID)

            lblTodayTx.Text = cmdTx.ExecuteScalar().ToString()

            Dim da As New MySqlDataAdapter(
"SELECT
    s.receipt_no AS 'Receipt',
    s.sale_date AS 'Date',
    SUM(sd.quantity*sd.unit_price-sd.discount) AS 'Total',
    s.amount_paid AS 'Paid',
    (s.amount_paid -
     SUM(sd.quantity*sd.unit_price-sd.discount)) AS 'Change'
FROM tbl_sales s
JOIN tbl_sale_details sd
    ON s.sale_id = sd.sale_id
WHERE s.employee_id=" & LoggedUserID & "
GROUP BY
    s.sale_id,
    s.receipt_no,
    s.sale_date,
    s.amount_paid
ORDER BY s.sale_date DESC
LIMIT 50", conn)

            Dim dt As New DataTable()
            da.Fill(dt)

            dgvRecent.DataSource = dt

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)

        Finally
            CloseConnection()
        End Try
    End Sub

End Class