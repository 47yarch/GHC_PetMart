Imports MySql.Data.MySqlClient

Public Class frmCustomers

    Private dgv As New DataGridView()
    Private txtSearch As New TextBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Customer Management"
        Me.Size = New Size(950, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke
        BuildUI()
        LoadCustomers()
    End Sub

    Private Sub BuildUI()
        Dim lblTitle As New Label() With {
            .Text = "Customer Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lS As New Label() With {
            .Text = "Search:",
            .AutoSize = True,
            .Location = New Point(15, 55)
        }
        Me.Controls.Add(lS)

        txtSearch.Size = New Size(250, 28)
        txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim btnS As New Button() With {
            .Text = "Search",
            .Size = New Size(80, 28),
            .Location = New Point(325, 52),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnS.Click, Sub(s, e) LoadCustomers()
        Me.Controls.Add(btnS)

        Dim btnAdd As New Button() With {
            .Text = "Add Customer",
            .Size = New Size(120, 32),
            .Location = New Point(15, 92),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdd.Click, AddressOf AddCustomer
        Me.Controls.Add(btnAdd)

        Dim btnEdit As New Button() With {
            .Text = "Edit",
            .Size = New Size(80, 32),
            .Location = New Point(145, 92),
            .BackColor = Color.FromArgb(255, 152, 0),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnEdit.Click, AddressOf EditCustomer
        Me.Controls.Add(btnEdit)

        Dim btnDel As New Button() With {
            .Text = "Delete",
            .Size = New Size(80, 32),
            .Location = New Point(235, 92),
            .BackColor = Color.FromArgb(244, 67, 54),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDel.Click, AddressOf DeleteCustomer
        Me.Controls.Add(btnDel)

        Dim btnHistory As New Button() With {
            .Text = "Purchase History",
            .Size = New Size(130, 32),
            .Location = New Point(325, 92),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnHistory.Click, AddressOf ViewHistory
        Me.Controls.Add(btnHistory)

        dgv.Location = New Point(15, 135)
        dgv.Size = New Size(910, 400)
        dgv.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or
                     AnchorStyles.Left Or AnchorStyles.Right
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)
    End Sub

    Public Sub LoadCustomers()
        Try
            Using con As MySqlConnection = GetConn()
                Dim search As String = "%" & txtSearch.Text.Trim() & "%"
                Dim cmd As New MySqlCommand(
                    "SELECT customer_id, " &
                    "CONCAT(first_name, ' ', last_name) AS 'Name', " &
                    "email AS 'Email', phone AS 'Phone', " &
                    "city AS 'City', pet_type AS 'Pet Type', " &
                    "created_at AS 'Registered' " &
                    "FROM tbl_customers " &
                    "WHERE CONCAT(first_name,' ',last_name) LIKE @s " &
                    "OR email LIKE @s OR phone LIKE @s " &
                    "ORDER BY first_name", con)
                cmd.Parameters.AddWithValue("@s", search)

                Dim dt As New DataTable()
                Dim da As New MySqlDataAdapter(cmd)
                da.Fill(dt)
                dgv.DataSource = dt

                If dgv.Columns.Contains("customer_id") Then
                    dgv.Columns("customer_id").Visible = False
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading customers: " & ex.Message)
        End Try
    End Sub

    Private Sub AddCustomer(sender As Object, e As EventArgs)
        Dim f As New frmCustomerForm()
        f.IsEditMode = False
        f.ShowDialog()
        LoadCustomers()
    End Sub

    Private Sub EditCustomer(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Select a customer to edit.") : Return
        End If
        Dim f As New frmCustomerForm()
        f.IsEditMode = True
        f.EditCustomerID = CInt(dgv.SelectedRows(0).Cells("customer_id").Value)
        f.ShowDialog()
        LoadCustomers()
    End Sub

    Private Sub DeleteCustomer(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Select a customer to delete.") : Return
        End If
        Dim name As String = dgv.SelectedRows(0).Cells("Name").Value.ToString()
        Dim id As Integer = CInt(dgv.SelectedRows(0).Cells("customer_id").Value)

        If MessageBox.Show("Delete customer '" & name & "'?", "Confirm Delete",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning) = DialogResult.No Then Return
        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand(
                    "DELETE FROM tbl_customers WHERE customer_id = @id", con)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
            End Using
            LoadCustomers()
        Catch ex As Exception
            MessageBox.Show("Cannot delete — customer has existing orders." &
                            vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub ViewHistory(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Select a customer.") : Return
        End If
        Dim custID As Integer = CInt(dgv.SelectedRows(0).Cells("customer_id").Value)
        Dim custName As String = dgv.SelectedRows(0).Cells("Name").Value.ToString()

        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand(
    "SELECT
        s.receipt_no AS 'Receipt',
        s.sale_date AS 'Date',

        SUM(sd.quantity * sd.unit_price - sd.discount)
            AS 'Total (P)',

        s.amount_paid AS 'Paid (P)',

        (
            s.amount_paid -
            SUM(sd.quantity * sd.unit_price - sd.discount)
        ) AS 'Change (P)',

        u.full_name AS 'Processed By'

    FROM tbl_sales s

    LEFT JOIN tbl_users u
        ON s.employee_id = u.user_id

    LEFT JOIN tbl_sale_details sd
        ON s.sale_id = sd.sale_id

    WHERE s.customer_id = @cid

    GROUP BY
        s.sale_id,
        s.receipt_no,
        s.sale_date,
        s.amount_paid,
        u.full_name

    ORDER BY s.sale_date DESC", con)
                cmd.Parameters.AddWithValue("@cid", custID)

                Dim dt As New DataTable()
                Dim da As New MySqlDataAdapter(cmd)
                da.Fill(dt)

                Dim frm As New Form() With {
                    .Text = "Purchase History — " & custName,
                    .Size = New Size(800, 450),
                    .StartPosition = FormStartPosition.CenterScreen
                }
                Dim g As New DataGridView() With {
                    .Dock = DockStyle.Fill,
                    .DataSource = dt,
                    .ReadOnly = True,
                    .AllowUserToAddRows = False,
                    .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                }
                frm.Controls.Add(g)
                frm.ShowDialog()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

End Class