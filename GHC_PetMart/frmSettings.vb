Imports MySql.Data.MySqlClient

Public Class frmSettings

    Private dgv As New DataGridView()
    Private txtSearch As New TextBox()
    Private cboRole As New ComboBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "User Management"
        Me.Size = New Size(950, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "User Account Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lS As New Label() With {.Text = "Search:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lS)
        txtSearch.Size = New Size(200, 28) : txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim lR As New Label() With {.Text = "Role:", .AutoSize = True, .Location = New Point(280, 55)}
        Me.Controls.Add(lR)
        cboRole.Location = New Point(315, 52) : cboRole.Size = New Size(130, 28)
        cboRole.DropDownStyle = ComboBoxStyle.DropDownList
        cboRole.Items.AddRange(New String() {"All", "Admin", "Employee", "Customer"})
        cboRole.SelectedIndex = 0
        Me.Controls.Add(cboRole)

        Dim btnLoad As New Button() With {
            .Text = "Load", .Size = New Size(75, 28), .Location = New Point(460, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnLoad.Click, AddressOf LoadUsers
        Me.Controls.Add(btnLoad)

        Dim btnToggle As New Button() With {
            .Text = "Activate/Deactivate", .Size = New Size(160, 32), .Location = New Point(15, 92),
            .BackColor = Color.FromArgb(255, 152, 0), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnToggle.Click, AddressOf ToggleUser
        Me.Controls.Add(btnToggle)

        Dim btnReset As New Button() With {
            .Text = "Reset Password to '1234'", .Size = New Size(190, 32), .Location = New Point(185, 92),
            .BackColor = Color.FromArgb(244, 67, 54), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnReset.Click, AddressOf ResetPassword
        Me.Controls.Add(btnReset)

        dgv.Location = New Point(15, 135)
        dgv.Size = New Size(910, 400)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)

        LoadUsers(Nothing, Nothing)
    End Sub

    Private Sub LoadUsers(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim search As String = "%" & txtSearch.Text.Trim() & "%"
            Dim sql As String =
                "SELECT user_id, email AS 'Email', full_name AS 'Name', role AS 'Role', " &
                "CASE WHEN is_active=1 THEN 'Active' ELSE 'Inactive' END AS 'Status', " &
                "created_at AS 'Created' FROM tbl_users WHERE (email LIKE @s OR full_name LIKE @s)"
            If cboRole.Text <> "All" Then sql &= " AND role=@r"
            sql &= " ORDER BY role, full_name"

            Dim cmd As New MySqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@s", search)
            If cboRole.Text <> "All" Then cmd.Parameters.AddWithValue("@r", cboRole.Text)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgv.DataSource = dt
            dgv.Columns("user_id").Visible = False
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub ToggleUser(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then MessageBox.Show("Select a user.") : Return
        Dim id As Integer = CInt(dgv.SelectedRows(0).Cells("user_id").Value)
        Dim status As String = dgv.SelectedRows(0).Cells("Status").Value.ToString()
        Dim newStatus As Integer = If(status = "Active", 0, 1)
        If Not OpenConnection() Then Return
        Dim cmd As New MySqlCommand("UPDATE tbl_users SET is_active=@s WHERE user_id=@id", conn)
        cmd.Parameters.AddWithValue("@s", newStatus)
        cmd.Parameters.AddWithValue("@id", id)
        cmd.ExecuteNonQuery()
        CloseConnection()
        LoadUsers(Nothing, Nothing)
    End Sub

    Private Sub ResetPassword(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then MessageBox.Show("Select a user.") : Return
        Dim id As Integer = CInt(dgv.SelectedRows(0).Cells("user_id").Value)
        If MessageBox.Show("Reset this user's password to '1234'?", "Confirm",
                           MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If Not OpenConnection() Then Return
            Dim cmd As New MySqlCommand("UPDATE tbl_users SET password=MD5('1234') WHERE user_id=@id", conn)
            cmd.Parameters.AddWithValue("@id", id)
            cmd.ExecuteNonQuery()
            CloseConnection()
            MessageBox.Show("Password reset to '1234'.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class