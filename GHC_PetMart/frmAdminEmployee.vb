Imports MySql.Data.MySqlClient
Imports System.Security.Cryptography
Imports System.Text

Public Class frmAdminEmployee

    Private dgvEmployees As New DataGridView()

    Private txtSearch As New TextBox()

    Private txtFullName As New TextBox()
    Private txtUsername As New TextBox()
    Private txtEmail As New TextBox()
    Private txtPassword As New TextBox()

    Private cboRole As New ComboBox()

    Private chkActive As New CheckBox()

    Private btnAdd As New Button()
    Private btnUpdate As New Button()
    Private btnDelete As New Button()
    Private btnClear As New Button()

    Private lblTotal As New Label()

    Private selectedID As Integer = 0

    Private Function GetMD5(input As String) As String

        Using md5 As MD5 = MD5.Create()

            Dim bytes = Encoding.UTF8.GetBytes(input)

            Dim hash = md5.ComputeHash(bytes)

            Dim sb As New StringBuilder()

            For Each b As Byte In hash
                sb.Append(b.ToString("x2"))
            Next

            Return sb.ToString()

        End Using
    End Function


    Private Sub frmAdminEmployee_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Employee Management"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Size = New Size(1200, 700)
        Me.BackColor = Color.WhiteSmoke

        BuildUI()

        LoadEmployees()

    End Sub

    Private Sub BuildUI()

        Dim lblTitle As New Label With {
                .Text = "Employee Management",
                .Font = New Font("Arial", 15, FontStyle.Bold),
                .AutoSize = True,
                .Location = New Point(20, 20)
            }

        Me.Controls.Add(lblTitle)

        Dim lblSearch As New Label With {
                .Text = "Search",
                .AutoSize = True,
                .Location = New Point(20, 65)
            }

        Me.Controls.Add(lblSearch)

        txtSearch.Location = New Point(80, 60)
        txtSearch.Size = New Size(250, 28)

        AddHandler txtSearch.TextChanged, AddressOf txtSearch_TextChanged

        Me.Controls.Add(txtSearch)

        lblTotal.Font = New Font("Arial", 10, FontStyle.Bold)
        lblTotal.AutoSize = True
        lblTotal.Location = New Point(360, 64)

        Me.Controls.Add(lblTotal)

        Dim pnl As New Panel()

        pnl.Location = New Point(20, 100)
        pnl.Size = New Size(350, 520)
        pnl.BackColor = Color.White
        pnl.BorderStyle = BorderStyle.FixedSingle

        Me.Controls.Add(pnl)

        Dim y As Integer = 20

        Dim lblFN As New Label With {
                .Text = "Full Name",
                .Location = New Point(20, y),
                .AutoSize = True
            }

        pnl.Controls.Add(lblFN)

        y += 22

        txtFullName.Location = New Point(20, y)
        txtFullName.Size = New Size(300, 28)

        pnl.Controls.Add(txtFullName)

        y += 45


        Dim lblEmail As New Label With {
                .Text = "Email",
                .Location = New Point(20, y),
                .AutoSize = True
            }

        pnl.Controls.Add(lblEmail)

        y += 22

        txtEmail.Location = New Point(20, y)
        txtEmail.Size = New Size(300, 28)

        pnl.Controls.Add(txtEmail)

        y += 45


        Dim lblUser As New Label With {
                .Text = "Username",
                .Location = New Point(20, y),
                .AutoSize = True
            }

        pnl.Controls.Add(lblUser)

        y += 22

        txtUsername.Location = New Point(20, y)
        txtUsername.Size = New Size(300, 28)

        pnl.Controls.Add(txtUsername)

        y += 45


        Dim lblPass As New Label With {
                .Text = "Password",
                .Location = New Point(20, y),
                .AutoSize = True
            }

        pnl.Controls.Add(lblPass)

        y += 22

        txtPassword.Location = New Point(20, y)
        txtPassword.Size = New Size(300, 28)
        txtPassword.UseSystemPasswordChar = True

        pnl.Controls.Add(txtPassword)

        y += 45

        Dim lblRole As New Label With {
                .Text = "Role",
                .Location = New Point(20, y),
                .AutoSize = True
            }

        pnl.Controls.Add(lblRole)

        y += 22

        cboRole.Location = New Point(20, y)
        cboRole.Size = New Size(300, 28)
        cboRole.DropDownStyle = ComboBoxStyle.DropDownList

        cboRole.Items.Add("Admin")
        cboRole.Items.Add("Employee")

        pnl.Controls.Add(cboRole)

        y += 45

        chkActive.Text = "Active"
        chkActive.Location = New Point(20, y)
        chkActive.Checked = True

        pnl.Controls.Add(chkActive)

        y += 45


        btnAdd.Text = "Add"
        btnAdd.Size = New Size(140, 40)
        btnAdd.Location = New Point(20, y)
        btnAdd.BackColor = Color.FromArgb(63, 81, 181)
        btnAdd.ForeColor = Color.White
        btnAdd.FlatStyle = FlatStyle.Flat

        AddHandler btnAdd.Click, AddressOf btnAdd_Click

        pnl.Controls.Add(btnAdd)

        btnUpdate.Text = "Update"
        btnUpdate.Size = New Size(140, 40)
        btnUpdate.Location = New Point(180, y)
        btnUpdate.BackColor = Color.FromArgb(0, 150, 136)
        btnUpdate.ForeColor = Color.White
        btnUpdate.FlatStyle = FlatStyle.Flat

        AddHandler btnUpdate.Click, AddressOf btnUpdate_Click

        pnl.Controls.Add(btnUpdate)

        y += 55

        btnDelete.Text = "Delete"
        btnDelete.Size = New Size(140, 40)
        btnDelete.Location = New Point(20, y)
        btnDelete.BackColor = Color.Firebrick
        btnDelete.ForeColor = Color.White
        btnDelete.FlatStyle = FlatStyle.Flat

        AddHandler btnDelete.Click, AddressOf btnDelete_Click

        pnl.Controls.Add(btnDelete)

        btnClear.Text = "Clear"
        btnClear.Size = New Size(140, 40)
        btnClear.Location = New Point(180, y)
        btnClear.BackColor = Color.Gray
        btnClear.ForeColor = Color.White
        btnClear.FlatStyle = FlatStyle.Flat

        AddHandler btnClear.Click, AddressOf btnClear_Click

        pnl.Controls.Add(btnClear)


        dgvEmployees.Location = New Point(390, 100)
        dgvEmployees.Size = New Size(780, 520)

        dgvEmployees.AllowUserToAddRows = False
        dgvEmployees.AllowUserToDeleteRows = False
        dgvEmployees.ReadOnly = True
        dgvEmployees.MultiSelect = False
        dgvEmployees.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvEmployees.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvEmployees.BackgroundColor = Color.White

        AddHandler dgvEmployees.CellClick, AddressOf dgvEmployees_CellClick

        Me.Controls.Add(dgvEmployees)

    End Sub


    Private Sub LoadEmployees()

        If Not OpenConnection() Then Exit Sub

        Try

            Dim sql As String =
    "SELECT
        user_id AS 'ID',
        full_name AS 'Full Name',
        email AS 'Email',
        username AS 'Username',
        role AS 'Role',
        CASE
            WHEN is_active = 1 THEN 'Active'
            ELSE 'Inactive'
        END AS 'Status',
        created_at AS 'Created'
    FROM tbl_users
    WHERE role IN ('Admin','Employee')
    ORDER BY full_name"

            Dim da As New MySqlDataAdapter(sql, conn)
            Dim dt As New DataTable()

            da.Fill(dt)

            dgvEmployees.DataSource = dt

            dgvEmployees.Columns("ID").Visible = False

            lblTotal.Text = "Employees : " & dt.Rows.Count

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs)

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand("
    SELECT
        user_id AS 'ID',
        full_name AS 'Full Name',
        email AS 'Email',
        username AS 'Username',
        role AS 'Role',
        CASE
            WHEN is_active = 1 THEN 'Active'
            ELSE 'Inactive'
        END AS 'Status',
        created_at AS 'Created'
    FROM tbl_users
    WHERE
    (
        full_name LIKE @search
        OR username LIKE @search
        OR email LIKE @search
    )
    AND role IN ('Admin','Employee')
    ORDER BY full_name", conn)

            cmd.Parameters.AddWithValue("@search", "%" & txtSearch.Text.Trim() & "%")

            Dim dt As New DataTable()

            Dim da As New MySqlDataAdapter(cmd)

            da.Fill(dt)

            dgvEmployees.DataSource = dt

            dgvEmployees.Columns("ID").Visible = False

            lblTotal.Text = "Employees : " & dt.Rows.Count

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub dgvEmployees_CellClick(sender As Object, e As DataGridViewCellEventArgs)

        If e.RowIndex < 0 Then Exit Sub

        Dim row As DataGridViewRow = dgvEmployees.Rows(e.RowIndex)

        selectedID = CInt(row.Cells("ID").Value)

        txtFullName.Text = row.Cells("Full Name").Value.ToString()
        txtEmail.Text = row.Cells("Email").Value.ToString()
        txtUsername.Text = row.Cells("Username").Value.ToString()

        txtPassword.Clear()

        cboRole.Text = row.Cells("Role").Value.ToString()

        chkActive.Checked =
                row.Cells("Status").Value.ToString() = "Active"

    End Sub


    Private Sub ClearFields()

        selectedID = 0

        txtFullName.Clear()
        txtEmail.Clear()
        txtUsername.Clear()
        txtPassword.Clear()

        cboRole.SelectedIndex = -1

        chkActive.Checked = True

        txtFullName.Focus()

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)

        ClearFields()

    End Sub

    Private Function ValidateInput() As Boolean

        If txtFullName.Text.Trim = "" Then
            MessageBox.Show("Please enter the employee's full name.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            txtFullName.Focus()
            Return False
        End If

        If txtEmail.Text.Trim = "" Then
            MessageBox.Show("Please enter an email address.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            txtEmail.Focus()
            Return False
        End If

        If txtUsername.Text.Trim = "" Then
            MessageBox.Show("Please enter a username.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            txtUsername.Focus()
            Return False
        End If

        If txtPassword.Text.Trim = "" Then
            MessageBox.Show("Please enter a password.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            txtPassword.Focus()
            Return False
        End If

        If cboRole.SelectedIndex = -1 Then
            MessageBox.Show("Please select a role.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
            cboRole.Focus()
            Return False
        End If

        Return True

    End Function


    Private Function EmployeeExists() As Boolean

        If Not OpenConnection() Then Return True

        Try

            Dim cmd As New MySqlCommand("
    SELECT COUNT(*)
    FROM tbl_users
    WHERE username=@username
    OR email=@email", conn)

            cmd.Parameters.AddWithValue("@username", txtUsername.Text.Trim())
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim())

            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0

        Finally

            CloseConnection()

        End Try

    End Function

    Private Sub btnAdd_Click(sender As Object, e As EventArgs)

        If Not ValidateInput() Then Exit Sub

        If EmployeeExists() Then

            MessageBox.Show("Username or Email already exists.",
                                "Duplicate Entry",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

            Exit Sub

        End If

        If Not OpenConnection() Then Exit Sub

        Try
            Dim cmd As New MySqlCommand("
    INSERT INTO tbl_users
    (
        email,
        username,
        password,
        full_name,
        role,
        is_active
    )
    VALUES
    (
        @email,
        @username,
        MD5(@password),
        @fullname,
        @role,
        @active
    )", conn)

            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
            cmd.Parameters.AddWithValue("@username", txtUsername.Text.Trim())
            cmd.Parameters.AddWithValue("@password", txtPassword.Text.Trim())

            cmd.Parameters.AddWithValue("@fullname", txtFullName.Text.Trim())
            cmd.Parameters.AddWithValue("@role", cboRole.Text)
            cmd.Parameters.AddWithValue("@active", chkActive.Checked)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Employee added successfully.",
                                "Success",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            LoadEmployees()
            ClearFields()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Function EmployeeExistsForUpdate() As Boolean

        If Not OpenConnection() Then Return True

        Try

            Dim cmd As New MySqlCommand("
    SELECT COUNT(*)
    FROM tbl_users
    WHERE
    (
        username=@username
        OR email=@email
    )
    AND user_id<>@id", conn)

            cmd.Parameters.AddWithValue("@username", txtUsername.Text.Trim())
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
            cmd.Parameters.AddWithValue("@id", selectedID)

            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0

        Finally

            CloseConnection()

        End Try

    End Function

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs)

        If selectedID = 0 Then

            MessageBox.Show("Please select an employee first.",
                                "Update",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
            Exit Sub

        End If

        If txtFullName.Text.Trim = "" _
                OrElse txtEmail.Text.Trim = "" _
                OrElse txtUsername.Text.Trim = "" _
                OrElse cboRole.SelectedIndex = -1 Then

            MessageBox.Show("Please complete all required fields.",
                                "Validation",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

            Exit Sub

        End If

        If EmployeeExistsForUpdate() Then

            MessageBox.Show("Username or Email already exists.",
                                "Duplicate",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)

            Exit Sub

        End If

        If MessageBox.Show("Update this employee?",
                               "Confirm Update",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) = DialogResult.No Then
            Exit Sub
        End If

        If Not OpenConnection() Then Exit Sub

        Try

            Dim sql As String

            If txtPassword.Text.Trim = "" Then

                sql =
        "UPDATE tbl_users
    SET
    full_name=@fullname,
    email=@email,
    username=@username,
    role=@role,
    is_active=@active
    WHERE user_id=@id"

            Else

                sql =
        "UPDATE tbl_users
    SET
    full_name=@fullname,
    email=@email,
    username=@username,
    password=MD5(@password),
    role=@role,
    is_active=@active
    WHERE user_id=@id"

            End If

            Dim cmd As New MySqlCommand(sql, conn)

            cmd.Parameters.AddWithValue("@fullname", txtFullName.Text.Trim())
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
            cmd.Parameters.AddWithValue("@username", txtUsername.Text.Trim())
            cmd.Parameters.AddWithValue("@role", cboRole.Text)
            cmd.Parameters.AddWithValue("@active", chkActive.Checked)
            cmd.Parameters.AddWithValue("@id", selectedID)

            If txtPassword.Text.Trim <> "" Then
                cmd.Parameters.AddWithValue("@password", txtPassword.Text.Trim())
            End If

            cmd.ExecuteNonQuery()

            MessageBox.Show("Employee updated successfully.",
                                "Success",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            LoadEmployees()
            ClearFields()

        Catch ex As Exception

            MessageBox.Show(ex.Message,
                                "Database Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub btnDelete_Click(sender As Object, e As EventArgs)

        If selectedID = 0 Then

            MessageBox.Show("Please select an employee first.",
                                "Deactivate Employee",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            Exit Sub

        End If

        If MessageBox.Show(
                "Deactivate this employee account?" & vbCrLf &
                vbCrLf &
                "The employee will no longer be able to login.",
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) = DialogResult.No Then

            Exit Sub

        End If

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand("
    UPDATE tbl_users
    SET is_active = 0
    WHERE user_id = @id", conn)

            cmd.Parameters.AddWithValue("@id", selectedID)

            cmd.ExecuteNonQuery()

            MessageBox.Show("Employee account has been deactivated.",
                                "Success",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)

            LoadEmployees()
            ClearFields()

        Catch ex As Exception

            MessageBox.Show(ex.Message,
                                "Database Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub ActivateEmployee()

        If selectedID = 0 Then Exit Sub

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand("
    UPDATE tbl_users
    SET is_active = 1
    WHERE user_id = @id", conn)

            cmd.Parameters.AddWithValue("@id", selectedID)

            cmd.ExecuteNonQuery()

            LoadEmployees()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub


    Private Sub RefreshEmployees()

        LoadEmployees()

    End Sub


    Private Sub frmAdminEmployee_FormClosed(sender As Object,
                                                e As FormClosedEventArgs) _
                                                Handles Me.FormClosed

        selectedID = 0

    End Sub

End Class