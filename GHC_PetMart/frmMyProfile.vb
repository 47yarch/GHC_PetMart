Imports MySql.Data.MySqlClient

Public Class frmMyProfile
    Inherits Form

    Private txtFirst As New TextBox()
    Private txtLast As New TextBox()
    Private txtEmail As New TextBox()
    Private txtPhone As New TextBox()
    Private txtAddress As New TextBox()
    Private txtCity As New TextBox()
    Private cboPet As New ComboBox()

    Private txtCurrentPass As New TextBox()
    Private txtNewPass As New TextBox()
    Private txtConfirmPass As New TextBox()

    Private chkShow As New CheckBox()

    Private btnSave As New Button()
    Private btnChangePass As New Button()
    Private btnClose As New Button()

    Private Sub frmMyProfile_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "My Profile"
        Me.Size = New Size(520, 760)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White
        Me.AutoScroll = True

        Dim y As Integer = 20

        Dim lblTitle As New Label() With {
            .Text = "MY PROFILE",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(165, y)
        }
        Me.Controls.Add(lblTitle)

        y += 50

        AddRow("First Name", txtFirst, y)
        y += 55

        AddRow("Last Name", txtLast, y)
        y += 55

        AddRow("Email", txtEmail, y)
        y += 55

        AddRow("Phone", txtPhone, y)
        y += 55

        AddRow("Address", txtAddress, y)
        y += 55

        AddRow("City", txtCity, y)
        y += 55

        Dim lblPet As New Label() With {
            .Text = "Pet Type:",
            .Location = New Point(25, y),
            .AutoSize = True
        }
        Me.Controls.Add(lblPet)

        cboPet.Location = New Point(25, y + 20)
        cboPet.Size = New Size(430, 28)
        cboPet.DropDownStyle = ComboBoxStyle.DropDownList
        cboPet.Items.AddRange(New String() {
            "Dog",
            "Cat",
            "Fish",
            "Bird",
            "Rabbit",
            "Other"
        })
        Me.Controls.Add(cboPet)

        y += 70

        Dim lblPass As New Label() With {
            .Text = "CHANGE PASSWORD",
            .Font = New Font("Arial", 12, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(25, y)
        }
        Me.Controls.Add(lblPass)

        y += 40

        AddRow("Current Password", txtCurrentPass, y)
        txtCurrentPass.PasswordChar = "*"c
        y += 55

        AddRow("New Password", txtNewPass, y)
        txtNewPass.PasswordChar = "*"c
        y += 55

        AddRow("Confirm Password", txtConfirmPass, y)
        txtConfirmPass.PasswordChar = "*"c
        y += 55

        chkShow.Text = "Show Password"
        chkShow.AutoSize = True
        chkShow.Location = New Point(25, y)
        AddHandler chkShow.CheckedChanged, AddressOf TogglePassword
        Me.Controls.Add(chkShow)

        y += 45

        btnSave.Text = "SAVE PROFILE"
        btnSave.Size = New Size(135, 40)
        btnSave.Location = New Point(25, y)
        btnSave.BackColor = Color.FromArgb(76, 175, 80)
        btnSave.ForeColor = Color.White
        btnSave.FlatStyle = FlatStyle.Flat
        AddHandler btnSave.Click, AddressOf SaveProfile
        Me.Controls.Add(btnSave)

        btnChangePass.Text = "CHANGE PASSWORD"
        btnChangePass.Size = New Size(170, 40)
        btnChangePass.Location = New Point(170, y)
        btnChangePass.BackColor = Color.FromArgb(63, 81, 181)
        btnChangePass.ForeColor = Color.White
        btnChangePass.FlatStyle = FlatStyle.Flat
        AddHandler btnChangePass.Click, AddressOf ChangePassword
        Me.Controls.Add(btnChangePass)

        btnClose.Text = "CLOSE"
        btnClose.Size = New Size(100, 40)
        btnClose.Location = New Point(350, y)
        btnClose.FlatStyle = FlatStyle.Flat
        AddHandler btnClose.Click,
            Sub()
                Me.Close()
            End Sub
        Me.Controls.Add(btnClose)

        LoadProfile()

    End Sub

    Private Sub AddRow(lbl As String, txt As TextBox, y As Integer)

        Dim l As New Label()

        l.Text = lbl & ":"
        l.AutoSize = True
        l.Location = New Point(25, y)

        Me.Controls.Add(l)

        txt.Location = New Point(25, y + 20)
        txt.Size = New Size(430, 28)

        Me.Controls.Add(txt)

    End Sub

    Private Sub TogglePassword(sender As Object, e As EventArgs)

        Dim ch As Char = If(chkShow.Checked, Chr(0), "*"c)

        txtCurrentPass.PasswordChar = ch
        txtNewPass.PasswordChar = ch
        txtConfirmPass.PasswordChar = ch

    End Sub

    Private Sub LoadProfile()

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand(
                "SELECT c.*, u.email
                 FROM tbl_customers c
                 INNER JOIN tbl_users u
                    ON c.user_id=u.user_id
                 WHERE c.user_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", LoggedUserID)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)

            If dt.Rows.Count = 0 Then
                MessageBox.Show("Unable to load your profile.")
                Exit Sub
            End If

            Dim r As DataRow = dt.Rows(0)

            txtFirst.Text = r("first_name").ToString()
            txtLast.Text = r("last_name").ToString()
            txtEmail.Text = r("email").ToString()
            txtPhone.Text = r("phone").ToString()
            txtAddress.Text = r("address").ToString()
            txtCity.Text = r("city").ToString()

            If Not IsDBNull(r("pet_type")) Then
                cboPet.Text = r("pet_type").ToString()
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Function PasswordCorrect() As Boolean

        If Not OpenConnection() Then Return False

        Try

            Dim cmd As New MySqlCommand(
                "SELECT COUNT(*)
                 FROM tbl_users
                 WHERE user_id=@id
                 AND password=MD5(@p)", conn)

            cmd.Parameters.AddWithValue("@id", LoggedUserID)
            cmd.Parameters.AddWithValue("@p", txtCurrentPass.Text.Trim())

            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0

        Catch

            Return False

        Finally

            CloseConnection()

        End Try

    End Function

    Private Function EmailExists(email As String) As Boolean

        If Not OpenConnection() Then Return False

        Try

            Dim cmd As New MySqlCommand(
                "SELECT COUNT(*)
                 FROM tbl_users
                 WHERE email=@e
                 AND user_id<>@id", conn)

            cmd.Parameters.AddWithValue("@e", email)
            cmd.Parameters.AddWithValue("@id", LoggedUserID)

            Return Convert.ToInt32(cmd.ExecuteScalar()) > 0

        Finally

            CloseConnection()

        End Try

    End Function

    Private Sub SaveProfile(sender As Object, e As EventArgs)

        If txtFirst.Text.Trim() = "" OrElse txtLast.Text.Trim() = "" Then
            MessageBox.Show("First Name and Last Name are required.")
            Exit Sub
        End If

        If txtEmail.Text.Trim() = "" Then
            MessageBox.Show("Email is required.")
            Exit Sub
        End If

        If EmailExists(txtEmail.Text.Trim()) Then
            MessageBox.Show("That email is already being used.")
            Exit Sub
        End If

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand(
                "UPDATE tbl_customers SET
                    first_name=@fn,
                    last_name=@ln,
                    email=@em,
                    phone=@ph,
                    address=@ad,
                    city=@cy,
                    pet_type=@pt,
                    updated_at=NOW()
                 WHERE user_id=@id", conn)

            cmd.Parameters.AddWithValue("@fn", txtFirst.Text.Trim())
            cmd.Parameters.AddWithValue("@ln", txtLast.Text.Trim())
            cmd.Parameters.AddWithValue("@em", txtEmail.Text.Trim())
            cmd.Parameters.AddWithValue("@ph", txtPhone.Text.Trim())
            cmd.Parameters.AddWithValue("@ad", txtAddress.Text.Trim())
            cmd.Parameters.AddWithValue("@cy", txtCity.Text.Trim())
            cmd.Parameters.AddWithValue("@pt", cboPet.Text)
            cmd.Parameters.AddWithValue("@id", LoggedUserID)

            cmd.ExecuteNonQuery()

            Dim cmdUser As New MySqlCommand(
                "UPDATE tbl_users
                 SET
                    full_name=@name,
                    email=@email
                 WHERE user_id=@id", conn)

            cmdUser.Parameters.AddWithValue("@name",
                txtFirst.Text.Trim() & " " & txtLast.Text.Trim())

            cmdUser.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
            cmdUser.Parameters.AddWithValue("@id", LoggedUserID)

            cmdUser.ExecuteNonQuery()

            LoggedFullName = txtFirst.Text.Trim() & " " & txtLast.Text.Trim()

            MessageBox.Show(
                "Profile updated successfully.",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub ChangePassword(sender As Object, e As EventArgs)

        If txtCurrentPass.Text.Trim() = "" Then
            MessageBox.Show("Enter your current password.")
            Exit Sub
        End If

        If txtNewPass.Text.Trim() = "" Then
            MessageBox.Show("Enter a new password.")
            Exit Sub
        End If

        If txtConfirmPass.Text.Trim() = "" Then
            MessageBox.Show("Confirm your new password.")
            Exit Sub
        End If

        If txtNewPass.Text <> txtConfirmPass.Text Then
            MessageBox.Show("New passwords do not match.")
            Exit Sub
        End If

        If Not PasswordCorrect() Then
            MessageBox.Show("Current password is incorrect.")
            Exit Sub
        End If

        If Not OpenConnection() Then Exit Sub

        Try

            Dim cmd As New MySqlCommand(
                "UPDATE tbl_users
                 SET password=MD5(@pass)
                 WHERE user_id=@id", conn)

            cmd.Parameters.AddWithValue("@pass", txtNewPass.Text.Trim())
            cmd.Parameters.AddWithValue("@id", LoggedUserID)

            cmd.ExecuteNonQuery()

            MessageBox.Show(
                "Password changed successfully.",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

            txtCurrentPass.Clear()
            txtNewPass.Clear()
            txtConfirmPass.Clear()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        Me.ClientSize = New System.Drawing.Size(282, 253)
        Me.Name = "frmMyProfile"
        Me.ResumeLayout(False)

    End Sub

    Private Sub frmMyProfile_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class