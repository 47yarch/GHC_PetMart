Imports MySql.Data.MySqlClient

Public Class frmRegister

    Private txtFN As New TextBox()
    Private txtLN As New TextBox()
    Private txtEmail As New TextBox()
    Private txtPhone As New TextBox()
    Private txtPass As New TextBox()
    Private txtConfirm As New TextBox()
    Private cboPet As New ComboBox()
    Private lblError As New Label()

    Private Sub InitRegister() Handles MyBase.Load
        Me.Text = "Create Account"
        Me.Size = New Size(420, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        Dim y As Integer = 20

        Dim lblTitle As New Label() With {
            .Text = "CREATE ACCOUNT",
            .Font = New Font("Arial", 16, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(105, y)
        }
        Me.Controls.Add(lblTitle)
        y += 50

        AddRow("First Name", txtFN, y) : y += 60
        AddRow("Last Name", txtLN, y) : y += 60
        AddRow("Email (must be @user.com)", txtEmail, y) : y += 60
        AddRow("Phone Number", txtPhone, y) : y += 60

        Dim lblPet As New Label() With {
            .Text = "Pet Type:",
            .AutoSize = True,
            .Location = New Point(50, y)
        }
        Me.Controls.Add(lblPet)
        y += 22

        cboPet.Location = New Point(50, y)
        cboPet.Size = New Size(300, 30)
        cboPet.DropDownStyle = ComboBoxStyle.DropDownList
        cboPet.Items.AddRange(New String() {"Dog", "Cat", "Fish", "Bird", "Other"})
        cboPet.SelectedIndex = 0
        Me.Controls.Add(cboPet)
        y += 55

        AddRow("Password", txtPass, y, True) : y += 60
        AddRow("Confirm Password", txtConfirm, y, True) : y += 60

        lblError.ForeColor = Color.Red
        lblError.Location = New Point(50, y)
        lblError.Size = New Size(300, 40)
        lblError.Visible = False
        Me.Controls.Add(lblError)
        y += 45

        Dim btnReg As New Button() With {
            .Text = "REGISTER",
            .Size = New Size(300, 42),
            .Location = New Point(50, y),
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 10, FontStyle.Bold)
        }
        AddHandler btnReg.Click, AddressOf DoRegister
        Me.Controls.Add(btnReg)
        y += 52

        Dim btnBack As New Button() With {
            .Text = "Back to Login",
            .Size = New Size(300, 35),
            .Location = New Point(50, y),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnBack.Click, AddressOf GoBack
        Me.Controls.Add(btnBack)
    End Sub

    Private Sub AddRow(label As String, txt As TextBox, y As Integer, Optional isPass As Boolean = False)
        Dim lbl As New Label() With {
            .Text = label & ":",
            .AutoSize = True,
            .Location = New Point(50, y)
        }
        Me.Controls.Add(lbl)

        txt.Size = New Size(300, 30)
        txt.Location = New Point(50, y + 22)
        txt.Font = New Font("Arial", 10)
        If isPass Then txt.PasswordChar = "*"c
        Me.Controls.Add(txt)
    End Sub

    Private Sub DoRegister(sender As Object, e As EventArgs)
        lblError.Visible = False

        If txtFN.Text.Trim() = "" OrElse txtLN.Text.Trim() = "" OrElse
           txtEmail.Text.Trim() = "" OrElse txtPass.Text.Trim() = "" Then
            ShowError("Please fill in all required fields.")
            Return
        End If

        If Not txtEmail.Text.Trim().ToLower().EndsWith("@user.com") Then
            ShowError("Email must end with @user.com")
            Return
        End If

        If txtPass.Text <> txtConfirm.Text Then
            ShowError("Passwords do not match.")
            Return
        End If

        If txtPass.Text.Length < 4 Then
            ShowError("Password must be at least 4 characters.")
            Return
        End If

        Try
            If Not OpenConnection() Then Return

            Dim chk As New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_users WHERE email=@e", conn)
            chk.Parameters.AddWithValue("@e", txtEmail.Text.Trim().ToLower())

            If Convert.ToInt32(chk.ExecuteScalar()) > 0 Then
                ShowError("Email is already registered.")
                CloseConnection()
                Return
            End If

            Dim c1 As New MySqlCommand(
                "INSERT INTO tbl_users(email, username, password, full_name, role) " &
                "VALUES(@e, @u, MD5(@p), @fn, 'Customer')", conn)
            c1.Parameters.AddWithValue("@e", txtEmail.Text.Trim().ToLower())
            c1.Parameters.AddWithValue("@u", txtFN.Text.Trim().ToLower() & txtLN.Text.Trim().ToLower())
            c1.Parameters.AddWithValue("@p", txtPass.Text.Trim())
            c1.Parameters.AddWithValue("@fn", txtFN.Text.Trim() & " " & txtLN.Text.Trim())
            c1.ExecuteNonQuery()

            Dim newID As Integer = Convert.ToInt32(
                New MySqlCommand("SELECT LAST_INSERT_ID()", conn).ExecuteScalar())

            Dim c2 As New MySqlCommand(
                "INSERT INTO tbl_customers(user_id, first_name, last_name, email, phone, pet_type) " &
                "VALUES(@uid, @fn, @ln, @e, @ph, @pet)", conn)
            c2.Parameters.AddWithValue("@uid", newID)
            c2.Parameters.AddWithValue("@fn", txtFN.Text.Trim())
            c2.Parameters.AddWithValue("@ln", txtLN.Text.Trim())
            c2.Parameters.AddWithValue("@e", txtEmail.Text.Trim().ToLower())
            c2.Parameters.AddWithValue("@ph", txtPhone.Text.Trim())
            c2.Parameters.AddWithValue("@pet", cboPet.Text)
            c2.ExecuteNonQuery()

            CloseConnection()

            MessageBox.Show("Account created successfully! You can now log in.",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Dim f As New frmLogin()
            f.Show()
            Me.Close()

        Catch ex As Exception
            ShowError("Error: " & ex.Message)
            CloseConnection()
        End Try
    End Sub

    Private Sub GoBack(sender As Object, e As EventArgs)
        Dim f As New frmLogin()
        f.Show()
        Me.Close()
    End Sub

    Private Sub ShowError(msg As String)
        lblError.Text = msg
        lblError.Visible = True
    End Sub

End Class