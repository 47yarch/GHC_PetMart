Imports MySql.Data.MySqlClient

Public Class frmLogin

    Private txtEmail As New TextBox()
    Private txtPassword As New TextBox()
    Private lblError As New Label()
    Private chkShow As New CheckBox()

    Private Sub InitLogin() Handles MyBase.Load
        Me.Text = "GHC Pet Mart - Login"
        Me.Size = New Size(420, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        Dim lblTitle As New Label() With {
            .Text = "GHC PET MART",
            .Font = New Font("Arial", 18, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(110, 30)
        }
        Me.Controls.Add(lblTitle)

        Dim lblSub As New Label() With {
            .Text = "Sign in to your account",
            .Font = New Font("Arial", 9),
            .ForeColor = Color.Gray,
            .AutoSize = True,
            .Location = New Point(130, 65)
        }
        Me.Controls.Add(lblSub)

        Dim lblE As New Label() With {
            .Text = "Email:",
            .AutoSize = True,
            .Location = New Point(50, 110)
        }
        Me.Controls.Add(lblE)

        txtEmail.Size = New Size(300, 30)
        txtEmail.Location = New Point(50, 130)
        txtEmail.Font = New Font("Arial", 10)
        Me.Controls.Add(txtEmail)

        Dim lblP As New Label() With {
            .Text = "Password:",
            .AutoSize = True,
            .Location = New Point(50, 175)
        }
        Me.Controls.Add(lblP)

        txtPassword.Size = New Size(300, 30)
        txtPassword.Location = New Point(50, 195)
        txtPassword.Font = New Font("Arial", 10)
        txtPassword.PasswordChar = "*"c
        Me.Controls.Add(txtPassword)

        chkShow.Text = "Show Password"
        chkShow.Location = New Point(50, 235)
        chkShow.AutoSize = True
        AddHandler chkShow.CheckedChanged, AddressOf TogglePassword
        Me.Controls.Add(chkShow)

        lblError.ForeColor = Color.Red
        lblError.Location = New Point(50, 265)
        lblError.Size = New Size(300, 40)
        lblError.Visible = False
        Me.Controls.Add(lblError)

        Dim btnLogin As New Button() With {
            .Text = "LOGIN",
            .Size = New Size(300, 45),
            .Location = New Point(50, 310),
            .BackColor = Color.FromArgb(63, 81, 181),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Arial", 11, FontStyle.Bold)
        }
        AddHandler btnLogin.Click, AddressOf DoLogin
        Me.Controls.Add(btnLogin)

        Dim lnkReg As New LinkLabel() With {
            .Text = "Don't have an account? Register here",
            .AutoSize = True,
            .Location = New Point(90, 368)
        }
        AddHandler lnkReg.LinkClicked, AddressOf GoRegister
        Me.Controls.Add(lnkReg)
    End Sub

    Private Sub TogglePassword(sender As Object, e As EventArgs)
        txtPassword.PasswordChar = If(chkShow.Checked, Chr(0), "*"c)
    End Sub

    Private Sub GoRegister(sender As Object, e As EventArgs)
        Dim f As New frmRegister()
        f.Show()
        Me.Hide()
    End Sub

    Private Sub DoLogin(sender As Object, e As EventArgs)
        lblError.Visible = False

        Dim email As String = txtEmail.Text.Trim().ToLower()
        Dim pass As String = txtPassword.Text.Trim()

        If email = "" OrElse pass = "" Then
            ShowError("Please enter email and password.")
            Return
        End If

        If Not email.Contains("@") Then
            ShowError("Please enter a valid email.")
            Return
        End If

        Dim user As DataRow = GetUser(email, pass)

        If user Is Nothing Then
            ShowError("Incorrect email or password.")
            txtPassword.Clear()
            Return
        End If

        If Convert.ToInt32(user("is_active")) = 0 Then
            ShowError("Account is deactivated. Contact admin.")
            Return
        End If

        LoggedUserID = Convert.ToInt32(user("user_id"))
        LoggedFullName = user("full_name").ToString()
        LoggedRole = user("role").ToString()

        If email.EndsWith("@admin.com") AndAlso LoggedRole = "Admin" Then
            Dim f As New frmMain()
            f.Show()
            Me.Hide()

        ElseIf email.EndsWith("@employee.com") AndAlso LoggedRole = "Employee" Then
            Dim f As New frmEmployee()
            f.Show()
            Me.Hide()

        ElseIf email.EndsWith("@user.com") AndAlso LoggedRole = "Customer" Then
            Dim f As New frmShop()
            f.Show()
            Me.Hide()

        Else
            ShowError("Unauthorized. Check your email domain.")
        End If
    End Sub

    Private Function GetUser(email As String, pass As String) As DataRow
        Try
            If Not OpenConnection() Then Return Nothing

            Dim cmd As New MySqlCommand(
                "SELECT * FROM tbl_users WHERE email=@e AND password=MD5(@p)", conn)
            cmd.Parameters.AddWithValue("@e", email)
            cmd.Parameters.AddWithValue("@p", pass)

            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            CloseConnection()

            If dt.Rows.Count > 0 Then
                Return dt.Rows(0)
            Else
                Return Nothing
            End If

        Catch ex As Exception
            ShowError("DB Error: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub ShowError(msg As String)
        lblError.Text = msg
        lblError.Visible = True
    End Sub

    Private Sub OnFormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub

End Class