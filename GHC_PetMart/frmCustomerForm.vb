Imports MySql.Data.MySqlClient

Public Class frmCustomerForm

    Public IsEditMode As Boolean = False
    Public EditCustomerID As Integer = 0

    Private txtFirst As New TextBox()
    Private txtLast As New TextBox()
    Private txtEmail As New TextBox()
    Private txtPhone As New TextBox()
    Private txtAddress As New TextBox()
    Private txtCity As New TextBox()
    Private txtPassword As New TextBox()
    Private cboPet As New ComboBox()
    Private txtNotes As New TextBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = If(IsEditMode, "Edit Customer", "Add Customer")
        Me.Size = New Size(420, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        Dim y As Integer = 20
        Dim lT As New Label() With {
            .Text = If(IsEditMode, "EDIT CUSTOMER", "ADD CUSTOMER"),
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(100, y)
        }
        Me.Controls.Add(lT) : y += 45

        AddRow("First Name *", txtFirst, y) : y += 55
        AddRow("Last Name *", txtLast, y) : y += 55
        AddRow("Email", txtEmail, y) : y += 55
        AddRow("Phone", txtPhone, y) : y += 55
        AddRow("Address", txtAddress, y) : y += 55
        AddRow("City", txtCity, y) : y += 55
        AddRow("Password", txtPassword, y)
        txtPassword.PasswordChar = "*"c
        y += 55

        Dim lP As New Label() With {.Text = "Pet Type:", .AutoSize = True, .Location = New Point(20, y)}
        Me.Controls.Add(lP) : y += 20
        cboPet.Location = New Point(20, y) : cboPet.Size = New Size(360, 28)
        cboPet.DropDownStyle = ComboBoxStyle.DropDownList
        cboPet.Items.AddRange(New String() {"Dog", "Cat", "Fish", "Bird", "Other"})
        cboPet.SelectedIndex = 0
        Me.Controls.Add(cboPet) : y += 50

        Dim btnSave As New Button() With {
            .Text = "SAVE", .Size = New Size(170, 40), .Location = New Point(20, y),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSave.Click, AddressOf SaveCustomer
        Me.Controls.Add(btnSave)

        Dim btnCancel As New Button() With {
            .Text = "CANCEL", .Size = New Size(170, 40), .Location = New Point(210, y),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnCancel.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnCancel)

        If IsEditMode Then LoadCustomerData()
    End Sub

    Private Sub AddRow(lbl As String, txt As TextBox, y As Integer)
        Dim l As New Label() With {.Text = lbl & ":", .AutoSize = True, .Location = New Point(20, y)}
        Me.Controls.Add(l)
        txt.Size = New Size(360, 28) : txt.Location = New Point(20, y + 20)
        Me.Controls.Add(txt)
    End Sub

    Private Sub LoadCustomerData()
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand("SELECT * FROM tbl_customers WHERE customer_id=@id", conn)
            cmd.Parameters.AddWithValue("@id", EditCustomerID)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            If dt.Rows.Count = 0 Then Return
            Dim r As DataRow = dt.Rows(0)
            txtFirst.Text = r("first_name").ToString()
            txtLast.Text = r("last_name").ToString()
            txtEmail.Text = r("email").ToString()
            txtPhone.Text = r("phone").ToString()
            txtAddress.Text = r("address").ToString()
            txtCity.Text = r("city").ToString()
            cboPet.Text = r("pet_type").ToString()

            Dim cmdPass As New MySqlCommand(
"SELECT password FROM tbl_users
INNER JOIN tbl_customers
ON tbl_users.user_id=tbl_customers.user_id
WHERE customer_id=@id", conn)

            cmdPass.Parameters.AddWithValue("@id", EditCustomerID)

            Dim encrypted As String = Convert.ToString(cmdPass.ExecuteScalar())

            txtPassword.Text = ""
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub SaveCustomer(sender As Object, e As EventArgs)

        If txtFirst.Text.Trim() = "" OrElse txtLast.Text.Trim() = "" Then
            MessageBox.Show("First and Last Name are required.")
            Exit Sub
        End If

        If Not OpenConnection() Then Exit Sub

        Try

            If IsEditMode Then

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
                 WHERE customer_id=@id", conn)

                cmd.Parameters.AddWithValue("@fn", txtFirst.Text.Trim())
                cmd.Parameters.AddWithValue("@ln", txtLast.Text.Trim())
                cmd.Parameters.AddWithValue("@em", txtEmail.Text.Trim())
                cmd.Parameters.AddWithValue("@ph", txtPhone.Text.Trim())
                cmd.Parameters.AddWithValue("@ad", txtAddress.Text.Trim())
                cmd.Parameters.AddWithValue("@cy", txtCity.Text.Trim())
                cmd.Parameters.AddWithValue("@pt", cboPet.Text)
                cmd.Parameters.AddWithValue("@id", EditCustomerID)

                cmd.ExecuteNonQuery()

                If txtPassword.Text.Trim() <> "" Then

                    Dim cmdPass As New MySqlCommand(
                    "UPDATE tbl_users
     INNER JOIN tbl_customers
     ON tbl_users.user_id=tbl_customers.user_id
     SET password=MD5(@pass)
     WHERE customer_id=@id", conn)

                    cmdPass.Parameters.AddWithValue("@pass", txtPassword.Text.Trim())
                    cmdPass.Parameters.AddWithValue("@id", EditCustomerID)

                    cmdPass.ExecuteNonQuery()

                End If

            Else

                Dim username As String

                If txtEmail.Text.Trim() <> "" Then
                    username = txtEmail.Text.Trim()
                Else
                    username = (txtFirst.Text.Trim() & txtLast.Text.Trim()).Replace(" ", "").ToLower()
                End If

                Dim chk As New MySqlCommand(
                    "SELECT COUNT(*) FROM tbl_users
                 WHERE username=@u OR email=@e", conn)

                chk.Parameters.AddWithValue("@u", username)
                chk.Parameters.AddWithValue("@e", txtEmail.Text.Trim())

                If Convert.ToInt32(chk.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Email or Username already exists.")
                    Exit Sub
                End If

                Dim cmdUser As New MySqlCommand(
                    "INSERT INTO tbl_users
                (email,username,password,full_name,role,is_active)
                VALUES
                (@email,@username,MD5(@pass),@fullname,'Customer',1)", conn)

                cmdUser.Parameters.AddWithValue("@email", txtEmail.Text.Trim())
                cmdUser.Parameters.AddWithValue("@username", username)
                cmdUser.Parameters.AddWithValue("@fullname",
                    txtFirst.Text.Trim() & " " & txtLast.Text.Trim())
                If txtPassword.Text.Trim() = "" Then
                    txtPassword.Text = "1234"
                End If

                cmdUser.Parameters.AddWithValue("@pass", txtPassword.Text.Trim())

                cmdUser.ExecuteNonQuery()

                Dim userID As Integer =
                    Convert.ToInt32(New MySqlCommand("SELECT LAST_INSERT_ID()", conn).ExecuteScalar())

                Dim cmdCustomer As New MySqlCommand(
                    "INSERT INTO tbl_customers
                (user_id,first_name,last_name,email,phone,address,city,pet_type)
                VALUES
                (@uid,@fn,@ln,@em,@ph,@ad,@cy,@pt)", conn)

                cmdCustomer.Parameters.AddWithValue("@uid", userID)
                cmdCustomer.Parameters.AddWithValue("@fn", txtFirst.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@ln", txtLast.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@em", txtEmail.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@ph", txtPhone.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@ad", txtAddress.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@cy", txtCity.Text.Trim())
                cmdCustomer.Parameters.AddWithValue("@pt", cboPet.Text)

                cmdCustomer.ExecuteNonQuery()

            End If

            MessageBox.Show("Customer saved successfully!",
                            "Success",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            CloseConnection()
        End Try

    End Sub
End Class