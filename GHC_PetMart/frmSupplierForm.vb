Imports MySql.Data.MySqlClient

Public Class frmSupplierForm

    Public IsEditMode As Boolean = False
    Public EditSupplierID As Integer = 0

    Private txtName As New TextBox()
    Private txtContact As New TextBox()
    Private txtEmail As New TextBox()
    Private txtPhone As New TextBox()
    Private txtAddress As New TextBox()
    Private txtCity As New TextBox()
    Private txtNotes As New TextBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = If(IsEditMode, "Edit Supplier", "Add Supplier")
        Me.Size = New Size(420, 540)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.BackColor = Color.White

        Dim y As Integer = 20
        Dim lT As New Label() With {
            .Text = If(IsEditMode, "EDIT SUPPLIER", "ADD SUPPLIER"),
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(100, y)
        }
        Me.Controls.Add(lT) : y += 45

        AddRow("Supplier Name *", txtName, y) : y += 55
        AddRow("Contact Person", txtContact, y) : y += 55
        AddRow("Email", txtEmail, y) : y += 55
        AddRow("Phone", txtPhone, y) : y += 55
        AddRow("Address", txtAddress, y) : y += 55
        AddRow("City", txtCity, y) : y += 55
        AddRow("Notes", txtNotes, y) : y += 55

        Dim btnSave As New Button() With {
            .Text = "SAVE", .Size = New Size(170, 40), .Location = New Point(20, y),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSave.Click, AddressOf SaveSupplier
        Me.Controls.Add(btnSave)

        Dim btnCancel As New Button() With {
            .Text = "CANCEL", .Size = New Size(170, 40), .Location = New Point(210, y),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnCancel.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnCancel)

        If IsEditMode Then LoadSupplierData()
    End Sub

    Private Sub AddRow(lbl As String, txt As TextBox, y As Integer)
        Dim l As New Label() With {.Text = lbl & ":", .AutoSize = True, .Location = New Point(20, y)}
        Me.Controls.Add(l)
        txt.Size = New Size(360, 28) : txt.Location = New Point(20, y + 20)
        Me.Controls.Add(txt)
    End Sub

    Private Sub LoadSupplierData()
        If Not OpenConnection() Then Return
        Try
            Dim cmd As New MySqlCommand("SELECT * FROM tbl_suppliers WHERE supplier_id=@id", conn)
            cmd.Parameters.AddWithValue("@id", EditSupplierID)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            If dt.Rows.Count = 0 Then Return
            Dim r As DataRow = dt.Rows(0)
            txtName.Text = r("supplier_name").ToString()
            txtContact.Text = r("contact_person").ToString()
            txtEmail.Text = r("email").ToString()
            txtPhone.Text = r("phone").ToString()
            txtAddress.Text = r("address").ToString()
            txtCity.Text = r("city").ToString()
            txtNotes.Text = r("notes").ToString()
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub SaveSupplier(sender As Object, e As EventArgs)
        If txtName.Text.Trim() = "" Then MessageBox.Show("Supplier name is required.") : Return
        If Not OpenConnection() Then Return
        Try
            If IsEditMode Then
                Dim cmd As New MySqlCommand(
                    "UPDATE tbl_suppliers SET supplier_name=@n,contact_person=@c,email=@e," &
                    "phone=@p,address=@a,city=@ci,notes=@no,updated_at=NOW() WHERE supplier_id=@id", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@c", txtContact.Text.Trim())
                cmd.Parameters.AddWithValue("@e", txtEmail.Text.Trim())
                cmd.Parameters.AddWithValue("@p", txtPhone.Text.Trim())
                cmd.Parameters.AddWithValue("@a", txtAddress.Text.Trim())
                cmd.Parameters.AddWithValue("@ci", txtCity.Text.Trim())
                cmd.Parameters.AddWithValue("@no", txtNotes.Text.Trim())
                cmd.Parameters.AddWithValue("@id", EditSupplierID)
                cmd.ExecuteNonQuery()
            Else
                Dim cmd As New MySqlCommand(
                    "INSERT INTO tbl_suppliers(supplier_name,contact_person,email,phone,address,city,notes) " &
                    "VALUES(@n,@c,@e,@p,@a,@ci,@no)", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@c", txtContact.Text.Trim())
                cmd.Parameters.AddWithValue("@e", txtEmail.Text.Trim())
                cmd.Parameters.AddWithValue("@p", txtPhone.Text.Trim())
                cmd.Parameters.AddWithValue("@a", txtAddress.Text.Trim())
                cmd.Parameters.AddWithValue("@ci", txtCity.Text.Trim())
                cmd.Parameters.AddWithValue("@no", txtNotes.Text.Trim())
                cmd.ExecuteNonQuery()
            End If
            CloseConnection()
            MessageBox.Show("Supplier saved!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message) : CloseConnection()
        End Try
    End Sub

End Class