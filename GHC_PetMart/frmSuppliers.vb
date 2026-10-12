Imports MySql.Data.MySqlClient

Public Class frmSuppliers

    Private dgv As New DataGridView()
    Private txtSearch As New TextBox()

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Supplier Management"
        Me.Size = New Size(950, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Supplier Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        Dim lS As New Label() With {.Text = "Search:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lS)
        txtSearch.Size = New Size(250, 28) : txtSearch.Location = New Point(65, 52)
        Me.Controls.Add(txtSearch)

        Dim btnS As New Button() With {
            .Text = "Search", .Size = New Size(80, 28), .Location = New Point(325, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnS.Click, AddressOf LoadSuppliers
        Me.Controls.Add(btnS)

        Dim btnAdd As New Button() With {
            .Text = "Add", .Size = New Size(80, 32), .Location = New Point(15, 92),
            .BackColor = Color.FromArgb(76, 175, 80), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdd.Click, Sub(s, e)
                                     Dim f As New frmSupplierForm() : f.IsEditMode = False : f.ShowDialog() : LoadSuppliers(Nothing, Nothing)
                                 End Sub
        Me.Controls.Add(btnAdd)

        Dim btnEdit As New Button() With {
            .Text = "Edit", .Size = New Size(80, 32), .Location = New Point(105, 92),
            .BackColor = Color.FromArgb(255, 152, 0), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnEdit.Click, AddressOf EditSupplier
        Me.Controls.Add(btnEdit)

        Dim btnDelete As New Button() With {
    .Text = "Delete",
    .Size = New Size(80, 32),
    .Location = New Point(195, 92),
    .BackColor = Color.FromArgb(244, 67, 54),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}
        AddHandler btnDelete.Click, AddressOf DeleteSupplier
        Me.Controls.Add(btnDelete)

        dgv.Location = New Point(15, 135)
        dgv.Size = New Size(910, 400)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)

        LoadSuppliers(Nothing, Nothing)
    End Sub

    Private Sub LoadSuppliers(sender As Object, e As EventArgs)
        If Not OpenConnection() Then Return
        Try
            Dim search As String = "%" & txtSearch.Text.Trim() & "%"
            Dim cmd As New MySqlCommand(
                "SELECT supplier_id, supplier_name AS 'Supplier', contact_person AS 'Contact', " &
                "phone AS 'Phone', email AS 'Email', city AS 'City' " &
                "FROM tbl_suppliers WHERE supplier_name LIKE @s OR phone LIKE @s ORDER BY supplier_name", conn)
            cmd.Parameters.AddWithValue("@s", search)
            Dim dt As New DataTable()
            Dim da As New MySqlDataAdapter(cmd)
            da.Fill(dt)
            dgv.DataSource = dt
            dgv.Columns("supplier_id").Visible = False
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub EditSupplier(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then MessageBox.Show("Select a supplier.") : Return
        Dim f As New frmSupplierForm()
        f.IsEditMode = True
        f.EditSupplierID = CInt(dgv.SelectedRows(0).Cells("supplier_id").Value)
        f.ShowDialog()
        LoadSuppliers(Nothing, Nothing)
    End Sub

    Private Sub DeleteSupplier(sender As Object, e As EventArgs)

        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a supplier.")
            Return
        End If

        Dim supplierID As Integer =
            CInt(dgv.SelectedRows(0).Cells("supplier_id").Value)

        Dim supplierName As String =
            dgv.SelectedRows(0).Cells("Supplier").Value.ToString()

        If MessageBox.Show(
            "Delete supplier '" & supplierName & "' ?" &
            vbCrLf & vbCrLf &
            "This action cannot be undone.",
            "Delete Supplier",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning) = DialogResult.No Then Exit Sub

        If Not OpenConnection() Then Return

        Try

            Dim chk As New MySqlCommand(
                "SELECT COUNT(*) FROM tbl_products WHERE supplier_id=@id", conn)

            chk.Parameters.AddWithValue("@id", supplierID)

            Dim used As Integer = Convert.ToInt32(chk.ExecuteScalar())

            If used > 0 Then

                MessageBox.Show(
                    "This supplier cannot be deleted." &
                    vbCrLf &
                    "There are products assigned to this supplier.",
                    "Delete Blocked",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

                CloseConnection()
                Return

            End If

            Dim cmd As New MySqlCommand(
                "DELETE FROM tbl_suppliers WHERE supplier_id=@id", conn)

            cmd.Parameters.AddWithValue("@id", supplierID)

            cmd.ExecuteNonQuery()

            MessageBox.Show(
                "Supplier deleted successfully.",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        Finally

            CloseConnection()

        End Try

        LoadSuppliers(Nothing, Nothing)

    End Sub
End Class