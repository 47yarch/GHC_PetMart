Imports MySql.Data.MySqlClient

Public Class frmCategories

    Private dgv As New DataGridView()
    Private txtName As New TextBox()
    Private txtDesc As New TextBox()
    Private isEdit As Boolean = False
    Private editID As Integer = 0

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Categories"
        Me.Size = New Size(750, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Category Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        dgv.Location = New Point(15, 55)
        dgv.Size = New Size(430, 390)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        AddHandler dgv.SelectionChanged, AddressOf OnSelect
        Me.Controls.Add(dgv)

        ' Right panel
        Dim pnl As New Panel() With {
            .Location = New Point(460, 55), .Size = New Size(270, 390),
            .BackColor = Color.White, .BorderStyle = BorderStyle.FixedSingle
        }

        Dim y As Integer = 15
        Dim lh As New Label() With {.Text = "Category Details", .Font = New Font("Arial", 11, FontStyle.Bold), .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(lh) : y += 35

        Dim l1 As New Label() With {.Text = "Name:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l1) : y += 20
        txtName.Size = New Size(240, 28) : txtName.Location = New Point(10, y)
        pnl.Controls.Add(txtName) : y += 40

        Dim l2 As New Label() With {.Text = "Description:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l2) : y += 20
        txtDesc.Size = New Size(240, 28) : txtDesc.Location = New Point(10, y)
        pnl.Controls.Add(txtDesc) : y += 50

        Dim btnSave As New Button() With {
            .Text = "SAVE", .Size = New Size(115, 35), .Location = New Point(10, y),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSave.Click, AddressOf SaveCategory
        pnl.Controls.Add(btnSave)

        Dim btnNew As New Button() With {
            .Text = "NEW", .Size = New Size(115, 35), .Location = New Point(135, y),
            .BackColor = Color.FromArgb(76, 175, 80), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnNew.Click, Sub(s, ev)
                                     isEdit = False : editID = 0
                                     txtName.Clear() : txtDesc.Clear() : txtName.Focus()
                                 End Sub
        pnl.Controls.Add(btnNew) : y += 45

        Dim btnDel As New Button() With {
            .Text = "DELETE SELECTED", .Size = New Size(240, 35), .Location = New Point(10, y),
            .BackColor = Color.FromArgb(244, 67, 54), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDel.Click, AddressOf DeleteCategory
        pnl.Controls.Add(btnDel)

        Me.Controls.Add(pnl)
        LoadCategories()
    End Sub

    Private Sub LoadCategories()
        If Not OpenConnection() Then Return
        Try
            Dim da As New MySqlDataAdapter(
                "SELECT category_id, category_name AS 'Name', description AS 'Description' FROM tbl_categories ORDER BY category_name", conn)
            Dim dt As New DataTable() : da.Fill(dt)
            dgv.DataSource = dt
            dgv.Columns("category_id").Visible = False
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub OnSelect(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then Return
        isEdit = True
        editID = CInt(dgv.SelectedRows(0).Cells("category_id").Value)
        txtName.Text = dgv.SelectedRows(0).Cells("Name").Value.ToString()
        txtDesc.Text = dgv.SelectedRows(0).Cells("Description").Value.ToString()
    End Sub

    Private Sub SaveCategory(sender As Object, e As EventArgs)
        If txtName.Text.Trim() = "" Then MessageBox.Show("Category name is required.") : Return
        If Not OpenConnection() Then Return
        Try
            If isEdit Then
                Dim cmd As New MySqlCommand(
                    "UPDATE tbl_categories SET category_name=@n, description=@d WHERE category_id=@id", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@d", txtDesc.Text.Trim())
                cmd.Parameters.AddWithValue("@id", editID)
                cmd.ExecuteNonQuery()
            Else
                Dim cmd As New MySqlCommand(
                    "INSERT INTO tbl_categories(category_name, description) VALUES(@n,@d)", conn)
                cmd.Parameters.AddWithValue("@n", txtName.Text.Trim())
                cmd.Parameters.AddWithValue("@d", txtDesc.Text.Trim())
                cmd.ExecuteNonQuery()
            End If
            CloseConnection()
            txtName.Clear() : txtDesc.Clear()
            isEdit = False : editID = 0
            LoadCategories()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message) : CloseConnection()
        End Try
    End Sub

    Private Sub DeleteCategory(sender As Object, e As EventArgs)
        If Not isEdit Then MessageBox.Show("Select a category first.") : Return
        If MessageBox.Show("Delete this category?", "Confirm",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
            If Not OpenConnection() Then Return
            Dim cmd As New MySqlCommand("DELETE FROM tbl_categories WHERE category_id=@id", conn)
            cmd.Parameters.AddWithValue("@id", editID)
            cmd.ExecuteNonQuery()
            CloseConnection()
            isEdit = False : editID = 0
            txtName.Clear() : txtDesc.Clear()
            LoadCategories()
        End If
    End Sub

End Class