Imports MySql.Data.MySqlClient

Public Class frmExpenses

    Private dgv As New DataGridView()
    Private dtpDate As New DateTimePicker()
    Private cboCategory As New ComboBox()
    Private txtDescription As New TextBox()
    Private txtAmount As New TextBox()
    Private lblRecordedByValue As New Label()

    Private isEdit As Boolean = False
    Private editID As Integer = 0

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Expense Management"
        Me.Size = New Size(950, 580)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
            .Text = "Expense Management",
            .Font = New Font("Arial", 14, FontStyle.Bold),
            .AutoSize = True, .Location = New Point(15, 15)
        }
        Me.Controls.Add(lblTitle)

        ' ── Grid (left) ─────────────────────────────
        dgv.Location = New Point(15, 55)
        dgv.Size = New Size(630, 470)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.MultiSelect = False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        AddHandler dgv.SelectionChanged, AddressOf OnSelect
        Me.Controls.Add(dgv)

        ' ── Right panel (form) ──────────────────────
        Dim pnl As New Panel() With {
            .Location = New Point(660, 55), .Size = New Size(265, 470),
            .BackColor = Color.White, .BorderStyle = BorderStyle.FixedSingle
        }

        Dim y As Integer = 15
        Dim lh As New Label() With {.Text = "Expense Details", .Font = New Font("Arial", 11, FontStyle.Bold), .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(lh) : y += 35

        Dim l1 As New Label() With {.Text = "Date:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l1) : y += 20
        dtpDate.Size = New Size(240, 28) : dtpDate.Location = New Point(10, y)
        dtpDate.Format = DateTimePickerFormat.Short
        pnl.Controls.Add(dtpDate) : y += 38

        Dim l2 As New Label() With {.Text = "Category:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l2) : y += 20
        cboCategory.Size = New Size(240, 28) : cboCategory.Location = New Point(10, y)
        cboCategory.DropDownStyle = ComboBoxStyle.DropDown
        cboCategory.Items.AddRange(New String() {
            "Rent", "Utilities", "Salaries & Wages", "Supplies",
            "Maintenance & Repairs", "Marketing", "Transportation", "Miscellaneous"
        })
        pnl.Controls.Add(cboCategory) : y += 38

        Dim l3 As New Label() With {.Text = "Description:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l3) : y += 20
        txtDescription.Size = New Size(240, 50)
        txtDescription.Multiline = True
        txtDescription.Location = New Point(10, y)
        pnl.Controls.Add(txtDescription) : y += 60

        Dim l4 As New Label() With {.Text = "Amount:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l4) : y += 20
        txtAmount.Size = New Size(240, 28) : txtAmount.Location = New Point(10, y)
        pnl.Controls.Add(txtAmount) : y += 38

        Dim l5 As New Label() With {.Text = "Recorded By:", .AutoSize = True, .Location = New Point(10, y)}
        pnl.Controls.Add(l5) : y += 20
        lblRecordedByValue.AutoSize = True
        lblRecordedByValue.Font = New Font("Arial", 9, FontStyle.Bold)
        lblRecordedByValue.Location = New Point(10, y)
        lblRecordedByValue.Text = LoggedFullName
        pnl.Controls.Add(lblRecordedByValue) : y += 40

        Dim btnSave As New Button() With {
            .Text = "SAVE", .Size = New Size(115, 35), .Location = New Point(10, y),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSave.Click, AddressOf SaveExpense
        pnl.Controls.Add(btnSave)

        Dim btnNew As New Button() With {
            .Text = "NEW", .Size = New Size(115, 35), .Location = New Point(135, y),
            .BackColor = Color.FromArgb(76, 175, 80), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnNew.Click, AddressOf ClearForm
        pnl.Controls.Add(btnNew) : y += 45

        Dim btnDel As New Button() With {
            .Text = "DELETE SELECTED", .Size = New Size(240, 35), .Location = New Point(10, y),
            .BackColor = Color.FromArgb(244, 67, 54), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDel.Click, AddressOf DeleteExpense
        pnl.Controls.Add(btnDel)

        Me.Controls.Add(pnl)

        ClearForm(Nothing, Nothing)
        LoadExpenses()
    End Sub

    ' ── READ ─────────────────────────────────────────
    Private Sub LoadExpenses()
        Try
            Using con As MySqlConnection = GetConn()
                Dim sql As String =
                    "SELECT e.expense_id AS 'ID', e.expense_date AS 'Date', e.category AS 'Category', " &
                    "e.description AS 'Description', e.amount AS 'Amount', u.full_name AS 'Recorded By' " &
                    "FROM tbl_expenses e " &
                    "INNER JOIN tbl_users u ON e.recorded_by = u.user_id " &
                    "ORDER BY e.expense_date DESC, e.expense_id DESC"
                Dim da As New MySqlDataAdapter(sql, con)
                Dim dt As New DataTable()
                da.Fill(dt)
                dgv.DataSource = dt
                dgv.Columns("ID").Visible = False
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading expenses: " & ex.Message)
        End Try
    End Sub

    Private Sub OnSelect(sender As Object, e As EventArgs)
        If dgv.SelectedRows.Count = 0 Then Return
        isEdit = True
        editID = CInt(dgv.SelectedRows(0).Cells("ID").Value)
        dtpDate.Value = Convert.ToDateTime(dgv.SelectedRows(0).Cells("Date").Value)
        cboCategory.Text = dgv.SelectedRows(0).Cells("Category").Value.ToString()
        txtDescription.Text = dgv.SelectedRows(0).Cells("Description").Value.ToString()
        txtAmount.Text = dgv.SelectedRows(0).Cells("Amount").Value.ToString()
        lblRecordedByValue.Text = dgv.SelectedRows(0).Cells("Recorded By").Value.ToString()
    End Sub

    ' ── CREATE / UPDATE ──────────────────────────────
    Private Sub SaveExpense(sender As Object, e As EventArgs)
        If cboCategory.Text.Trim() = "" Then
            MessageBox.Show("Category is required.")
            Return
        End If

        Dim amt As Decimal
        If Not Decimal.TryParse(txtAmount.Text.Trim(), amt) OrElse amt <= 0 Then
            MessageBox.Show("Enter a valid amount greater than 0.")
            Return
        End If

        Try
            Using con As MySqlConnection = GetConn()
                If isEdit Then
                    Dim cmd As New MySqlCommand(
                        "UPDATE tbl_expenses SET expense_date=@d, category=@c, description=@desc, amount=@a " &
                        "WHERE expense_id=@id", con)
                    cmd.Parameters.AddWithValue("@d", dtpDate.Value.Date)
                    cmd.Parameters.AddWithValue("@c", cboCategory.Text.Trim())
                    cmd.Parameters.AddWithValue("@desc", txtDescription.Text.Trim())
                    cmd.Parameters.AddWithValue("@a", amt)
                    cmd.Parameters.AddWithValue("@id", editID)
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Expense updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Dim cmd As New MySqlCommand(
                        "INSERT INTO tbl_expenses(expense_date, category, description, amount, recorded_by) " &
                        "VALUES(@d, @c, @desc, @a, @u)", con)
                    cmd.Parameters.AddWithValue("@d", dtpDate.Value.Date)
                    cmd.Parameters.AddWithValue("@c", cboCategory.Text.Trim())
                    cmd.Parameters.AddWithValue("@desc", txtDescription.Text.Trim())
                    cmd.Parameters.AddWithValue("@a", amt)
                    cmd.Parameters.AddWithValue("@u", LoggedUserID)
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Expense recorded successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

            ClearForm(Nothing, Nothing)
            LoadExpenses()
        Catch ex As Exception
            MessageBox.Show("Error saving expense: " & ex.Message)
        End Try
    End Sub

    ' ── DELETE ───────────────────────────────────────
    Private Sub DeleteExpense(sender As Object, e As EventArgs)
        If Not isEdit Then
            MessageBox.Show("Select an expense first.")
            Return
        End If

        If MessageBox.Show(
            "Delete this expense?" & vbCrLf & vbCrLf & "This action cannot be undone.",
            "Delete Expense",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning) = DialogResult.No Then Exit Sub

        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand("DELETE FROM tbl_expenses WHERE expense_id=@id", con)
                cmd.Parameters.AddWithValue("@id", editID)
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Expense deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ClearForm(Nothing, Nothing)
            LoadExpenses()
        Catch ex As Exception
            MessageBox.Show("Error deleting expense: " & ex.Message)
        End Try
    End Sub

    ' ── RESET FORM ───────────────────────────────────
    Private Sub ClearForm(sender As Object, e As EventArgs)
        isEdit = False
        editID = 0
        dtpDate.Value = DateTime.Now
        cboCategory.Text = ""
        txtDescription.Clear()
        txtAmount.Clear()
        lblRecordedByValue.Text = LoggedFullName
        cboCategory.Focus()
    End Sub

End Class