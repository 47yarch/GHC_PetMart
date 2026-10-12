Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports MySql.Data.MySqlClient
Imports Document = iTextSharp.text.Document
Imports FontFactory = iTextSharp.text.FontFactory
Imports PageSize = iTextSharp.text.PageSize
Imports Paragraph = iTextSharp.text.Paragraph
Imports PdfPTable = iTextSharp.text.pdf.PdfPTable
Imports PdfWriter = iTextSharp.text.pdf.PdfWriter
Imports Phrase = iTextSharp.text.Phrase

Public Class frmReports

    Private cboType As New ComboBox()
    Private cboPeriod As New ComboBox()
    Private dtpFrom As New DateTimePicker()
    Private dtpTo As New DateTimePicker()
    Private dgv As New DataGridView()
    Private lblSummary As New Label()
    Private gross As Decimal = 0D
    Private expenses As Decimal = 0D
    Private income As Decimal = 0D

    Private Sub InitForm() Handles MyBase.Load
        Me.Text = "Reports"
        Me.Size = New Size(1100, 660)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke

        Dim lblTitle As New Label() With {
    .Text = "Report Generation",
    .Font = New System.Drawing.Font("Arial", 14, System.Drawing.FontStyle.Bold),
    .AutoSize = True,
    .Location = New Point(15, 15)
}
        Me.Controls.Add(lblTitle)

        Dim lT As New Label() With {.Text = "Report Type:", .AutoSize = True, .Location = New Point(15, 55)}
        Me.Controls.Add(lT)
        cboType.Location = New Point(100, 52) : cboType.Size = New Size(160, 28)
        cboType.DropDownStyle = ComboBoxStyle.DropDownList
        cboType.Items.AddRange(New String() {"Sales", "Inventory", "Customers", "Suppliers", "Total Income"})
        cboType.SelectedIndex = 0
        Me.Controls.Add(cboType)

        Dim lP As New Label() With {.Text = "Period:", .AutoSize = True, .Location = New Point(275, 55)}
        Me.Controls.Add(lP)
        cboPeriod.Location = New Point(320, 52) : cboPeriod.Size = New Size(130, 28)
        cboPeriod.DropDownStyle = ComboBoxStyle.DropDownList
        cboPeriod.Items.AddRange(New String() {"Daily", "Weekly", "Monthly", "Annual", "Custom"})
        cboPeriod.SelectedIndex = 0
        AddHandler cboPeriod.SelectedIndexChanged, AddressOf OnPeriodChange
        Me.Controls.Add(cboPeriod)

        Dim lF As New Label() With {.Text = "From:", .AutoSize = True, .Location = New Point(465, 55)}
        Me.Controls.Add(lF)
        dtpFrom.Location = New Point(500, 52) : dtpFrom.Size = New Size(140, 28)
        dtpFrom.Enabled = False
        Me.Controls.Add(dtpFrom)

        Dim lTo As New Label() With {.Text = "To:", .AutoSize = True, .Location = New Point(655, 55)}
        Me.Controls.Add(lTo)
        dtpTo.Location = New Point(675, 52) : dtpTo.Size = New Size(140, 28)
        dtpTo.Enabled = False
        Me.Controls.Add(dtpTo)

        Dim btnGen As New Button() With {
            .Text = "GENERATE", .Size = New Size(110, 28), .Location = New Point(830, 52),
            .BackColor = Color.FromArgb(63, 81, 181), .ForeColor = Color.White, .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnGen.Click, AddressOf GenerateReport
        Me.Controls.Add(btnGen)

        Dim btnPDF As New Button() With {
    .Text = "EXPORT PDF",
    .Size = New Size(120, 28),
    .Location = New Point(950, 52),
    .BackColor = Color.FromArgb(76, 175, 80),
    .ForeColor = Color.White,
    .FlatStyle = FlatStyle.Flat
}

        AddHandler btnPDF.Click, AddressOf ExportPDF

        Me.Controls.Add(btnPDF)

        lblSummary.AutoSize = True
        lblSummary.Font = New System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
        lblSummary.ForeColor = Color.FromArgb(63, 81, 181)
        lblSummary.Location = New Point(15, 92)
        Me.Controls.Add(lblSummary)

        dgv.Location = New Point(15, 120)
        dgv.Size = New Size(1060, 480)
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.BackgroundColor = Color.White
        Me.Controls.Add(dgv)
    End Sub

    Private Sub OnPeriodChange(sender As Object, e As EventArgs)
        Dim isCustom As Boolean = cboPeriod.Text = "Custom"
        dtpFrom.Enabled = isCustom
        dtpTo.Enabled = isCustom
    End Sub

    Private Sub GenerateReport(sender As Object, e As EventArgs)
        Dim dateFrom As String = ""
        Dim dateTo As String = ""

        Select Case cboPeriod.Text
            Case "Daily"
                dateFrom = DateTime.Now.ToString("yyyy-MM-dd")
                dateTo = dateFrom
            Case "Weekly"
                dateFrom = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd")
                dateTo = DateTime.Now.ToString("yyyy-MM-dd")
            Case "Monthly"
                dateFrom = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).ToString("yyyy-MM-dd")
                dateTo = DateTime.Now.ToString("yyyy-MM-dd")
            Case "Annual"
                dateFrom = New DateTime(DateTime.Now.Year, 1, 1).ToString("yyyy-MM-dd")
                dateTo = DateTime.Now.ToString("yyyy-MM-dd")
            Case "Custom"
                dateFrom = dtpFrom.Value.ToString("yyyy-MM-dd")
                dateTo = dtpTo.Value.ToString("yyyy-MM-dd")
        End Select

        If Not OpenConnection() Then Return
        Try
            Dim dt As New DataTable()
            Dim da As MySqlDataAdapter

            Select Case cboType.Text
                Case "Sales"
                    da = New MySqlDataAdapter(
"SELECT
s.receipt_no AS 'Receipt',

DATE(s.sale_date) AS 'Date',

COALESCE(CONCAT(c.first_name,' ',c.last_name),'Walk-in') AS 'Customer',

u.full_name AS 'Employee',

SUM(sd.quantity) AS 'Items',

SUM(sd.quantity * sd.unit_price) AS 'Gross',

SUM(sd.quantity * sd.unit_price - sd.discount) AS 'Net Total',

s.amount_paid AS 'Paid',

s.amount_paid -
SUM(sd.quantity * sd.unit_price - sd.discount)
AS 'Change'

FROM tbl_sales s

INNER JOIN tbl_sale_details sd
ON s.sale_id = sd.sale_id

LEFT JOIN tbl_customers c
ON s.customer_id=c.customer_id

INNER JOIN tbl_users u
ON s.employee_id=u.user_id

WHERE DATE(s.sale_date)
BETWEEN '" & dateFrom & "' AND '" & dateTo & "'

GROUP BY
s.sale_id,
s.receipt_no,
s.sale_date,
c.first_name,
c.last_name,
u.full_name,
s.amount_paid

ORDER BY s.sale_date DESC", conn)

                Case "Inventory"
                    da = New MySqlDataAdapter(
                        "SELECT

p.product_name AS 'Product',

c.category_name AS 'Category',

sup.supplier_name AS 'Supplier',

p.stock_quantity AS 'Stock',

p.regular_price AS 'Price',

(p.stock_quantity * p.regular_price)
AS 'Inventory Value'

FROM tbl_products p

INNER JOIN tbl_categories c
ON p.category_id=c.category_id

INNER JOIN tbl_suppliers sup
ON p.supplier_id=sup.supplier_id

ORDER BY c.category_name,p.product_name", conn)

                Case "Customers"
                    da = New MySqlDataAdapter(
                        "SELECT

CONCAT(c.first_name,' ',c.last_name)
AS 'Customer',

COUNT(DISTINCT s.sale_id)
AS 'Orders',

SUM(sd.quantity)
AS 'Items Purchased',

SUM(sd.quantity*sd.unit_price-sd.discount)
AS 'Total Spent'

FROM tbl_customers c

INNER JOIN tbl_sales s
ON c.customer_id=s.customer_id

INNER JOIN tbl_sale_details sd
ON s.sale_id=sd.sale_id

GROUP BY
c.customer_id

ORDER BY
SUM(sd.quantity*sd.unit_price-sd.discount) DESC", conn)

                Case "Total Income"

                    Dim sql As String =
                    "SELECT
    DATE(s.sale_date) AS 'Date',

    IFNULL(SUM(sd.quantity * sd.unit_price),0)
    AS 'Gross Sales',

    (
        IFNULL(SUM(sd.quantity * sd.unit_price),0)
    )
    AS 'Net Sales',

    (
        SELECT IFNULL(SUM(amount),0)
        FROM tbl_expenses e
        WHERE DATE(e.expense_date)=DATE(s.sale_date)
    )
    AS 'Expenses',

    (
        IFNULL(SUM(sd.quantity * sd.unit_price),0)
        -
        (
            SELECT IFNULL(SUM(amount),0)
            FROM tbl_expenses e
            WHERE DATE(e.expense_date)=DATE(s.sale_date)
        )
    )
    AS 'Income'

FROM tbl_sales s

INNER JOIN tbl_sale_details sd
ON s.sale_id = sd.sale_id

WHERE DATE(s.sale_date)
BETWEEN '" & dateFrom & "'
AND '" & dateTo & "'

GROUP BY DATE(s.sale_date)

ORDER BY DATE(s.sale_date)"

                    da = New MySqlDataAdapter(sql, conn)

                Case Else
                    da = New MySqlDataAdapter(
                        "SELECT

sup.supplier_name
AS 'Supplier',

COUNT(p.product_id)
AS 'Products Supplied',

SUM(p.stock_quantity)
AS 'Total Stocks',

SUM(p.stock_quantity*p.regular_price)
AS 'Inventory Value'

FROM tbl_suppliers sup

INNER JOIN tbl_products p
ON sup.supplier_id=p.supplier_id

GROUP BY
sup.supplier_id

ORDER BY
sup.supplier_name", conn)
            End Select

            da.Fill(dt)
            dgv.DataSource = dt
            If cboType.Text = "Total Income" Then

                gross = 0D
                expenses = 0D
                income = 0D

                For Each r As DataRow In dt.Rows

                    gross += Convert.ToDecimal(r("Net Sales"))
                    expenses += Convert.ToDecimal(r("Expenses"))
                    income += Convert.ToDecimal(r("Income"))

                Next

                'Add TOTAL row
                Dim totalTable As New DataTable()

                'Create columns manually so Date is TEXT
                totalTable.Columns.Add("Date", GetType(String))
                totalTable.Columns.Add("Gross Sales", GetType(Decimal))
                totalTable.Columns.Add("Net Sales", GetType(Decimal))
                totalTable.Columns.Add("Expenses", GetType(Decimal))
                totalTable.Columns.Add("Income", GetType(Decimal))

                'Copy existing rows
                For Each r As DataRow In dt.Rows
                    totalTable.Rows.Add(
        Convert.ToDateTime(r("Date")).ToString("yyyy-MM-dd"),
        r("Gross Sales"),
        r("Net Sales"),
        r("Expenses"),
        r("Income")
    )
                Next

                'Add TOTAL row
                totalTable.Rows.Add(
    "TOTAL",
    DBNull.Value,
    gross,
    expenses,
    income
)

                dgv.DataSource = totalTable

                Dim lastRow As Integer = dgv.Rows.Count - 1

                dgv.Rows(lastRow).DefaultCellStyle.Font =
    New System.Drawing.Font(dgv.Font, FontStyle.Bold)

                dgv.Rows(lastRow).DefaultCellStyle.BackColor =
    Color.LightYellow

                lblSummary.Text =
        "Gross Sales : ₱" &
        gross.ToString("N2") &
        "     Expenses : ₱" &
        expenses.ToString("N2") &
        "     TOTAL INCOME : ₱" &
        income.ToString("N2")

            Else

                lblSummary.Text =
        cboType.Text &
        " Report | Period : " &
        cboPeriod.Text &
        " | Records : " &
        dt.Rows.Count

            End If
        Finally
            CloseConnection()
        End Try
    End Sub

    Private Sub ExportPDF(sender As Object, e As EventArgs)

        If dgv.Rows.Count = 0 Then
            MessageBox.Show("Generate a report first.")
            Exit Sub
        End If

        Dim sfd As New SaveFileDialog()

        sfd.Filter = "PDF Files|*.pdf"

        sfd.FileName =
            cboType.Text &
            "_" &
            Date.Now.ToString("yyyyMMdd_HHmmss") &
            ".pdf"

        If sfd.ShowDialog <> DialogResult.OK Then Exit Sub

        Try

            Dim doc As New Document(PageSize.A4.Rotate(), 20, 20, 20, 20)

            PdfWriter.GetInstance(doc,
                                  New FileStream(sfd.FileName,
                                                 FileMode.Create))

            doc.Open()

            Try

                Dim logoPath As String =
    System.IO.Path.Combine(
        Application.StartupPath,
        "Images",
        "logo.png")

                If System.IO.File.Exists(logoPath) Then

                    Dim logo As iTextSharp.text.Image =
            iTextSharp.text.Image.GetInstance(logoPath)

                    logo.ScaleToFit(80, 80)

                    logo.Alignment = Element.ALIGN_CENTER

                    doc.Add(logo)

                End If

            Catch ex As Exception

            End Try

            Dim titleFont As iTextSharp.text.Font =
    FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18)

            Dim normalFont As iTextSharp.text.Font =
    FontFactory.GetFont(FontFactory.HELVETICA, 10)

            Dim title As New Paragraph(
"GHC PET MART",
titleFont)

            title.Alignment = Element.ALIGN_CENTER

            doc.Add(title)

            Dim subtitle As New Paragraph(
"Inventory Management and Point of Sale System",
normalFont)

            subtitle.Alignment = Element.ALIGN_CENTER

            doc.Add(subtitle)

            Dim report As New Paragraph(
cboType.Text.ToUpper() & " REPORT",
titleFont)

            report.Alignment = Element.ALIGN_CENTER

            doc.Add(report)

            Dim period As New Paragraph(
"Period : " & cboPeriod.Text,
normalFont)

            period.Alignment = Element.ALIGN_CENTER

            doc.Add(period)

            Dim generated As New Paragraph(
"Generated on : " &
Date.Now.ToString("MMMM dd, yyyy hh:mm tt"),
normalFont)

            generated.Alignment = Element.ALIGN_CENTER

            doc.Add(generated)

            doc.Add(New Paragraph(" "))

            Dim table As New PdfPTable(dgv.Columns.Count)

            table.WidthPercentage = 100

            For Each col As DataGridViewColumn In dgv.Columns

                table.AddCell(New Phrase(col.HeaderText))

            Next

            For Each row As DataGridViewRow In dgv.Rows

                If row.IsNewRow Then Continue For

                For Each cell As DataGridViewCell In row.Cells

                    table.AddCell(cell.Value.ToString())

                Next

            Next

            doc.Add(table)

            If cboType.Text = "Total Income" Then

                doc.Add(New Paragraph(" "))

                Dim totalFont As iTextSharp.text.Font =
        FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12)

                doc.Add(New Paragraph("TOTAL INCOME SUMMARY", totalFont))
                doc.Add(New Paragraph(""))

                doc.Add(New Paragraph("Gross Sales : ₱" & gross.ToString("N2"), totalFont))
                doc.Add(New Paragraph("Expenses : ₱" & expenses.ToString("N2"), totalFont))
                doc.Add(New Paragraph("Total Income : ₱" & income.ToString("N2"), totalFont))

            End If

            doc.Add(New Paragraph(" "))

            Dim totalRows As Integer = dgv.Rows.Count

            doc.Add(New Paragraph(
"Total Records : " & totalRows,
normalFont))

            doc.Add(New Paragraph(
"Generated By : " & LoggedFullName,
normalFont))

            doc.Add(New Paragraph(
"Generated On : " &
Date.Now.ToString("MMMM dd, yyyy hh:mm tt"),
normalFont))

            doc.Add(New Paragraph(
"System : GHC Pet Mart Inventory & POS",
normalFont))

            doc.Close()

            MessageBox.Show("PDF exported successfully!")

        Catch ex As Exception

            MessageBox.Show(ex.Message)

        End Try

    End Sub

End Class