Imports System.IO
Imports System.Diagnostics

Imports iTextSharp.text
Imports iTextSharp.text.pdf

Imports MySql.Data.MySqlClient

Public Class ReceiptGenerator

    Public Shared ReceiptFolder As String =
        Path.Combine(Application.StartupPath, "Receipts")

    Public Shared Sub GenerateReceipt(ByVal SaleID As Integer)

        Try

            'Create Receipts folder if it does not exist
            If Not Directory.Exists(ReceiptFolder) Then
                Directory.CreateDirectory(ReceiptFolder)
            End If

            Dim pdfFile As String =
                Path.Combine(ReceiptFolder,
                             "Receipt_" &
                             SaleID.ToString() &
                             ".pdf")

            Dim doc As New Document(PageSize.A6, 20, 20, 20, 20)

            Dim writer As PdfWriter =
                PdfWriter.GetInstance(doc,
                                      New FileStream(pdfFile,
                                      FileMode.Create))

            doc.Open()

            '==========================
            ' Fonts
            '==========================
            Dim titleFont As Font =
                FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16)

            Dim headerFont As Font =
                FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 11)

            Dim normalFont As Font =
                FontFactory.GetFont(FontFactory.HELVETICA, 10)

            Dim boldFont As Font =
                FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)

            '==========================
            ' Store Name
            '==========================
            Dim p As New Paragraph("GHC PET MART", titleFont)
            p.Alignment = Element.ALIGN_CENTER
            doc.Add(p)

            p = New Paragraph("Pet Supplies and Accessories", normalFont)
            p.Alignment = Element.ALIGN_CENTER
            doc.Add(p)

            doc.Add(New Paragraph(" "))

            '==========================
            ' Receipt Header
            '==========================

            Using con As MySqlConnection = GetConn()

                Dim sql As String =
    "SELECT s.sale_id,
            s.receipt_no,
            s.sale_date,
            s.amount_paid,

            c.first_name,
            c.last_name,

            u.full_name,

            SUM((sd.unit_price * sd.quantity) - sd.discount) AS Total

     FROM tbl_sales s

     LEFT JOIN tbl_customers c
            ON s.customer_id = c.customer_id

     INNER JOIN tbl_users u
            ON s.employee_id = u.user_id

     INNER JOIN tbl_sale_details sd
            ON s.sale_id = sd.sale_id

     WHERE s.sale_id=@id

     GROUP BY s.sale_id"

                Dim cmd As New MySqlCommand(sql, con)
                cmd.Parameters.AddWithValue("@id", SaleID)

                Dim dr As MySqlDataReader = cmd.ExecuteReader()

                If dr.Read() Then

                    Dim total As Decimal = CDec(dr("Total"))
                    Dim paid As Decimal = CDec(dr("amount_paid"))
                    Dim change As Decimal = paid - total

                    doc.Add(New Paragraph("----------------------------------", normalFont))

                    doc.Add(New Paragraph("Receipt No : " &
                              dr("receipt_no").ToString(), normalFont))

                    doc.Add(New Paragraph("Date       : " &
                              CDate(dr("sale_date")).ToString("MMMM dd, yyyy hh:mm tt"),
                              normalFont))

                    doc.Add(New Paragraph("Cashier    : " &
                              dr("full_name").ToString(),
                              normalFont))

                    doc.Add(New Paragraph("Customer   : " &
                              dr("first_name").ToString() &
                              " " &
                              dr("last_name").ToString(),
                              normalFont))

                    doc.Add(New Paragraph("----------------------------------", normalFont))

                End If

                dr.Close()

            End Using

            Using con As MySqlConnection = GetConn()

                Dim sql As String =
    "SELECT
        p.product_name,
        sd.quantity,
        sd.unit_price,
        sd.discount
    FROM tbl_sale_details sd
    INNER JOIN tbl_products p
        ON sd.product_id = p.product_id
    WHERE sd.sale_id=@id"

                Dim cmd As New MySqlCommand(sql, con)
                cmd.Parameters.AddWithValue("@id", SaleID)

                Dim dr As MySqlDataReader = cmd.ExecuteReader()

                Dim table As New PdfPTable(4)

                table.WidthPercentage = 100

                table.SetWidths(New Single() {45, 15, 20, 20})

                table.AddCell(New Phrase("Product", boldFont))
                table.AddCell(New Phrase("Qty", boldFont))
                table.AddCell(New Phrase("Price", boldFont))
                table.AddCell(New Phrase("Amount", boldFont))

                While dr.Read()

                    Dim qty As Integer = CInt(dr("quantity"))
                    Dim price As Decimal = CDec(dr("unit_price"))
                    Dim discount As Decimal = CDec(dr("discount"))

                    Dim amount As Decimal =
            (qty * price) - discount

                    table.AddCell(New Phrase(
            dr("product_name").ToString(),
            normalFont))

                    table.AddCell(New Phrase(
            qty.ToString(),
            normalFont))

                    table.AddCell(New Phrase(
            "₱" & price.ToString("N2"),
            normalFont))

                    table.AddCell(New Phrase(
            "₱" & amount.ToString("N2"),
            normalFont))

                End While

                dr.Close()

                doc.Add(table)

            End Using

            doc.Add(New Paragraph(" "))

            Using con As MySqlConnection = GetConn()

                Dim sql As String =
    "SELECT
        s.amount_paid,
        SUM((sd.unit_price * sd.quantity) - sd.discount) AS Total
     FROM tbl_sales s
     INNER JOIN tbl_sale_details sd
        ON s.sale_id = sd.sale_id
     WHERE s.sale_id=@id
     GROUP BY s.sale_id"

                Dim cmd As New MySqlCommand(sql, con)
                cmd.Parameters.AddWithValue("@id", SaleID)

                Dim dr As MySqlDataReader = cmd.ExecuteReader()

                If dr.Read() Then

                    Dim total As Decimal = CDec(dr("Total"))
                    Dim paid As Decimal = CDec(dr("amount_paid"))
                    Dim change As Decimal = paid - total

                    doc.Add(New Paragraph("------------------------------------------", normalFont))

                    Dim tblTotals As New PdfPTable(2)
                    tblTotals.WidthPercentage = 100
                    tblTotals.SetWidths(New Single() {60, 40})

                    Dim c1 As PdfPCell
                    Dim c2 As PdfPCell

                    c1 = New PdfPCell(New Phrase("TOTAL", boldFont))
                    c1.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c1)

                    c2 = New PdfPCell(New Phrase("₱ " & total.ToString("N2"), boldFont))
                    c2.HorizontalAlignment = Element.ALIGN_RIGHT
                    c2.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c2)

                    c1 = New PdfPCell(New Phrase("PAID", boldFont))
                    c1.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c1)

                    c2 = New PdfPCell(New Phrase("₱ " & paid.ToString("N2"), boldFont))
                    c2.HorizontalAlignment = Element.ALIGN_RIGHT
                    c2.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c2)

                    c1 = New PdfPCell(New Phrase("CHANGE", boldFont))
                    c1.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c1)

                    c2 = New PdfPCell(New Phrase("₱ " & change.ToString("N2"), boldFont))
                    c2.HorizontalAlignment = Element.ALIGN_RIGHT
                    c2.Border = Rectangle.NO_BORDER
                    tblTotals.AddCell(c2)

                    doc.Add(tblTotals)

                    doc.Add(New Paragraph("------------------------------------------", normalFont))
                    doc.Add(New Paragraph(" "))

                    Dim thanks As New Paragraph("THANK YOU FOR SHOPPING!", headerFont)
                    thanks.Alignment = Element.ALIGN_CENTER
                    doc.Add(thanks)

                    Dim visit As New Paragraph("Please come again!", normalFont)
                    visit.Alignment = Element.ALIGN_CENTER
                    doc.Add(visit)

                    Dim store As New Paragraph("GHC PET MART", boldFont)
                    store.Alignment = Element.ALIGN_CENTER
                    doc.Add(store)

                End If

                dr.Close()

            End Using

            doc.Close()

            writer.Close()

            Process.Start(New ProcessStartInfo(pdfFile) With {
                .UseShellExecute = True
            })

        Catch ex As Exception

            MessageBox.Show(
                "Receipt Error:" &
                vbCrLf &
                ex.Message)

        End Try

    End Sub

End Class