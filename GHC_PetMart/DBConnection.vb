Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Module DBConnection

    Public Const SERVER As String = "localhost"
    Public Const DATABASE As String = "ghc_petmart_3nf"
    Public Const DBUSER As String = "root"
    Public Const DBPASS As String = ""

    Public connStr As String =
        "server=" & SERVER &
        ";database=" & DATABASE &
        ";uid=" & DBUSER &
        ";pwd=" & DBPASS & ";"

    ' Session variables set after login
    Public LoggedUserID As Integer = 0
    Public LoggedUsername As String = ""
    Public LoggedFullName As String = ""
    Public LoggedRole As String = ""

    ' ── Always returns a fresh open connection ──────────────
    ' Use this in ALL forms — wraps in Using block
    Public Function GetConn() As MySqlConnection
        Dim c As New MySqlConnection(connStr)
        c.Open()
        Return c
    End Function

    ' ── Get the Online Store employee ID ───────────────────
    ' Used when customer places order (no employee present)
    Public Function GetOnlineEmployeeID() As Integer
        Try
            Using con As MySqlConnection = GetConn()
                Dim cmd As New MySqlCommand(
                    "SELECT user_id FROM tbl_users WHERE username = 'online' LIMIT 1", con)
                Dim result As Object = cmd.ExecuteScalar()
                Return If(result IsNot Nothing, Convert.ToInt32(result), 1)
            End Using
        Catch
            Return 1
        End Try
    End Function

    ' ── Legacy (kept for backward compatibility) ────────────
    Public conn As New MySqlConnection(connStr)

    Public Function OpenConnection() As Boolean
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
            conn.Open()
            Return True
        Catch ex As Exception
            MessageBox.Show("DB Error: " & ex.Message)
            Return False
        End Try
    End Function

    Public Sub CloseConnection()
        Try
            If conn.State = ConnectionState.Open Then conn.Close()
        Catch
        End Try
    End Sub

End Module