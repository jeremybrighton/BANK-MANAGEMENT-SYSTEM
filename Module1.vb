' ============================================================
'  BANK MANAGEMENT SYSTEM - Module1.vb
'  Global declarations, constants, shared state, and helpers
' ============================================================
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Module Module1

    ' ── Application Constants ─────────────────────────────────
    Public Const APP_TITLE As String = "NexaBank Management System"
    Public Const DB_FOLDER As String = "BankData"
    Public Const CUSTOMERS_FILE As String = "customers.csv"
    Public Const EMPLOYEES_FILE As String = "employees.csv"
    Public Const TRANSACTIONS_FILE As String = "transactions.csv"
    Public Const AUDIT_FILE As String = "audit_log.txt"
    Public Const BANK_NAME As String = "NexaBank"
    Public Const CURRENCY As String = "KES"

    ' ── Session State ─────────────────────────────────────────
    Public CurrentEmployee As EmployeeRecord = Nothing
    Public IsAdminSession As Boolean = False

    ' ── Data Paths ────────────────────────────────────────────
    Public ReadOnly Property DataFolder As String
        Get
            Dim base As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "NexaBank")
            Return base
        End Get
    End Property

    Public ReadOnly Property CustomersPath As String
        Get
            Return Path.Combine(DataFolder, CUSTOMERS_FILE)
        End Get
    End Property

    Public ReadOnly Property EmployeesPath As String
        Get
            Return Path.Combine(DataFolder, EMPLOYEES_FILE)
        End Get
    End Property

    Public ReadOnly Property TransactionsPath As String
        Get
            Return Path.Combine(DataFolder, TRANSACTIONS_FILE)
        End Get
    End Property

    Public ReadOnly Property AuditPath As String
        Get
            Return Path.Combine(DataFolder, AUDIT_FILE)
        End Get
    End Property

    ' ── Initialise Storage ────────────────────────────────────
    Public Sub InitialiseStorage()
        If Not Directory.Exists(DataFolder) Then
            Directory.CreateDirectory(DataFolder)
        End If

        ' Seed CSV headers if files don't exist
        If Not File.Exists(CustomersPath) Then
            File.WriteAllText(CustomersPath,
                "AccountNumber,FullName,PhoneNumber,Address,Email,IDNumber,Balance,DateOpened,Status,AccountType" & vbCrLf)
        End If

        If Not File.Exists(EmployeesPath) Then
            ' Write header then seed a default admin
            File.WriteAllText(EmployeesPath,
                "EmployeeID,FullName,Role,PinHash,Email,Phone,DateRegistered,IsActive" & vbCrLf)
            ' Create default admin: PIN = 1234
            Dim admin As New EmployeeRecord With {
                .EmployeeID = "EMP-001",
                .FullName = "System Administrator",
                .Role = "Admin",
                .PinHash = HashPin("1234"),
                .Email = "admin@nexabank.com",
                .Phone = "+254700000000",
                .DateRegistered = DateTime.Now,
                .IsActive = True
            }
            DataAccess.SaveEmployee(admin)
        End If

        If Not File.Exists(TransactionsPath) Then
            File.WriteAllText(TransactionsPath,
                "TransactionID,AccountNumber,Type,Amount,BalanceBefore,BalanceAfter,Timestamp,PerformedBy,Notes" & vbCrLf)
        End If
    End Sub

    ' ── Security: SHA-256 PIN Hashing ─────────────────────────
    Public Function HashPin(pin As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes() As Byte = sha.ComputeHash(Encoding.UTF8.GetBytes(pin & "NEXABANK_SALT"))
            Dim sb As New StringBuilder()
            For Each b As Byte In bytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function

    ' ── Audit Logger ──────────────────────────────────────────
    Public Sub WriteAudit(action As String, detail As String)
        Try
            Dim who As String = If(CurrentEmployee IsNot Nothing, CurrentEmployee.EmployeeID, "SYSTEM")
            Dim entry As String = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{who}] {action}: {detail}"
            File.AppendAllText(AuditPath, entry & vbCrLf)
        Catch
            ' Audit should never crash the app
        End Try
    End Sub

    ' ── Generate Unique IDs ───────────────────────────────────
    Public Function GenerateAccountNumber() As String
        Dim rnd As New Random()
        Dim num As String = rnd.Next(10000000, 99999999).ToString()
        ' Ensure uniqueness
        Dim customers = DataAccess.LoadAllCustomers()
        Do While customers.Exists(Function(c) c.AccountNumber = num)
            num = rnd.Next(10000000, 99999999).ToString()
        Loop
        Return num
    End Function

    Public Function GenerateEmployeeID() As String
        Dim employees = DataAccess.LoadAllEmployees()
        Dim nextNum As Integer = employees.Count + 1
        Return $"EMP-{nextNum:000}"
    End Function

    Public Function GenerateTransactionID() As String
        Return $"TXN-{DateTime.Now:yyyyMMddHHmmssfff}"
    End Function

    ' ── Format Currency ───────────────────────────────────────
    Public Function FormatMoney(amount As Decimal) As String
        Return $"{CURRENCY} {amount:N2}"
    End Function

    ' ── CSV Escaping ──────────────────────────────────────────
    Public Function EscapeCsv(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        If value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCrLf) Then
            Return """" & value.Replace("""", """""") & """"
        End If
        Return value
    End Function

    Public Function UnescapeCsv(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        If value.StartsWith("""") AndAlso value.EndsWith("""") Then
            Return value.Substring(1, value.Length - 2).Replace("""""", """")
        End If
        Return value
    End Function

    ' ── Parse CSV Line ────────────────────────────────────────
    Public Function ParseCsvLine(line As String) As List(Of String)
        Dim fields As New List(Of String)()
        Dim sb As New StringBuilder()
        Dim inQuotes As Boolean = False

        For i As Integer = 0 To line.Length - 1
            Dim c As Char = line(i)
            If c = """"c Then
                If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    sb.Append(""""c)
                    i += 1
                Else
                    inQuotes = Not inQuotes
                End If
            ElseIf c = ","c AndAlso Not inQuotes Then
                fields.Add(sb.ToString())
                sb.Clear()
            Else
                sb.Append(c)
            End If
        Next
        fields.Add(sb.ToString())
        Return fields
    End Function

End Module
