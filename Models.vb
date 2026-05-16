' ============================================================
'  BANK MANAGEMENT SYSTEM - Models.vb
'  All data record structures used throughout the application
' ============================================================

' ── Customer / Account Record ─────────────────────────────────
Public Class CustomerRecord
    Public Property AccountNumber As String
    Public Property FullName As String
    Public Property PhoneNumber As String
    Public Property Address As String
    Public Property Email As String
    Public Property IDNumber As String
    Public Property Balance As Decimal
    Public Property DateOpened As DateTime
    Public Property Status As String      ' Active / Frozen / Closed
    Public Property AccountType As String ' Savings / Current / Fixed

    Public Overrides Function ToString() As String
        Return $"{AccountNumber} - {FullName} ({AccountType})"
    End Function

    Public Function ToCsvLine() As String
        Return String.Join(",",
            EscapeCsv(AccountNumber),
            EscapeCsv(FullName),
            EscapeCsv(PhoneNumber),
            EscapeCsv(Address),
            EscapeCsv(Email),
            EscapeCsv(IDNumber),
            Balance.ToString("F2"),
            DateOpened.ToString("yyyy-MM-dd HH:mm:ss"),
            EscapeCsv(Status),
            EscapeCsv(AccountType))
    End Function

    Public Shared Function FromCsvLine(line As String) As CustomerRecord
        Dim f = ParseCsvLine(line)
        If f.Count < 10 Then Return Nothing
        Dim rec As New CustomerRecord()
        rec.AccountNumber = f(0)
        rec.FullName = f(1)
        rec.PhoneNumber = f(2)
        rec.Address = f(3)
        rec.Email = f(4)
        rec.IDNumber = f(5)
        Decimal.TryParse(f(6), rec.Balance)
        DateTime.TryParse(f(7), rec.DateOpened)
        rec.Status = f(8)
        rec.AccountType = f(9)
        Return rec
    End Function
End Class

' ── Employee Record ────────────────────────────────────────────
Public Class EmployeeRecord
    Public Property EmployeeID As String
    Public Property FullName As String
    Public Property Role As String        ' Admin / Teller / Manager
    Public Property PinHash As String
    Public Property Email As String
    Public Property Phone As String
    Public Property DateRegistered As DateTime
    Public Property IsActive As Boolean

    Public Overrides Function ToString() As String
        Return $"{EmployeeID} - {FullName} [{Role}]"
    End Function

    Public Function ToCsvLine() As String
        Return String.Join(",",
            EscapeCsv(EmployeeID),
            EscapeCsv(FullName),
            EscapeCsv(Role),
            EscapeCsv(PinHash),
            EscapeCsv(Email),
            EscapeCsv(Phone),
            DateRegistered.ToString("yyyy-MM-dd HH:mm:ss"),
            IsActive.ToString())
    End Function

    Public Shared Function FromCsvLine(line As String) As EmployeeRecord
        Dim f = ParseCsvLine(line)
        If f.Count < 8 Then Return Nothing
        Dim rec As New EmployeeRecord()
        rec.EmployeeID = f(0)
        rec.FullName = f(1)
        rec.Role = f(2)
        rec.PinHash = f(3)
        rec.Email = f(4)
        rec.Phone = f(5)
        DateTime.TryParse(f(6), rec.DateRegistered)
        Boolean.TryParse(f(7), rec.IsActive)
        Return rec
    End Function
End Class

' ── Transaction Record ─────────────────────────────────────────
Public Class TransactionRecord
    Public Property TransactionID As String
    Public Property AccountNumber As String
    Public Property TransactionType As String  ' Deposit / Withdrawal / Transfer
    Public Property Amount As Decimal
    Public Property BalanceBefore As Decimal
    Public Property BalanceAfter As Decimal
    Public Property Timestamp As DateTime
    Public Property PerformedBy As String
    Public Property Notes As String

    Public Function ToCsvLine() As String
        Return String.Join(",",
            EscapeCsv(TransactionID),
            EscapeCsv(AccountNumber),
            EscapeCsv(TransactionType),
            Amount.ToString("F2"),
            BalanceBefore.ToString("F2"),
            BalanceAfter.ToString("F2"),
            Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
            EscapeCsv(PerformedBy),
            EscapeCsv(Notes))
    End Function

    Public Shared Function FromCsvLine(line As String) As TransactionRecord
        Dim f = ParseCsvLine(line)
        If f.Count < 9 Then Return Nothing
        Dim rec As New TransactionRecord()
        rec.TransactionID = f(0)
        rec.AccountNumber = f(1)
        rec.TransactionType = f(2)
        Decimal.TryParse(f(3), rec.Amount)
        Decimal.TryParse(f(4), rec.BalanceBefore)
        Decimal.TryParse(f(5), rec.BalanceAfter)
        DateTime.TryParse(f(6), rec.Timestamp)
        rec.PerformedBy = f(7)
        rec.Notes = f(8)
        Return rec
    End Function
End Class

' ── Report Summary ─────────────────────────────────────────────
Public Class ReportSummary
    Public Property TotalCustomers As Integer
    Public Property ActiveAccounts As Integer
    Public Property FrozenAccounts As Integer
    Public Property ClosedAccounts As Integer
    Public Property TotalDeposits As Decimal
    Public Property TotalWithdrawals As Decimal
    Public Property TotalBalance As Decimal
    Public Property TransactionCount As Integer
    Public Property GeneratedOn As DateTime = DateTime.Now
End Class
