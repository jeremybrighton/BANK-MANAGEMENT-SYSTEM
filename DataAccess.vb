' ============================================================
'  BANK MANAGEMENT SYSTEM - DataAccess.vb
'  All CSV file read/write operations centralised here
' ============================================================
Imports System.IO

Public Module DataAccess

    ' ════════════════════════════════════════════════════════
    '  CUSTOMER OPERATIONS
    ' ════════════════════════════════════════════════════════

    Public Function LoadAllCustomers() As List(Of CustomerRecord)
        Dim list As New List(Of CustomerRecord)()
        If Not File.Exists(CustomersPath) Then Return list
        Try
            Dim lines() As String = File.ReadAllLines(CustomersPath)
            For i As Integer = 1 To lines.Length - 1  ' skip header
                Dim line As String = lines(i).Trim()
                If String.IsNullOrEmpty(line) Then Continue For
                Dim rec = CustomerRecord.FromCsvLine(line)
                If rec IsNot Nothing Then list.Add(rec)
            Next
        Catch ex As Exception
            WriteAudit("ERROR", $"LoadAllCustomers: {ex.Message}")
        End Try
        Return list
    End Function

    Public Function FindCustomerByAccount(accountNumber As String) As CustomerRecord
        Dim all = LoadAllCustomers()
        Return all.Find(Function(c) c.AccountNumber = accountNumber)
    End Function

    Public Function FindCustomersByName(name As String) As List(Of CustomerRecord)
        Dim all = LoadAllCustomers()
        Return all.FindAll(Function(c) c.FullName.ToLower().Contains(name.ToLower()))
    End Function

    Public Function SaveCustomer(customer As CustomerRecord) As Boolean
        Try
            Dim all = LoadAllCustomers()
            Dim idx As Integer = all.FindIndex(Function(c) c.AccountNumber = customer.AccountNumber)
            If idx >= 0 Then
                all(idx) = customer
            Else
                all.Add(customer)
            End If
            WriteAllCustomers(all)
            Return True
        Catch ex As Exception
            WriteAudit("ERROR", $"SaveCustomer: {ex.Message}")
            Return False
        End Try
    End Function

    Public Function DeleteCustomer(accountNumber As String) As Boolean
        Try
            Dim all = LoadAllCustomers()
            Dim removed = all.RemoveAll(Function(c) c.AccountNumber = accountNumber)
            If removed > 0 Then
                WriteAllCustomers(all)
                WriteAudit("DELETE_CUSTOMER", $"AccountNumber={accountNumber}")
                Return True
            End If
            Return False
        Catch ex As Exception
            WriteAudit("ERROR", $"DeleteCustomer: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub WriteAllCustomers(list As List(Of CustomerRecord))
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("AccountNumber,FullName,PhoneNumber,Address,Email,IDNumber,Balance,DateOpened,Status,AccountType")
        For Each c In list
            sb.AppendLine(c.ToCsvLine())
        Next
        File.WriteAllText(CustomersPath, sb.ToString())
    End Sub

    ' ════════════════════════════════════════════════════════
    '  EMPLOYEE OPERATIONS
    ' ════════════════════════════════════════════════════════

    Public Function LoadAllEmployees() As List(Of EmployeeRecord)
        Dim list As New List(Of EmployeeRecord)()
        If Not File.Exists(EmployeesPath) Then Return list
        Try
            Dim lines() As String = File.ReadAllLines(EmployeesPath)
            For i As Integer = 1 To lines.Length - 1
                Dim line As String = lines(i).Trim()
                If String.IsNullOrEmpty(line) Then Continue For
                Dim rec = EmployeeRecord.FromCsvLine(line)
                If rec IsNot Nothing Then list.Add(rec)
            Next
        Catch ex As Exception
            WriteAudit("ERROR", $"LoadAllEmployees: {ex.Message}")
        End Try
        Return list
    End Function

    Public Function FindEmployeeByID(empID As String) As EmployeeRecord
        Dim all = LoadAllEmployees()
        Return all.Find(Function(e) e.EmployeeID = empID)
    End Function

    Public Function AuthenticateEmployee(empID As String, pin As String) As EmployeeRecord
        Dim emp = FindEmployeeByID(empID)
        If emp Is Nothing OrElse Not emp.IsActive Then Return Nothing
        If emp.PinHash = HashPin(pin) Then Return emp
        Return Nothing
    End Function

    Public Function SaveEmployee(emp As EmployeeRecord) As Boolean
        Try
            Dim all = LoadAllEmployees()
            Dim idx As Integer = all.FindIndex(Function(e) e.EmployeeID = emp.EmployeeID)
            If idx >= 0 Then
                all(idx) = emp
            Else
                all.Add(emp)
            End If
            WriteAllEmployees(all)
            Return True
        Catch ex As Exception
            WriteAudit("ERROR", $"SaveEmployee: {ex.Message}")
            Return False
        End Try
    End Function

    Public Function DeleteEmployee(empID As String) As Boolean
        Try
            Dim all = LoadAllEmployees()
            Dim removed = all.RemoveAll(Function(e) e.EmployeeID = empID)
            If removed > 0 Then
                WriteAllEmployees(all)
                WriteAudit("DELETE_EMPLOYEE", $"EmployeeID={empID}")
                Return True
            End If
            Return False
        Catch ex As Exception
            WriteAudit("ERROR", $"DeleteEmployee: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub WriteAllEmployees(list As List(Of EmployeeRecord))
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("EmployeeID,FullName,Role,PinHash,Email,Phone,DateRegistered,IsActive")
        For Each e In list
            sb.AppendLine(e.ToCsvLine())
        Next
        File.WriteAllText(EmployeesPath, sb.ToString())
    End Sub

    ' ════════════════════════════════════════════════════════
    '  TRANSACTION OPERATIONS
    ' ════════════════════════════════════════════════════════

    Public Function LoadAllTransactions() As List(Of TransactionRecord)
        Dim list As New List(Of TransactionRecord)()
        If Not File.Exists(TransactionsPath) Then Return list
        Try
            Dim lines() As String = File.ReadAllLines(TransactionsPath)
            For i As Integer = 1 To lines.Length - 1
                Dim line As String = lines(i).Trim()
                If String.IsNullOrEmpty(line) Then Continue For
                Dim rec = TransactionRecord.FromCsvLine(line)
                If rec IsNot Nothing Then list.Add(rec)
            Next
        Catch ex As Exception
            WriteAudit("ERROR", $"LoadAllTransactions: {ex.Message}")
        End Try
        Return list
    End Function

    Public Function LoadTransactionsByAccount(accountNumber As String) As List(Of TransactionRecord)
        Dim all = LoadAllTransactions()
        Return all.FindAll(Function(t) t.AccountNumber = accountNumber)
    End Function

    Public Function SaveTransaction(txn As TransactionRecord) As Boolean
        Try
            File.AppendAllText(TransactionsPath, txn.ToCsvLine() & vbCrLf)
            Return True
        Catch ex As Exception
            WriteAudit("ERROR", $"SaveTransaction: {ex.Message}")
            Return False
        End Try
    End Function

    ' ── Perform Deposit (atomic) ───────────────────────────────
    Public Function Deposit(accountNumber As String, amount As Decimal, notes As String) As (Success As Boolean, Message As String)
        If amount <= 0 Then Return (False, "Amount must be greater than zero.")
        Dim customer = FindCustomerByAccount(accountNumber)
        If customer Is Nothing Then Return (False, "Account not found.")
        If customer.Status <> "Active" Then Return (False, $"Account is {customer.Status}. Transactions not allowed.")

        Dim before = customer.Balance
        customer.Balance += amount

        Dim txn As New TransactionRecord With {
            .TransactionID = GenerateTransactionID(),
            .AccountNumber = accountNumber,
            .TransactionType = "Deposit",
            .Amount = amount,
            .BalanceBefore = before,
            .BalanceAfter = customer.Balance,
            .Timestamp = DateTime.Now,
            .PerformedBy = If(CurrentEmployee IsNot Nothing, CurrentEmployee.EmployeeID, "SYSTEM"),
            .Notes = notes
        }

        SaveCustomer(customer)
        SaveTransaction(txn)
        WriteAudit("DEPOSIT", $"Acc={accountNumber}, Amount={FormatMoney(amount)}, NewBal={FormatMoney(customer.Balance)}")
        Return (True, $"Deposit successful. New balance: {FormatMoney(customer.Balance)}")
    End Function

    ' ── Perform Withdrawal (atomic) ───────────────────────────
    Public Function Withdraw(accountNumber As String, amount As Decimal, notes As String) As (Success As Boolean, Message As String)
        If amount <= 0 Then Return (False, "Amount must be greater than zero.")
        Dim customer = FindCustomerByAccount(accountNumber)
        If customer Is Nothing Then Return (False, "Account not found.")
        If customer.Status <> "Active" Then Return (False, $"Account is {customer.Status}. Transactions not allowed.")
        If customer.Balance < amount Then Return (False, $"Insufficient funds. Available: {FormatMoney(customer.Balance)}")

        Dim before = customer.Balance
        customer.Balance -= amount

        Dim txn As New TransactionRecord With {
            .TransactionID = GenerateTransactionID(),
            .AccountNumber = accountNumber,
            .TransactionType = "Withdrawal",
            .Amount = amount,
            .BalanceBefore = before,
            .BalanceAfter = customer.Balance,
            .Timestamp = DateTime.Now,
            .PerformedBy = If(CurrentEmployee IsNot Nothing, CurrentEmployee.EmployeeID, "SYSTEM"),
            .Notes = notes
        }

        SaveCustomer(customer)
        SaveTransaction(txn)
        WriteAudit("WITHDRAWAL", $"Acc={accountNumber}, Amount={FormatMoney(amount)}, NewBal={FormatMoney(customer.Balance)}")
        Return (True, $"Withdrawal successful. New balance: {FormatMoney(customer.Balance)}")
    End Function

    ' ── Perform Transfer (atomic) ──────────────────────────────
    Public Function Transfer(fromAccount As String, toAccount As String, amount As Decimal, notes As String) As (Success As Boolean, Message As String)
        If fromAccount = toAccount Then Return (False, "Cannot transfer to the same account.")
        If amount <= 0 Then Return (False, "Amount must be greater than zero.")

        Dim sender = FindCustomerByAccount(fromAccount)
        Dim receiver = FindCustomerByAccount(toAccount)

        If sender Is Nothing Then Return (False, "Source account not found.")
        If receiver Is Nothing Then Return (False, "Destination account not found.")
        If sender.Status <> "Active" Then Return (False, "Source account is not active.")
        If receiver.Status <> "Active" Then Return (False, "Destination account is not active.")
        If sender.Balance < amount Then Return (False, $"Insufficient funds. Available: {FormatMoney(sender.Balance)}")

        Dim senderBefore = sender.Balance
        Dim receiverBefore = receiver.Balance
        sender.Balance -= amount
        receiver.Balance += amount

        Dim who = If(CurrentEmployee IsNot Nothing, CurrentEmployee.EmployeeID, "SYSTEM")
        Dim stamp = DateTime.Now

        Dim txnOut As New TransactionRecord With {
            .TransactionID = GenerateTransactionID(),
            .AccountNumber = fromAccount,
            .TransactionType = "Transfer-Out",
            .Amount = amount,
            .BalanceBefore = senderBefore,
            .BalanceAfter = sender.Balance,
            .Timestamp = stamp,
            .PerformedBy = who,
            .Notes = $"Transfer to {toAccount}. {notes}"
        }

        Dim txnIn As New TransactionRecord With {
            .TransactionID = GenerateTransactionID(),
            .AccountNumber = toAccount,
            .TransactionType = "Transfer-In",
            .Amount = amount,
            .BalanceBefore = receiverBefore,
            .BalanceAfter = receiver.Balance,
            .Timestamp = stamp,
            .PerformedBy = who,
            .Notes = $"Transfer from {fromAccount}. {notes}"
        }

        SaveCustomer(sender)
        SaveCustomer(receiver)
        SaveTransaction(txnOut)
        SaveTransaction(txnIn)
        WriteAudit("TRANSFER", $"From={fromAccount}, To={toAccount}, Amount={FormatMoney(amount)}")
        Return (True, $"Transfer successful. Your new balance: {FormatMoney(sender.Balance)}")
    End Function

    ' ── Generate Report Summary ────────────────────────────────
    Public Function GenerateSummaryReport() As ReportSummary
        Dim customers = LoadAllCustomers()
        Dim transactions = LoadAllTransactions()
        Dim rpt As New ReportSummary()

        rpt.TotalCustomers = customers.Count
        rpt.ActiveAccounts = customers.Count(Function(c) c.Status = "Active")
        rpt.FrozenAccounts = customers.Count(Function(c) c.Status = "Frozen")
        rpt.ClosedAccounts = customers.Count(Function(c) c.Status = "Closed")
        rpt.TotalBalance = customers.Sum(Function(c) c.Balance)
        rpt.TotalDeposits = transactions.Where(Function(t) t.TransactionType = "Deposit" OrElse t.TransactionType = "Transfer-In").Sum(Function(t) t.Amount)
        rpt.TotalWithdrawals = transactions.Where(Function(t) t.TransactionType = "Withdrawal" OrElse t.TransactionType = "Transfer-Out").Sum(Function(t) t.Amount)
        rpt.TransactionCount = transactions.Count
        Return rpt
    End Function

End Module
