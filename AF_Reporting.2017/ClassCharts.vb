Imports System.Data
Imports System.Data.SqlClient

Friend Class ClassCharts
    Implements IDisposable

#Region "  Class Members  "

    Private _fiscalyear As Int32
    'database vars
    Private _connString As String
    Private cn As SqlConnection

#End Region

#Region "  Initialisation & Disposal  "

    Private disposed As Boolean     'flags whether object is disposed

    Sub New()
        Dim obj As AF_Master.Authuser
        Try
            _connString = obj.ConnectionString
            _fiscalyear = obj.FiscalYear
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw
        Finally
            obj = Nothing
        End Try
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        Dispose(True)
    End Sub

    Protected Sub Dispose(ByVal disposing As Boolean)
        'exit routine if object already disposed
        If disposed Then Exit Sub
        If disposing Then
            'disposal occurs here, not finalization; access to other objects allowed
            If Not cn Is Nothing Then cn.Dispose()
        End If
        'dispose of objects irrelevant of finalization mode
        disposed = True
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

#End Region

#Region "  Properties "

    Private ReadOnly Property FiscalYear() As Int32       'current year
        Get
            Return _fiscalyear
        End Get
    End Property

#End Region

#Region "  Retrieval Methods "

    Friend Function GetMonthlyAcctBalances(ByVal bankaccountnumber As String) As DataTable

        'For Chart - Transactions - Receipts
        'this method retrieves all transaction details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_month_balance,af_mtd_receipts,af_mtd_expend," _
        & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust) as Total" _
        & " From acct_info " _
        & " WHERE Bank_acct_num = @p1" _
        & " and (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust)  <> '0'" _
        & " and af_status = 'O'" _
        & " ORDER BY total asc"




        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            MsgBox(ex.ToString)


        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMonthlyAcctBalancesReceipts(ByVal bankaccountnumber As String) As DataTable
        'For Chart - Transactions - Receipts
        'this method retrieves all transaction details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_month_balance,af_mtd_receipts,af_mtd_expend," _
        & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust) as Total" _
        & " From acct_info " _
        & " WHERE Bank_acct_num = @p1" _
        & " and af_mtd_receipts  <> '0' " _
        & " and af_status = 'O'" _
        & " ORDER BY af_mtd_receipts asc"

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            MsgBox(ex.ToString)


        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMonthlyAcctBalancesExpense(ByVal bankaccountnumber As String) As DataTable
        'For Chart - Transactions - Checks
        'this method retrieves all transaction details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_month_balance,af_mtd_receipts,af_mtd_expend," _
        & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust) as Total" _
        & " From acct_info " _
        & " WHERE Bank_acct_num = @p1" _
        & " and af_mtd_expend  <> '0' " _
        & " and af_status = 'O'" _
        & " ORDER BY af_mtd_expend asc"

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            MsgBox(ex.ToString)


        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMonthlyAcctBalancesAdjustments(ByVal bankaccountnumber As String) As DataTable
        'For Chart - Transactions - Adjustments
        'this method retrieves all transaction details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_month_balance,af_mtd_receipts,af_mtd_expend," _
        & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust) as Total" _
        & " From acct_info " _
        & " WHERE Bank_acct_num = @p1" _
        & " and af_mtd_adjust  <> '0' " _
        & " and af_status = 'O'" _
        & " ORDER BY af_mtd_adjust asc"

        '

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            MsgBox(ex.ToString)


        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDAcctBalances(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_year_balance,af_ytd_receipts,af_ytd_expend," _
        & " af_ytd_adjust, (af_beg_year_balance + af_ytd_receipts - " _
        & " af_ytd_expend + af_ytd_adjust) as Total " _
        & " From acct_info " _
        & " WHERE Bank_acct_num = @p1 " _
        & " and af_status = 'O' " _
        & " ORDER BY af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMonthlyAcctBalancesAllBanks(ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details for All banks
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_month_balance,af_mtd_receipts,af_mtd_expend," _
        & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - " _
        & " af_mtd_expend + af_mtd_adjust) as Total " _
        & " From acct_info " _
        & " WHERE af_status = 'O' " _
        & " ORDER BY bank_acct_num,af_acct_num"

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreportallbanks")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDAcctBalancesAllBanks(ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details for All banks
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " af_status,af_beg_year_balance,af_ytd_receipts,af_ytd_expend," _
        & " af_ytd_adjust, (af_beg_year_balance + af_ytd_receipts - " _
        & " af_ytd_expend + af_ytd_adjust) as Total " _
        & " From acct_info " _
        & " WHERE af_status = 'O' " _
        & " ORDER BY bank_acct_num,af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreportallbanks")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try
    End Function

    Friend Function GetMonthlyAcctSubBalances(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
            & "	asub.as_acct_num,asub.as_acct_name," _
            & " ai.af_status,ai.af_beg_month_balance,ai.af_mtd_receipts,ai.af_mtd_expend," _
            & " ai.af_mtd_adjust, (ai.af_beg_month_balance + ai.af_mtd_receipts - " _
            & " ai.af_mtd_expend + ai.af_mtd_adjust) as Totalai," _
            & " asub.as_beg_month_balance, asub.as_mtd_receipts, " _
            & " asub.as_mtd_expend, asub.as_mtd_adjust," _
            & " (asub.as_beg_month_balance + asub.as_mtd_receipts - " _
            & " asub.as_mtd_expend + asub.as_mtd_adjust) as Totalasub," _
            & " ai.af_transdate,asub.as_transdate " _
            & " FROM acct_info as ai,acct_sub as asub " _
            & " WHERE ai.af_status = 'O' " _
            & " AND ai.bank_acct_num = @p1" _
            & " AND ai.bank_acct_num = asub.bank_acct_num" _
            & " AND ai.af_acct_num = asub.af_acct_num" _
            & " ORDER BY ai.bank_acct_num,ai.af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDAcctSubBalances(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
      & "	asub.as_acct_num,asub.as_acct_name," _
      & " ai.af_status,ai.af_beg_year_balance,ai.af_ytd_receipts,ai.af_ytd_expend," _
      & " ai.af_ytd_adjust, (ai.af_beg_year_balance + ai.af_ytd_receipts - " _
      & " ai.af_ytd_expend + ai.af_ytd_adjust) as Totalai," _
      & " asub.as_beg_year_balance, asub.as_ytd_receipts, " _
      & " asub.as_ytd_expend, asub.as_ytd_adjust," _
      & " (asub.as_beg_year_balance + asub.as_ytd_receipts - " _
      & " asub.as_ytd_expend + asub.as_ytd_adjust) as Totalasub," _
      & " ai.af_transdate,asub.as_transdate " _
      & " FROM acct_info as ai,acct_sub as asub " _
      & " WHERE ai.af_status = 'O' " _
      & " AND ai.bank_acct_num = asub.bank_acct_num" _
      & " AND ai.bank_acct_num = @p1" _
      & " AND ai.af_acct_num = asub.af_acct_num" _
      & " ORDER BY ai.bank_acct_num,ai.af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMonthlyAcctSubBalancesAllBanks(ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details for All banks
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
           & "	asub.as_acct_num,asub.as_acct_name," _
           & " ai.af_status,ai.af_beg_month_balance,ai.af_mtd_receipts,ai.af_mtd_expend," _
           & " ai.af_mtd_adjust, (ai.af_beg_month_balance + ai.af_mtd_receipts - " _
           & " ai.af_mtd_expend + ai.af_mtd_adjust) as Totalai," _
           & " asub.as_beg_month_balance, asub.as_mtd_receipts, " _
           & " asub.as_mtd_expend, asub.as_mtd_adjust," _
           & " (asub.as_beg_month_balance + asub.as_mtd_receipts - " _
           & " asub.as_mtd_expend + asub.as_mtd_adjust) as Totalasub," _
           & " ai.af_transdate,asub.as_transdate " _
           & " FROM acct_info as ai,acct_sub as asub " _
           & " WHERE ai.af_status = 'O' " _
           & " AND ai.bank_acct_num = asub.bank_acct_num" _
           & " AND ai.af_acct_num = asub.af_acct_num" _
           & " ORDER BY ai.bank_acct_num,ai.af_acct_num"

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("monthaccountreportallbanks")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDAcctSubBalancesAllBanks(ByVal fiscalyear As Int32) As Object

        'this method retrieves all account details for All banks
        'for viewing & returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
        & "	asub.as_acct_num,asub.as_acct_name," _
        & " ai.af_status,ai.af_beg_year_balance,ai.af_ytd_receipts,ai.af_ytd_expend," _
        & " ai.af_ytd_adjust, (ai.af_beg_year_balance + ai.af_ytd_receipts - " _
        & " ai.af_ytd_expend + ai.af_ytd_adjust) as Totalai," _
        & " asub.as_beg_year_balance, asub.as_ytd_receipts, " _
        & " asub.as_ytd_expend, asub.as_ytd_adjust," _
        & " (asub.as_beg_year_balance + asub.as_ytd_receipts - " _
        & " asub.as_ytd_expend + asub.as_ytd_adjust) as Totalasub," _
        & " ai.af_transdate,asub.as_transdate " _
        & " FROM acct_info as ai,acct_sub as asub " _
        & " WHERE ai.af_status = 'O' " _
        & " AND ai.bank_acct_num = asub.bank_acct_num" _
        & " AND ai.af_acct_num = asub.af_acct_num" _
        & " ORDER BY ai.bank_acct_num,ai.af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("ytdaccountsubreportallbanks")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try
    End Function

    Friend Function GetYTDtransactions(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32) As Object

        'this method retrieves all receipts,checks & adjustments transactions
        'from ceratin bank account number
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd WHERE rd.rcpt_num = ri.rcpt_num " _
        & " AND ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " and ri.bank_acct_num = @p2 and ri.rcpt_fisyr = @p1" _
        & " ORDER BY ri.rcpt_applied_date"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE(ci.bank_acct_num = cd.bank_acct_num) " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & " and ci.bank_acct_num = @p2 and ci.chks_fisyr = @p1 " _
        & " ORDER BY ci.chks_applied_date"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 and tr.bank_acct_num = @p2" _
        & " ORDER BY tr.tran_applied_date"


        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsAllBanks(ByVal fiscalyear As Int32) As Object

        'this method retrieves all receipts,checks & adjustments transactions
        'from all banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE(ci.bank_acct_num = cd.bank_acct_num) " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " ORDER BY  tr.bank_acct_num"


        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMTDtransactions(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments transactions
        'from certain bank account number for the current month
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " and ri.bank_acct_num = @p2 and ri.rcpt_fisyr = @p1" _
        & " AND ri.rcpt_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY ri.rcpt_applied_date"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE(ci.bank_acct_num = cd.bank_acct_num) " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & " and ci.chks_status = 'O' " _
        & " and ci.bank_acct_num = @p2 and ci.chks_fisyr = @p1 " _
        & " and ci.chks_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY ci.chks_applied_date"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 and tr.bank_acct_num = @p2" _
        & " AND tr.tran_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY tr.tran_applied_date"

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", bankaccountnumber)
        cmd.Parameters.Add("@p3", monthbegindate)
        cmd.Parameters.Add("@p4", monthenddate)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMTDtransactionsAllBanks(ByVal fiscalyear As Int32, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments transactions
        'from all banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND ri.rcpt_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE(ci.bank_acct_num = cd.bank_acct_num) " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & " and ci.chks_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.tran_applied_date BETWEEN @p3 and @p4 " _
        & " ORDER BY  tr.bank_acct_num"


        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p3", monthbegindate)
        cmd.Parameters.Add("@p4", monthenddate)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsByOcas(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for certain bank account
        'for viewing & returns a dataset cast as a generic object
        'MULTI BANK - YEARLY TRANSACTIONS BY OCASCODE



        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND rd.bank_acct_num = @p4 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and cd.bank_acct_num = @p4 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocrv_code LIKE @p2 " _
        & " AND tr.bank_acct_num = @p4 " _
        & " ORDER BY  tr.bank_acct_num"
        '& " AND tr.ocex_code LIKE @p3 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p4", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsByOcasExpense(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for certain bank account
        'for viewing & returns a dataset cast as a generic object
        'MULTI BANK - YEARLY TRANSACTIONS BY OCASCODE



        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND rd.bank_acct_num = @p4 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and cd.bank_acct_num = @p4 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocex_code LIKE @p3 " _
        & " AND tr.bank_acct_num = @p4 " _
        & " ORDER BY  tr.bank_acct_num"

        '& " AND tr.ocrv_code LIKE @p2 " _
        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p4", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsAllBanksByOcas(ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code for all banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocrv_code LIKE @p2 " _
        & " ORDER BY  tr.bank_acct_num"
        '& " AND tr.ocex_code LIKE @p3 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsAllBanksByOcasExpense(ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code for all banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocex_code LIKE @p3 " _
        & " ORDER BY  tr.bank_acct_num"

        '  & " AND tr.ocrv_code LIKE @p2 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsByOcasByMonth(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for current Month 
        '& for certain Bank account
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND rd.bank_acct_num = @p4 " _
        & " AND ri.rcpt_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and cd.bank_acct_num = @p4 " _
        & "  and ci.chks_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocrv_code LIKE @p2 " _
        & " AND tr.bank_acct_num = @p4 " _
        & " AND tr.tran_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY tr.tran_applied_date"

        '& " AND tr.ocex_code LIKE @p3 " _

        '@p3 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p4", bankaccountnumber)
        cmd.Parameters.Add("@p5", monthbegindate)
        cmd.Parameters.Add("@p6", monthenddate)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery1")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsByOcasByMonthExpense(ByVal bankaccountnumber As String, ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for current Month 
        '& for certain Bank account
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND rd.bank_acct_num = @p4 " _
        & " AND ri.rcpt_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and cd.bank_acct_num = @p4 " _
        & "  and ci.chks_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocex_code LIKE @p3 " _
        & " AND tr.bank_acct_num = @p4 " _
        & " AND tr.tran_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY tr.tran_applied_date"


        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p4", bankaccountnumber)
        cmd.Parameters.Add("@p5", monthbegindate)
        cmd.Parameters.Add("@p6", monthenddate)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery1")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsAllBanksByOcasByMonth(ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for current Month
        'For All Banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND ri.rcpt_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and ci.chks_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocrv_code LIKE @p2 " _
        & " AND tr.tran_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY  tr.bank_acct_num"
        ' & " AND tr.ocex_code LIKE @p3 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p5", monthbegindate)
        cmd.Parameters.Add("@p6", monthenddate)


        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetYTDtransactionsAllBanksByOcasByMonthExpense(ByVal fiscalyear As Int32, ByVal ocasrevenuecode As String, ByVal ocasexpcode As String, ByVal monthbegindate As Date, ByVal monthenddate As Date) As Object

        'this method retrieves all receipts,checks & adjustments
        'By Ocas Code & for current Month
        'For All Banks
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL, SSQL1, SSQL2 As String
        SSQL = " Select ri.rcpt_applied_date,rd.af_acct_num,rd.as_acct_num," _
        & " ri.rcpt_num, rd.ocrv_code,ri.rcpt_rcvd_from, rd.rcdt_remarks, " _
        & " ri.rcpt_status," _
        & " ri.bank_acct_num,rd.rcdt_amount, (SELECT sum(rd.rcdt_amount)" _
        & " as TotalAmount " _
        & " FROM receipt_detl as rd where rd.rcpt_num = ri.rcpt_num " _
        & " and ri.bank_acct_num = rd.bank_acct_num group by " _
        & " ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_transdate," _
        & " ri.rcpt_rcvd_from,ri.rcpt_status, ri.vend_number, ri.rcpt_recon_sw," _
        & " bi.bank_cur_balance) FROM receipt_info as ri ,receipt_detl as rd," _
        & " bank_info as bi WHERE ri.bank_acct_num = rd.bank_acct_num " _
        & " and ri.rcpt_fisyr = rd.rcpt_fisyr and rd.rcpt_num = ri.rcpt_num " _
        & " and rd.bank_acct_num = bi.bank_acct_num and ri.rcpt_status = 'O' " _
        & " AND ri.rcpt_fisyr = @p1" _
        & " AND rd.ocrv_code LIKE @p2" _
        & " AND ri.rcpt_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ri.bank_acct_num"

        SSQL1 = " Select ci.chks_applied_date,cd.af_acct_num,cd.as_acct_num," _
        & " ci.chks_num,cd.ocex_code," _
        & " ci.chks_payee_name,ci.chks_status, ci.po_num, " _
        & " ci.chks_descr, ci.bank_acct_num,cd.ckdt_amount, " _
        & " ci.chks_amount,cd.chks_autoinc_key,ci.chks_autoinc_key " _
        & " FROM chks_detl as cd,chks_info as ci " _
        & " WHERE ci.bank_acct_num = cd.bank_acct_num " _
        & " AND cd.chks_autoinc_key = ci.chks_autoinc_key " _
        & "  and ci.chks_status = 'O' " _
        & "  and ci.chks_fisyr = @p1 " _
        & "  and cd.ocex_code LIKE @p3 " _
        & "  and ci.chks_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY ci.bank_acct_num"

        SSQL2 = " SELECT tr.tran_applied_date, tr.af_acct_num, tr.as_acct_num, tr.tran_type, tr.ocex_code," _
        & " tr.ocrv_code, tr.tran_descr, tr.tran_amt, tr.tran_fisyr, tr.bank_acct_num" _
        & " FROM transactions as tr WHERE tr.tran_fisyr = @p1 " _
        & " AND tr.ocex_code LIKE @p3 " _
        & " AND tr.tran_applied_date BETWEEN @p5 and @p6 " _
        & " ORDER BY  tr.bank_acct_num"
        ' & " AND tr.ocrv_code LIKE @p2 " _

        SSQL &= (SSQL1 & SSQL2)

        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", fiscalyear)
        cmd.Parameters.Add("@p2", ocasrevenuecode)
        cmd.Parameters.Add("@p3", ocasexpcode)
        cmd.Parameters.Add("@p5", monthbegindate)
        cmd.Parameters.Add("@p6", monthenddate)


        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("ocasquery")

        Try
            da.Fill(ds)
            Return ds

        Catch ex As Exception
            MsgBox(ex.ToString)
            Throw
        Finally
            cn.Close()
            da.Dispose()
            ds.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMTDAcctFilterAllBanks(ByVal acctfinrange1 As String, ByVal acctfinrange2 As String) As DataTable
        'this method retrieves all subaccounts within accounts 
        'within certain account range
        'returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
            & "	asub.as_acct_num,asub.as_acct_name," _
            & " ai.af_status,ai.af_beg_month_balance,ai.af_mtd_receipts,ai.af_mtd_expend," _
            & " ai.af_mtd_adjust, (ai.af_beg_month_balance + ai.af_mtd_receipts - " _
            & " ai.af_mtd_expend + ai.af_mtd_adjust) as Totalai," _
            & " asub.as_beg_month_balance, asub.as_mtd_receipts, " _
            & " asub.as_mtd_expend, asub.as_mtd_adjust," _
            & " (asub.as_beg_month_balance + asub.as_mtd_receipts - " _
            & " asub.as_mtd_expend + asub.as_mtd_adjust) as Totalasub," _
            & " ai.af_transdate,asub.as_transdate " _
            & " FROM acct_info as ai,acct_sub as asub " _
            & " WHERE ai.af_status = 'O' " _
            & " AND ai.af_acct_num BETWEEN @p1 and @p2" _
            & " AND ai.bank_acct_num = asub.bank_acct_num" _
            & " AND ai.af_acct_num = asub.af_acct_num" _
            & " ORDER BY ai.bank_acct_num,ai.af_acct_num"


        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", acctfinrange1)
        cmd.Parameters.Add("@p2", acctfinrange2)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("finsearchreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

    Friend Function GetMTDAcctFilter(ByVal bankaccountnumber As String, ByVal acctfinrange1 As String, ByVal acctfinrange2 As String) As DataTable
        'this method retrieves all subaccounts within accounts and
        'within certain account range
        'returns a datatable cast as a generic object

        Dim SSQL As String

        SSQL = "SELECT ai.bank_acct_num, ai.af_acct_num, ai.af_acct_name," _
            & "	asub.as_acct_num,asub.as_acct_name," _
            & " ai.af_status,ai.af_beg_month_balance,ai.af_mtd_receipts,ai.af_mtd_expend," _
            & " ai.af_mtd_adjust, (ai.af_beg_month_balance + ai.af_mtd_receipts - " _
            & " ai.af_mtd_expend + ai.af_mtd_adjust) as Totalai," _
            & " asub.as_beg_month_balance, asub.as_mtd_receipts, " _
            & " asub.as_mtd_expend, asub.as_mtd_adjust," _
            & " (asub.as_beg_month_balance + asub.as_mtd_receipts - " _
            & " asub.as_mtd_expend + asub.as_mtd_adjust) as Totalasub," _
            & " ai.af_transdate,asub.as_transdate " _
            & " FROM acct_info as ai,acct_sub as asub " _
            & " WHERE ai.af_status = 'O' " _
            & " AND ai.af_acct_num BETWEEN @p1 and @p2" _
            & " AND asub.bank_acct_num = @p3" _
            & " AND ai.bank_acct_num = asub.bank_acct_num" _
            & " AND ai.af_acct_num = asub.af_acct_num" _
            & " ORDER BY ai.bank_acct_num,ai.af_acct_num"



        cn = New SqlConnection(_connString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", acctfinrange1)
        cmd.Parameters.Add("@p2", acctfinrange2)
        cmd.Parameters.Add("@p3", bankaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("finsearchreport")

        Try
            da.Fill(tbl)
            Return tbl                  'return the datatable as an object

        Catch ex As Exception
            Throw

        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

    End Function

#End Region

End Class

