Imports Microsoft.Win32
Imports System.Data
Imports System.Data.SqlClient

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region "  Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Dim authobj As AF_Master.Authuser
        Try
            'handle permissions;
            Me.Administrator = authobj.IsAdministrator
            Me.UAccess = Module1.GetPermissions(authobj.ConnectionString, authobj.UserKey, 820, Me.Administrator)
            If Me.UAccess < 1 Then Throw New ArgumentException("Permission denied.")
            '
            Me.ConnectionString = authobj.ConnectionString
            Call GetBanks()
            Call GetRegistryEntries()
            '
            Me.FiscalYear = authobj.FiscalYear
            Me.LastDepositNumber = authobj.NextDepositNumber
            Me.NextPONumber = authobj.NextPurchaseOrderNumber
            Me.CurrentMonthBeginning = authobj.CurrentMonthBeginning
            Me.CurrentMonthEnding = authobj.CurrentMonthEnding
            Me.CurrentMonthString = authobj.CurrentMonthString
            Me.SchoolNumber = authobj.SchoolNumber
            Me.SiteNumber = authobj.SchoolDatabaseNumber
            Me.lblSchoolName.Text = authobj.SchoolName
            'if ocas is disabled then change color of the bold code button;
            If Not authobj.UseOCAS Then Me.lblBoldCodeMessage.Visible = True
        Catch ex As Exception
            Throw
        End Try

        Try
            Call InitForm()
            Call GetFiscalYears()
            Call GetAdjustNumbers()
            'update the statusbar;
            Me.StatusBar1.Panels(0).Text = "FY-" & Me.FiscalYear.ToString
            Me.StatusBar1.Panels(1).Text = Me.CurrentMonthString
            Me.StatusBar1.Panels(2).Text = Me.BankAccountNumber
        Catch ex As Exception
            Throw
        End Try

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rdoAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpenditure As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRevenue As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancials As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAdjustments As System.Windows.Forms.RadioButton
    Friend WithEvents rdoVendors As System.Windows.Forms.RadioButton
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents grpAccounts As System.Windows.Forms.GroupBox
    Friend WithEvents grpAdjustments As System.Windows.Forms.GroupBox
    Friend WithEvents grpExpenditures As System.Windows.Forms.GroupBox
    Friend WithEvents grpRevenue As System.Windows.Forms.GroupBox
    Friend WithEvents grpFinancials As System.Windows.Forms.GroupBox
    Friend WithEvents grpVendors As System.Windows.Forms.GroupBox
    Friend WithEvents rdoBoldCode As System.Windows.Forms.RadioButton
    Friend WithEvents grpBoldCode As System.Windows.Forms.GroupBox
    Friend WithEvents cboFiscalYears As System.Windows.Forms.ComboBox
    Friend WithEvents chkAccountsIncludeSubaccounts As System.Windows.Forms.CheckBox
    Friend WithEvents rdoAccountsChartOfAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents grpNumberRange As System.Windows.Forms.GroupBox
    Friend WithEvents dtBeginningDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtEndingDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtBeginningNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtEndingNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents rdoRevenueReceiptRegister As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRevenueReceiptTicket As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRevenueDailyDeposit As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRevenueDepositSummary As System.Windows.Forms.RadioButton
    Friend WithEvents chkUseDate As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseNumber As System.Windows.Forms.CheckBox
    Friend WithEvents panelOptions As System.Windows.Forms.Panel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents btnCancelPreview As System.Windows.Forms.Button
    Friend WithEvents StatusBarPanel4 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents chkRevenuesPrintDepositTicket As System.Windows.Forms.CheckBox
    Friend WithEvents chkRevenuesIncludeCreditCards As System.Windows.Forms.CheckBox
    Friend WithEvents lblSchoolName As System.Windows.Forms.Label
    Friend WithEvents txtDepositBegNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtDepositEndNumber As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cboAccountsMonthListing As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents panelAccountsSelectMonth As System.Windows.Forms.Panel
    Friend WithEvents rdoAccountsMTDSummaryOfAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAccountsYTDSummaryOfAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsDetailOfAccountMTD As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsDetailOfAccountYTD As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsDetailOfAccountPeriodical As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsSelectAccount As System.Windows.Forms.RadioButton
    Friend WithEvents lblAccountNumber As System.Windows.Forms.Label
    Friend WithEvents lblAccountName As System.Windows.Forms.Label
    Friend WithEvents panelFinancialsSelectAccount As System.Windows.Forms.Panel
    Friend WithEvents rdoFinancialsAllAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents grpReconciliation As System.Windows.Forms.GroupBox
    Friend WithEvents panelReconciliationTrialBalance As System.Windows.Forms.Panel
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtReconBankStatementBalance As System.Windows.Forms.TextBox
    Friend WithEvents txtReconInterestNotYetPosted As System.Windows.Forms.TextBox
    Friend WithEvents txtReconExpensesNotYetPosted As System.Windows.Forms.TextBox
    Friend WithEvents txtReconInvestments As System.Windows.Forms.TextBox
    Friend WithEvents rdoReconciliation As System.Windows.Forms.RadioButton
    Friend WithEvents dtReconTrialBalanceDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents rdoReconTrialBalance As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresCheckRegister As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresPrintPurchaseOrder As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsNoMtdDetail As System.Windows.Forms.RadioButton
    Friend WithEvents rdoFinancialsNoYtdDetail As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresPoRegister As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresPrintVoidChecks As System.Windows.Forms.RadioButton
    Friend WithEvents rdoRevenuePrintVoidReceipts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAdjustmentsPrintAdjustmentRegister As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAdjustmentsPrintTransferRegister As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresPrintOutstandingChecks As System.Windows.Forms.RadioButton
    Friend WithEvents chkExpendituresAllFiscalYears As System.Windows.Forms.CheckBox
    Friend WithEvents chkRevenuesAllFiscalYears As System.Windows.Forms.CheckBox
    Friend WithEvents rdoBoldCodeListingByExpenditures As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBoldCodeListingByRevenues As System.Windows.Forms.RadioButton
    Friend WithEvents chkBoldCodeDetailErrorsOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkBoldCodeSortDetailByCoding As System.Windows.Forms.CheckBox
    Friend WithEvents lblBoldCodeMessage As System.Windows.Forms.Label
    Friend WithEvents rdoVendorListing As System.Windows.Forms.RadioButton
    Friend WithEvents rdoVendorExpenses As System.Windows.Forms.RadioButton
    Friend WithEvents chkVendorsIncludeZeroBalances As System.Windows.Forms.CheckBox
    Friend WithEvents rdoVendor1099Listing As System.Windows.Forms.RadioButton
    Friend WithEvents cboCalendarYears As System.Windows.Forms.ComboBox
    Friend WithEvents rdoAccountsHistoricalMTDSummaryOfAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoAccountsHistoricalYTDSummaryOfAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents cboVendors As System.Windows.Forms.ComboBox
    Friend WithEvents chkVendorsSelectSingleVendor As System.Windows.Forms.CheckBox
    Friend WithEvents chkVendorsUse600Minimum As System.Windows.Forms.CheckBox
    Friend WithEvents rdoAccountsBalanceSheet As System.Windows.Forms.RadioButton
    Friend WithEvents chkVendorsIncludeSSN As System.Windows.Forms.CheckBox
    Friend WithEvents chkVendorsUseFiscalYear As System.Windows.Forms.CheckBox
    Friend WithEvents panelAccountsSelectAccountRange As System.Windows.Forms.Panel
    Friend WithEvents txtAccountsAcctNumberTo As System.Windows.Forms.TextBox
    Friend WithEvents txtAccountsAcctNumberFrom As System.Windows.Forms.TextBox
    Friend WithEvents chkAccountsUseAccountRange As System.Windows.Forms.CheckBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents rdoClassification As System.Windows.Forms.RadioButton
    Friend WithEvents grpClassification As System.Windows.Forms.GroupBox
    Friend WithEvents rdoExpendituresEncumbranceAccounts As System.Windows.Forms.RadioButton
    Friend WithEvents chkFinancialsEncumbranceDetail As System.Windows.Forms.CheckBox
    Friend WithEvents rdoRevenuePrintOutstandingReceipts As System.Windows.Forms.RadioButton
    Friend WithEvents rdoExpendituresPendingInvoice As System.Windows.Forms.RadioButton
    Friend WithEvents txtDim1 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim2 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim3 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim4 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim5 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim6 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim7 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim8 As System.Windows.Forms.TextBox
    Friend WithEvents txtDim9 As System.Windows.Forms.TextBox
    Friend WithEvents rdoClassificationYTDRevenue As System.Windows.Forms.RadioButton
    Friend WithEvents rdoClassificationYTDExpend As System.Windows.Forms.RadioButton
    Friend WithEvents rdoClassificationMTDRevenue As System.Windows.Forms.RadioButton
    Friend WithEvents rdoClassificationMTDExpend As System.Windows.Forms.RadioButton
    Friend WithEvents chkClassificationUseCodeRange As System.Windows.Forms.CheckBox
    Friend WithEvents panelClassificationCodes As System.Windows.Forms.Panel
    Friend WithEvents panelRevenueSearchReceived As System.Windows.Forms.Panel
    Friend WithEvents chkRevenuesSearchReceived As System.Windows.Forms.CheckBox
    Friend WithEvents txtRevenueSearch As System.Windows.Forms.TextBox
    Friend WithEvents panelRevenueAllFiscalYears As System.Windows.Forms.Panel
    Friend WithEvents panelRevenue1098T As System.Windows.Forms.Panel
    Friend WithEvents panelRevenueDailyDeposit As System.Windows.Forms.Panel
    Friend WithEvents rdoRevenue1098T As System.Windows.Forms.RadioButton
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtRevenueSubaccountNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtRevenueAccountNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents chkExpendituresOutstandingInvoices As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cboBanks As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents grpDateRange As System.Windows.Forms.GroupBox
    Friend WithEvents chkAccountsSuppressZeros As System.Windows.Forms.CheckBox
    Friend WithEvents chk1099ByEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents chk1099Summary As System.Windows.Forms.CheckBox
    Friend WithEvents rdoClassificationExpenditureCodes As System.Windows.Forms.RadioButton
    Friend WithEvents rdoClassificationRevenueCodes As System.Windows.Forms.RadioButton
    Friend WithEvents rdoVendorAudit As System.Windows.Forms.RadioButton
    Friend WithEvents lblVendorCalendar As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents chkAccountsIncludeEncumbrances As System.Windows.Forms.CheckBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents rdoExpendituresPositivePay As System.Windows.Forms.RadioButton
    Friend WithEvents lblExpendituresPosPay As System.Windows.Forms.Label
    Friend WithEvents grpCheck As System.Windows.Forms.GroupBox
    Friend WithEvents btnVerifyRev As System.Windows.Forms.Button
    Friend WithEvents btnVerifyExp As System.Windows.Forms.Button
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents txtC1SiteExp As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Job As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Subject As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1ProgramExp As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Object As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Function As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1ProjectExp As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1FundExp As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1YearExp As C1.Win.C1Input.C1TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtC1Site As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Program As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Source As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Project As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Fund As C1.Win.C1Input.C1TextBox
    Friend WithEvents txtC1Year As C1.Win.C1Input.C1TextBox
    Friend WithEvents lblBoldRev As System.Windows.Forms.Label
    Friend WithEvents lblBoldExp As System.Windows.Forms.Label
    Friend WithEvents chkChkaCode As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.cboFiscalYears = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rdoAccounts = New System.Windows.Forms.RadioButton()
        Me.rdoExpenditure = New System.Windows.Forms.RadioButton()
        Me.rdoRevenue = New System.Windows.Forms.RadioButton()
        Me.rdoFinancials = New System.Windows.Forms.RadioButton()
        Me.rdoAdjustments = New System.Windows.Forms.RadioButton()
        Me.rdoVendors = New System.Windows.Forms.RadioButton()
        Me.StatusBar1 = New System.Windows.Forms.StatusBar()
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel()
        Me.StatusBarPanel4 = New System.Windows.Forms.StatusBarPanel()
        Me.grpAccounts = New System.Windows.Forms.GroupBox()
        Me.chkAccountsSuppressZeros = New System.Windows.Forms.CheckBox()
        Me.chkAccountsIncludeEncumbrances = New System.Windows.Forms.CheckBox()
        Me.chkAccountsUseAccountRange = New System.Windows.Forms.CheckBox()
        Me.rdoAccountsBalanceSheet = New System.Windows.Forms.RadioButton()
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts = New System.Windows.Forms.RadioButton()
        Me.rdoAccountsYTDSummaryOfAccounts = New System.Windows.Forms.RadioButton()
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts = New System.Windows.Forms.RadioButton()
        Me.rdoAccountsMTDSummaryOfAccounts = New System.Windows.Forms.RadioButton()
        Me.chkAccountsIncludeSubaccounts = New System.Windows.Forms.CheckBox()
        Me.rdoAccountsChartOfAccounts = New System.Windows.Forms.RadioButton()
        Me.panelAccountsSelectAccountRange = New System.Windows.Forms.Panel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtAccountsAcctNumberTo = New System.Windows.Forms.TextBox()
        Me.txtAccountsAcctNumberFrom = New System.Windows.Forms.TextBox()
        Me.panelAccountsSelectMonth = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboAccountsMonthListing = New System.Windows.Forms.ComboBox()
        Me.grpAdjustments = New System.Windows.Forms.GroupBox()
        Me.rdoAdjustmentsPrintTransferRegister = New System.Windows.Forms.RadioButton()
        Me.rdoAdjustmentsPrintAdjustmentRegister = New System.Windows.Forms.RadioButton()
        Me.grpExpenditures = New System.Windows.Forms.GroupBox()
        Me.lblExpendituresPosPay = New System.Windows.Forms.Label()
        Me.rdoExpendituresPositivePay = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresPendingInvoice = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresEncumbranceAccounts = New System.Windows.Forms.RadioButton()
        Me.chkExpendituresOutstandingInvoices = New System.Windows.Forms.CheckBox()
        Me.chkExpendituresAllFiscalYears = New System.Windows.Forms.CheckBox()
        Me.rdoExpendituresPrintOutstandingChecks = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresPrintVoidChecks = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresPoRegister = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresPrintPurchaseOrder = New System.Windows.Forms.RadioButton()
        Me.rdoExpendituresCheckRegister = New System.Windows.Forms.RadioButton()
        Me.grpRevenue = New System.Windows.Forms.GroupBox()
        Me.panelRevenue1098T = New System.Windows.Forms.Panel()
        Me.txtRevenueSubaccountNumber = New System.Windows.Forms.TextBox()
        Me.txtRevenueAccountNumber = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.panelRevenueAllFiscalYears = New System.Windows.Forms.Panel()
        Me.chkRevenuesAllFiscalYears = New System.Windows.Forms.CheckBox()
        Me.panelRevenueSearchReceived = New System.Windows.Forms.Panel()
        Me.txtRevenueSearch = New System.Windows.Forms.TextBox()
        Me.chkRevenuesSearchReceived = New System.Windows.Forms.CheckBox()
        Me.panelRevenueDailyDeposit = New System.Windows.Forms.Panel()
        Me.txtDepositEndNumber = New System.Windows.Forms.TextBox()
        Me.chkRevenuesIncludeCreditCards = New System.Windows.Forms.CheckBox()
        Me.chkRevenuesPrintDepositTicket = New System.Windows.Forms.CheckBox()
        Me.txtDepositBegNumber = New System.Windows.Forms.TextBox()
        Me.rdoRevenuePrintOutstandingReceipts = New System.Windows.Forms.RadioButton()
        Me.rdoRevenuePrintVoidReceipts = New System.Windows.Forms.RadioButton()
        Me.rdoRevenueDepositSummary = New System.Windows.Forms.RadioButton()
        Me.rdoRevenueDailyDeposit = New System.Windows.Forms.RadioButton()
        Me.rdoRevenueReceiptTicket = New System.Windows.Forms.RadioButton()
        Me.rdoRevenueReceiptRegister = New System.Windows.Forms.RadioButton()
        Me.rdoRevenue1098T = New System.Windows.Forms.RadioButton()
        Me.grpFinancials = New System.Windows.Forms.GroupBox()
        Me.chkFinancialsEncumbranceDetail = New System.Windows.Forms.CheckBox()
        Me.rdoFinancialsNoYtdDetail = New System.Windows.Forms.RadioButton()
        Me.rdoFinancialsNoMtdDetail = New System.Windows.Forms.RadioButton()
        Me.panelFinancialsSelectAccount = New System.Windows.Forms.Panel()
        Me.lblAccountName = New System.Windows.Forms.Label()
        Me.lblAccountNumber = New System.Windows.Forms.Label()
        Me.rdoFinancialsAllAccounts = New System.Windows.Forms.RadioButton()
        Me.rdoFinancialsSelectAccount = New System.Windows.Forms.RadioButton()
        Me.rdoFinancialsDetailOfAccountPeriodical = New System.Windows.Forms.RadioButton()
        Me.rdoFinancialsDetailOfAccountYTD = New System.Windows.Forms.RadioButton()
        Me.rdoFinancialsDetailOfAccountMTD = New System.Windows.Forms.RadioButton()
        Me.grpVendors = New System.Windows.Forms.GroupBox()
        Me.rdoVendorAudit = New System.Windows.Forms.RadioButton()
        Me.chkVendorsUseFiscalYear = New System.Windows.Forms.CheckBox()
        Me.chkVendorsIncludeSSN = New System.Windows.Forms.CheckBox()
        Me.chkVendorsUse600Minimum = New System.Windows.Forms.CheckBox()
        Me.chkVendorsSelectSingleVendor = New System.Windows.Forms.CheckBox()
        Me.cboVendors = New System.Windows.Forms.ComboBox()
        Me.cboCalendarYears = New System.Windows.Forms.ComboBox()
        Me.rdoVendor1099Listing = New System.Windows.Forms.RadioButton()
        Me.chkVendorsIncludeZeroBalances = New System.Windows.Forms.CheckBox()
        Me.rdoVendorExpenses = New System.Windows.Forms.RadioButton()
        Me.chk1099ByEmployee = New System.Windows.Forms.CheckBox()
        Me.rdoVendorListing = New System.Windows.Forms.RadioButton()
        Me.lblVendorCalendar = New System.Windows.Forms.Label()
        Me.chk1099Summary = New System.Windows.Forms.CheckBox()
        Me.rdoBoldCode = New System.Windows.Forms.RadioButton()
        Me.grpBoldCode = New System.Windows.Forms.GroupBox()
        Me.chkChkaCode = New System.Windows.Forms.CheckBox()
        Me.grpCheck = New System.Windows.Forms.GroupBox()
        Me.btnVerifyRev = New System.Windows.Forms.Button()
        Me.btnVerifyExp = New System.Windows.Forms.Button()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.txtC1SiteExp = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Job = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Subject = New C1.Win.C1Input.C1TextBox()
        Me.txtC1ProgramExp = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Object = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Function = New C1.Win.C1Input.C1TextBox()
        Me.txtC1ProjectExp = New C1.Win.C1Input.C1TextBox()
        Me.txtC1FundExp = New C1.Win.C1Input.C1TextBox()
        Me.txtC1YearExp = New C1.Win.C1Input.C1TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtC1Site = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Program = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Source = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Project = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Fund = New C1.Win.C1Input.C1TextBox()
        Me.txtC1Year = New C1.Win.C1Input.C1TextBox()
        Me.lblBoldRev = New System.Windows.Forms.Label()
        Me.lblBoldExp = New System.Windows.Forms.Label()
        Me.lblBoldCodeMessage = New System.Windows.Forms.Label()
        Me.chkBoldCodeSortDetailByCoding = New System.Windows.Forms.CheckBox()
        Me.chkBoldCodeDetailErrorsOnly = New System.Windows.Forms.CheckBox()
        Me.rdoBoldCodeListingByRevenues = New System.Windows.Forms.RadioButton()
        Me.rdoBoldCodeListingByExpenditures = New System.Windows.Forms.RadioButton()
        Me.grpDateRange = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtEndingDate = New System.Windows.Forms.DateTimePicker()
        Me.dtBeginningDate = New System.Windows.Forms.DateTimePicker()
        Me.grpNumberRange = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtEndingNumber = New System.Windows.Forms.TextBox()
        Me.txtBeginningNumber = New System.Windows.Forms.TextBox()
        Me.chkUseDate = New System.Windows.Forms.CheckBox()
        Me.chkUseNumber = New System.Windows.Forms.CheckBox()
        Me.panelOptions = New System.Windows.Forms.Panel()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.btnCancelPreview = New System.Windows.Forms.Button()
        Me.lblSchoolName = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.rdoReconTrialBalance = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationRevenueCodes = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationExpenditureCodes = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationYTDRevenue = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationYTDExpend = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationMTDRevenue = New System.Windows.Forms.RadioButton()
        Me.rdoClassificationMTDExpend = New System.Windows.Forms.RadioButton()
        Me.panelReconciliationTrialBalance = New System.Windows.Forms.Panel()
        Me.dtReconTrialBalanceDate = New System.Windows.Forms.DateTimePicker()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtReconInvestments = New System.Windows.Forms.TextBox()
        Me.txtReconExpensesNotYetPosted = New System.Windows.Forms.TextBox()
        Me.txtReconInterestNotYetPosted = New System.Windows.Forms.TextBox()
        Me.txtReconBankStatementBalance = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.rdoReconciliation = New System.Windows.Forms.RadioButton()
        Me.grpReconciliation = New System.Windows.Forms.GroupBox()
        Me.rdoClassification = New System.Windows.Forms.RadioButton()
        Me.grpClassification = New System.Windows.Forms.GroupBox()
        Me.chkClassificationUseCodeRange = New System.Windows.Forms.CheckBox()
        Me.panelClassificationCodes = New System.Windows.Forms.Panel()
        Me.txtDim9 = New System.Windows.Forms.TextBox()
        Me.txtDim1 = New System.Windows.Forms.TextBox()
        Me.txtDim7 = New System.Windows.Forms.TextBox()
        Me.txtDim6 = New System.Windows.Forms.TextBox()
        Me.txtDim5 = New System.Windows.Forms.TextBox()
        Me.txtDim4 = New System.Windows.Forms.TextBox()
        Me.txtDim8 = New System.Windows.Forms.TextBox()
        Me.txtDim2 = New System.Windows.Forms.TextBox()
        Me.txtDim3 = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cboBanks = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAccounts.SuspendLayout()
        Me.panelAccountsSelectAccountRange.SuspendLayout()
        Me.panelAccountsSelectMonth.SuspendLayout()
        Me.grpAdjustments.SuspendLayout()
        Me.grpExpenditures.SuspendLayout()
        Me.grpRevenue.SuspendLayout()
        Me.panelRevenue1098T.SuspendLayout()
        Me.panelRevenueAllFiscalYears.SuspendLayout()
        Me.panelRevenueSearchReceived.SuspendLayout()
        Me.panelRevenueDailyDeposit.SuspendLayout()
        Me.grpFinancials.SuspendLayout()
        Me.panelFinancialsSelectAccount.SuspendLayout()
        Me.grpVendors.SuspendLayout()
        Me.grpBoldCode.SuspendLayout()
        Me.grpCheck.SuspendLayout()
        CType(Me.txtC1SiteExp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Job, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Subject, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1ProgramExp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Object, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Function, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1ProjectExp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1FundExp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1YearExp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Site, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Program, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Source, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Project, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Fund, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtC1Year, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDateRange.SuspendLayout()
        Me.grpNumberRange.SuspendLayout()
        Me.panelOptions.SuspendLayout()
        Me.panelReconciliationTrialBalance.SuspendLayout()
        Me.grpReconciliation.SuspendLayout()
        Me.grpClassification.SuspendLayout()
        Me.panelClassificationCodes.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboFiscalYears
        '
        Me.cboFiscalYears.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFiscalYears.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFiscalYears.Location = New System.Drawing.Point(240, 48)
        Me.cboFiscalYears.Name = "cboFiscalYears"
        Me.cboFiscalYears.Size = New System.Drawing.Size(64, 21)
        Me.cboFiscalYears.TabIndex = 0
        Me.cboFiscalYears.TabStop = False
        Me.ToolTip1.SetToolTip(Me.cboFiscalYears, " The currently selected fiscal year ")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(240, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Fiscal year:"
        '
        'rdoAccounts
        '
        Me.rdoAccounts.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccounts.Location = New System.Drawing.Point(16, 80)
        Me.rdoAccounts.Name = "rdoAccounts"
        Me.rdoAccounts.Size = New System.Drawing.Size(112, 16)
        Me.rdoAccounts.TabIndex = 2
        Me.rdoAccounts.Text = "Accounts"
        '
        'rdoExpenditure
        '
        Me.rdoExpenditure.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpenditure.Location = New System.Drawing.Point(16, 176)
        Me.rdoExpenditure.Name = "rdoExpenditure"
        Me.rdoExpenditure.Size = New System.Drawing.Size(112, 16)
        Me.rdoExpenditure.TabIndex = 3
        Me.rdoExpenditure.Text = "Expenditures"
        '
        'rdoRevenue
        '
        Me.rdoRevenue.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoRevenue.Location = New System.Drawing.Point(16, 248)
        Me.rdoRevenue.Name = "rdoRevenue"
        Me.rdoRevenue.Size = New System.Drawing.Size(112, 16)
        Me.rdoRevenue.TabIndex = 4
        Me.rdoRevenue.Text = "Revenues"
        '
        'rdoFinancials
        '
        Me.rdoFinancials.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoFinancials.Location = New System.Drawing.Point(16, 200)
        Me.rdoFinancials.Name = "rdoFinancials"
        Me.rdoFinancials.Size = New System.Drawing.Size(112, 16)
        Me.rdoFinancials.TabIndex = 5
        Me.rdoFinancials.Text = "Financials"
        '
        'rdoAdjustments
        '
        Me.rdoAdjustments.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAdjustments.Location = New System.Drawing.Point(16, 104)
        Me.rdoAdjustments.Name = "rdoAdjustments"
        Me.rdoAdjustments.Size = New System.Drawing.Size(112, 16)
        Me.rdoAdjustments.TabIndex = 6
        Me.rdoAdjustments.Text = "Adjustments"
        '
        'rdoVendors
        '
        Me.rdoVendors.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoVendors.Location = New System.Drawing.Point(16, 272)
        Me.rdoVendors.Name = "rdoVendors"
        Me.rdoVendors.Size = New System.Drawing.Size(112, 16)
        Me.rdoVendors.TabIndex = 7
        Me.rdoVendors.Text = "Vendors"
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 396)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2, Me.StatusBarPanel3, Me.StatusBarPanel4})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(754, 20)
        Me.StatusBar1.SizingGrip = False
        Me.StatusBar1.TabIndex = 8
        Me.StatusBar1.Text = "StatusBar1"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel1.Icon = CType(resources.GetObject("StatusBarPanel1.Icon"), System.Drawing.Icon)
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 31
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel2.Icon = CType(resources.GetObject("StatusBarPanel2.Icon"), System.Drawing.Icon)
        Me.StatusBarPanel2.Name = "StatusBarPanel2"
        Me.StatusBarPanel2.Width = 31
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel3.Icon = CType(resources.GetObject("StatusBarPanel3.Icon"), System.Drawing.Icon)
        Me.StatusBarPanel3.Name = "StatusBarPanel3"
        Me.StatusBarPanel3.Width = 31
        '
        'StatusBarPanel4
        '
        Me.StatusBarPanel4.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel4.Name = "StatusBarPanel4"
        Me.StatusBarPanel4.Width = 661
        '
        'grpAccounts
        '
        Me.grpAccounts.Controls.Add(Me.chkAccountsSuppressZeros)
        Me.grpAccounts.Controls.Add(Me.chkAccountsIncludeEncumbrances)
        Me.grpAccounts.Controls.Add(Me.chkAccountsUseAccountRange)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsBalanceSheet)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsHistoricalYTDSummaryOfAccounts)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsYTDSummaryOfAccounts)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsHistoricalMTDSummaryOfAccounts)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsMTDSummaryOfAccounts)
        Me.grpAccounts.Controls.Add(Me.chkAccountsIncludeSubaccounts)
        Me.grpAccounts.Controls.Add(Me.rdoAccountsChartOfAccounts)
        Me.grpAccounts.Controls.Add(Me.panelAccountsSelectAccountRange)
        Me.grpAccounts.Controls.Add(Me.panelAccountsSelectMonth)
        Me.grpAccounts.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpAccounts.Location = New System.Drawing.Point(1440, 80)
        Me.grpAccounts.Name = "grpAccounts"
        Me.grpAccounts.Size = New System.Drawing.Size(448, 208)
        Me.grpAccounts.TabIndex = 10
        Me.grpAccounts.TabStop = False
        Me.grpAccounts.Text = " Accounts"
        '
        'chkAccountsSuppressZeros
        '
        Me.chkAccountsSuppressZeros.BackColor = System.Drawing.SystemColors.Control
        Me.chkAccountsSuppressZeros.Enabled = False
        Me.chkAccountsSuppressZeros.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountsSuppressZeros.Location = New System.Drawing.Point(48, 158)
        Me.chkAccountsSuppressZeros.Name = "chkAccountsSuppressZeros"
        Me.chkAccountsSuppressZeros.Size = New System.Drawing.Size(160, 16)
        Me.chkAccountsSuppressZeros.TabIndex = 14
        Me.chkAccountsSuppressZeros.TabStop = False
        Me.chkAccountsSuppressZeros.Text = "Suppress zero amounts"
        Me.ToolTip1.SetToolTip(Me.chkAccountsSuppressZeros, " Include subaccount information for report ")
        Me.chkAccountsSuppressZeros.UseVisualStyleBackColor = False
        '
        'chkAccountsIncludeEncumbrances
        '
        Me.chkAccountsIncludeEncumbrances.BackColor = System.Drawing.SystemColors.Control
        Me.chkAccountsIncludeEncumbrances.Enabled = False
        Me.chkAccountsIncludeEncumbrances.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountsIncludeEncumbrances.Location = New System.Drawing.Point(48, 138)
        Me.chkAccountsIncludeEncumbrances.Name = "chkAccountsIncludeEncumbrances"
        Me.chkAccountsIncludeEncumbrances.Size = New System.Drawing.Size(160, 16)
        Me.chkAccountsIncludeEncumbrances.TabIndex = 13
        Me.chkAccountsIncludeEncumbrances.TabStop = False
        Me.chkAccountsIncludeEncumbrances.Text = "Include encumbrance"
        Me.ToolTip1.SetToolTip(Me.chkAccountsIncludeEncumbrances, " Include subaccount information for report ")
        Me.chkAccountsIncludeEncumbrances.UseVisualStyleBackColor = False
        '
        'chkAccountsUseAccountRange
        '
        Me.chkAccountsUseAccountRange.BackColor = System.Drawing.SystemColors.Control
        Me.chkAccountsUseAccountRange.Enabled = False
        Me.chkAccountsUseAccountRange.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountsUseAccountRange.Location = New System.Drawing.Point(48, 118)
        Me.chkAccountsUseAccountRange.Name = "chkAccountsUseAccountRange"
        Me.chkAccountsUseAccountRange.Size = New System.Drawing.Size(160, 16)
        Me.chkAccountsUseAccountRange.TabIndex = 12
        Me.chkAccountsUseAccountRange.TabStop = False
        Me.chkAccountsUseAccountRange.Text = "Use account range"
        Me.ToolTip1.SetToolTip(Me.chkAccountsUseAccountRange, " Include subaccount information for report ")
        Me.chkAccountsUseAccountRange.UseVisualStyleBackColor = False
        '
        'rdoAccountsBalanceSheet
        '
        Me.rdoAccountsBalanceSheet.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsBalanceSheet.Location = New System.Drawing.Point(232, 24)
        Me.rdoAccountsBalanceSheet.Name = "rdoAccountsBalanceSheet"
        Me.rdoAccountsBalanceSheet.Size = New System.Drawing.Size(184, 16)
        Me.rdoAccountsBalanceSheet.TabIndex = 11
        Me.rdoAccountsBalanceSheet.Text = "Statement of change"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsBalanceSheet, " Provides average daily balance by account ")
        '
        'rdoAccountsHistoricalYTDSummaryOfAccounts
        '
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Location = New System.Drawing.Point(232, 72)
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Name = "rdoAccountsHistoricalYTDSummaryOfAccounts"
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Size = New System.Drawing.Size(184, 16)
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.TabIndex = 10
        Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Text = "Historical YTD summary"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsHistoricalYTDSummaryOfAccounts, " Historical summary of account/subaccount balances ")
        '
        'rdoAccountsYTDSummaryOfAccounts
        '
        Me.rdoAccountsYTDSummaryOfAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsYTDSummaryOfAccounts.Location = New System.Drawing.Point(24, 72)
        Me.rdoAccountsYTDSummaryOfAccounts.Name = "rdoAccountsYTDSummaryOfAccounts"
        Me.rdoAccountsYTDSummaryOfAccounts.Size = New System.Drawing.Size(192, 16)
        Me.rdoAccountsYTDSummaryOfAccounts.TabIndex = 8
        Me.rdoAccountsYTDSummaryOfAccounts.Text = "Summary of accounts (YTD)"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsYTDSummaryOfAccounts, " Summary of account/subaccount balances ")
        '
        'rdoAccountsHistoricalMTDSummaryOfAccounts
        '
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Location = New System.Drawing.Point(232, 48)
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Name = "rdoAccountsHistoricalMTDSummaryOfAccounts"
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Size = New System.Drawing.Size(184, 16)
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.TabIndex = 4
        Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Text = "Historical MTD summary"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsHistoricalMTDSummaryOfAccounts, " Historical summary of account/subaccount balances ")
        '
        'rdoAccountsMTDSummaryOfAccounts
        '
        Me.rdoAccountsMTDSummaryOfAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsMTDSummaryOfAccounts.Location = New System.Drawing.Point(24, 48)
        Me.rdoAccountsMTDSummaryOfAccounts.Name = "rdoAccountsMTDSummaryOfAccounts"
        Me.rdoAccountsMTDSummaryOfAccounts.Size = New System.Drawing.Size(192, 16)
        Me.rdoAccountsMTDSummaryOfAccounts.TabIndex = 2
        Me.rdoAccountsMTDSummaryOfAccounts.Text = "Summary of accounts (MTD)"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsMTDSummaryOfAccounts, " Summary of account/subaccount balances ")
        '
        'chkAccountsIncludeSubaccounts
        '
        Me.chkAccountsIncludeSubaccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAccountsIncludeSubaccounts.Location = New System.Drawing.Point(48, 98)
        Me.chkAccountsIncludeSubaccounts.Name = "chkAccountsIncludeSubaccounts"
        Me.chkAccountsIncludeSubaccounts.Size = New System.Drawing.Size(160, 16)
        Me.chkAccountsIncludeSubaccounts.TabIndex = 1
        Me.chkAccountsIncludeSubaccounts.TabStop = False
        Me.chkAccountsIncludeSubaccounts.Text = "Include subaccounts"
        Me.ToolTip1.SetToolTip(Me.chkAccountsIncludeSubaccounts, " Include subaccount information for report ")
        '
        'rdoAccountsChartOfAccounts
        '
        Me.rdoAccountsChartOfAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAccountsChartOfAccounts.Location = New System.Drawing.Point(24, 24)
        Me.rdoAccountsChartOfAccounts.Name = "rdoAccountsChartOfAccounts"
        Me.rdoAccountsChartOfAccounts.Size = New System.Drawing.Size(192, 16)
        Me.rdoAccountsChartOfAccounts.TabIndex = 0
        Me.rdoAccountsChartOfAccounts.Text = "Chart of accounts"
        Me.ToolTip1.SetToolTip(Me.rdoAccountsChartOfAccounts, " Listing of accounts ")
        '
        'panelAccountsSelectAccountRange
        '
        Me.panelAccountsSelectAccountRange.Controls.Add(Me.Label13)
        Me.panelAccountsSelectAccountRange.Controls.Add(Me.Label12)
        Me.panelAccountsSelectAccountRange.Controls.Add(Me.txtAccountsAcctNumberTo)
        Me.panelAccountsSelectAccountRange.Controls.Add(Me.txtAccountsAcctNumberFrom)
        Me.panelAccountsSelectAccountRange.Location = New System.Drawing.Point(240, 144)
        Me.panelAccountsSelectAccountRange.Name = "panelAccountsSelectAccountRange"
        Me.panelAccountsSelectAccountRange.Size = New System.Drawing.Size(168, 56)
        Me.panelAccountsSelectAccountRange.TabIndex = 9
        Me.panelAccountsSelectAccountRange.Visible = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(73, 28)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(19, 11)
        Me.Label13.TabIndex = 8
        Me.Label13.Text = "To"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(14, 8)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(128, 16)
        Me.Label12.TabIndex = 7
        Me.Label12.Text = "Account range:"
        '
        'txtAccountsAcctNumberTo
        '
        Me.txtAccountsAcctNumberTo.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccountsAcctNumberTo.Location = New System.Drawing.Point(96, 24)
        Me.txtAccountsAcctNumberTo.Name = "txtAccountsAcctNumberTo"
        Me.txtAccountsAcctNumberTo.Size = New System.Drawing.Size(46, 20)
        Me.txtAccountsAcctNumberTo.TabIndex = 1
        Me.txtAccountsAcctNumberTo.Text = "0001"
        Me.txtAccountsAcctNumberTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtAccountsAcctNumberFrom
        '
        Me.txtAccountsAcctNumberFrom.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAccountsAcctNumberFrom.Location = New System.Drawing.Point(24, 24)
        Me.txtAccountsAcctNumberFrom.Name = "txtAccountsAcctNumberFrom"
        Me.txtAccountsAcctNumberFrom.Size = New System.Drawing.Size(46, 20)
        Me.txtAccountsAcctNumberFrom.TabIndex = 0
        Me.txtAccountsAcctNumberFrom.Text = "0001"
        Me.txtAccountsAcctNumberFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'panelAccountsSelectMonth
        '
        Me.panelAccountsSelectMonth.Controls.Add(Me.Label6)
        Me.panelAccountsSelectMonth.Controls.Add(Me.cboAccountsMonthListing)
        Me.panelAccountsSelectMonth.Location = New System.Drawing.Point(240, 120)
        Me.panelAccountsSelectMonth.Name = "panelAccountsSelectMonth"
        Me.panelAccountsSelectMonth.Size = New System.Drawing.Size(168, 56)
        Me.panelAccountsSelectMonth.TabIndex = 7
        Me.panelAccountsSelectMonth.Visible = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(7, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(144, 16)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "Select reporting month:"
        '
        'cboAccountsMonthListing
        '
        Me.cboAccountsMonthListing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccountsMonthListing.Location = New System.Drawing.Point(24, 24)
        Me.cboAccountsMonthListing.Name = "cboAccountsMonthListing"
        Me.cboAccountsMonthListing.Size = New System.Drawing.Size(96, 22)
        Me.cboAccountsMonthListing.TabIndex = 5
        '
        'grpAdjustments
        '
        Me.grpAdjustments.Controls.Add(Me.rdoAdjustmentsPrintTransferRegister)
        Me.grpAdjustments.Controls.Add(Me.rdoAdjustmentsPrintAdjustmentRegister)
        Me.grpAdjustments.Location = New System.Drawing.Point(1440, 80)
        Me.grpAdjustments.Name = "grpAdjustments"
        Me.grpAdjustments.Size = New System.Drawing.Size(448, 208)
        Me.grpAdjustments.TabIndex = 11
        Me.grpAdjustments.TabStop = False
        Me.grpAdjustments.Text = " Adjustments"
        '
        'rdoAdjustmentsPrintTransferRegister
        '
        Me.rdoAdjustmentsPrintTransferRegister.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAdjustmentsPrintTransferRegister.Location = New System.Drawing.Point(24, 48)
        Me.rdoAdjustmentsPrintTransferRegister.Name = "rdoAdjustmentsPrintTransferRegister"
        Me.rdoAdjustmentsPrintTransferRegister.Size = New System.Drawing.Size(176, 16)
        Me.rdoAdjustmentsPrintTransferRegister.TabIndex = 8
        Me.rdoAdjustmentsPrintTransferRegister.Text = "Print transfer register"
        '
        'rdoAdjustmentsPrintAdjustmentRegister
        '
        Me.rdoAdjustmentsPrintAdjustmentRegister.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoAdjustmentsPrintAdjustmentRegister.Location = New System.Drawing.Point(24, 24)
        Me.rdoAdjustmentsPrintAdjustmentRegister.Name = "rdoAdjustmentsPrintAdjustmentRegister"
        Me.rdoAdjustmentsPrintAdjustmentRegister.Size = New System.Drawing.Size(176, 16)
        Me.rdoAdjustmentsPrintAdjustmentRegister.TabIndex = 7
        Me.rdoAdjustmentsPrintAdjustmentRegister.Text = "Print adjustment register"
        '
        'grpExpenditures
        '
        Me.grpExpenditures.Controls.Add(Me.lblExpendituresPosPay)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPositivePay)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPendingInvoice)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresEncumbranceAccounts)
        Me.grpExpenditures.Controls.Add(Me.chkExpendituresOutstandingInvoices)
        Me.grpExpenditures.Controls.Add(Me.chkExpendituresAllFiscalYears)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPrintOutstandingChecks)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPrintVoidChecks)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPoRegister)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresPrintPurchaseOrder)
        Me.grpExpenditures.Controls.Add(Me.rdoExpendituresCheckRegister)
        Me.grpExpenditures.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpExpenditures.Location = New System.Drawing.Point(144, 80)
        Me.grpExpenditures.Name = "grpExpenditures"
        Me.grpExpenditures.Size = New System.Drawing.Size(448, 208)
        Me.grpExpenditures.TabIndex = 12
        Me.grpExpenditures.TabStop = False
        Me.grpExpenditures.Text = " Expenditures"
        '
        'lblExpendituresPosPay
        '
        Me.lblExpendituresPosPay.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpendituresPosPay.Location = New System.Drawing.Point(216, 96)
        Me.lblExpendituresPosPay.Name = "lblExpendituresPosPay"
        Me.lblExpendituresPosPay.Size = New System.Drawing.Size(216, 48)
        Me.lblExpendituresPosPay.TabIndex = 10
        Me.lblExpendituresPosPay.Text = "Enter the beginning and ending check number for the selected fiscal year."
        Me.lblExpendituresPosPay.Visible = False
        '
        'rdoExpendituresPositivePay
        '
        Me.rdoExpendituresPositivePay.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPositivePay.Location = New System.Drawing.Point(200, 72)
        Me.rdoExpendituresPositivePay.Name = "rdoExpendituresPositivePay"
        Me.rdoExpendituresPositivePay.Size = New System.Drawing.Size(224, 16)
        Me.rdoExpendituresPositivePay.TabIndex = 9
        Me.rdoExpendituresPositivePay.Text = "Positive pay file (Bank Standard)"
        '
        'rdoExpendituresPendingInvoice
        '
        Me.rdoExpendituresPendingInvoice.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPendingInvoice.Location = New System.Drawing.Point(200, 48)
        Me.rdoExpendituresPendingInvoice.Name = "rdoExpendituresPendingInvoice"
        Me.rdoExpendituresPendingInvoice.Size = New System.Drawing.Size(192, 16)
        Me.rdoExpendituresPendingInvoice.TabIndex = 8
        Me.rdoExpendituresPendingInvoice.Text = "Outstanding invoice report"
        '
        'rdoExpendituresEncumbranceAccounts
        '
        Me.rdoExpendituresEncumbranceAccounts.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresEncumbranceAccounts.Location = New System.Drawing.Point(200, 24)
        Me.rdoExpendituresEncumbranceAccounts.Name = "rdoExpendituresEncumbranceAccounts"
        Me.rdoExpendituresEncumbranceAccounts.Size = New System.Drawing.Size(192, 16)
        Me.rdoExpendituresEncumbranceAccounts.TabIndex = 7
        Me.rdoExpendituresEncumbranceAccounts.Text = "Encumbrance account"
        '
        'chkExpendituresOutstandingInvoices
        '
        Me.chkExpendituresOutstandingInvoices.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExpendituresOutstandingInvoices.Location = New System.Drawing.Point(216, 104)
        Me.chkExpendituresOutstandingInvoices.Name = "chkExpendituresOutstandingInvoices"
        Me.chkExpendituresOutstandingInvoices.Size = New System.Drawing.Size(200, 16)
        Me.chkExpendituresOutstandingInvoices.TabIndex = 6
        Me.chkExpendituresOutstandingInvoices.Text = "Outstanding encumbrances"
        Me.ToolTip1.SetToolTip(Me.chkExpendituresOutstandingInvoices, " Show open purchase orders only ")
        Me.chkExpendituresOutstandingInvoices.Visible = False
        '
        'chkExpendituresAllFiscalYears
        '
        Me.chkExpendituresAllFiscalYears.Checked = True
        Me.chkExpendituresAllFiscalYears.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkExpendituresAllFiscalYears.Font = New System.Drawing.Font("Lucida Console", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExpendituresAllFiscalYears.Location = New System.Drawing.Point(216, 128)
        Me.chkExpendituresAllFiscalYears.Name = "chkExpendituresAllFiscalYears"
        Me.chkExpendituresAllFiscalYears.Size = New System.Drawing.Size(136, 16)
        Me.chkExpendituresAllFiscalYears.TabIndex = 5
        Me.chkExpendituresAllFiscalYears.Text = "All fiscal years"
        Me.chkExpendituresAllFiscalYears.Visible = False
        '
        'rdoExpendituresPrintOutstandingChecks
        '
        Me.rdoExpendituresPrintOutstandingChecks.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPrintOutstandingChecks.Location = New System.Drawing.Point(24, 96)
        Me.rdoExpendituresPrintOutstandingChecks.Name = "rdoExpendituresPrintOutstandingChecks"
        Me.rdoExpendituresPrintOutstandingChecks.Size = New System.Drawing.Size(168, 16)
        Me.rdoExpendituresPrintOutstandingChecks.TabIndex = 4
        Me.rdoExpendituresPrintOutstandingChecks.Text = "Outstanding checks"
        '
        'rdoExpendituresPrintVoidChecks
        '
        Me.rdoExpendituresPrintVoidChecks.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPrintVoidChecks.Location = New System.Drawing.Point(24, 120)
        Me.rdoExpendituresPrintVoidChecks.Name = "rdoExpendituresPrintVoidChecks"
        Me.rdoExpendituresPrintVoidChecks.Size = New System.Drawing.Size(168, 16)
        Me.rdoExpendituresPrintVoidChecks.TabIndex = 3
        Me.rdoExpendituresPrintVoidChecks.Text = "Void checks"
        '
        'rdoExpendituresPoRegister
        '
        Me.rdoExpendituresPoRegister.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPoRegister.Location = New System.Drawing.Point(24, 48)
        Me.rdoExpendituresPoRegister.Name = "rdoExpendituresPoRegister"
        Me.rdoExpendituresPoRegister.Size = New System.Drawing.Size(168, 16)
        Me.rdoExpendituresPoRegister.TabIndex = 2
        Me.rdoExpendituresPoRegister.Text = "Purchase order register"
        '
        'rdoExpendituresPrintPurchaseOrder
        '
        Me.rdoExpendituresPrintPurchaseOrder.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresPrintPurchaseOrder.Location = New System.Drawing.Point(24, 72)
        Me.rdoExpendituresPrintPurchaseOrder.Name = "rdoExpendituresPrintPurchaseOrder"
        Me.rdoExpendituresPrintPurchaseOrder.Size = New System.Drawing.Size(168, 16)
        Me.rdoExpendituresPrintPurchaseOrder.TabIndex = 1
        Me.rdoExpendituresPrintPurchaseOrder.Text = "Purchase order"
        '
        'rdoExpendituresCheckRegister
        '
        Me.rdoExpendituresCheckRegister.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoExpendituresCheckRegister.Location = New System.Drawing.Point(24, 24)
        Me.rdoExpendituresCheckRegister.Name = "rdoExpendituresCheckRegister"
        Me.rdoExpendituresCheckRegister.Size = New System.Drawing.Size(168, 16)
        Me.rdoExpendituresCheckRegister.TabIndex = 0
        Me.rdoExpendituresCheckRegister.Text = "Check register"
        '
        'grpRevenue
        '
        Me.grpRevenue.Controls.Add(Me.panelRevenue1098T)
        Me.grpRevenue.Controls.Add(Me.panelRevenueAllFiscalYears)
        Me.grpRevenue.Controls.Add(Me.panelRevenueSearchReceived)
        Me.grpRevenue.Controls.Add(Me.panelRevenueDailyDeposit)
        Me.grpRevenue.Controls.Add(Me.rdoRevenuePrintOutstandingReceipts)
        Me.grpRevenue.Controls.Add(Me.rdoRevenuePrintVoidReceipts)
        Me.grpRevenue.Controls.Add(Me.rdoRevenueDepositSummary)
        Me.grpRevenue.Controls.Add(Me.rdoRevenueDailyDeposit)
        Me.grpRevenue.Controls.Add(Me.rdoRevenueReceiptTicket)
        Me.grpRevenue.Controls.Add(Me.rdoRevenueReceiptRegister)
        Me.grpRevenue.Controls.Add(Me.rdoRevenue1098T)
        Me.grpRevenue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grpRevenue.Location = New System.Drawing.Point(1440, 80)
        Me.grpRevenue.Name = "grpRevenue"
        Me.grpRevenue.Size = New System.Drawing.Size(448, 208)
        Me.grpRevenue.TabIndex = 13
        Me.grpRevenue.TabStop = False
        Me.grpRevenue.Text = " Revenues"
        '
        'panelRevenue1098T
        '
        Me.panelRevenue1098T.BackColor = System.Drawing.Color.White
        Me.panelRevenue1098T.Controls.Add(Me.txtRevenueSubaccountNumber)
        Me.panelRevenue1098T.Controls.Add(Me.txtRevenueAccountNumber)
        Me.panelRevenue1098T.Controls.Add(Me.Label14)
        Me.panelRevenue1098T.Controls.Add(Me.Label15)
        Me.panelRevenue1098T.Location = New System.Drawing.Point(198, 72)
        Me.panelRevenue1098T.Name = "panelRevenue1098T"
        Me.panelRevenue1098T.Size = New System.Drawing.Size(240, 60)
        Me.panelRevenue1098T.TabIndex = 238
        Me.panelRevenue1098T.Visible = False
        '
        'txtRevenueSubaccountNumber
        '
        Me.txtRevenueSubaccountNumber.Location = New System.Drawing.Point(184, 32)
        Me.txtRevenueSubaccountNumber.Name = "txtRevenueSubaccountNumber"
        Me.txtRevenueSubaccountNumber.Size = New System.Drawing.Size(32, 20)
        Me.txtRevenueSubaccountNumber.TabIndex = 1
        Me.txtRevenueSubaccountNumber.Text = "001"
        Me.txtRevenueSubaccountNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtRevenueAccountNumber
        '
        Me.txtRevenueAccountNumber.Location = New System.Drawing.Point(136, 32)
        Me.txtRevenueAccountNumber.Name = "txtRevenueAccountNumber"
        Me.txtRevenueAccountNumber.Size = New System.Drawing.Size(40, 20)
        Me.txtRevenueAccountNumber.TabIndex = 0
        Me.txtRevenueAccountNumber.Text = "0001"
        Me.txtRevenueAccountNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 34)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(128, 16)
        Me.Label14.TabIndex = 7
        Me.Label14.Text = "Enter account number:"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Verdana", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Firebrick
        Me.Label15.Location = New System.Drawing.Point(8, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(216, 16)
        Me.Label15.TabIndex = 8
        Me.Label15.Text = "Choose calendar year at top left"
        '
        'panelRevenueAllFiscalYears
        '
        Me.panelRevenueAllFiscalYears.BackColor = System.Drawing.Color.White
        Me.panelRevenueAllFiscalYears.Controls.Add(Me.chkRevenuesAllFiscalYears)
        Me.panelRevenueAllFiscalYears.Location = New System.Drawing.Point(1980, 98)
        Me.panelRevenueAllFiscalYears.Name = "panelRevenueAllFiscalYears"
        Me.panelRevenueAllFiscalYears.Size = New System.Drawing.Size(240, 60)
        Me.panelRevenueAllFiscalYears.TabIndex = 237
        Me.panelRevenueAllFiscalYears.Visible = False
        '
        'chkRevenuesAllFiscalYears
        '
        Me.chkRevenuesAllFiscalYears.Checked = True
        Me.chkRevenuesAllFiscalYears.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRevenuesAllFiscalYears.Location = New System.Drawing.Point(16, 8)
        Me.chkRevenuesAllFiscalYears.Name = "chkRevenuesAllFiscalYears"
        Me.chkRevenuesAllFiscalYears.Size = New System.Drawing.Size(112, 16)
        Me.chkRevenuesAllFiscalYears.TabIndex = 232
        Me.chkRevenuesAllFiscalYears.Text = "All fiscal years"
        '
        'panelRevenueSearchReceived
        '
        Me.panelRevenueSearchReceived.BackColor = System.Drawing.Color.White
        Me.panelRevenueSearchReceived.Controls.Add(Me.txtRevenueSearch)
        Me.panelRevenueSearchReceived.Controls.Add(Me.chkRevenuesSearchReceived)
        Me.panelRevenueSearchReceived.Location = New System.Drawing.Point(1980, 98)
        Me.panelRevenueSearchReceived.Name = "panelRevenueSearchReceived"
        Me.panelRevenueSearchReceived.Size = New System.Drawing.Size(240, 60)
        Me.panelRevenueSearchReceived.TabIndex = 235
        Me.panelRevenueSearchReceived.Visible = False
        '
        'txtRevenueSearch
        '
        Me.txtRevenueSearch.Location = New System.Drawing.Point(8, 32)
        Me.txtRevenueSearch.Name = "txtRevenueSearch"
        Me.txtRevenueSearch.Size = New System.Drawing.Size(224, 20)
        Me.txtRevenueSearch.TabIndex = 7
        Me.txtRevenueSearch.Visible = False
        '
        'chkRevenuesSearchReceived
        '
        Me.chkRevenuesSearchReceived.Location = New System.Drawing.Point(8, 8)
        Me.chkRevenuesSearchReceived.Name = "chkRevenuesSearchReceived"
        Me.chkRevenuesSearchReceived.Size = New System.Drawing.Size(160, 16)
        Me.chkRevenuesSearchReceived.TabIndex = 234
        Me.chkRevenuesSearchReceived.Text = "Search received field"
        '
        'panelRevenueDailyDeposit
        '
        Me.panelRevenueDailyDeposit.BackColor = System.Drawing.Color.White
        Me.panelRevenueDailyDeposit.Controls.Add(Me.txtDepositEndNumber)
        Me.panelRevenueDailyDeposit.Controls.Add(Me.chkRevenuesIncludeCreditCards)
        Me.panelRevenueDailyDeposit.Controls.Add(Me.chkRevenuesPrintDepositTicket)
        Me.panelRevenueDailyDeposit.Controls.Add(Me.txtDepositBegNumber)
        Me.panelRevenueDailyDeposit.Location = New System.Drawing.Point(1980, 98)
        Me.panelRevenueDailyDeposit.Name = "panelRevenueDailyDeposit"
        Me.panelRevenueDailyDeposit.Size = New System.Drawing.Size(240, 60)
        Me.panelRevenueDailyDeposit.TabIndex = 8
        Me.panelRevenueDailyDeposit.Visible = False
        '
        'txtDepositEndNumber
        '
        Me.txtDepositEndNumber.Location = New System.Drawing.Point(8, 32)
        Me.txtDepositEndNumber.Name = "txtDepositEndNumber"
        Me.txtDepositEndNumber.Size = New System.Drawing.Size(48, 20)
        Me.txtDepositEndNumber.TabIndex = 7
        Me.txtDepositEndNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDepositEndNumber.Visible = False
        '
        'chkRevenuesIncludeCreditCards
        '
        Me.chkRevenuesIncludeCreditCards.Checked = True
        Me.chkRevenuesIncludeCreditCards.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRevenuesIncludeCreditCards.Location = New System.Drawing.Point(64, 8)
        Me.chkRevenuesIncludeCreditCards.Name = "chkRevenuesIncludeCreditCards"
        Me.chkRevenuesIncludeCreditCards.Size = New System.Drawing.Size(168, 16)
        Me.chkRevenuesIncludeCreditCards.TabIndex = 6
        Me.chkRevenuesIncludeCreditCards.Text = "Include credit card receipts"
        Me.chkRevenuesIncludeCreditCards.Visible = False
        '
        'chkRevenuesPrintDepositTicket
        '
        Me.chkRevenuesPrintDepositTicket.Location = New System.Drawing.Point(64, 32)
        Me.chkRevenuesPrintDepositTicket.Name = "chkRevenuesPrintDepositTicket"
        Me.chkRevenuesPrintDepositTicket.Size = New System.Drawing.Size(128, 16)
        Me.chkRevenuesPrintDepositTicket.TabIndex = 5
        Me.chkRevenuesPrintDepositTicket.Text = "Print deposit ticket"
        Me.chkRevenuesPrintDepositTicket.Visible = False
        '
        'txtDepositBegNumber
        '
        Me.txtDepositBegNumber.Location = New System.Drawing.Point(8, 8)
        Me.txtDepositBegNumber.Name = "txtDepositBegNumber"
        Me.txtDepositBegNumber.Size = New System.Drawing.Size(48, 20)
        Me.txtDepositBegNumber.TabIndex = 4
        Me.txtDepositBegNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtDepositBegNumber.Visible = False
        '
        'rdoRevenuePrintOutstandingReceipts
        '
        Me.rdoRevenuePrintOutstandingReceipts.Location = New System.Drawing.Point(24, 120)
        Me.rdoRevenuePrintOutstandingReceipts.Name = "rdoRevenuePrintOutstandingReceipts"
        Me.rdoRevenuePrintOutstandingReceipts.Size = New System.Drawing.Size(160, 16)
        Me.rdoRevenuePrintOutstandingReceipts.TabIndex = 233
        Me.rdoRevenuePrintOutstandingReceipts.Text = "Print outstanding receipts"
        '
        'rdoRevenuePrintVoidReceipts
        '
        Me.rdoRevenuePrintVoidReceipts.Location = New System.Drawing.Point(200, 24)
        Me.rdoRevenuePrintVoidReceipts.Name = "rdoRevenuePrintVoidReceipts"
        Me.rdoRevenuePrintVoidReceipts.Size = New System.Drawing.Size(136, 16)
        Me.rdoRevenuePrintVoidReceipts.TabIndex = 9
        Me.rdoRevenuePrintVoidReceipts.Text = "Print void receipts"
        '
        'rdoRevenueDepositSummary
        '
        Me.rdoRevenueDepositSummary.Location = New System.Drawing.Point(24, 96)
        Me.rdoRevenueDepositSummary.Name = "rdoRevenueDepositSummary"
        Me.rdoRevenueDepositSummary.Size = New System.Drawing.Size(136, 16)
        Me.rdoRevenueDepositSummary.TabIndex = 3
        Me.rdoRevenueDepositSummary.Text = "Print deposit summary"
        '
        'rdoRevenueDailyDeposit
        '
        Me.rdoRevenueDailyDeposit.Location = New System.Drawing.Point(24, 72)
        Me.rdoRevenueDailyDeposit.Name = "rdoRevenueDailyDeposit"
        Me.rdoRevenueDailyDeposit.Size = New System.Drawing.Size(112, 16)
        Me.rdoRevenueDailyDeposit.TabIndex = 2
        Me.rdoRevenueDailyDeposit.Text = "Print daily deposit"
        '
        'rdoRevenueReceiptTicket
        '
        Me.rdoRevenueReceiptTicket.Location = New System.Drawing.Point(24, 48)
        Me.rdoRevenueReceiptTicket.Name = "rdoRevenueReceiptTicket"
        Me.rdoRevenueReceiptTicket.Size = New System.Drawing.Size(112, 16)
        Me.rdoRevenueReceiptTicket.TabIndex = 1
        Me.rdoRevenueReceiptTicket.Text = "Print receipt ticket"
        '
        'rdoRevenueReceiptRegister
        '
        Me.rdoRevenueReceiptRegister.Location = New System.Drawing.Point(24, 24)
        Me.rdoRevenueReceiptRegister.Name = "rdoRevenueReceiptRegister"
        Me.rdoRevenueReceiptRegister.Size = New System.Drawing.Size(128, 16)
        Me.rdoRevenueReceiptRegister.TabIndex = 0
        Me.rdoRevenueReceiptRegister.Text = "Print receipt register"
        '
        'rdoRevenue1098T
        '
        Me.rdoRevenue1098T.Location = New System.Drawing.Point(200, 45)
        Me.rdoRevenue1098T.Name = "rdoRevenue1098T"
        Me.rdoRevenue1098T.Size = New System.Drawing.Size(136, 16)
        Me.rdoRevenue1098T.TabIndex = 236
        Me.rdoRevenue1098T.Text = "Form 1098-T"
        '
        'grpFinancials
        '
        Me.grpFinancials.Controls.Add(Me.chkFinancialsEncumbranceDetail)
        Me.grpFinancials.Controls.Add(Me.rdoFinancialsNoYtdDetail)
        Me.grpFinancials.Controls.Add(Me.rdoFinancialsNoMtdDetail)
        Me.grpFinancials.Controls.Add(Me.panelFinancialsSelectAccount)
        Me.grpFinancials.Controls.Add(Me.rdoFinancialsDetailOfAccountPeriodical)
        Me.grpFinancials.Controls.Add(Me.rdoFinancialsDetailOfAccountYTD)
        Me.grpFinancials.Controls.Add(Me.rdoFinancialsDetailOfAccountMTD)
        Me.grpFinancials.Location = New System.Drawing.Point(1440, 80)
        Me.grpFinancials.Name = "grpFinancials"
        Me.grpFinancials.Size = New System.Drawing.Size(448, 208)
        Me.grpFinancials.TabIndex = 14
        Me.grpFinancials.TabStop = False
        Me.grpFinancials.Text = " Financials"
        '
        'chkFinancialsEncumbranceDetail
        '
        Me.chkFinancialsEncumbranceDetail.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFinancialsEncumbranceDetail.Location = New System.Drawing.Point(240, 24)
        Me.chkFinancialsEncumbranceDetail.Name = "chkFinancialsEncumbranceDetail"
        Me.chkFinancialsEncumbranceDetail.Size = New System.Drawing.Size(144, 16)
        Me.chkFinancialsEncumbranceDetail.TabIndex = 8
        Me.chkFinancialsEncumbranceDetail.Text = "Include encumbrances"
        Me.ToolTip1.SetToolTip(Me.chkFinancialsEncumbranceDetail, " Includes the encumbrance detail listing ")
        '
        'rdoFinancialsNoYtdDetail
        '
        Me.rdoFinancialsNoYtdDetail.Location = New System.Drawing.Point(24, 120)
        Me.rdoFinancialsNoYtdDetail.Name = "rdoFinancialsNoYtdDetail"
        Me.rdoFinancialsNoYtdDetail.Size = New System.Drawing.Size(168, 16)
        Me.rdoFinancialsNoYtdDetail.TabIndex = 6
        Me.rdoFinancialsNoYtdDetail.Text = "No YTD account activity"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsNoYtdDetail, " List accounts with no year-to-date activity ")
        '
        'rdoFinancialsNoMtdDetail
        '
        Me.rdoFinancialsNoMtdDetail.Location = New System.Drawing.Point(24, 96)
        Me.rdoFinancialsNoMtdDetail.Name = "rdoFinancialsNoMtdDetail"
        Me.rdoFinancialsNoMtdDetail.Size = New System.Drawing.Size(168, 16)
        Me.rdoFinancialsNoMtdDetail.TabIndex = 5
        Me.rdoFinancialsNoMtdDetail.Text = "No MTD account activity"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsNoMtdDetail, " List accounts with no month-to-date activity ")
        '
        'panelFinancialsSelectAccount
        '
        Me.panelFinancialsSelectAccount.Controls.Add(Me.lblAccountName)
        Me.panelFinancialsSelectAccount.Controls.Add(Me.lblAccountNumber)
        Me.panelFinancialsSelectAccount.Controls.Add(Me.rdoFinancialsAllAccounts)
        Me.panelFinancialsSelectAccount.Controls.Add(Me.rdoFinancialsSelectAccount)
        Me.panelFinancialsSelectAccount.Location = New System.Drawing.Point(240, 59)
        Me.panelFinancialsSelectAccount.Name = "panelFinancialsSelectAccount"
        Me.panelFinancialsSelectAccount.Size = New System.Drawing.Size(168, 97)
        Me.panelFinancialsSelectAccount.TabIndex = 4
        '
        'lblAccountName
        '
        Me.lblAccountName.Location = New System.Drawing.Point(24, 76)
        Me.lblAccountName.Name = "lblAccountName"
        Me.lblAccountName.Size = New System.Drawing.Size(136, 16)
        Me.lblAccountName.TabIndex = 6
        '
        'lblAccountNumber
        '
        Me.lblAccountNumber.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAccountNumber.Location = New System.Drawing.Point(24, 57)
        Me.lblAccountNumber.Name = "lblAccountNumber"
        Me.lblAccountNumber.Size = New System.Drawing.Size(64, 16)
        Me.lblAccountNumber.TabIndex = 5
        '
        'rdoFinancialsAllAccounts
        '
        Me.rdoFinancialsAllAccounts.Location = New System.Drawing.Point(16, 32)
        Me.rdoFinancialsAllAccounts.Name = "rdoFinancialsAllAccounts"
        Me.rdoFinancialsAllAccounts.Size = New System.Drawing.Size(136, 16)
        Me.rdoFinancialsAllAccounts.TabIndex = 4
        Me.rdoFinancialsAllAccounts.Text = "All accounts"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsAllAccounts, " Click here to select all accounts ")
        '
        'rdoFinancialsSelectAccount
        '
        Me.rdoFinancialsSelectAccount.Location = New System.Drawing.Point(16, 8)
        Me.rdoFinancialsSelectAccount.Name = "rdoFinancialsSelectAccount"
        Me.rdoFinancialsSelectAccount.Size = New System.Drawing.Size(136, 16)
        Me.rdoFinancialsSelectAccount.TabIndex = 3
        Me.rdoFinancialsSelectAccount.Text = "Select account"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsSelectAccount, " Click here to select an account ")
        '
        'rdoFinancialsDetailOfAccountPeriodical
        '
        Me.rdoFinancialsDetailOfAccountPeriodical.Location = New System.Drawing.Point(24, 72)
        Me.rdoFinancialsDetailOfAccountPeriodical.Name = "rdoFinancialsDetailOfAccountPeriodical"
        Me.rdoFinancialsDetailOfAccountPeriodical.Size = New System.Drawing.Size(168, 16)
        Me.rdoFinancialsDetailOfAccountPeriodical.TabIndex = 2
        Me.rdoFinancialsDetailOfAccountPeriodical.Text = "Detail of account (Periodical)"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsDetailOfAccountPeriodical, " Print detail of account report for selected period ")
        '
        'rdoFinancialsDetailOfAccountYTD
        '
        Me.rdoFinancialsDetailOfAccountYTD.Location = New System.Drawing.Point(24, 48)
        Me.rdoFinancialsDetailOfAccountYTD.Name = "rdoFinancialsDetailOfAccountYTD"
        Me.rdoFinancialsDetailOfAccountYTD.Size = New System.Drawing.Size(168, 16)
        Me.rdoFinancialsDetailOfAccountYTD.TabIndex = 1
        Me.rdoFinancialsDetailOfAccountYTD.Text = "Detail of account (YTD)"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsDetailOfAccountYTD, " Print year-to-date detail of account report ")
        '
        'rdoFinancialsDetailOfAccountMTD
        '
        Me.rdoFinancialsDetailOfAccountMTD.Location = New System.Drawing.Point(24, 24)
        Me.rdoFinancialsDetailOfAccountMTD.Name = "rdoFinancialsDetailOfAccountMTD"
        Me.rdoFinancialsDetailOfAccountMTD.Size = New System.Drawing.Size(168, 16)
        Me.rdoFinancialsDetailOfAccountMTD.TabIndex = 0
        Me.rdoFinancialsDetailOfAccountMTD.Text = "Detail of account (MTD)"
        Me.ToolTip1.SetToolTip(Me.rdoFinancialsDetailOfAccountMTD, " Print month-to-date detail of account report ")
        '
        'grpVendors
        '
        Me.grpVendors.Controls.Add(Me.rdoVendorAudit)
        Me.grpVendors.Controls.Add(Me.chkVendorsUseFiscalYear)
        Me.grpVendors.Controls.Add(Me.chkVendorsIncludeSSN)
        Me.grpVendors.Controls.Add(Me.chkVendorsUse600Minimum)
        Me.grpVendors.Controls.Add(Me.chkVendorsSelectSingleVendor)
        Me.grpVendors.Controls.Add(Me.cboVendors)
        Me.grpVendors.Controls.Add(Me.cboCalendarYears)
        Me.grpVendors.Controls.Add(Me.rdoVendor1099Listing)
        Me.grpVendors.Controls.Add(Me.chkVendorsIncludeZeroBalances)
        Me.grpVendors.Controls.Add(Me.rdoVendorExpenses)
        Me.grpVendors.Controls.Add(Me.chk1099ByEmployee)
        Me.grpVendors.Controls.Add(Me.rdoVendorListing)
        Me.grpVendors.Controls.Add(Me.lblVendorCalendar)
        Me.grpVendors.Controls.Add(Me.chk1099Summary)
        Me.grpVendors.Location = New System.Drawing.Point(1440, 80)
        Me.grpVendors.Name = "grpVendors"
        Me.grpVendors.Size = New System.Drawing.Size(448, 208)
        Me.grpVendors.TabIndex = 15
        Me.grpVendors.TabStop = False
        Me.grpVendors.Text = " Vendors"
        '
        'rdoVendorAudit
        '
        Me.rdoVendorAudit.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoVendorAudit.Location = New System.Drawing.Point(24, 96)
        Me.rdoVendorAudit.Name = "rdoVendorAudit"
        Me.rdoVendorAudit.Size = New System.Drawing.Size(200, 16)
        Me.rdoVendorAudit.TabIndex = 22
        Me.rdoVendorAudit.Text = "Vendor audit report"
        Me.ToolTip1.SetToolTip(Me.rdoVendorAudit, " Show detail listing of expenditures by vendor ")
        '
        'chkVendorsUseFiscalYear
        '
        Me.chkVendorsUseFiscalYear.Enabled = False
        Me.chkVendorsUseFiscalYear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorsUseFiscalYear.Location = New System.Drawing.Point(272, 160)
        Me.chkVendorsUseFiscalYear.Name = "chkVendorsUseFiscalYear"
        Me.chkVendorsUseFiscalYear.Size = New System.Drawing.Size(144, 16)
        Me.chkVendorsUseFiscalYear.TabIndex = 20
        Me.chkVendorsUseFiscalYear.Text = "Use fiscal year"
        Me.ToolTip1.SetToolTip(Me.chkVendorsUseFiscalYear, " Use fiscal year date range ")
        '
        'chkVendorsIncludeSSN
        '
        Me.chkVendorsIncludeSSN.Enabled = False
        Me.chkVendorsIncludeSSN.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorsIncludeSSN.Location = New System.Drawing.Point(272, 144)
        Me.chkVendorsIncludeSSN.Name = "chkVendorsIncludeSSN"
        Me.chkVendorsIncludeSSN.Size = New System.Drawing.Size(144, 16)
        Me.chkVendorsIncludeSSN.TabIndex = 19
        Me.chkVendorsIncludeSSN.Text = "Include SSN"
        Me.ToolTip1.SetToolTip(Me.chkVendorsIncludeSSN, " 1099 must be enabled for SSN ")
        '
        'chkVendorsUse600Minimum
        '
        Me.chkVendorsUse600Minimum.Enabled = False
        Me.chkVendorsUse600Minimum.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorsUse600Minimum.Location = New System.Drawing.Point(272, 112)
        Me.chkVendorsUse600Minimum.Name = "chkVendorsUse600Minimum"
        Me.chkVendorsUse600Minimum.Size = New System.Drawing.Size(144, 16)
        Me.chkVendorsUse600Minimum.TabIndex = 18
        Me.chkVendorsUse600Minimum.Text = "Use $600 minimum"
        Me.ToolTip1.SetToolTip(Me.chkVendorsUse600Minimum, " Show vendors with a minimum $600 and above ")
        '
        'chkVendorsSelectSingleVendor
        '
        Me.chkVendorsSelectSingleVendor.Enabled = False
        Me.chkVendorsSelectSingleVendor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorsSelectSingleVendor.Location = New System.Drawing.Point(272, 128)
        Me.chkVendorsSelectSingleVendor.Name = "chkVendorsSelectSingleVendor"
        Me.chkVendorsSelectSingleVendor.Size = New System.Drawing.Size(144, 16)
        Me.chkVendorsSelectSingleVendor.TabIndex = 17
        Me.chkVendorsSelectSingleVendor.Text = "Select vendor"
        Me.ToolTip1.SetToolTip(Me.chkVendorsSelectSingleVendor, " Select a single vendor ")
        '
        'cboVendors
        '
        Me.cboVendors.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVendors.Location = New System.Drawing.Point(16, 128)
        Me.cboVendors.Name = "cboVendors"
        Me.cboVendors.Size = New System.Drawing.Size(240, 22)
        Me.cboVendors.TabIndex = 16
        Me.cboVendors.Visible = False
        '
        'cboCalendarYears
        '
        Me.cboCalendarYears.Enabled = False
        Me.cboCalendarYears.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCalendarYears.Location = New System.Drawing.Point(352, 28)
        Me.cboCalendarYears.Name = "cboCalendarYears"
        Me.cboCalendarYears.Size = New System.Drawing.Size(64, 21)
        Me.cboCalendarYears.TabIndex = 13
        Me.cboCalendarYears.TabStop = False
        Me.ToolTip1.SetToolTip(Me.cboCalendarYears, " The currently selected calendar year ")
        '
        'rdoVendor1099Listing
        '
        Me.rdoVendor1099Listing.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoVendor1099Listing.Location = New System.Drawing.Point(24, 72)
        Me.rdoVendor1099Listing.Name = "rdoVendor1099Listing"
        Me.rdoVendor1099Listing.Size = New System.Drawing.Size(200, 16)
        Me.rdoVendor1099Listing.TabIndex = 11
        Me.rdoVendor1099Listing.Text = "Vendor 1099 detail report"
        Me.ToolTip1.SetToolTip(Me.rdoVendor1099Listing, " Show detail listing of expenditures by vendor ")
        '
        'chkVendorsIncludeZeroBalances
        '
        Me.chkVendorsIncludeZeroBalances.Enabled = False
        Me.chkVendorsIncludeZeroBalances.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVendorsIncludeZeroBalances.Location = New System.Drawing.Point(272, 96)
        Me.chkVendorsIncludeZeroBalances.Name = "chkVendorsIncludeZeroBalances"
        Me.chkVendorsIncludeZeroBalances.Size = New System.Drawing.Size(144, 16)
        Me.chkVendorsIncludeZeroBalances.TabIndex = 10
        Me.chkVendorsIncludeZeroBalances.Text = "Include zero balances"
        Me.ToolTip1.SetToolTip(Me.chkVendorsIncludeZeroBalances, " Include vendors with no activity ")
        '
        'rdoVendorExpenses
        '
        Me.rdoVendorExpenses.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoVendorExpenses.Location = New System.Drawing.Point(24, 48)
        Me.rdoVendorExpenses.Name = "rdoVendorExpenses"
        Me.rdoVendorExpenses.Size = New System.Drawing.Size(200, 16)
        Me.rdoVendorExpenses.TabIndex = 9
        Me.rdoVendorExpenses.Text = "Vendor expense summary"
        Me.ToolTip1.SetToolTip(Me.rdoVendorExpenses, " Summarise vendors by expenditure amount ")
        '
        'chk1099ByEmployee
        '
        Me.chk1099ByEmployee.Enabled = False
        Me.chk1099ByEmployee.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk1099ByEmployee.Location = New System.Drawing.Point(272, 72)
        Me.chk1099ByEmployee.Name = "chk1099ByEmployee"
        Me.chk1099ByEmployee.Size = New System.Drawing.Size(160, 16)
        Me.chk1099ByEmployee.TabIndex = 8
        Me.chk1099ByEmployee.Text = "1099 vendor by employee"
        Me.ToolTip1.SetToolTip(Me.chk1099ByEmployee, " Show 1099 vendors only ")
        '
        'rdoVendorListing
        '
        Me.rdoVendorListing.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoVendorListing.Location = New System.Drawing.Point(24, 24)
        Me.rdoVendorListing.Name = "rdoVendorListing"
        Me.rdoVendorListing.Size = New System.Drawing.Size(200, 16)
        Me.rdoVendorListing.TabIndex = 7
        Me.rdoVendorListing.Text = "Vendor listing"
        Me.ToolTip1.SetToolTip(Me.rdoVendorListing, " List active vendors")
        '
        'lblVendorCalendar
        '
        Me.lblVendorCalendar.Enabled = False
        Me.lblVendorCalendar.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorCalendar.Location = New System.Drawing.Point(256, 24)
        Me.lblVendorCalendar.Name = "lblVendorCalendar"
        Me.lblVendorCalendar.Size = New System.Drawing.Size(96, 32)
        Me.lblVendorCalendar.TabIndex = 14
        Me.lblVendorCalendar.Text = "Select or enter calendar year:"
        '
        'chk1099Summary
        '
        Me.chk1099Summary.Enabled = False
        Me.chk1099Summary.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk1099Summary.Location = New System.Drawing.Point(272, 56)
        Me.chk1099Summary.Name = "chk1099Summary"
        Me.chk1099Summary.Size = New System.Drawing.Size(152, 16)
        Me.chk1099Summary.TabIndex = 21
        Me.chk1099Summary.Text = "1099 vendor summary"
        Me.ToolTip1.SetToolTip(Me.chk1099Summary, " Use fiscal year date range ")
        '
        'rdoBoldCode
        '
        Me.rdoBoldCode.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoBoldCode.Location = New System.Drawing.Point(16, 128)
        Me.rdoBoldCode.Name = "rdoBoldCode"
        Me.rdoBoldCode.Size = New System.Drawing.Size(112, 16)
        Me.rdoBoldCode.TabIndex = 16
        Me.rdoBoldCode.Text = "Bold Code"
        '
        'grpBoldCode
        '
        Me.grpBoldCode.Controls.Add(Me.chkChkaCode)
        Me.grpBoldCode.Controls.Add(Me.grpCheck)
        Me.grpBoldCode.Controls.Add(Me.lblBoldCodeMessage)
        Me.grpBoldCode.Controls.Add(Me.chkBoldCodeSortDetailByCoding)
        Me.grpBoldCode.Controls.Add(Me.chkBoldCodeDetailErrorsOnly)
        Me.grpBoldCode.Controls.Add(Me.rdoBoldCodeListingByRevenues)
        Me.grpBoldCode.Controls.Add(Me.rdoBoldCodeListingByExpenditures)
        Me.grpBoldCode.Location = New System.Drawing.Point(1430, 80)
        Me.grpBoldCode.Name = "grpBoldCode"
        Me.grpBoldCode.Size = New System.Drawing.Size(601, 224)
        Me.grpBoldCode.TabIndex = 17
        Me.grpBoldCode.TabStop = False
        Me.grpBoldCode.Text = " Bold code"
        '
        'chkChkaCode
        '
        Me.chkChkaCode.Location = New System.Drawing.Point(480, 16)
        Me.chkChkaCode.Name = "chkChkaCode"
        Me.chkChkaCode.Size = New System.Drawing.Size(104, 24)
        Me.chkChkaCode.TabIndex = 73
        Me.chkChkaCode.Text = "Check a Code"
        '
        'grpCheck
        '
        Me.grpCheck.Controls.Add(Me.btnVerifyRev)
        Me.grpCheck.Controls.Add(Me.btnVerifyExp)
        Me.grpCheck.Controls.Add(Me.Label23)
        Me.grpCheck.Controls.Add(Me.Label24)
        Me.grpCheck.Controls.Add(Me.Label25)
        Me.grpCheck.Controls.Add(Me.Label26)
        Me.grpCheck.Controls.Add(Me.Label27)
        Me.grpCheck.Controls.Add(Me.Label28)
        Me.grpCheck.Controls.Add(Me.Label29)
        Me.grpCheck.Controls.Add(Me.Label30)
        Me.grpCheck.Controls.Add(Me.Label31)
        Me.grpCheck.Controls.Add(Me.txtC1SiteExp)
        Me.grpCheck.Controls.Add(Me.txtC1Job)
        Me.grpCheck.Controls.Add(Me.txtC1Subject)
        Me.grpCheck.Controls.Add(Me.txtC1ProgramExp)
        Me.grpCheck.Controls.Add(Me.txtC1Object)
        Me.grpCheck.Controls.Add(Me.txtC1Function)
        Me.grpCheck.Controls.Add(Me.txtC1ProjectExp)
        Me.grpCheck.Controls.Add(Me.txtC1FundExp)
        Me.grpCheck.Controls.Add(Me.txtC1YearExp)
        Me.grpCheck.Controls.Add(Me.Label17)
        Me.grpCheck.Controls.Add(Me.Label18)
        Me.grpCheck.Controls.Add(Me.Label19)
        Me.grpCheck.Controls.Add(Me.Label20)
        Me.grpCheck.Controls.Add(Me.Label21)
        Me.grpCheck.Controls.Add(Me.Label22)
        Me.grpCheck.Controls.Add(Me.txtC1Site)
        Me.grpCheck.Controls.Add(Me.txtC1Program)
        Me.grpCheck.Controls.Add(Me.txtC1Source)
        Me.grpCheck.Controls.Add(Me.txtC1Project)
        Me.grpCheck.Controls.Add(Me.txtC1Fund)
        Me.grpCheck.Controls.Add(Me.txtC1Year)
        Me.grpCheck.Controls.Add(Me.lblBoldRev)
        Me.grpCheck.Controls.Add(Me.lblBoldExp)
        Me.grpCheck.Location = New System.Drawing.Point(3, 94)
        Me.grpCheck.Name = "grpCheck"
        Me.grpCheck.Size = New System.Drawing.Size(589, 104)
        Me.grpCheck.TabIndex = 72
        Me.grpCheck.TabStop = False
        Me.grpCheck.Text = "Edit Check a Code"
        Me.grpCheck.Visible = False
        '
        'btnVerifyRev
        '
        Me.btnVerifyRev.Image = CType(resources.GetObject("btnVerifyRev.Image"), System.Drawing.Image)
        Me.btnVerifyRev.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnVerifyRev.Location = New System.Drawing.Point(506, 65)
        Me.btnVerifyRev.Name = "btnVerifyRev"
        Me.btnVerifyRev.Size = New System.Drawing.Size(80, 26)
        Me.btnVerifyRev.TabIndex = 81
        Me.btnVerifyRev.Text = "Test"
        '
        'btnVerifyExp
        '
        Me.btnVerifyExp.Image = CType(resources.GetObject("btnVerifyExp.Image"), System.Drawing.Image)
        Me.btnVerifyExp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnVerifyExp.Location = New System.Drawing.Point(506, 27)
        Me.btnVerifyExp.Name = "btnVerifyExp"
        Me.btnVerifyExp.Size = New System.Drawing.Size(80, 24)
        Me.btnVerifyExp.TabIndex = 96
        Me.btnVerifyExp.Text = "Test"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(468, 13)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(32, 16)
        Me.Label23.TabIndex = 105
        Me.Label23.Text = "Site"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(436, 13)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(24, 16)
        Me.Label24.TabIndex = 104
        Me.Label24.Text = "Job"
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(396, 13)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(32, 16)
        Me.Label25.TabIndex = 103
        Me.Label25.Text = "Subj"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(364, 13)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(32, 16)
        Me.Label26.TabIndex = 102
        Me.Label26.Text = "Pgm"
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(260, 13)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(24, 16)
        Me.Label27.TabIndex = 101
        Me.Label27.Text = "Prj"
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(332, 13)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(24, 16)
        Me.Label28.TabIndex = 100
        Me.Label28.Text = "Obj"
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(292, 13)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(32, 16)
        Me.Label29.TabIndex = 99
        Me.Label29.Text = "Func"
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(227, 13)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(24, 16)
        Me.Label30.TabIndex = 98
        Me.Label30.Text = "Fd"
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(210, 13)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(16, 16)
        Me.Label31.TabIndex = 97
        Me.Label31.Text = "Yr"
        '
        'txtC1SiteExp
        '
        Me.txtC1SiteExp.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1SiteExp.CustomFormat = "000;"
        Me.txtC1SiteExp.DataType = GetType(Integer)
        Me.txtC1SiteExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1SiteExp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1SiteExp.Location = New System.Drawing.Point(468, 29)
        Me.txtC1SiteExp.MaxLength = 3
        Me.txtC1SiteExp.Name = "txtC1SiteExp"
        Me.txtC1SiteExp.Size = New System.Drawing.Size(32, 21)
        Me.txtC1SiteExp.TabIndex = 95
        Me.txtC1SiteExp.Tag = Nothing
        Me.txtC1SiteExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Job
        '
        Me.txtC1Job.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Job.CustomFormat = "000;"
        Me.txtC1Job.DataType = GetType(Integer)
        Me.txtC1Job.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Job.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Job.Location = New System.Drawing.Point(436, 29)
        Me.txtC1Job.MaxLength = 3
        Me.txtC1Job.Name = "txtC1Job"
        Me.txtC1Job.Size = New System.Drawing.Size(32, 21)
        Me.txtC1Job.TabIndex = 94
        Me.txtC1Job.Tag = Nothing
        Me.txtC1Job.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Subject
        '
        Me.txtC1Subject.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Subject.CustomFormat = "0000;"
        Me.txtC1Subject.DataType = GetType(Integer)
        Me.txtC1Subject.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Subject.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Subject.Location = New System.Drawing.Point(388, 29)
        Me.txtC1Subject.MaxLength = 4
        Me.txtC1Subject.Name = "txtC1Subject"
        Me.txtC1Subject.Size = New System.Drawing.Size(40, 21)
        Me.txtC1Subject.TabIndex = 93
        Me.txtC1Subject.Tag = Nothing
        Me.txtC1Subject.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1ProgramExp
        '
        Me.txtC1ProgramExp.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1ProgramExp.CustomFormat = "000;"
        Me.txtC1ProgramExp.DataType = GetType(Integer)
        Me.txtC1ProgramExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1ProgramExp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1ProgramExp.Location = New System.Drawing.Point(356, 29)
        Me.txtC1ProgramExp.MaxLength = 3
        Me.txtC1ProgramExp.Name = "txtC1ProgramExp"
        Me.txtC1ProgramExp.Size = New System.Drawing.Size(32, 21)
        Me.txtC1ProgramExp.TabIndex = 92
        Me.txtC1ProgramExp.Tag = Nothing
        Me.txtC1ProgramExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Object
        '
        Me.txtC1Object.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Object.CustomFormat = "000;"
        Me.txtC1Object.DataType = GetType(Integer)
        Me.txtC1Object.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Object.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Object.Location = New System.Drawing.Point(324, 29)
        Me.txtC1Object.MaxLength = 3
        Me.txtC1Object.Name = "txtC1Object"
        Me.txtC1Object.Size = New System.Drawing.Size(32, 21)
        Me.txtC1Object.TabIndex = 91
        Me.txtC1Object.Tag = Nothing
        Me.txtC1Object.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Function
        '
        Me.txtC1Function.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Function.CustomFormat = "0000;"
        Me.txtC1Function.DataType = GetType(Integer)
        Me.txtC1Function.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Function.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Function.Location = New System.Drawing.Point(284, 29)
        Me.txtC1Function.MaxLength = 4
        Me.txtC1Function.Name = "txtC1Function"
        Me.txtC1Function.Size = New System.Drawing.Size(40, 21)
        Me.txtC1Function.TabIndex = 89
        Me.txtC1Function.Tag = Nothing
        Me.txtC1Function.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1ProjectExp
        '
        Me.txtC1ProjectExp.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1ProjectExp.CustomFormat = "000;"
        Me.txtC1ProjectExp.DataType = GetType(Integer)
        Me.txtC1ProjectExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1ProjectExp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1ProjectExp.Location = New System.Drawing.Point(252, 29)
        Me.txtC1ProjectExp.MaxLength = 3
        Me.txtC1ProjectExp.Name = "txtC1ProjectExp"
        Me.txtC1ProjectExp.Size = New System.Drawing.Size(32, 21)
        Me.txtC1ProjectExp.TabIndex = 86
        Me.txtC1ProjectExp.Tag = Nothing
        Me.txtC1ProjectExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1FundExp
        '
        Me.txtC1FundExp.BackColor = System.Drawing.Color.Ivory
        Me.txtC1FundExp.CustomFormat = "00;"
        Me.txtC1FundExp.DataType = GetType(Integer)
        Me.txtC1FundExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1FundExp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1FundExp.Location = New System.Drawing.Point(227, 29)
        Me.txtC1FundExp.MaxLength = 2
        Me.txtC1FundExp.Name = "txtC1FundExp"
        Me.txtC1FundExp.Size = New System.Drawing.Size(24, 21)
        Me.txtC1FundExp.TabIndex = 84
        Me.txtC1FundExp.Tag = Nothing
        Me.txtC1FundExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1YearExp
        '
        Me.txtC1YearExp.AcceptsReturn = True
        Me.txtC1YearExp.BackColor = System.Drawing.Color.Ivory
        Me.txtC1YearExp.CustomFormat = "0;"
        Me.txtC1YearExp.DataType = GetType(Integer)
        Me.txtC1YearExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1YearExp.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1YearExp.Location = New System.Drawing.Point(202, 29)
        Me.txtC1YearExp.MaxLength = 1
        Me.txtC1YearExp.Name = "txtC1YearExp"
        Me.txtC1YearExp.Size = New System.Drawing.Size(24, 21)
        Me.txtC1YearExp.TabIndex = 82
        Me.txtC1YearExp.Tag = Nothing
        Me.txtC1YearExp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(364, 53)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(32, 14)
        Me.Label17.TabIndex = 90
        Me.Label17.Text = "Site"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(326, 53)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(32, 14)
        Me.Label18.TabIndex = 88
        Me.Label18.Text = "Pgm"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(260, 53)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(24, 14)
        Me.Label19.TabIndex = 87
        Me.Label19.Text = "Prj"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(292, 53)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(28, 14)
        Me.Label20.TabIndex = 85
        Me.Label20.Text = "Src"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(228, 53)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(24, 14)
        Me.Label21.TabIndex = 83
        Me.Label21.Text = "Fd"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(202, 53)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(16, 14)
        Me.Label22.TabIndex = 80
        Me.Label22.Text = "Yr"
        '
        'txtC1Site
        '
        Me.txtC1Site.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Site.CustomFormat = "000;"
        Me.txtC1Site.DataType = GetType(Integer)
        Me.txtC1Site.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Site.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Site.Location = New System.Drawing.Point(356, 69)
        Me.txtC1Site.MaxLength = 3
        Me.txtC1Site.Name = "txtC1Site"
        Me.txtC1Site.Size = New System.Drawing.Size(32, 21)
        Me.txtC1Site.TabIndex = 79
        Me.txtC1Site.Tag = Nothing
        Me.txtC1Site.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Program
        '
        Me.txtC1Program.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Program.CustomFormat = "000;"
        Me.txtC1Program.DataType = GetType(Integer)
        Me.txtC1Program.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Program.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Program.Location = New System.Drawing.Point(324, 69)
        Me.txtC1Program.MaxLength = 3
        Me.txtC1Program.Name = "txtC1Program"
        Me.txtC1Program.Size = New System.Drawing.Size(32, 21)
        Me.txtC1Program.TabIndex = 78
        Me.txtC1Program.Tag = Nothing
        Me.txtC1Program.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Source
        '
        Me.txtC1Source.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Source.CustomFormat = "0000;"
        Me.txtC1Source.DataType = GetType(Integer)
        Me.txtC1Source.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Source.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Source.Location = New System.Drawing.Point(284, 69)
        Me.txtC1Source.MaxLength = 4
        Me.txtC1Source.Name = "txtC1Source"
        Me.txtC1Source.Size = New System.Drawing.Size(40, 21)
        Me.txtC1Source.TabIndex = 77
        Me.txtC1Source.Tag = Nothing
        Me.txtC1Source.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Project
        '
        Me.txtC1Project.BackColor = System.Drawing.SystemColors.Window
        Me.txtC1Project.CustomFormat = "000;"
        Me.txtC1Project.DataType = GetType(Integer)
        Me.txtC1Project.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Project.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Project.Location = New System.Drawing.Point(252, 69)
        Me.txtC1Project.MaxLength = 3
        Me.txtC1Project.Name = "txtC1Project"
        Me.txtC1Project.Size = New System.Drawing.Size(32, 21)
        Me.txtC1Project.TabIndex = 76
        Me.txtC1Project.Tag = Nothing
        Me.txtC1Project.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Fund
        '
        Me.txtC1Fund.BackColor = System.Drawing.Color.Ivory
        Me.txtC1Fund.CustomFormat = "00;"
        Me.txtC1Fund.DataType = GetType(Integer)
        Me.txtC1Fund.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Fund.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Fund.Location = New System.Drawing.Point(226, 69)
        Me.txtC1Fund.MaxLength = 2
        Me.txtC1Fund.Name = "txtC1Fund"
        Me.txtC1Fund.Size = New System.Drawing.Size(24, 21)
        Me.txtC1Fund.TabIndex = 75
        Me.txtC1Fund.Tag = Nothing
        Me.txtC1Fund.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtC1Year
        '
        Me.txtC1Year.BackColor = System.Drawing.Color.Ivory
        Me.txtC1Year.CustomFormat = "0;"
        Me.txtC1Year.DataType = GetType(Integer)
        Me.txtC1Year.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtC1Year.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.txtC1Year.Location = New System.Drawing.Point(202, 69)
        Me.txtC1Year.MaxLength = 1
        Me.txtC1Year.Name = "txtC1Year"
        Me.txtC1Year.Size = New System.Drawing.Size(24, 21)
        Me.txtC1Year.TabIndex = 74
        Me.txtC1Year.Tag = Nothing
        Me.txtC1Year.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblBoldRev
        '
        Me.lblBoldRev.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoldRev.Location = New System.Drawing.Point(2, 66)
        Me.lblBoldRev.Name = "lblBoldRev"
        Me.lblBoldRev.Size = New System.Drawing.Size(184, 16)
        Me.lblBoldRev.TabIndex = 73
        Me.lblBoldRev.Text = "Edit Check For Revenue:"
        Me.lblBoldRev.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBoldExp
        '
        Me.lblBoldExp.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoldExp.Location = New System.Drawing.Point(2, 31)
        Me.lblBoldExp.Name = "lblBoldExp"
        Me.lblBoldExp.Size = New System.Drawing.Size(184, 16)
        Me.lblBoldExp.TabIndex = 72
        Me.lblBoldExp.Text = "Edit Check For Expenditure:"
        Me.lblBoldExp.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBoldCodeMessage
        '
        Me.lblBoldCodeMessage.BackColor = System.Drawing.Color.Red
        Me.lblBoldCodeMessage.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoldCodeMessage.ForeColor = System.Drawing.Color.White
        Me.lblBoldCodeMessage.Location = New System.Drawing.Point(8, 201)
        Me.lblBoldCodeMessage.Name = "lblBoldCodeMessage"
        Me.lblBoldCodeMessage.Size = New System.Drawing.Size(416, 16)
        Me.lblBoldCodeMessage.TabIndex = 10
        Me.lblBoldCodeMessage.Text = "Bold code reporting invalid for this site."
        Me.lblBoldCodeMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblBoldCodeMessage.Visible = False
        '
        'chkBoldCodeSortDetailByCoding
        '
        Me.chkBoldCodeSortDetailByCoding.Enabled = False
        Me.chkBoldCodeSortDetailByCoding.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBoldCodeSortDetailByCoding.Location = New System.Drawing.Point(48, 97)
        Me.chkBoldCodeSortDetailByCoding.Name = "chkBoldCodeSortDetailByCoding"
        Me.chkBoldCodeSortDetailByCoding.Size = New System.Drawing.Size(224, 16)
        Me.chkBoldCodeSortDetailByCoding.TabIndex = 9
        Me.chkBoldCodeSortDetailByCoding.TabStop = False
        Me.chkBoldCodeSortDetailByCoding.Text = "Sort detail by coding (Date is default)"
        Me.chkBoldCodeSortDetailByCoding.Visible = False
        '
        'chkBoldCodeDetailErrorsOnly
        '
        Me.chkBoldCodeDetailErrorsOnly.Enabled = False
        Me.chkBoldCodeDetailErrorsOnly.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBoldCodeDetailErrorsOnly.Location = New System.Drawing.Point(48, 73)
        Me.chkBoldCodeDetailErrorsOnly.Name = "chkBoldCodeDetailErrorsOnly"
        Me.chkBoldCodeDetailErrorsOnly.Size = New System.Drawing.Size(176, 16)
        Me.chkBoldCodeDetailErrorsOnly.TabIndex = 8
        Me.chkBoldCodeDetailErrorsOnly.TabStop = False
        Me.chkBoldCodeDetailErrorsOnly.Text = "Bold code detail errors only"
        Me.chkBoldCodeDetailErrorsOnly.Visible = False
        '
        'rdoBoldCodeListingByRevenues
        '
        Me.rdoBoldCodeListingByRevenues.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoBoldCodeListingByRevenues.Location = New System.Drawing.Point(24, 48)
        Me.rdoBoldCodeListingByRevenues.Name = "rdoBoldCodeListingByRevenues"
        Me.rdoBoldCodeListingByRevenues.Size = New System.Drawing.Size(216, 16)
        Me.rdoBoldCodeListingByRevenues.TabIndex = 7
        Me.rdoBoldCodeListingByRevenues.Text = "Bold code listing by revenues"
        '
        'rdoBoldCodeListingByExpenditures
        '
        Me.rdoBoldCodeListingByExpenditures.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoBoldCodeListingByExpenditures.Location = New System.Drawing.Point(24, 24)
        Me.rdoBoldCodeListingByExpenditures.Name = "rdoBoldCodeListingByExpenditures"
        Me.rdoBoldCodeListingByExpenditures.Size = New System.Drawing.Size(216, 16)
        Me.rdoBoldCodeListingByExpenditures.TabIndex = 6
        Me.rdoBoldCodeListingByExpenditures.Text = "Bold code listing by expenditures"
        '
        'grpDateRange
        '
        Me.grpDateRange.Controls.Add(Me.Label3)
        Me.grpDateRange.Controls.Add(Me.Label2)
        Me.grpDateRange.Controls.Add(Me.dtEndingDate)
        Me.grpDateRange.Controls.Add(Me.dtBeginningDate)
        Me.grpDateRange.Location = New System.Drawing.Point(144, 307)
        Me.grpDateRange.Name = "grpDateRange"
        Me.grpDateRange.Size = New System.Drawing.Size(224, 56)
        Me.grpDateRange.TabIndex = 18
        Me.grpDateRange.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Ending date:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Beginning date:"
        '
        'dtEndingDate
        '
        Me.dtEndingDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtEndingDate.Location = New System.Drawing.Point(120, 24)
        Me.dtEndingDate.MaxDate = New Date(2099, 12, 31, 0, 0, 0, 0)
        Me.dtEndingDate.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtEndingDate.Name = "dtEndingDate"
        Me.dtEndingDate.Size = New System.Drawing.Size(88, 20)
        Me.dtEndingDate.TabIndex = 1
        '
        'dtBeginningDate
        '
        Me.dtBeginningDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtBeginningDate.Location = New System.Drawing.Point(16, 24)
        Me.dtBeginningDate.MaxDate = New Date(2099, 12, 31, 0, 0, 0, 0)
        Me.dtBeginningDate.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtBeginningDate.Name = "dtBeginningDate"
        Me.dtBeginningDate.Size = New System.Drawing.Size(88, 20)
        Me.dtBeginningDate.TabIndex = 0
        '
        'grpNumberRange
        '
        Me.grpNumberRange.Controls.Add(Me.Label5)
        Me.grpNumberRange.Controls.Add(Me.Label4)
        Me.grpNumberRange.Controls.Add(Me.txtEndingNumber)
        Me.grpNumberRange.Controls.Add(Me.txtBeginningNumber)
        Me.grpNumberRange.Location = New System.Drawing.Point(144, 307)
        Me.grpNumberRange.Name = "grpNumberRange"
        Me.grpNumberRange.Size = New System.Drawing.Size(224, 56)
        Me.grpNumberRange.TabIndex = 19
        Me.grpNumberRange.TabStop = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(120, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Ending number:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Beginning number:"
        '
        'txtEndingNumber
        '
        Me.txtEndingNumber.Location = New System.Drawing.Point(120, 24)
        Me.txtEndingNumber.Name = "txtEndingNumber"
        Me.txtEndingNumber.Size = New System.Drawing.Size(88, 20)
        Me.txtEndingNumber.TabIndex = 1
        Me.txtEndingNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBeginningNumber
        '
        Me.txtBeginningNumber.Location = New System.Drawing.Point(16, 24)
        Me.txtBeginningNumber.Name = "txtBeginningNumber"
        Me.txtBeginningNumber.Size = New System.Drawing.Size(88, 20)
        Me.txtBeginningNumber.TabIndex = 0
        Me.txtBeginningNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'chkUseDate
        '
        Me.chkUseDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseDate.Location = New System.Drawing.Point(16, 8)
        Me.chkUseDate.Name = "chkUseDate"
        Me.chkUseDate.Size = New System.Drawing.Size(96, 16)
        Me.chkUseDate.TabIndex = 4
        Me.chkUseDate.TabStop = False
        Me.chkUseDate.Text = "Use date"
        '
        'chkUseNumber
        '
        Me.chkUseNumber.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseNumber.Location = New System.Drawing.Point(16, 32)
        Me.chkUseNumber.Name = "chkUseNumber"
        Me.chkUseNumber.Size = New System.Drawing.Size(96, 16)
        Me.chkUseNumber.TabIndex = 6
        Me.chkUseNumber.TabStop = False
        Me.chkUseNumber.Text = "Use number"
        '
        'panelOptions
        '
        Me.panelOptions.Controls.Add(Me.chkUseDate)
        Me.panelOptions.Controls.Add(Me.chkUseNumber)
        Me.panelOptions.Enabled = False
        Me.panelOptions.Location = New System.Drawing.Point(8, 304)
        Me.panelOptions.Name = "panelOptions"
        Me.panelOptions.Size = New System.Drawing.Size(128, 56)
        Me.panelOptions.TabIndex = 21
        '
        'btnExit
        '
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.Location = New System.Drawing.Point(496, 323)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(96, 32)
        Me.btnExit.TabIndex = 227
        Me.btnExit.TabStop = False
        Me.btnExit.Text = "  Exit"
        '
        'btnPreview
        '
        Me.btnPreview.Image = CType(resources.GetObject("btnPreview.Image"), System.Drawing.Image)
        Me.btnPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPreview.Location = New System.Drawing.Point(376, 323)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(96, 32)
        Me.btnPreview.TabIndex = 226
        Me.btnPreview.Text = "Preview"
        '
        'btnCancelPreview
        '
        Me.btnCancelPreview.Image = CType(resources.GetObject("btnCancelPreview.Image"), System.Drawing.Image)
        Me.btnCancelPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancelPreview.Location = New System.Drawing.Point(496, 323)
        Me.btnCancelPreview.Name = "btnCancelPreview"
        Me.btnCancelPreview.Size = New System.Drawing.Size(96, 32)
        Me.btnCancelPreview.TabIndex = 228
        Me.btnCancelPreview.Text = "Cancel"
        Me.btnCancelPreview.Visible = False
        '
        'lblSchoolName
        '
        Me.lblSchoolName.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSchoolName.Location = New System.Drawing.Point(320, 48)
        Me.lblSchoolName.Name = "lblSchoolName"
        Me.lblSchoolName.Size = New System.Drawing.Size(344, 16)
        Me.lblSchoolName.TabIndex = 229
        Me.lblSchoolName.Text = "School name"
        Me.lblSchoolName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'rdoReconTrialBalance
        '
        Me.rdoReconTrialBalance.Checked = True
        Me.rdoReconTrialBalance.Location = New System.Drawing.Point(24, 24)
        Me.rdoReconTrialBalance.Name = "rdoReconTrialBalance"
        Me.rdoReconTrialBalance.Size = New System.Drawing.Size(88, 16)
        Me.rdoReconTrialBalance.TabIndex = 5
        Me.rdoReconTrialBalance.TabStop = True
        Me.rdoReconTrialBalance.Text = "Trial balance"
        Me.ToolTip1.SetToolTip(Me.rdoReconTrialBalance, " Print trial balance reconciliation ")
        '
        'rdoClassificationRevenueCodes
        '
        Me.rdoClassificationRevenueCodes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationRevenueCodes.Location = New System.Drawing.Point(240, 56)
        Me.rdoClassificationRevenueCodes.Name = "rdoClassificationRevenueCodes"
        Me.rdoClassificationRevenueCodes.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationRevenueCodes.TabIndex = 27
        Me.rdoClassificationRevenueCodes.Text = "Revenue code listing"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationRevenueCodes, " Lists all revenue codes for the selected fiscal year ")
        '
        'rdoClassificationExpenditureCodes
        '
        Me.rdoClassificationExpenditureCodes.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationExpenditureCodes.Location = New System.Drawing.Point(240, 32)
        Me.rdoClassificationExpenditureCodes.Name = "rdoClassificationExpenditureCodes"
        Me.rdoClassificationExpenditureCodes.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationExpenditureCodes.TabIndex = 26
        Me.rdoClassificationExpenditureCodes.Text = "Expenditure code listing"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationExpenditureCodes, " Lists all expenditure codes for the selected fiscal year ")
        '
        'rdoClassificationYTDRevenue
        '
        Me.rdoClassificationYTDRevenue.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationYTDRevenue.Location = New System.Drawing.Point(24, 104)
        Me.rdoClassificationYTDRevenue.Name = "rdoClassificationYTDRevenue"
        Me.rdoClassificationYTDRevenue.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationYTDRevenue.TabIndex = 14
        Me.rdoClassificationYTDRevenue.Text = "YTD revenue by code"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationYTDRevenue, " Lists revenue by code or code range for the fiscal year ")
        '
        'rdoClassificationYTDExpend
        '
        Me.rdoClassificationYTDExpend.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationYTDExpend.Location = New System.Drawing.Point(24, 56)
        Me.rdoClassificationYTDExpend.Name = "rdoClassificationYTDExpend"
        Me.rdoClassificationYTDExpend.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationYTDExpend.TabIndex = 13
        Me.rdoClassificationYTDExpend.Text = "YTD expense by code"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationYTDExpend, " Lists expenses by code or code range for the fiscal year ")
        '
        'rdoClassificationMTDRevenue
        '
        Me.rdoClassificationMTDRevenue.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationMTDRevenue.Location = New System.Drawing.Point(24, 80)
        Me.rdoClassificationMTDRevenue.Name = "rdoClassificationMTDRevenue"
        Me.rdoClassificationMTDRevenue.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationMTDRevenue.TabIndex = 7
        Me.rdoClassificationMTDRevenue.Text = "MTD revenue by code"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationMTDRevenue, " Lists revenue by code or code range for the current month ")
        '
        'rdoClassificationMTDExpend
        '
        Me.rdoClassificationMTDExpend.Checked = True
        Me.rdoClassificationMTDExpend.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassificationMTDExpend.Location = New System.Drawing.Point(24, 32)
        Me.rdoClassificationMTDExpend.Name = "rdoClassificationMTDExpend"
        Me.rdoClassificationMTDExpend.Size = New System.Drawing.Size(192, 16)
        Me.rdoClassificationMTDExpend.TabIndex = 6
        Me.rdoClassificationMTDExpend.TabStop = True
        Me.rdoClassificationMTDExpend.Text = "MTD expense by code"
        Me.ToolTip1.SetToolTip(Me.rdoClassificationMTDExpend, " Lists expenses by code or code range for the current month ")
        '
        'panelReconciliationTrialBalance
        '
        Me.panelReconciliationTrialBalance.Controls.Add(Me.dtReconTrialBalanceDate)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.Label11)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.Label10)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.Label9)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.Label8)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.txtReconInvestments)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.txtReconExpensesNotYetPosted)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.txtReconInterestNotYetPosted)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.txtReconBankStatementBalance)
        Me.panelReconciliationTrialBalance.Controls.Add(Me.Label7)
        Me.panelReconciliationTrialBalance.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panelReconciliationTrialBalance.Location = New System.Drawing.Point(200, 8)
        Me.panelReconciliationTrialBalance.Name = "panelReconciliationTrialBalance"
        Me.panelReconciliationTrialBalance.Size = New System.Drawing.Size(232, 144)
        Me.panelReconciliationTrialBalance.TabIndex = 6
        '
        'dtReconTrialBalanceDate
        '
        Me.dtReconTrialBalanceDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtReconTrialBalanceDate.Location = New System.Drawing.Point(144, 8)
        Me.dtReconTrialBalanceDate.MaxDate = New Date(2099, 12, 31, 0, 0, 0, 0)
        Me.dtReconTrialBalanceDate.MinDate = New Date(2001, 1, 1, 0, 0, 0, 0)
        Me.dtReconTrialBalanceDate.Name = "dtReconTrialBalanceDate"
        Me.dtReconTrialBalanceDate.Size = New System.Drawing.Size(80, 20)
        Me.dtReconTrialBalanceDate.TabIndex = 9
        Me.dtReconTrialBalanceDate.TabStop = False
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(128, 16)
        Me.Label11.TabIndex = 8
        Me.Label11.Text = "Reconciliation date"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 112)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(128, 16)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "Investments"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 88)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(128, 16)
        Me.Label9.TabIndex = 6
        Me.Label9.Text = "Expenses not yet posted"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 64)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(128, 16)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Interest not yet posted"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtReconInvestments
        '
        Me.txtReconInvestments.Location = New System.Drawing.Point(144, 112)
        Me.txtReconInvestments.MaxLength = 15
        Me.txtReconInvestments.Name = "txtReconInvestments"
        Me.txtReconInvestments.Size = New System.Drawing.Size(80, 20)
        Me.txtReconInvestments.TabIndex = 4
        Me.txtReconInvestments.Text = "0.00"
        Me.txtReconInvestments.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtReconExpensesNotYetPosted
        '
        Me.txtReconExpensesNotYetPosted.Location = New System.Drawing.Point(144, 88)
        Me.txtReconExpensesNotYetPosted.MaxLength = 15
        Me.txtReconExpensesNotYetPosted.Name = "txtReconExpensesNotYetPosted"
        Me.txtReconExpensesNotYetPosted.Size = New System.Drawing.Size(80, 20)
        Me.txtReconExpensesNotYetPosted.TabIndex = 3
        Me.txtReconExpensesNotYetPosted.Text = "0.00"
        Me.txtReconExpensesNotYetPosted.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtReconInterestNotYetPosted
        '
        Me.txtReconInterestNotYetPosted.Location = New System.Drawing.Point(144, 64)
        Me.txtReconInterestNotYetPosted.MaxLength = 15
        Me.txtReconInterestNotYetPosted.Name = "txtReconInterestNotYetPosted"
        Me.txtReconInterestNotYetPosted.Size = New System.Drawing.Size(80, 20)
        Me.txtReconInterestNotYetPosted.TabIndex = 2
        Me.txtReconInterestNotYetPosted.Text = "0.00"
        Me.txtReconInterestNotYetPosted.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtReconBankStatementBalance
        '
        Me.txtReconBankStatementBalance.Location = New System.Drawing.Point(144, 40)
        Me.txtReconBankStatementBalance.MaxLength = 15
        Me.txtReconBankStatementBalance.Name = "txtReconBankStatementBalance"
        Me.txtReconBankStatementBalance.Size = New System.Drawing.Size(80, 20)
        Me.txtReconBankStatementBalance.TabIndex = 1
        Me.txtReconBankStatementBalance.Text = "0.00"
        Me.txtReconBankStatementBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(128, 16)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Bank statement balance"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'rdoReconciliation
        '
        Me.rdoReconciliation.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoReconciliation.Location = New System.Drawing.Point(16, 224)
        Me.rdoReconciliation.Name = "rdoReconciliation"
        Me.rdoReconciliation.Size = New System.Drawing.Size(112, 16)
        Me.rdoReconciliation.TabIndex = 230
        Me.rdoReconciliation.Text = "Reconciliation"
        '
        'grpReconciliation
        '
        Me.grpReconciliation.Controls.Add(Me.rdoReconTrialBalance)
        Me.grpReconciliation.Controls.Add(Me.panelReconciliationTrialBalance)
        Me.grpReconciliation.Location = New System.Drawing.Point(1440, 80)
        Me.grpReconciliation.Name = "grpReconciliation"
        Me.grpReconciliation.Size = New System.Drawing.Size(448, 208)
        Me.grpReconciliation.TabIndex = 231
        Me.grpReconciliation.TabStop = False
        Me.grpReconciliation.Text = " Reconciliation"
        '
        'rdoClassification
        '
        Me.rdoClassification.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdoClassification.Location = New System.Drawing.Point(16, 152)
        Me.rdoClassification.Name = "rdoClassification"
        Me.rdoClassification.Size = New System.Drawing.Size(112, 16)
        Me.rdoClassification.TabIndex = 232
        Me.rdoClassification.Text = "Classification"
        '
        'grpClassification
        '
        Me.grpClassification.Controls.Add(Me.rdoClassificationRevenueCodes)
        Me.grpClassification.Controls.Add(Me.rdoClassificationExpenditureCodes)
        Me.grpClassification.Controls.Add(Me.chkClassificationUseCodeRange)
        Me.grpClassification.Controls.Add(Me.panelClassificationCodes)
        Me.grpClassification.Controls.Add(Me.rdoClassificationYTDRevenue)
        Me.grpClassification.Controls.Add(Me.rdoClassificationYTDExpend)
        Me.grpClassification.Controls.Add(Me.rdoClassificationMTDRevenue)
        Me.grpClassification.Controls.Add(Me.rdoClassificationMTDExpend)
        Me.grpClassification.Location = New System.Drawing.Point(1440, 80)
        Me.grpClassification.Name = "grpClassification"
        Me.grpClassification.Size = New System.Drawing.Size(448, 208)
        Me.grpClassification.TabIndex = 233
        Me.grpClassification.TabStop = False
        Me.grpClassification.Text = " Classification"
        '
        'chkClassificationUseCodeRange
        '
        Me.chkClassificationUseCodeRange.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkClassificationUseCodeRange.Location = New System.Drawing.Point(48, 136)
        Me.chkClassificationUseCodeRange.Name = "chkClassificationUseCodeRange"
        Me.chkClassificationUseCodeRange.Size = New System.Drawing.Size(264, 16)
        Me.chkClassificationUseCodeRange.TabIndex = 25
        Me.chkClassificationUseCodeRange.TabStop = False
        Me.chkClassificationUseCodeRange.Text = "Select classification code or range"
        '
        'panelClassificationCodes
        '
        Me.panelClassificationCodes.Controls.Add(Me.txtDim9)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim1)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim7)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim6)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim5)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim4)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim8)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim2)
        Me.panelClassificationCodes.Controls.Add(Me.txtDim3)
        Me.panelClassificationCodes.Location = New System.Drawing.Point(16, 160)
        Me.panelClassificationCodes.Name = "panelClassificationCodes"
        Me.panelClassificationCodes.Size = New System.Drawing.Size(416, 40)
        Me.panelClassificationCodes.TabIndex = 24
        Me.panelClassificationCodes.Visible = False
        '
        'txtDim9
        '
        Me.txtDim9.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim9.Location = New System.Drawing.Point(344, 8)
        Me.txtDim9.MaxLength = 3
        Me.txtDim9.Name = "txtDim9"
        Me.txtDim9.Size = New System.Drawing.Size(32, 21)
        Me.txtDim9.TabIndex = 23
        Me.txtDim9.Text = "***"
        Me.txtDim9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim1
        '
        Me.txtDim1.Enabled = False
        Me.txtDim1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim1.Location = New System.Drawing.Point(32, 8)
        Me.txtDim1.Name = "txtDim1"
        Me.txtDim1.Size = New System.Drawing.Size(16, 21)
        Me.txtDim1.TabIndex = 15
        Me.txtDim1.Text = "*"
        Me.txtDim1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim7
        '
        Me.txtDim7.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim7.Location = New System.Drawing.Point(256, 8)
        Me.txtDim7.MaxLength = 4
        Me.txtDim7.Name = "txtDim7"
        Me.txtDim7.Size = New System.Drawing.Size(40, 21)
        Me.txtDim7.TabIndex = 21
        Me.txtDim7.Text = "****"
        Me.txtDim7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim6
        '
        Me.txtDim6.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim6.Location = New System.Drawing.Point(216, 8)
        Me.txtDim6.MaxLength = 3
        Me.txtDim6.Name = "txtDim6"
        Me.txtDim6.Size = New System.Drawing.Size(32, 21)
        Me.txtDim6.TabIndex = 20
        Me.txtDim6.Text = "***"
        Me.txtDim6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim5
        '
        Me.txtDim5.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim5.Location = New System.Drawing.Point(176, 8)
        Me.txtDim5.MaxLength = 3
        Me.txtDim5.Name = "txtDim5"
        Me.txtDim5.Size = New System.Drawing.Size(32, 21)
        Me.txtDim5.TabIndex = 19
        Me.txtDim5.Text = "***"
        Me.txtDim5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim4
        '
        Me.txtDim4.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim4.Location = New System.Drawing.Point(128, 8)
        Me.txtDim4.MaxLength = 4
        Me.txtDim4.Name = "txtDim4"
        Me.txtDim4.Size = New System.Drawing.Size(40, 21)
        Me.txtDim4.TabIndex = 18
        Me.txtDim4.Text = "****"
        Me.txtDim4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim8
        '
        Me.txtDim8.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim8.Location = New System.Drawing.Point(304, 8)
        Me.txtDim8.MaxLength = 3
        Me.txtDim8.Name = "txtDim8"
        Me.txtDim8.Size = New System.Drawing.Size(32, 21)
        Me.txtDim8.TabIndex = 22
        Me.txtDim8.Text = "***"
        Me.txtDim8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim2
        '
        Me.txtDim2.Enabled = False
        Me.txtDim2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim2.Location = New System.Drawing.Point(56, 8)
        Me.txtDim2.Name = "txtDim2"
        Me.txtDim2.Size = New System.Drawing.Size(24, 21)
        Me.txtDim2.TabIndex = 16
        Me.txtDim2.Text = "60"
        Me.txtDim2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtDim3
        '
        Me.txtDim3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDim3.Location = New System.Drawing.Point(88, 8)
        Me.txtDim3.MaxLength = 3
        Me.txtDim3.Name = "txtDim3"
        Me.txtDim3.Size = New System.Drawing.Size(32, 21)
        Me.txtDim3.TabIndex = 17
        Me.txtDim3.Text = "***"
        Me.txtDim3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.cboBanks)
        Me.Panel1.Controls.Add(Me.cboFiscalYears)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.lblSchoolName)
        Me.Panel1.Location = New System.Drawing.Point(-16, -16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(760, 88)
        Me.Panel1.TabIndex = 234
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(32, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(144, 16)
        Me.Label16.TabIndex = 231
        Me.Label16.Text = "Select bank account:"
        '
        'cboBanks
        '
        Me.cboBanks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboBanks.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboBanks.Location = New System.Drawing.Point(32, 48)
        Me.cboBanks.Name = "cboBanks"
        Me.cboBanks.Size = New System.Drawing.Size(200, 21)
        Me.cboBanks.TabIndex = 230
        '
        'Button1
        '
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(496, 363)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 32)
        Me.Button1.TabIndex = 235
        Me.Button1.Text = "Test"
        Me.Button1.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(754, 416)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.rdoClassification)
        Me.Controls.Add(Me.grpReconciliation)
        Me.Controls.Add(Me.rdoReconciliation)
        Me.Controls.Add(Me.grpAccounts)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.btnCancelPreview)
        Me.Controls.Add(Me.panelOptions)
        Me.Controls.Add(Me.grpNumberRange)
        Me.Controls.Add(Me.grpDateRange)
        Me.Controls.Add(Me.grpBoldCode)
        Me.Controls.Add(Me.rdoBoldCode)
        Me.Controls.Add(Me.grpVendors)
        Me.Controls.Add(Me.grpFinancials)
        Me.Controls.Add(Me.grpRevenue)
        Me.Controls.Add(Me.grpExpenditures)
        Me.Controls.Add(Me.grpAdjustments)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.rdoVendors)
        Me.Controls.Add(Me.rdoAdjustments)
        Me.Controls.Add(Me.rdoFinancials)
        Me.Controls.Add(Me.rdoRevenue)
        Me.Controls.Add(Me.rdoExpenditure)
        Me.Controls.Add(Me.rdoAccounts)
        Me.Controls.Add(Me.grpClassification)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund Reporting "
        Me.ToolTip1.SetToolTip(Me, "Reporting Module")
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAccounts.ResumeLayout(False)
        Me.panelAccountsSelectAccountRange.ResumeLayout(False)
        Me.panelAccountsSelectAccountRange.PerformLayout()
        Me.panelAccountsSelectMonth.ResumeLayout(False)
        Me.grpAdjustments.ResumeLayout(False)
        Me.grpExpenditures.ResumeLayout(False)
        Me.grpRevenue.ResumeLayout(False)
        Me.panelRevenue1098T.ResumeLayout(False)
        Me.panelRevenue1098T.PerformLayout()
        Me.panelRevenueAllFiscalYears.ResumeLayout(False)
        Me.panelRevenueSearchReceived.ResumeLayout(False)
        Me.panelRevenueSearchReceived.PerformLayout()
        Me.panelRevenueDailyDeposit.ResumeLayout(False)
        Me.panelRevenueDailyDeposit.PerformLayout()
        Me.grpFinancials.ResumeLayout(False)
        Me.panelFinancialsSelectAccount.ResumeLayout(False)
        Me.grpVendors.ResumeLayout(False)
        Me.grpBoldCode.ResumeLayout(False)
        Me.grpCheck.ResumeLayout(False)
        CType(Me.txtC1SiteExp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Job, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Subject, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1ProgramExp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Object, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Function, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1ProjectExp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1FundExp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1YearExp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Site, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Program, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Source, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Project, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Fund, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtC1Year, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDateRange.ResumeLayout(False)
        Me.grpNumberRange.ResumeLayout(False)
        Me.grpNumberRange.PerformLayout()
        Me.panelOptions.ResumeLayout(False)
        Me.panelReconciliationTrialBalance.ResumeLayout(False)
        Me.panelReconciliationTrialBalance.PerformLayout()
        Me.grpReconciliation.ResumeLayout(False)
        Me.grpClassification.ResumeLayout(False)
        Me.panelClassificationCodes.ResumeLayout(False)
        Me.panelClassificationCodes.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  Button Events "

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Call SetRegistryEntries()
        Me.Close()
    End Sub

    Private Sub btnPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPreview.Click
        Try
            'get the selected fiscal year;
            Application.DoEvents()
            Me.FiscalYearSelected = CInt(Me.cboFiscalYears.SelectedItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Exit Sub
        Finally

        End Try

        Try
            Me.Cursor = Cursors.WaitCursor
            Me.btnPreview.Enabled = False
            Application.DoEvents()
            Select Case True
                Case Me.rdoAccounts.Checked()
                    Call DoAccounts()
                Case Me.rdoAdjustments.Checked()
                    Call DoAdjustments()
                Case Me.rdoBoldCode.Checked()
                    Call DoBoldCode()
                Case Me.rdoClassification.Checked
                    Call DoClassification()
                Case Me.rdoExpenditure.Checked()
                    Call DoExpenditures()
                Case Me.rdoFinancials.Checked()
                    Call DoFinancials()
                Case Me.rdoReconciliation.Checked
                    Call DoReconciliation()
                Case Me.rdoRevenue.Checked()
                    Call DoRevenues()
                Case Me.rdoVendors.Checked()
                    Call DoVendors()
                Case Else
                    MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
        Finally
            Me.Cursor = Cursors.Default
            Me.btnPreview.Enabled = True
        End Try
    End Sub

#End Region

#Region "  Checkbox Events "

    Private Sub chk1099Summary_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk1099Summary.CheckedChanged
        If Me.chk1099Summary.Checked = True Then
            Me.chk1099ByEmployee.Checked = False
            Me.chk1099ByEmployee.Enabled = False
            Me.chkVendorsIncludeSSN.Checked = False
            Me.chkVendorsIncludeSSN.Enabled = False
            Me.chkVendorsIncludeZeroBalances.Checked = False
            Me.chkVendorsIncludeZeroBalances.Enabled = False
            Me.chkVendorsSelectSingleVendor.Checked = False
            Me.chkVendorsSelectSingleVendor.Enabled = False
            Me.chkVendorsUseFiscalYear.Checked = False
            Me.chkVendorsUseFiscalYear.Enabled = False
            Me.chkVendorsUse600Minimum.Checked = False
            Me.chkVendorsUse600Minimum.Enabled = False
        Else
            Me.chk1099ByEmployee.Enabled = True
            Me.chkVendorsIncludeSSN.Enabled = True
            Me.chkVendorsIncludeZeroBalances.Enabled = True
            Me.chkVendorsSelectSingleVendor.Enabled = True
            Me.chkVendorsUseFiscalYear.Enabled = True
            Me.chkVendorsUse600Minimum.Enabled = True
        End If
    End Sub

    Private Sub chkAccountsUseAccountRange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAccountsUseAccountRange.CheckedChanged
        Dim checked As Boolean
        Me.panelAccountsSelectAccountRange.Visible = False
        checked = Me.chkAccountsUseAccountRange.Checked
        Me.panelAccountsSelectAccountRange.Visible = checked
        If checked Then Me.txtAccountsAcctNumberFrom.Focus()
    End Sub

    Private Sub chkExpendituresAllFiscalYears_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkExpendituresAllFiscalYears.Click
        If chkExpendituresAllFiscalYears.Checked = True Then
            Me.panelOptions.Enabled = False
            Me.chkUseDate.Checked = False
            Me.chkUseNumber.Checked = False
        Else
            Me.panelOptions.Enabled = True
        End If
    End Sub

    Private Sub chkHandleClassificationCheckboxes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkClassificationUseCodeRange.CheckedChanged
        Dim checked As Boolean
        checked = Me.chkClassificationUseCodeRange.Checked
        If checked Then
            Me.panelClassificationCodes.Visible = True
            Me.txtDim3.Focus()
        Else
            Me.panelClassificationCodes.Visible = False
        End If
    End Sub

    Private Sub chkHandleBoldCodeCheckboxes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBoldCodeDetailErrorsOnly.CheckedChanged, chkBoldCodeSortDetailByCoding.CheckedChanged
        Dim checked As Boolean
        If sender Is Me.chkBoldCodeDetailErrorsOnly Then
            checked = Me.chkBoldCodeDetailErrorsOnly.Checked
            If checked Then
                Me.chkBoldCodeSortDetailByCoding.Enabled = True
            Else
                Me.chkBoldCodeSortDetailByCoding.Enabled = False
                Me.chkBoldCodeSortDetailByCoding.Checked = False
            End If
        End If
    End Sub

    Private Sub chkHandleUseDateNumber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseDate.CheckedChanged, chkUseNumber.CheckedChanged
        Dim checked As Boolean
        Me.grpDateRange.Visible = False
        Me.grpNumberRange.Visible = False
        If sender Is Me.chkUseDate Then
            checked = Me.chkUseDate.Checked
            If Not checked Then Exit Sub
            Me.chkUseNumber.Checked = False
            If checked Then Me.grpDateRange.Visible = True
        End If
        If sender Is Me.chkUseNumber Then
            checked = Me.chkUseNumber.Checked
            If Not checked Then Exit Sub
            Me.chkUseDate.Checked = False
            Me.grpNumberRange.Visible = True
            Me.txtBeginningNumber.Focus()
        End If
        Application.DoEvents()
        Call SetRegistryEntries()
    End Sub

    Private Sub chkRevenuesSearchReceived_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRevenuesSearchReceived.CheckedChanged
        'receipt register search by received from field;
        Dim checked As Boolean
        checked = Me.chkRevenuesSearchReceived.Checked
        If checked Then
            Me.txtRevenueSearch.Visible = True
            Me.txtRevenueSearch.Focus()
        Else
            Me.txtRevenueSearch.Visible = False
        End If
    End Sub

    Private Sub chkVendorsSelectSingleVendor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVendorsSelectSingleVendor.CheckedChanged
        Try
            Dim checked As Boolean = Me.chkVendorsSelectSingleVendor.Checked
            Me.cboVendors.Visible = checked
            If checked Then
                Me.chk1099ByEmployee.Checked = False
                Me.chk1099ByEmployee.Enabled = False
            Else
                Me.chk1099ByEmployee.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkVendorsUseFiscalYear_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkVendorsUseFiscalYear.CheckedChanged
        Try
            Dim checked As Boolean = Me.chkVendorsUseFiscalYear.Checked
            If checked Then
                'disable calendar year
                Me.cboCalendarYears.Enabled = False
            Else
                Me.cboCalendarYears.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "  Class Description "

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Fred 2005.08.17
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'This form will eventually replace the current AF_Reporting.frmMainMenuReports
    'interface.  The current interface has too many controls, object creation
    'is global when it's not needed (lots of overhead).  Also, most of the reports
    'are using the C1 render tables, which have some major performance issues and
    'most reports (even small reports 20-40 pages) are very time and resource
    'consuming.  These reports will be rewritten and then added to this form.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#End Region

#Region "  Class Members "

    Private AccountNumber As String
    Private AccountName As String
    Private AppliedDate As Date
    Private BankAccountNumber As String
    Private CheckBeg As String = ""
    Private CheckEnd As String = ""
    Private ConnectionString As String
    Private CurrentMonthBeginning As Date
    Private CurrentMonthEnding As Date
    Private CurrentMonthString As String
    Private FiscalYear As Int32
    Private FiscalYearSelected As Int32
    Private IsVendorLoaded As Boolean = False
    Private Const MSGTITLE As String = "Activity Fund Reporting"
    Private NextAdjustmentNumber As String
    Private NextCheckNumber As String
    Private NextPONumber As String
    Private NextReceiptNumber As String
    Private NextTransferNumber As String
    Private SchoolNumber As String
    Private SiteNumber As String
    Private SubaccountNumber As String
    Private SubaccountName As String
    Private UseDate As Int32
    Private _lastdepositnumber As String
    '
    Private Administrator As Boolean
    Private UAccess As Int32
    '
    Private cn As SqlConnection
    'bold Validation Members:
    'dims;
    Public _editchk As Boolean
    Private _year As String
    Private _fund As String
    Private _project As String
    Private _function As String
    Private _source As String
    Private _object As String
    Private _program As String
    Private _subject As String
    Private _job As String
    Private _site As String
    Private _code As String

#End Region

#Region "  Combobox Events "

    Private Sub cboBanks_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBanks.SelectedIndexChanged
        With Me.cboBanks
            Try
                'get the bank account number;
                Dim selitem As System.Data.DataRowView
                Dim index As Int32
                index = .SelectedIndex
                If index < 0 Then Exit Sub
                selitem = CType(.Items.Item(index), DataRowView)
                Me.BankAccountNumber = DirectCast(selitem.Item(0), String)
                Me.NextCheckNumber = DirectCast(selitem.Item(7), String)
                Me.NextReceiptNumber = DirectCast(selitem.Item(8), String)
                Me.StatusBar1.Panels(2).Text = Me.BankAccountNumber
            Catch ex As Exception
                Throw
            End Try
        End With
    End Sub

    Private Sub cboFiscalYears_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFiscalYears.SelectedIndexChanged
        Dim selected As Int32 = CInt(Me.cboFiscalYears.SelectedItem)
        Me.txtDim1.Text = selected.ToString.Substring(3)
        Call GetCheckNumbers()
    End Sub

#End Region

#Region "  Delegates & Events "

    Public Delegate Sub RecordCountEvents()
    Public Event RecordStatus(ByVal erecords As Int32, ByVal erecordtotal As Int32)

    Private Sub UpdateStatus(ByVal records As Int32, ByVal recordstotal As Int32)
        'With Me.pBar1
        '    StatusBarPanel6.Text = "Processing Records : " & records.ToString
        '    .Visible = True
        '    .Maximum = recordstotal
        '    .Minimum = 0
        '    .Step = 1
        '    .PerformStep()
        'End With
    End Sub

#End Region

#Region "  Groupbox Events "

    Private Sub grpRevenue_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grpRevenue.VisibleChanged
        If Not Me.grpRevenue.Visible Then Exit Sub
        If Me.rdoRevenueReceiptTicket.Checked Then Me.panelOptions.Enabled = True : Me.chkUseNumber.Checked = True
    End Sub

#End Region

#Region "  Methods "

    Private Sub InitForm()


        'locate the groups
        Me.grpAccounts.Location = New Point(144, 80)
        Me.grpAdjustments.Location = Me.grpAccounts.Location
        Me.grpBoldCode.Location = Me.grpAccounts.Location
        Me.grpClassification.Location = Me.grpAccounts.Location
        Me.grpExpenditures.Location = Me.grpAccounts.Location
        Me.grpFinancials.Location = Me.grpAccounts.Location
        Me.grpReconciliation.Location = Me.grpAccounts.Location
        Me.grpRevenue.Location = Me.grpAccounts.Location
        Me.grpVendors.Location = Me.grpAccounts.Location
        'turn off options
        Me.chkUseNumber.Checked = False
        Me.chkUseDate.Checked = False
        Me.grpDateRange.Visible = False
        Me.grpNumberRange.Visible = False
        'choose the account button first
        Me.rdoAccounts.Checked = True
        'hide all but the first group
        Call ShowGroupBoxes(True, False, False, False, False, False, False, False, False)
        'assign default values to the form controls
        Me.dtBeginningDate.Value = Me.CurrentMonthBeginning
        Me.dtEndingDate.Value = Me.CurrentMonthEnding
        'load months into month combo-box on the accounts tab
        With Me.cboAccountsMonthListing
            .Items.Clear()
            .Items.Add("January")
            .Items.Add("February")
            .Items.Add("March")
            .Items.Add("April")
            .Items.Add("May")
            .Items.Add("June")
            .Items.Add("July")
            .Items.Add("August")
            .Items.Add("September")
            .Items.Add("October")
            .Items.Add("November")
            .Items.Add("December")
        End With

        'Bold Validation default
        Me.txtC1Year.Value = CInt(Me.FiscalYear.ToString.Substring(3, 1))
        Me.txtC1YearExp.Value = CInt(Me.FiscalYear.ToString.Substring(3, 1))

    End Sub

    Private Sub SetCtrlBank(ByVal ebankaccountnumber As String)
        Try
            Dim index As Int32
            Dim s As String
            Dim item As System.Data.DataRowView
            With Me.cboBanks
                For index = 0 To .Items.Count - 1
                    item = CType(.Items.Item(index), DataRowView)
                    s = CStr(item(0))
                    If ebankaccountnumber = s Then .SelectedIndex = index
                Next
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
        End Try
    End Sub

    Private Sub ShowGroupBoxes(ByVal eAccounts As Boolean, ByVal eAdjustments As Boolean, ByVal eBoldCode As Boolean, ByVal eClassification As Boolean, ByVal eExpenditures As Boolean, ByVal eFinancials As Boolean, ByVal eReconciliation As Boolean, ByVal eRevenues As Boolean, ByVal eVendors As Boolean)
        Me.grpAccounts.Visible = eAccounts
        Me.grpAdjustments.Visible = eAdjustments
        Me.grpBoldCode.Visible = eBoldCode
        Me.grpClassification.Visible = eClassification
        Me.grpExpenditures.Visible = eExpenditures
        Me.grpFinancials.Visible = eFinancials
        Me.grpReconciliation.Visible = eReconciliation
        Me.grpRevenue.Visible = eRevenues
        Me.grpVendors.Visible = eVendors
    End Sub

#End Region

#Region "  Methods Logging "

    '''''Private Sub WriteLogEntry(ByVal methodname As String, ByVal message As String)
    '''''    Dim path As String = "C:\Fred\fredlog.txt"
    '''''    Dim fi As New FileStream(path, FileMode.Append, FileAccess.Write)
    '''''    Dim sw As New StreamWriter(fi)
    '''''    Try
    '''''        Dim line1 As String
    '''''        line1 = Now.ToString.PadRight(25)
    '''''        line1 += methodname.ToLower.PadRight(35)
    '''''        line1 += message.PadRight(50)
    '''''        sw.WriteLine(line1)
    '''''    Catch ex As Exception
    '''''        Throw
    '''''    Finally
    '''''        sw.Close()
    '''''        fi.Close()
    '''''    End Try
    '''''End Sub

#End Region

#Region "  Methods Reports "

    Private Sub DoAccounts()
        'check if a selection has been checked;
        Select Case True
            Case Me.rdoAccountsBalanceSheet.Checked
            Case Me.rdoAccountsChartOfAccounts.Checked
            Case Me.rdoAccountsMTDSummaryOfAccounts.Checked, Me.rdoAccountsYTDSummaryOfAccounts.Checked
                If Me.FiscalYear <> Me.FiscalYearSelected Then
                    MsgBox("This report is only available for current fiscal year.  If you would like to run a prior year, please select the Historical Report.", MsgBoxStyle.Information, MSGTITLE)
                    Me.cboFiscalYears.Focus()
                    Exit Sub
                End If
            Case Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Checked
            Case Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Checked
            Case Else
                MsgBox("Please select a report to preview.", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Note:  Added zero amount suppression to the MTD & YTD Summary on 2009.08.26, requested by Joan;
        '       Also, the zero suppression is only valid for all accounts; partial account reports will
        '       not zero suppress; - Fred
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''balance sheet replaced by statement of change (2011.02.01, Fred);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'do statement of change;
        If Me.rdoAccountsBalanceSheet.Checked Then
            'If Me.FiscalYearSelected <> Me.FiscalYear Then
            'MsgBox("This report is only available for the current fiscal year.", MsgBoxStyle.Exclamation, MSGTITLE)
            'Me.cboFiscalYears.Focus()
            'Exit Sub
            'End If

            Dim obj As New AF_Reporting.frmAccountsReports
            If Me.FiscalYearSelected <> Me.FiscalYear Then
                Try
                    obj.GenerateStatementOfChangePriorYear(Me.FiscalYearSelected, Me.BankAccountNumber, True, True)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Else
                Try
                    obj.GenerateStatementOfChange(Me.FiscalYearSelected, Me.BankAccountNumber, True, True)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            End If
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''balance sheet removed on 2011.02.01, will restore if being used;''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''do balance sheet;
        '''''If Me.rdoAccountsBalanceSheet.Checked Then
        '''''    Dim obj As New AF_Reporting.frmAccountsReports
        '''''    Try
        '''''        obj.GenerateBalanceSheet(Me.BankAccountNumber, Me.chkAccountsIncludeSubaccounts.Checked)
        '''''    Catch ex As Exception
        '''''        MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
        '''''    Finally
        '''''        obj.Dispose()
        '''''    End Try
        '''''End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Chart of Accounts;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsChartOfAccounts.Checked Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateChartOfAccounts(Me.BankAccountNumber, Me.chkAccountsIncludeSubaccounts.Checked)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MTD Summary of Accounts (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsMTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = False And _
        Me.chkAccountsIncludeEncumbrances.Checked = False Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateMTDSummaryOfAccounts(Me.BankAccountNumber, Me.chkAccountsSuppressZeros.Checked, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MTD Summary of Accounts with Encumbrance (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsMTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeEncumbrances.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = False Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateMTDSummaryOfAccountsWithEncumbrance(Me.BankAccountNumber, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MTD Summary of Subaccounts (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsMTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = True And _
        Me.chkAccountsIncludeEncumbrances.Checked = False Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateMTDSummaryOfSubaccounts(Me.BankAccountNumber, Me.chkAccountsSuppressZeros.Checked, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MTD Summary of Subaccounts with Encumbrance (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsMTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = True And _
        Me.chkAccountsIncludeEncumbrances.Checked = True Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateMTDSummaryOfSubaccountsWithEncumbrance(Me.BankAccountNumber, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If




        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'YTD Summary of Accounts (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsYTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = False And _
        Me.chkAccountsIncludeEncumbrances.Checked = False Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateYTDSummaryOfAccounts(Me.BankAccountNumber, Me.chkAccountsSuppressZeros.Checked, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'YTD Summary of Accounts with Encumbrance (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsYTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = False And _
        Me.chkAccountsIncludeEncumbrances.Checked = True Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateYTDSummaryOfAccountsWithEncumbrance(Me.BankAccountNumber, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'YTD Summary of Subaccounts (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsYTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = True And _
        Me.chkAccountsIncludeEncumbrances.Checked = False Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateYTDSummaryOfSubaccounts(Me.BankAccountNumber, Me.chkAccountsSuppressZeros.Checked, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'YTD Summary of Subaccounts with Encumbrance (All or Range);
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Me.rdoAccountsYTDSummaryOfAccounts.Checked = True And _
        Me.chkAccountsIncludeSubaccounts.Checked = True And _
        Me.chkAccountsIncludeEncumbrances.Checked = True Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                obj.GenerateYTDSummaryOfSubaccountsWithEncumbrance(Me.BankAccountNumber, Me.chkAccountsUseAccountRange.Checked, Me.txtAccountsAcctNumberFrom.Text, Me.txtAccountsAcctNumberTo.Text)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'do historical mtd summary of accounts;
        If Me.rdoAccountsHistoricalMTDSummaryOfAccounts.Checked Then
            Dim selmonth As String = CType(Me.cboAccountsMonthListing.SelectedItem, String)
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                'check if a month has been selected
                If selmonth Is Nothing OrElse selmonth.Trim.Length = 0 Then
                    MsgBox("No reporting month has been selected for this report.", MsgBoxStyle.Information, MSGTITLE)
                    Me.cboAccountsMonthListing.Focus()
                    Exit Sub
                End If
                If Me.chkAccountsIncludeSubaccounts.Checked Then
                    obj.GenerateHistoricalMTDSummaryOfSubaccounts(Me.BankAccountNumber, Me.FiscalYearSelected, selmonth)
                Else
                    obj.GenerateHistoricalMTDSummaryOfAccounts(Me.BankAccountNumber, Me.FiscalYearSelected, selmonth)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'do historical ytd summary of accounts;
        If Me.rdoAccountsHistoricalYTDSummaryOfAccounts.Checked Then
            Dim obj As New AF_Reporting.frmAccountsReports
            Try
                If Me.chkAccountsIncludeSubaccounts.Checked Then
                    obj.GenerateHistoricalYTDSummaryOfSubaccounts(Me.BankAccountNumber, Me.FiscalYearSelected)
                Else
                    obj.GenerateHistoricalYTDSummaryOfAccounts(Me.BankAccountNumber, Me.FiscalYearSelected)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If
    End Sub

    Private Sub DoAdjustments()
        Dim usedate, usenumber As Boolean
        usedate = Me.chkUseDate.Checked
        usenumber = Me.chkUseNumber.Checked
        If usenumber = False And usedate = False Then
            Me.chkUseDate.Checked = True
            usedate = True
            Application.DoEvents()
        End If

        Select Case True
            Case Me.rdoAdjustmentsPrintAdjustmentRegister.Checked
                'adjustment register
                Dim obj As New AF_Reporting.frmFinancialReports
                Try
                    obj.GenerateAdjustmentRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Case Me.rdoAdjustmentsPrintTransferRegister.Checked
                'transfer register
                Dim obj As New AF_Reporting.frmFinancialReports
                Try
                    obj.GenerateTransferRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Case Else
                MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select
    End Sub

    Private Sub DoBoldCode()
        'only for edit checking 4.13.2017, Mark;
        _editchk = False
        'flag for error report
        Dim chkerronly As Boolean = Me.chkBoldCodeDetailErrorsOnly.Checked
        'flag for sort
        Dim sortbycoding As Boolean = Me.chkBoldCodeSortDetailByCoding.Checked

        Application.DoEvents()

        Select Case True
            Case Me.rdoBoldCodeListingByExpenditures.Checked
                'bold code expenditures
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    obj.GenerateBoldCodeExpenditure(chkerronly, sortbycoding, Me.FiscalYearSelected)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Case Me.rdoBoldCodeListingByRevenues.Checked
                'bold code revenues
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    Application.DoEvents()
                    obj.GenerateBoldCodeRevenue(chkerronly, sortbycoding, Me.FiscalYearSelected)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Case Else
                MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select
    End Sub

    Private Sub DoClassification()
        Application.DoEvents()
        Dim mtdbegdate, mtdenddate, ytdbegdate, ytdenddate As Date
        Dim classfilter As String
        Dim reporttype As Int32

        Try
            'get the date vars;
            mtdbegdate = Me.CurrentMonthBeginning
            mtdenddate = Me.CurrentMonthEnding
            'ytdbegdate = CDate("07/01/" & Me.FiscalYear - 1)
            ytdbegdate = CDate("07/01/" & Me.FiscalYearSelected - 1)
            ytdenddate = ytdbegdate.AddYears(1)
            ytdenddate = ytdenddate.AddMilliseconds(-1)
            'get the filter
            Select Case True
                Case Me.rdoClassificationMTDExpend.Checked
                    reporttype = 1
                    classfilter = Me.txtDim1.Text & Me.txtDim2.Text & Me.txtDim3.Text & Me.txtDim4.Text & Me.txtDim5.Text & Me.txtDim6.Text & Me.txtDim7.Text & Me.txtDim8.Text & Me.txtDim9.Text
                Case Me.rdoClassificationYTDExpend.Checked
                    reporttype = 2
                    classfilter = Me.txtDim1.Text & Me.txtDim2.Text & Me.txtDim3.Text & Me.txtDim4.Text & Me.txtDim5.Text & Me.txtDim6.Text & Me.txtDim7.Text & Me.txtDim8.Text & Me.txtDim9.Text
                Case Me.rdoClassificationMTDRevenue.Checked
                    reporttype = 1
                    classfilter = Me.txtDim1.Text & Me.txtDim2.Text & Me.txtDim3.Text & Me.txtDim4.Text & Me.txtDim5.Text & Me.txtDim6.Text
                Case Me.rdoClassificationYTDRevenue.Checked
                    reporttype = 2
                    classfilter = Me.txtDim1.Text & Me.txtDim2.Text & Me.txtDim3.Text & Me.txtDim4.Text & Me.txtDim5.Text & Me.txtDim6.Text
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
            Exit Sub
        End Try


        Select Case True
            Case Me.rdoClassificationMTDExpend.Checked, Me.rdoClassificationYTDExpend.Checked
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    Application.DoEvents()
                    If reporttype = 1 Then
                        'mtd report;
                        obj.GenerateClassificationExpenditureDetailListing(Me.BankAccountNumber, Me.FiscalYearSelected, mtdbegdate, mtdenddate, reporttype, Me.chkClassificationUseCodeRange.Checked, classfilter)
                    End If

                    If reporttype = 2 Then
                        'ytd report;
                        obj.GenerateClassificationExpenditureDetailListing(Me.BankAccountNumber, Me.FiscalYearSelected, ytdbegdate, ytdenddate, reporttype, Me.chkClassificationUseCodeRange.Checked, classfilter)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoClassificationMTDRevenue.Checked, Me.rdoClassificationYTDRevenue.Checked
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    Application.DoEvents()
                    If reporttype = 1 Then
                        'mtd report;
                        obj.GenerateClassificationRevenueDetailListing(Me.BankAccountNumber, Me.FiscalYearSelected, mtdbegdate, mtdenddate, reporttype, Me.chkClassificationUseCodeRange.Checked, classfilter)
                    End If

                    If reporttype = 2 Then
                        'ytd report;
                        obj.GenerateClassificationRevenueDetailListing(Me.BankAccountNumber, Me.FiscalYearSelected, ytdbegdate, ytdenddate, reporttype, Me.chkClassificationUseCodeRange.Checked, classfilter)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoClassificationExpenditureCodes.Checked
                'list all expenditure codes for the selected fiscal year;
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    'allow for valid codes only;
                    obj.GenerateExpenditureCodes(Me.FiscalYearSelected, True)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoClassificationRevenueCodes.Checked
                'list all revenue codes for the selected fiscal year;
                Dim obj As New AF_Reporting.frmBoldCode
                Try
                    'allow for valid codes only;
                    obj.GenerateRevenueCodes(Me.FiscalYearSelected, True)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Else
                MsgBox("Please select a report to preview.", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select

    End Sub

    Private Sub DoExpenditures()
        Dim usedate, usenumber As Boolean

        usedate = Me.chkUseDate.Checked
        usenumber = Me.chkUseNumber.Checked

        Select Case True
            Case Me.rdoExpendituresPrintOutstandingChecks.Checked, Me.rdoExpendituresEncumbranceAccounts.Checked, Me.rdoExpendituresPendingInvoice.Checked()
            Case Else
                If usenumber = False And usedate = False Then
                    Me.chkUseDate.Checked = True
                    usedate = True
                    Application.DoEvents()
                End If
        End Select

        Select Case True
            Case Me.rdoExpendituresCheckRegister.Checked
                'check register
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GenerateCheckRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresPoRegister.Checked
                'purchase order register
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GeneratePurchaseOrderRegisterByInvoice(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text, Me.chkExpendituresOutstandingInvoices.Checked)
                    'obj.GeneratePurchaseOrderRegisterByAccount(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text, Me.chkExpendituresOutstandingInvoices.Checked)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresPrintPurchaseOrder.Checked
                'print purchase order;
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    If usedate Then
                        obj.GeneratePurchaseOrder(Me.BankAccountNumber, Me.FiscalYearSelected, Me.dtBeginningDate.Value, Me.dtEndingDate.Value)
                    End If
                    If usenumber Then
                        obj.GeneratePurchaseOrder(Me.BankAccountNumber, Me.FiscalYearSelected, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresPrintVoidChecks.Checked
                'print void check register
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GenerateVoidCheckRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresPrintOutstandingChecks.Checked
                'print outstanding check register;
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GeneratechecksOutstanding(Me.BankAccountNumber, Me.FiscalYearSelected, Me.chkExpendituresAllFiscalYears.Checked, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresEncumbranceAccounts.Checked()
                'print encumbrance outstanding report;
                Dim obj As New AF_Reporting.frmAccountsReports
                Try
                    obj.GenerateExpenditureYTDEncumbranceBalances(Me.BankAccountNumber, Me.FiscalYearSelected)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

                'this is a test for kick back doc
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Case Me.rdoExpendituresPendingInvoice.Checked
                '    Dim obj As New AF_Reporting.frmExpenditureReports
                '    Try
                '        obj.PrintPurchaseOrderKick()
                '    Catch ex As Exception
                '        MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                '    Finally
                '        obj.Dispose()
                '   End Try
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Case Me.rdoExpendituresPendingInvoice.Checked
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GenerateInvoicePending(Me.BankAccountNumber, Me.FiscalYear)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoExpendituresPositivePay.Checked
                Dim obj As New AF_Reporting.frmExpenditureReports
                Try
                    obj.GeneratePositivePayFile(Me.BankAccountNumber, Me.FiscalYearSelected, CType(Me.txtBeginningNumber.Text, Int32), CType(Me.txtEndingNumber.Text, Int32))
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Else
                'No report selected;
                MsgBox("Please select a report to preview.", MsgBoxStyle.Information, MSGTITLE)
        End Select
    End Sub

    Private Sub DoFinancials()
        Dim IsAll As Boolean = Me.rdoFinancialsAllAccounts.Checked

        'check if a selection has been checked;
        Select Case True
            Case Me.rdoFinancialsDetailOfAccountMTD.Checked
            Case Me.rdoFinancialsDetailOfAccountYTD.Checked
            Case Me.rdoFinancialsDetailOfAccountPeriodical.Checked
            Case Me.rdoFinancialsNoMtdDetail.Checked()
            Case Me.rdoFinancialsNoYtdDetail.Checked()
            Case Else
                MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select

        'mtd detail;
        If Me.rdoFinancialsDetailOfAccountMTD.Checked And Not (Me.chkFinancialsEncumbranceDetail.Checked) Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                If IsAll Then
                    obj.GenerateDetailOfAccountsMTDAllAccounts(Me.BankAccountNumber, Me.FiscalYearSelected, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.CurrentMonthString, 2)
                Else
                    obj.GenerateDetailOfAccountsMTDSingleAccount(Me.BankAccountNumber, Me.FiscalYearSelected, Me.AccountNumber, Me.SubaccountNumber, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.CurrentMonthString, 2)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'ytd detail;
        If Me.rdoFinancialsDetailOfAccountYTD.Checked And Not (Me.chkFinancialsEncumbranceDetail.Checked) Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                If IsAll Then
                    obj.GenerateDetailOfAccountsYTDAllAccounts(Me.BankAccountNumber, Me.FiscalYearSelected)
                Else
                    obj.GenerateDetailOfAccountsYTDSingleAccount(Me.BankAccountNumber, Me.FiscalYearSelected, Me.AccountNumber, Me.SubaccountNumber)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'periodical detail;
        If Me.rdoFinancialsDetailOfAccountPeriodical.Checked Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                If IsAll Then
                    obj.GenerateDetailOfAccountsMTDAllAccounts(Me.BankAccountNumber, Me.FiscalYearSelected, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, "Periodical", 3)
                Else
                    obj.GenerateDetailOfAccountsMTDSingleAccount(Me.BankAccountNumber, Me.FiscalYearSelected, Me.AccountNumber, Me.SubaccountNumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, "Periodical", 3)
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'no activity mtd;
        If Me.rdoFinancialsNoMtdDetail.Checked Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                obj.GenerateNoMTDDetailOfAccounts(Me.BankAccountNumber, Me.FiscalYear, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.CurrentMonthString)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'no activity ytd;
        If Me.rdoFinancialsNoYtdDetail.Checked Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                obj.GenerateNoYTDDetailOfAccounts(Me.BankAccountNumber, Me.FiscalYear)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        'encumbrance detail;
        If Me.chkFinancialsEncumbranceDetail.Checked Then
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                If IsAll Then
                    If Me.rdoFinancialsDetailOfAccountMTD.Checked Then
                        obj.GenerateEncumbranceDetailOfAccountsMTDAllAccounts(Me.BankAccountNumber, Me.FiscalYearSelected, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.CurrentMonthString, 2)
                    End If
                    If Me.rdoFinancialsDetailOfAccountYTD.Checked Then
                        obj.GenerateEncumbranceDetailOfAccountsYTDAllAccounts(Me.BankAccountNumber, Me.FiscalYearSelected)
                    End If
                Else
                    If Me.rdoFinancialsDetailOfAccountMTD.Checked Then
                        obj.GenerateEncumbranceDetailOfAccountsMTDSingleAccount(Me.BankAccountNumber, Me.FiscalYearSelected, Me.AccountNumber, Me.SubaccountNumber, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.CurrentMonthString, 2)
                    End If
                    If Me.rdoFinancialsDetailOfAccountYTD.Checked Then
                        obj.GenerateEncumbranceDetailOfAccountsYTDSingleAccount(Me.BankAccountNumber, Me.FiscalYearSelected, Me.AccountNumber, Me.SubaccountNumber)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

    End Sub

    Private Sub DoReconciliation()
        If Me.rdoReconTrialBalance.Checked Then
            'collect the currency values
            Dim balance, interest, expenditure, investment As Double
            Try
                Me.Cursor = Cursors.WaitCursor
                balance = CDbl(Me.txtReconBankStatementBalance.Text)
                interest = CDbl(Me.txtReconInterestNotYetPosted.Text)
                expenditure = CDbl(Me.txtReconExpensesNotYetPosted.Text)
                investment = CDbl(Me.txtReconInvestments.Text)
                Application.DoEvents()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End Try

            'use the new recon
            Dim obj As New AF_Reporting.frmFinancialReports
            Try
                obj.GenerateTrialBalance(Me.BankAccountNumber, Me.FiscalYear, Me.CurrentMonthBeginning, Me.CurrentMonthEnding, Me.dtReconTrialBalanceDate.Value, balance, interest, expenditure, investment, True)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
        End If

        ''''''use the old recon (available for testing only);
        '''''Dim obj As New AF_Reporting.frmFinancials
        '''''Try
        '''''    If dodetail Then
        '''''        obj.PrintFinStTrialBalanceDetailed(Me.BankAccountNumber, Me.FiscalYear, balance, interest, expenditure, investment, Me.dtReconTrialBalanceDate.Value)
        '''''    Else
        '''''        obj.PrintFinStTrialBalance(Me.BankAccountNumber, Me.FiscalYear, balance, interest, expenditure, investment, Me.dtReconTrialBalanceDate.Value)
        '''''    End If
        '''''Catch ex As Exception
        '''''    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
        '''''Finally
        '''''    obj.Dispose()
        '''''End Try

    End Sub

    Private Sub DoRevenues()
        Dim usedate, usenumber As Boolean
        'check if a selection has been checked;
        Select Case True
            Case Me.rdoRevenueDailyDeposit.Checked
                'daily deposit report;
                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateDeposit(Me.BankAccountNumber, Me.FiscalYearSelected, Me.txtDepositBegNumber.Text.Trim, Me.chkRevenuesIncludeCreditCards.Checked, Me.chkRevenuesPrintDepositTicket.Checked)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenueDepositSummary.Checked
                'deposit summary;
                Dim begnum, endnum As String
                Try
                    begnum = Me.txtDepositBegNumber.Text.Trim
                    endnum = Me.txtDepositEndNumber.Text.Trim
                    'verify that numbers are available;
                    If begnum Is Nothing OrElse begnum.Length = 0 Then
                        MsgBox("No beginning number is available.", MsgBoxStyle.Information, MSGTITLE)
                        Exit Sub
                    End If
                    If endnum Is Nothing OrElse endnum.Length = 0 Then
                        MsgBox("No ending number is available.", MsgBoxStyle.Information, MSGTITLE)
                        Exit Sub
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                    Exit Sub
                End Try

                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateDepositSummary(Me.BankAccountNumber, Me.FiscalYearSelected, begnum, endnum)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenueReceiptRegister.Checked
                'receipt register;
                Dim usesearch As Boolean
                Dim searchstring As String
                usesearch = Me.chkRevenuesSearchReceived.Checked
                searchstring = Me.txtRevenueSearch.Text.Trim
                If usesearch And searchstring.Length = 0 Then
                    MsgBox("You must enter a search value to continue.", MsgBoxStyle.Information, MSGTITLE)
                    Me.txtRevenueSearch.Focus()
                    Exit Sub
                End If

                usedate = Me.chkUseDate.Checked
                usenumber = Me.chkUseNumber.Checked
                If usenumber = False And usedate = False Then
                    Me.chkUseDate.Checked = True
                    usedate = True
                    Application.DoEvents()
                End If
                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateReceiptRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text, usesearch, searchstring)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenueReceiptTicket.Checked
                'receipt ticket;
                usedate = Me.chkUseDate.Checked
                usenumber = Me.chkUseNumber.Checked
                If usenumber = False And usedate = False Then
                    Me.chkUseDate.Checked = True
                    usedate = True
                    Application.DoEvents()
                End If

                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateReceiptTickets(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenuePrintOutstandingReceipts.Checked
                'use all fiscal years;
                usedate = Me.chkRevenuesAllFiscalYears.Checked
                'print outstanding check register;
                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateOutstandingReceiptsRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenuePrintVoidReceipts.Checked
                'void receipt register;
                usedate = Me.chkUseDate.Checked
                usenumber = Me.chkUseNumber.Checked
                If usenumber = False And usedate = False Then
                    Me.chkUseDate.Checked = True
                    usedate = True
                    Application.DoEvents()
                End If
                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    obj.GenerateVoidReceiptRegister(Me.BankAccountNumber, Me.FiscalYearSelected, usedate, usenumber, Me.dtBeginningDate.Value, Me.dtEndingDate.Value, Me.txtBeginningNumber.Text, Me.txtEndingNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoRevenue1098T.Checked
                '1098-T report;
                Dim obj As New AF_Reporting.frmRevenueReports
                Try
                    'note: this is a calendar based report;
                    obj.Generate1098TReport(Me.FiscalYearSelected, Me.txtRevenueAccountNumber.Text, Me.txtRevenueSubaccountNumber.Text)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
            Case Else
                MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select
    End Sub

    Private Sub DoVendors()
        Dim selyear As Int32

        Try
            selyear = CInt(Me.cboCalendarYears.Text)
            If Me.chkVendorsUseFiscalYear.Checked Then selyear = Me.FiscalYearSelected
        Catch ex As Exception
            Application.DoEvents()
            Me.cboCalendarYears.SelectedIndex = 0
            selyear = CInt(Me.cboCalendarYears.Text)
        End Try

        If Me.chk1099Summary.Checked = True Then
            Dim obj As New AF_Reporting.frmVendorReports
            Try
                obj.Generate1099Vendors(selyear)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
            Finally
                obj.Dispose()
            End Try
            Exit Sub
        End If

        Select Case True
            Case Me.rdoVendorListing.Checked
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'vendor listing;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim obj As New AF_Reporting.frmVendorReports
                Try
                    obj.GenerateVendorListing()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'vendor listing;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case Me.rdoVendorExpenses.Checked
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'vendor expense summary - 1099 summary report;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim obj As New AF_Reporting.frmVendorReports
                Try
                    obj.GenerateVendorExpenditureSummary(selyear, Me.chk1099ByEmployee.Checked, Me.chkVendorsIncludeZeroBalances.Checked, Me.chkVendorsUse600Minimum.Checked, Me.chkVendorsUseFiscalYear.Checked)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoVendor1099Listing.Checked
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '1099 reporting, by summary or detail;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim usesingle As Boolean = Me.chkVendorsSelectSingleVendor.Checked
                Dim index, key As Integer
                Try
                    If usesingle Then
                        Dim selecteditem As System.Data.DataRowView
                        With Me.cboVendors
                            index = Me.cboVendors.SelectedIndex
                            'convert combobox object row to a datarowview
                            selecteditem = CType(.Items.Item(index), DataRowView)
                            'extract the columns from the datarowview
                            key = CType(selecteditem.Item(0), Int32)
                        End With
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
                    Exit Sub
                End Try

                Dim obj As New AF_Reporting.frmVendorReports
                Try
                    If usesingle Then
                        'vendor expense 1099 detail report (single vendor)
                        obj.GenerateVendorExpenditureDetailSingle(key, selyear, Me.chkVendorsIncludeSSN.Checked, Me.chkVendorsUseFiscalYear.Checked)
                    Else
                        'vendor expense 1099 detail report (all vendors)
                        obj.GenerateVendorExpenditureDetail(selyear, chk1099Summary.Checked, Me.chkVendorsIncludeSSN.Checked, Me.chkVendorsUseFiscalYear.Checked)
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Me.rdoVendorAudit.Checked
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'vendor audit report ~ by fiscal year;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim obj As New AF_Reporting.frmVendorReports
                Try
                    obj.GenerateVendorAudit(Me.FiscalYearSelected)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
                Finally
                    obj.Dispose()
                End Try

            Case Else
                MsgBox("Please select a report to preview...", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
        End Select
    End Sub

#End Region

#Region "  Methods Retrieval "

    Private Sub GetAdjustNumbers()
        Dim SSQL As String
        SSQL = "SELECT ISNULL(MAX(tran_autoinc_key), 0) FROM transactions; "
        SSQL += "SELECT ISNULL(MAX(trx_autoinc_key), 0) FROM transfers; "
        SSQL += "SELECT MIN(af_acct_num), MAX(af_acct_num) FROM acct_info"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("keys")
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try
        Dim trankey, trxkey As Int32
        With ds
            trankey = CInt(.Tables(0).Rows(0).Item(0))
            trxkey = CInt(.Tables(1).Rows(0).Item(0))
            Me.NextAdjustmentNumber = trankey.ToString.Format("{0:D8}", trankey)
            Me.NextTransferNumber = trxkey.ToString.Format("{0:D8}", trxkey)
            Me.txtAccountsAcctNumberFrom.Text = CStr(.Tables(2).Rows(0).Item(0))
            Me.txtAccountsAcctNumberTo.Text = CStr(.Tables(2).Rows(0).Item(1))
        End With
    End Sub

    Private Sub GetBanks()
        ''''''''''''''''''''''''''''''''' GridBank '''''''''''''''''''''''''''''''''''''
        '       0       1      2       3       4        5        6       7        8  
        '    number   name  status  begbal  netbal   curbal   assets  nextchk  nextrcpt
        '       9      10     11      12      13       14       15      16
        '    addr1   addr2   addr3   city    state     zip    zipext  phone1 
        '      17      18     19      20      21       22       23      24
        '   ph1ext  phone2  ph2ext   fax   contact1 contact2  cust1   cust2
        '      25     26      27  
        '    site  sitename  key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method collects all bank information for the activity fund system; 
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim index As Int32
        Try
            'show all banks;
            SSQL = "SELECT bank_acct_num, bank_name, bank_status, bank_beg_balance," _
            & " bank_net_balance, bank_cur_balance, bank_frozen_assets," _
            & " bank_next_check, bank_next_receipt, bank_addr1, bank_addr2," _
            & " bank_addr3, bank_city, bank_state, bank_zip, bank_zip_ext," _
            & " bank_phone1, bank_phone1_ext, bank_phone2, bank_phone2_ext," _
            & " bank_fax, bank_contact1, bank_contact2, bank_custodian1," _
            & " bank_custodian2, site_num, site_name, bank_autoinc_key" _
            & " FROM bank_info" _
            & " WHERE bank_status = 'O' OR bank_status = 'I'" _
            & " ORDER BY bank_acct_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("bank_info")
            da.Fill(tbl)
            With Me.cboBanks
                .DataSource = tbl
                .DisplayMember = "bank_acct_num"
                .ValueMember = "bank_acct_num"
                'force an index change;
                .SelectedIndex = -1
                'Me.IsLoading = False
                .SelectedIndex = 0
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try
    End Sub

    Private Sub GetFiscalYears()
        Dim SSQL As String
        SSQL = "SELECT DISTINCT ocrv_fisyr FROM ocas_rev ORDER BY ocrv_fisyr"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("years")
        Try
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'load the fiscal years
            Dim row As DataRow
            For Each row In tbl.Rows
                Me.cboFiscalYears.Items.Add(row.Item(0))
            Next
            Me.cboFiscalYears.SelectedIndex = Me.cboFiscalYears.Items.Count - 1
        Catch ex As Exception

        End Try

        Try
            'load the calendar years
            Dim year As Int32
            For year = (Me.FiscalYear - 1) To Me.FiscalYear
                Me.cboCalendarYears.Items.Add(year)
            Next
            Me.cboCalendarYears.SelectedIndex = 0



        Catch ex As Exception

        End Try
    End Sub

    Private Sub GetCheckNumbers()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Collect the beginning and ending check numbers for the selected fiscal year;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Try
            'Ensure the fiscal year is available;
            Me.FiscalYearSelected = CInt(Me.cboFiscalYears.SelectedItem)
            If Me.FiscalYearSelected = 0 Then Me.FiscalYearSelected = Me.FiscalYear
            SSQL = "SELECT ISNULL(MIN(CHKS_NUM), 0), ISNULL(MAX(CHKS_NUM), 0) FROM chks_info WHERE bank_acct_num = @p1 AND chks_fisyr = @p2"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", Me.BankAccountNumber)
            cmd.Parameters.Add("@p2", Me.FiscalYearSelected)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("Series")
            da.Fill(tbl)
            cn.Close()
            Me.CheckBeg = CType(tbl.Rows(0).Item(0), String)
            Me.CheckEnd = CType(tbl.Rows(0).Item(1), String)
            If Me.rdoExpenditure.Checked = True And Me.rdoExpendituresPositivePay.Checked = True Then
                Me.txtBeginningNumber.Text = Me.CheckBeg
                Me.txtEndingNumber.Text = Me.CheckEnd
            End If
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
        End Try
    End Sub

    Private Sub GetVendors()
        'don't refresh if already loaded;
        If Me.IsVendorLoaded Then Exit Sub
        Dim SSQL As String
        SSQL = "SELECT vend_autoinc_key, vend_number, vend_name" _
        & " FROM vend_info" _
        & " WHERE (vend_number <> '00000')" _
        & " AND (vend_status <> 'D')" _
        & " AND (vend_name <> '')" _
        & " ORDER BY vend_name"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("vendors")
        Try
            Me.Cursor = Cursors.WaitCursor
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
            Me.Cursor = Cursors.Default
        End Try

        Try
            Me.Cursor = Cursors.WaitCursor
            If tbl.Rows.Count < 1 Then Exit Sub
            'bind the combobox to the tbl
            With Me.cboVendors
                .DisplayMember = "vend_name"
                .ValueMember = "vend_autoinc_key"
                .DataSource = tbl
                .SelectedIndex = 0
            End With
        Catch ex As Exception
            Throw
        Finally
            Me.IsVendorLoaded = True
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub SelectAccount()
        Dim obj As New AF_Accounts.SelectAccount(Me.BankAccountNumber)
        Try
            obj.ShowDialog()
            If obj.HasSelected Then
                Me.AccountNumber = obj.AccountNumber
                Me.AccountName = obj.AccountName
                Me.SubaccountNumber = obj.SubAccountNumber
                Me.SubaccountName = obj.SubAccountName
            Else
                Me.AccountNumber = ""
                Me.AccountName = ""
                Me.SubaccountNumber = ""
                Me.SubaccountName = ""
            End If
        Catch ex As Exception
            Throw
        Finally
            obj.Dispose()
        End Try
    End Sub

#End Region

#Region "  Properties "

    Private Property LastDepositNumber() As String
        Get
            Return _lastdepositnumber
        End Get
        Set(ByVal Value As String)
            Dim tempnum As Int32
            Try
                tempnum = CInt(Value)
                tempnum -= 1
                _lastdepositnumber = CStr(tempnum)
            Catch ex As Exception
                _lastdepositnumber = Value
            End Try
        End Set
    End Property

#End Region

#Region "  Radiobutton Events "

    Private Sub rdoFinancialsSelectAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoFinancialsSelectAccount.Click, rdoFinancialsAllAccounts.Click
        If sender Is Me.rdoFinancialsSelectAccount Then
            Try
                Call SelectAccount()
                If Me.AccountNumber.Length = 0 Then
                    Me.lblAccountNumber.Text = ""
                    Me.lblAccountName.Text = ""
                Else
                    Me.lblAccountNumber.Text = Me.AccountNumber & " - " & Me.SubaccountNumber
                    Me.lblAccountName.Text = Me.AccountName
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            End Try
        End If
        If sender Is Me.rdoFinancialsAllAccounts Then
            Me.lblAccountNumber.Text = ""
            Me.lblAccountName.Text = ""
        End If
    End Sub

    Private Sub rdoHandleAccounts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAccountsChartOfAccounts.Click, rdoAccountsMTDSummaryOfAccounts.Click, rdoAccountsYTDSummaryOfAccounts.Click, rdoAccountsHistoricalMTDSummaryOfAccounts.Click, rdoAccountsHistoricalYTDSummaryOfAccounts.Click, rdoAccountsBalanceSheet.Click
        'disable options by default;
        Me.panelAccountsSelectMonth.Visible = False
        Me.chkAccountsUseAccountRange.Enabled = False
        Me.chkAccountsIncludeEncumbrances.Enabled = False
        Me.chkAccountsSuppressZeros.Enabled = False
        'uncheck options by default;
        Me.chkAccountsUseAccountRange.Checked = False
        Me.chkAccountsIncludeEncumbrances.Checked = False
        Me.chkAccountsSuppressZeros.Checked = False
        If sender Is Me.rdoAccountsHistoricalMTDSummaryOfAccounts Then
            Me.panelAccountsSelectMonth.Visible = True
        End If
        If sender Is Me.rdoAccountsMTDSummaryOfAccounts Then
            Me.chkAccountsUseAccountRange.Enabled = True
            Me.chkAccountsIncludeEncumbrances.Enabled = True
            Me.chkAccountsSuppressZeros.Enabled = True
        End If
        If sender Is Me.rdoAccountsYTDSummaryOfAccounts Then
            'me.chkAccountsIncludeDecliningEncumbrances.Enabled = True     ---- was commented out
            Me.chkAccountsIncludeEncumbrances.Enabled = True
            Me.chkAccountsUseAccountRange.Enabled = True
            Me.chkAccountsSuppressZeros.Enabled = True
        End If
    End Sub

    Private Sub rdoHandleAdjustments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoAdjustmentsPrintAdjustmentRegister.Click, rdoAdjustmentsPrintTransferRegister.Click
        'set default settings for the adjustment tab
        If sender Is Me.rdoAdjustmentsPrintAdjustmentRegister Then
            Me.panelOptions.Enabled = True
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextAdjustmentNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        If sender Is Me.rdoAdjustmentsPrintTransferRegister Then
            Me.panelOptions.Enabled = True
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextTransferNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
    End Sub

    Private Sub rdoHandleBoldCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoBoldCodeListingByExpenditures.Click, rdoBoldCodeListingByRevenues.Click
        'set default settings for the bold code tab
        Me.panelOptions.Enabled = False
        Me.chkUseDate.Checked = False
        Me.chkUseNumber.Checked = False
        Me.chkBoldCodeDetailErrorsOnly.Enabled = False
        If sender Is Me.rdoBoldCodeListingByExpenditures Then

        End If
        If sender Is Me.rdoBoldCodeListingByRevenues Then

        End If
    End Sub

    Private Sub rdoHandleClassification_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoClassificationMTDExpend.Click, rdoClassificationMTDRevenue.Click, rdoClassificationYTDExpend.Click, rdoClassificationYTDRevenue.Click, rdoClassificationExpenditureCodes.Click, rdoClassificationRevenueCodes.Click
        If (sender Is Me.rdoClassificationMTDExpend) Or (sender Is Me.rdoClassificationYTDExpend) Then
            Me.txtDim7.Enabled = True
            Me.txtDim8.Enabled = True
            Me.txtDim9.Enabled = True
            Me.chkClassificationUseCodeRange.Enabled = True
            If Me.chkClassificationUseCodeRange.Checked Then Me.txtDim3.Focus()
        End If
        If (sender Is Me.rdoClassificationMTDRevenue) Or (sender Is Me.rdoClassificationYTDRevenue) Then
            Me.txtDim7.Enabled = False
            Me.txtDim8.Enabled = False
            Me.txtDim9.Enabled = False
            Me.txtDim7.Text = "****"
            Me.txtDim8.Text = "***"
            Me.txtDim9.Text = "***"
            Me.chkClassificationUseCodeRange.Enabled = True
            If Me.chkClassificationUseCodeRange.Checked Then Me.txtDim3.Focus()
        End If

        If (sender Is Me.rdoClassificationExpenditureCodes) Or (sender Is Me.rdoClassificationRevenueCodes) Then
            'disable & hide the range box;
            Me.chkClassificationUseCodeRange.Enabled = False
            Me.chkClassificationUseCodeRange.Checked = False
        End If
    End Sub

    Private Sub rdoHandleExpenditures_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoExpendituresCheckRegister.Click, rdoExpendituresPoRegister.Click, rdoExpendituresPrintPurchaseOrder.Click, rdoExpendituresPrintVoidChecks.Click, rdoExpendituresPrintOutstandingChecks.Click, rdoExpendituresEncumbranceAccounts.Click, rdoExpendituresPendingInvoice.Click, rdoExpendituresPositivePay.Click
        'set default settings for the expenditure tab;
        Me.chkExpendituresAllFiscalYears.Visible = False
        Me.chkExpendituresOutstandingInvoices.Visible = False
        Me.lblExpendituresPosPay.Visible = False
        '
        'Print check register;
        If sender Is Me.rdoExpendituresCheckRegister Then
            Me.panelOptions.Enabled = True
            Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextCheckNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        'Print po register;
        If sender Is Me.rdoExpendituresPoRegister Then
            Me.panelOptions.Enabled = True
            Me.chkExpendituresOutstandingInvoices.Visible = True
            Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextPONumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        'Print purchase order ticket;
        If sender Is Me.rdoExpendituresPrintPurchaseOrder Then
            Me.panelOptions.Enabled = True
            Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextPONumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        'Print outstanding checks report;
        If sender Is Me.rdoExpendituresPrintOutstandingChecks Then
            Me.panelOptions.Enabled = False
            Me.chkUseNumber.Checked = False
            Me.chkUseDate.Checked = False
            Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextCheckNumber
            Me.chkExpendituresAllFiscalYears.Visible = True
        End If
        'Print void check report;
        If sender Is Me.rdoExpendituresPrintVoidChecks Then
            Me.panelOptions.Enabled = True
            Me.txtBeginningNumber.Text = "00000001"
            Me.txtEndingNumber.Text = Me.NextCheckNumber
            If (Me.chkUseDate.Checked = False) And (Me.chkUseNumber.Checked = False) Then Me.chkUseDate.Checked = True
        End If
        If sender Is Me.rdoExpendituresEncumbranceAccounts Then
            Me.chkUseNumber.Checked = False
            Me.chkUseDate.Checked = False
            Me.panelOptions.Enabled = False
        End If
        If sender Is Me.rdoExpendituresPendingInvoice Then
            Me.chkUseNumber.Checked = False
            Me.chkUseDate.Checked = False
            Me.panelOptions.Enabled = False
        End If
        'Print positive pay file (BOK bank file);
        If sender Is Me.rdoExpendituresPositivePay Then
            'Refresh the register number collection;
            Call GetCheckNumbers()
            Me.chkUseNumber.Checked = True
            Me.lblExpendituresPosPay.Visible = True
            Me.txtBeginningNumber.Text = Me.CheckBeg
            Me.txtEndingNumber.Text = Me.CheckEnd
            Me.txtBeginningNumber.SelectAll()
            Me.txtBeginningNumber.Focus()
        End If
    End Sub

    Private Sub rdoHandleFinancials_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoFinancialsDetailOfAccountMTD.Click, rdoFinancialsDetailOfAccountYTD.Click, rdoFinancialsDetailOfAccountPeriodical.Click, rdoFinancialsNoMtdDetail.Click, rdoFinancialsNoYtdDetail.Click
        'set default settings for the financials tab
        Me.chkUseDate.Checked = False
        Me.panelFinancialsSelectAccount.Enabled = True
        Me.rdoFinancialsAllAccounts.Enabled = True
        Me.chkFinancialsEncumbranceDetail.Enabled = False
        Me.chkFinancialsEncumbranceDetail.Checked = False

        If sender Is Me.rdoFinancialsDetailOfAccountMTD Then
            Me.chkFinancialsEncumbranceDetail.Enabled = True
        End If
        If sender Is Me.rdoFinancialsDetailOfAccountYTD Then
            Me.chkFinancialsEncumbranceDetail.Enabled = True
        End If
        If sender Is Me.rdoFinancialsDetailOfAccountPeriodical Then
            Me.chkUseDate.Checked = True
        End If
        If sender Is Me.rdoFinancialsNoMtdDetail Then
            Me.panelFinancialsSelectAccount.Enabled = False
        End If
        If sender Is Me.rdoFinancialsNoYtdDetail Then
            Me.panelFinancialsSelectAccount.Enabled = False
        End If
    End Sub

    Private Sub rdoHandleReports_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoAccounts.Click, rdoAdjustments.Click, rdoBoldCode.Click, rdoClassification.Click, rdoExpenditure.Click, rdoFinancials.Click, rdoReconciliation.Click, rdoRevenue.Click, rdoVendors.Click
        'turn off options 
        Me.panelOptions.Enabled = False

        If sender Is Me.rdoAccounts Then
            ShowGroupBoxes(True, False, False, False, False, False, False, False, False)
        End If
        If sender Is Me.rdoAdjustments Then
            ShowGroupBoxes(False, True, False, False, False, False, False, False, False)
        End If
        If sender Is Me.rdoBoldCode Then
            ShowGroupBoxes(False, False, True, False, False, False, False, False, False)
        End If
        If sender Is Me.rdoClassification Then
            ShowGroupBoxes(False, False, False, True, False, False, False, False, False)
        End If
        If sender Is Me.rdoExpenditure Then
            ShowGroupBoxes(False, False, False, False, True, False, False, False, False)
        End If
        If sender Is Me.rdoFinancials Then
            ShowGroupBoxes(False, False, False, False, False, True, False, False, False)
        End If
        If sender Is Me.rdoReconciliation Then
            ShowGroupBoxes(False, False, False, False, False, False, True, False, False)
            Me.txtReconBankStatementBalance.Focus()
        End If
        If sender Is Me.rdoRevenue Then
            ShowGroupBoxes(False, False, False, False, False, False, False, True, False)
        End If
        If sender Is Me.rdoVendors Then
            ShowGroupBoxes(False, False, False, False, False, False, False, False, True)
            Application.DoEvents()
            Call GetVendors()
        End If
    End Sub

    Private Sub rdoHandleRevenue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoRevenueDailyDeposit.Click, rdoRevenueDepositSummary.Click, rdoRevenueReceiptRegister.Click, rdoRevenueReceiptTicket.Click, rdoRevenuePrintVoidReceipts.Click, rdoRevenuePrintOutstandingReceipts.Click, rdoRevenue1098T.Click
        'set default settings for the revenue tab;
        Me.txtDepositBegNumber.Text = "00001"
        Me.txtDepositEndNumber.Text = Me.LastDepositNumber
        Me.chkRevenuesIncludeCreditCards.Visible = False
        Me.chkRevenuesPrintDepositTicket.Visible = False
        'all panels invisible;
        Me.panelRevenueAllFiscalYears.Visible = False
        Me.panelRevenueDailyDeposit.Visible = False
        Me.panelRevenue1098T.Visible = False
        Me.panelRevenueSearchReceived.Visible = False
        'set panel coordinates;
        Me.panelRevenueDailyDeposit.Location = New Point(198, 98)
        Me.panelRevenueAllFiscalYears.Location = Me.panelRevenueDailyDeposit.Location
        Me.panelRevenueSearchReceived.Location = Me.panelRevenueDailyDeposit.Location
        Me.panelRevenue1098T.Location = Me.panelRevenueDailyDeposit.Location
        'print receipt register;
        If sender Is Me.rdoRevenueReceiptRegister Then
            Me.panelOptions.Enabled = True
            Me.panelRevenueSearchReceived.Visible = True
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Me.txtBeginningNumber.Text = "00000001"
            If Me.txtEndingNumber.Text.Trim.Length = 0 Then Me.txtEndingNumber.Text = Me.NextReceiptNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        'print receipt ticket;
        If sender Is Me.rdoRevenueReceiptTicket Then
            Me.panelOptions.Enabled = True
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Me.txtBeginningNumber.Text = "00000001"
            If Me.txtEndingNumber.Text.Trim.Length = 0 Then Me.txtEndingNumber.Text = Me.NextReceiptNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        'daily deposit;
        If sender Is Me.rdoRevenueDailyDeposit Then
            Me.panelRevenueDailyDeposit.Visible = True
            Me.panelOptions.Enabled = False
            Me.chkRevenuesIncludeCreditCards.Visible = True
            Me.chkRevenuesPrintDepositTicket.Visible = True
            Me.txtDepositBegNumber.Visible = True
            Me.txtDepositEndNumber.Visible = False
            Me.txtDepositBegNumber.Text = Me.LastDepositNumber
            Me.txtDepositBegNumber.Focus()
        End If
        'deposit summary;
        If sender Is Me.rdoRevenueDepositSummary Then
            Me.panelRevenueDailyDeposit.Visible = True
            Me.panelOptions.Enabled = False
            Me.txtDepositBegNumber.Visible = True
            Me.txtDepositEndNumber.Visible = True
            Me.txtDepositBegNumber.Focus()
        End If
        'outstanding receipts report;
        If sender Is Me.rdoRevenuePrintOutstandingReceipts Then
            Me.panelRevenueAllFiscalYears.Visible = True
            Me.panelOptions.Enabled = False
        End If
        'void receipt report;
        If sender Is Me.rdoRevenuePrintVoidReceipts Then
            Me.panelOptions.Enabled = True
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Me.txtBeginningNumber.Text = "00000001"
            If Me.txtEndingNumber.Text.Trim.Length = 0 Then Me.txtEndingNumber.Text = Me.NextReceiptNumber
            If Me.UseDate = 1 Then Me.chkUseDate.Checked = True Else Me.chkUseNumber.Checked = True
        End If
        '1098-T report/excel file (1098T is an IRS form for tuition disbursements);
        If sender Is Me.rdoRevenue1098T Then
            Me.panelRevenue1098T.Visible = True
            Me.panelOptions.Enabled = False
            Me.txtRevenueAccountNumber.Focus()
        End If
    End Sub

    Private Sub rdoHandleVendors_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoVendorListing.Click, rdoVendorExpenses.Click, rdoVendor1099Listing.Click, rdoVendorAudit.Click
        'set default settings for the vendor tab
        Me.panelOptions.Enabled = False
        Me.cboCalendarYears.Enabled = False
        Me.chkVendorsSelectSingleVendor.Checked = False
        Me.chkVendorsSelectSingleVendor.Enabled = False
        Me.chkVendorsUse600Minimum.Checked = False
        Me.chkVendorsUse600Minimum.Enabled = False
        Me.chkVendorsIncludeSSN.Checked = False
        Me.chkVendorsIncludeSSN.Enabled = False
        Me.chkVendorsUseFiscalYear.Checked = False
        Me.chkVendorsUseFiscalYear.Enabled = False
        Me.chkVendorsIncludeZeroBalances.Checked = False
        Me.chkVendorsIncludeZeroBalances.Enabled = False
        Me.chk1099ByEmployee.Checked = False
        Me.chk1099ByEmployee.Enabled = False
        Me.chk1099Summary.Checked = False
        Me.chk1099Summary.Enabled = False
        Me.lblVendorCalendar.Enabled = False

        'vendor listing;
        If sender Is Me.rdoVendorListing Then
            Me.chkVendorsIncludeZeroBalances.Checked = False
            Me.chkVendorsIncludeZeroBalances.Enabled = False
        End If

        'vendor expense summary;
        If sender Is Me.rdoVendorExpenses Then
            Me.chkVendorsIncludeZeroBalances.Enabled = True
            Me.chk1099ByEmployee.Enabled = True
            Me.cboCalendarYears.Enabled = True
            Me.chkVendorsUse600Minimum.Enabled = True
            Me.chkVendorsUseFiscalYear.Enabled = True
            Me.lblVendorCalendar.Enabled = True
        End If

        'vendor expense detail (1099 report);
        If sender Is Me.rdoVendor1099Listing Then
            Me.chk1099ByEmployee.Enabled = True
            Me.chk1099Summary.Enabled = True
            Me.cboCalendarYears.Enabled = True
            Me.chkVendorsIncludeSSN.Enabled = True
            Me.chkVendorsIncludeZeroBalances.Enabled = True
            Me.chkVendorsSelectSingleVendor.Enabled = True
            Me.chkVendorsUseFiscalYear.Enabled = True
            Me.chkVendorsUse600Minimum.Enabled = True
            Me.lblVendorCalendar.Enabled = True
        End If

        'vendor audit report;

    End Sub

#End Region

#Region "  Registry Win32 "

    Private Sub GetRegistryEntries()
        Dim rgSoftware As RegistryKey
        Dim rgModule As RegistryKey
        Dim tempdate As String
        Dim _rg32node As String = "ADPC\Activity Fund.Net\AF_Reporting"
        Dim _rg64node As String = "Wow6432Node\ADPC\Activity Fund.Net\AF_Reporting"

        Try
            Dim banknum As String

            'get key to HKEY_LOCAL_MACHINE/SOFTWARE/;
            rgSoftware = Registry.LocalMachine.OpenSubKey("SOFTWARE", False)
            If rgSoftware Is Nothing Then Exit Sub
            'Test for 64 bit.
            rgModule = rgSoftware.OpenSubKey(_rg64node, False)
            If rgModule Is Nothing Then
                'Test for 32 bit and exit if not found.
                rgModule = rgSoftware.OpenSubKey(_rg32node, False)
                If rgModule Is Nothing Then Exit Sub
            End If

            'get the values from the registry;
            Me.UseDate = CInt(rgModule.GetValue("UseDate", 1))
            banknum = CStr(rgModule.GetValue("BankAccount", ""))
            If banknum.Length > 0 Then Call SetCtrlBank(banknum)

            'Close resource.
            If Not rgSoftware Is Nothing Then rgSoftware.Close()
            If Not rgModule Is Nothing Then rgModule.Close()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub SetRegistryEntries()
        Dim rgSoftware As RegistryKey
        Dim rgModule As RegistryKey
        Dim _rg32node As String = "ADPC\Activity Fund.Net\AF_Reporting"
        Dim _rg64node As String = "Wow6432Node\ADPC\Activity Fund.Net\AF_Reporting"

        Try
            'get key to HKEY_LOCAL_MACHINE/SOFTWARE/;
            rgSoftware = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            If rgSoftware Is Nothing Then Exit Sub
            'Test for 64 bit.
            rgModule = rgSoftware.OpenSubKey(_rg64node, True)
            If rgModule Is Nothing Then
                'Test for 32 bit and exit if not found.
                rgModule = rgSoftware.OpenSubKey(_rg32node, True)
                If rgModule Is Nothing Then Exit Sub
            End If

            'set some entries for this module;
            Dim val As Int32
            If Me.chkUseDate.Checked Then Me.UseDate = 1 Else Me.UseDate = 0
            rgModule.SetValue("UseDate", Me.UseDate)
            rgModule.SetValue("BankAccount", Me.BankAccountNumber)

            'Close resource.
            If Not rgSoftware Is Nothing Then rgSoftware.Close()
            If Not rgModule Is Nothing Then rgModule.Close()
        Catch ex As Exception
            Throw
        Finally

        End Try
    End Sub

#End Region

#Region "  Statusbar Events "

    Public StatusBar As New ProgressStatus

    Private Sub InitialiseStatusBar()
        Dim info As StatusBarPanel = New System.Windows.Forms.StatusBarPanel
        Dim progress As StatusBarPanel = New System.Windows.Forms.StatusBarPanel
        info.Text = "Processing..."
        info.Width = 120
        progress.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring

        Try
            With StatusBar
                .Height = 20
                .Panels.Add(info)
                .Panels.Add(progress)
                .Panel = 1
                .ProgressBar.Minimum = 0
                .ProgressBar.Maximum = 100
                .ShowPanels = True
                .SizingGrip = False
            End With
            Me.Controls.Add(StatusBar)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

#End Region

#Region "  Textbox Events "

    Private Sub txtHandleDims_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDim3.Enter, txtDim4.Enter, txtDim5.Enter, txtDim6.Enter, txtDim7.Enter, txtDim8.Enter, txtDim9.Enter
        If sender Is Me.txtDim3 Then Me.txtDim3.SelectAll()
        If sender Is Me.txtDim4 Then Me.txtDim4.SelectAll()
        If sender Is Me.txtDim5 Then Me.txtDim5.SelectAll()
        If sender Is Me.txtDim6 Then Me.txtDim6.SelectAll()
        If sender Is Me.txtDim7 Then Me.txtDim7.SelectAll()
        If sender Is Me.txtDim8 Then Me.txtDim8.SelectAll()
        If sender Is Me.txtDim9 Then Me.txtDim9.SelectAll()
    End Sub

    Private Sub txtDimHandle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDim3.KeyPress
        If (Char.IsNumber(e.KeyChar) = True) Then Exit Sub
        If (e.KeyChar = "*") Then Exit Sub
        If (Char.IsControl(e.KeyChar) = True) Then Exit Sub
        e.Handled = True
    End Sub

    Private Sub txtHandleDims_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDim3.KeyUp, txtDim4.KeyUp, txtDim5.KeyUp, txtDim6.KeyUp, txtDim7.KeyUp, txtDim8.KeyUp, txtDim9.KeyUp
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtHandleDims_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDim3.Leave, txtDim4.Leave, txtDim5.Leave, txtDim6.Leave, txtDim7.Leave, txtDim8.Leave, txtDim9.Leave
        Dim newvalue, value As String
        Dim c As Char = "*"c
        If sender Is Me.txtDim3 Then
            value = Me.txtDim3.Text.Trim
            newvalue = value.PadRight(3, c)
            Me.txtDim3.Text = newvalue
        End If
        If sender Is Me.txtDim4 Then
            value = Me.txtDim4.Text.Trim
            newvalue = value.PadRight(4, c)
            Me.txtDim4.Text = newvalue
        End If
        If sender Is Me.txtDim5 Then
            value = Me.txtDim5.Text.Trim
            newvalue = value.PadRight(3, c)
            Me.txtDim5.Text = newvalue
        End If
        If sender Is Me.txtDim6 Then
            value = Me.txtDim6.Text.Trim
            newvalue = value.PadRight(3, c)
            Me.txtDim6.Text = newvalue
        End If
        If sender Is Me.txtDim7 Then
            value = Me.txtDim7.Text.Trim
            newvalue = value.PadRight(4, c)
            Me.txtDim7.Text = newvalue
        End If
        If sender Is Me.txtDim8 Then
            value = Me.txtDim8.Text.Trim
            newvalue = value.PadRight(3, c)
            Me.txtDim8.Text = newvalue
        End If
        If sender Is Me.txtDim9 Then
            value = Me.txtDim9.Text.Trim
            newvalue = value.PadRight(3, c)
            Me.txtDim9.Text = newvalue
        End If
    End Sub

    Private Sub txtHandleCurrency_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtReconBankStatementBalance.KeyPress, txtReconExpensesNotYetPosted.KeyPress, txtReconInterestNotYetPosted.KeyPress, txtReconInvestments.KeyPress
        If Char.IsControl(e.KeyChar) Then Exit Sub
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = "."c Then Exit Sub
            If e.KeyChar = "-"c Then Exit Sub
            e.Handled = True
        End If
    End Sub

    Private Sub txtHandleCurrency_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReconBankStatementBalance.KeyUp, txtReconExpensesNotYetPosted.KeyUp, txtReconInterestNotYetPosted.KeyUp, txtReconInvestments.KeyUp
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtHandleCurrency_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReconBankStatementBalance.Leave, txtReconExpensesNotYetPosted.Leave, txtReconInterestNotYetPosted.Leave, txtReconInvestments.Leave
        Dim strvalue As String
        If sender Is Me.txtReconBankStatementBalance Then strvalue = Me.txtReconBankStatementBalance.Text
        If sender Is Me.txtReconExpensesNotYetPosted Then strvalue = Me.txtReconExpensesNotYetPosted.Text
        If sender Is Me.txtReconInterestNotYetPosted Then strvalue = Me.txtReconInterestNotYetPosted.Text
        If sender Is Me.txtReconInvestments Then strvalue = Me.txtReconInvestments.Text
        Dim value As Double
        Try
            Me.Cursor = Cursors.WaitCursor
            value = CDbl(strvalue)
        Catch ex As Exception
            value = 0.0
        Finally
            Me.Cursor = Cursors.Default
        End Try

        If sender Is Me.txtReconBankStatementBalance Then Me.txtReconBankStatementBalance.Text = value.ToString.Format("{0:F2}", value)
        If sender Is Me.txtReconExpensesNotYetPosted Then Me.txtReconExpensesNotYetPosted.Text = value.ToString.Format("{0:F2}", value)
        If sender Is Me.txtReconInterestNotYetPosted Then Me.txtReconInterestNotYetPosted.Text = value.ToString.Format("{0:F2}", value)
        If sender Is Me.txtReconInvestments Then Me.txtReconInvestments.Text = value.ToString.Format("{0:F2}", value)

    End Sub

    Private Sub txtBeginningNumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBeginningNumber.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtBeginningNumber.Text.Trim.Length = 0 Then Exit Sub
            Me.txtBeginningNumber.Text = Format8CharNumber(Me.txtBeginningNumber.Text)
        Catch ex As Exception
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtHandleNumber_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBeginningNumber.KeyUp, txtEndingNumber.KeyUp
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtHandleNumber_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDepositBegNumber.Leave, txtDepositEndNumber.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If sender Is Me.txtDepositBegNumber Then
                If Me.txtDepositBegNumber.Text.Trim.Length = 0 Then Me.txtDepositBegNumber.Text = "00001"
                Me.txtDepositBegNumber.Text = Format5CharNumber(Me.txtDepositBegNumber.Text)
            End If
            If sender Is Me.txtDepositEndNumber Then
                If Me.txtDepositEndNumber.Text.Trim.Length = 0 Then Me.txtDepositEndNumber.Text = Me.LastDepositNumber
                Me.txtDepositEndNumber.Text = Format5CharNumber(Me.txtDepositEndNumber.Text)
            End If
        Catch ex As Exception
            Me.txtDepositBegNumber.Text = "00001"
            Me.txtDepositEndNumber.Text = Me.LastDepositNumber
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtEndingNumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEndingNumber.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.txtEndingNumber.Text.Trim.Length = 0 Then Exit Sub
            Me.txtEndingNumber.Text = Format8CharNumber(Me.txtEndingNumber.Text)
        Catch ex As Exception
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtAccountNumber_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAccountsAcctNumberFrom.KeyUp, txtAccountsAcctNumberTo.KeyUp
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtAccountNumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAccountsAcctNumberFrom.Leave, txtAccountsAcctNumberTo.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If sender Is Me.txtAccountsAcctNumberFrom Then
                If Me.txtAccountsAcctNumberFrom.Text.Trim.Length = 0 Then Me.txtAccountsAcctNumberFrom.Text = "0001"
                Me.txtAccountsAcctNumberFrom.Text = Format4CharNumber(Me.txtAccountsAcctNumberFrom.Text)
            End If
            If sender Is Me.txtAccountsAcctNumberTo Then
                If Me.txtAccountsAcctNumberTo.Text.Trim.Length = 0 Then Me.txtAccountsAcctNumberTo.Text = "9999"
                Me.txtAccountsAcctNumberTo.Text = Format4CharNumber(Me.txtAccountsAcctNumberTo.Text)
            End If
        Catch ex As Exception
            Me.txtAccountsAcctNumberFrom.Text = "0001"
            Me.txtAccountsAcctNumberTo.Text = "9999"
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtRevenueAccount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevenueAccountNumber.Enter, txtRevenueSubaccountNumber.Enter
        If sender Is Me.txtRevenueAccountNumber Then Me.txtRevenueAccountNumber.SelectAll()
        If sender Is Me.txtRevenueSubaccountNumber Then Me.txtRevenueSubaccountNumber.SelectAll()
    End Sub

    Private Sub txtRevenueAccount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRevenueAccountNumber.KeyPress, txtRevenueSubaccountNumber.KeyPress
        If Char.IsControl(e.KeyChar) Then Exit Sub
        If Char.IsNumber(e.KeyChar) = False Then
            If e.KeyChar = "."c Then Exit Sub
            If e.KeyChar = "-"c Then Exit Sub
            e.Handled = True
        End If
    End Sub

    Private Sub txtRevenueAccount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRevenueAccountNumber.Leave, txtRevenueSubaccountNumber.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If sender Is Me.txtRevenueAccountNumber Then
                If Me.txtRevenueAccountNumber.Text.Trim.Length = 0 Then Me.txtRevenueAccountNumber.Text = "0001"
                Me.txtRevenueAccountNumber.Text = Format4CharNumber(Me.txtRevenueAccountNumber.Text)
            End If
            If sender Is Me.txtRevenueSubaccountNumber Then
                If Me.txtRevenueSubaccountNumber.Text.Trim.Length = 0 Then Me.txtRevenueSubaccountNumber.Text = "001"
                Me.txtRevenueSubaccountNumber.Text = Format3CharNumber(Me.txtRevenueSubaccountNumber.Text)
            End If
        Catch ex As Exception
            Me.txtRevenueAccountNumber.Text = "0001"
            Me.txtRevenueSubaccountNumber.Text = "001"
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub txtRevenueAccount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRevenueAccountNumber.KeyUp, txtRevenueSubaccountNumber.KeyUp
        If e.KeyCode = Keys.Enter Then SendKeys.Send("{TAB}")
    End Sub

#End Region

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Application.DoEvents()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NEED REPORT OPTIONS FOR MTD SUMMARY INCLUDING:
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MTD SUMMARY                                    
        'MTD SUMMARY ACCOUNT RANGE
        'MTD SUMMARY W/ENCUMBRANCE
        'MTD SUMMARY ACCOUNT RANGE W/ENCUMBRANCE
        '
        'MTD SUBACCOUNT SUMMARY
        'MTD SUBACCOUNT SUMMARY ACCOUNT RANGE
        'MTD SUBACCOUNT SUMMARY W/ENCUMBRANCE
        'MTD SUBACCOUNT SUMMARY ACCOUNT RANGE W/ENCUMBRANCE
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim obj As New AF_Reporting.frmBoldCode
        Try
            'obj.GenerateYTDSummaryOfAccountsWithEncumbrance(Me.BankAccountNumber, False, "0800", "0912")
            'obj.GenerateYTDSummaryOfSubaccountsWithEncumbrance(Me.BankAccountNumber, False, "0001", "0020")
            Application.DoEvents()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
        Finally
            obj.Dispose()
        End Try
    End Sub

    Private Sub rdoExpendituresCheckRegister_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoExpendituresCheckRegister.CheckedChanged

    End Sub

    Private Sub rdoExpenditure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoExpenditure.CheckedChanged

    End Sub

    Private Sub btnVerifyExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVerifyExp.Click
        'only for edit checking 4.13.2017, Mark;
        _editchk = True

        _year = txtC1YearExp.Text
        _fund = txtC1FundExp.Text
        _project = txtC1ProjectExp.Text
        _function = txtC1Function.Text
        _object = txtC1Object.Text
        _program = txtC1ProgramExp.Text
        _subject = txtC1Subject.Text
        _job = txtC1Job.Text
        _site = txtC1SiteExp.Text

        Application.DoEvents()

        'Validate bold code expenditures
        Dim obj As New AF_Reporting.frmBoldCode
        Try
            obj._editchk = True
            obj._expendcode = DirectCast(_year + _fund + _project + _function + _object + _program + _subject + _job + _site, String)
            obj.GenerateBoldCodeExpenditure(False, True, Me.FiscalYearSelected)
            'rev:  obj.GenerateBoldCodeRevenue(chkerronly, sortbycoding, Me.FiscalYearSelected)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
        Finally
            obj._editchk = False
            obj.Dispose()
        End Try



    End Sub

    Private Sub btnVerifyRev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVerifyRev.Click
        'only for edit checking 4.13.2017, Mark;
        _editchk = True

        _year = txtC1Year.Text
        _fund = txtC1Fund.Text
        _project = txtC1Project.Text
        _source = txtC1Source.Text
        _program = txtC1Program.Text
        _site = txtC1Site.Text

        Application.DoEvents()

        'Validate bold code expenditures
        Dim obj As New AF_Reporting.frmBoldCode
        Try
            obj._editchk = True
            obj._Revenuecode = DirectCast(_year + _fund + _project + _source + _program + _site, String)

            obj.GenerateBoldCodeRevenue(False, False, Me.FiscalYearSelected)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, MSGTITLE)
        Finally
            obj._editchk = False
            obj.Dispose()
        End Try
    End Sub


    
    Private Sub chkChkaCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkChkaCode.CheckedChanged
        If chkChkaCode.Checked = True Then
            grpCheck.Visible = True
        Else
            grpCheck.Visible = False
        End If
    End Sub
End Class


