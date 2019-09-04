Imports C1.Win.C1Chart3D
Imports C1.Win.C1Chart3D.Chart3DInteraction

Public Class frmChartsMain___NOTUSED
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal bankaccountnum As String)

        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim authobj As AF_Master.Authuser

        If authobj.PermReports < 1 Then Throw New ArgumentException("Access Denied...")
        Me.SchoolName = authobj.SchoolName
        Me.SchoolAddress1 = authobj.SchoolAddress1
        Me.SchoolCity = authobj.SchoolCity
        Me.SchoolState = authobj.SchoolState
        Me.SchoolTelephone = authobj.SchoolTelephone1
        Me.SchoolZip = authobj.SchoolZipCode
        Me.SchoolFax = authobj.SchoolFax
        '  Me.Bankaccountnum = authobj.BankAccountNumber
        Me.currentmonth = authobj.CurrentMonthString()
        Me.fiscalyear = authobj.FiscalYear
        Me.monthbegindate = authobj.CurrentMonthBeginning
        Me.monthenddate = authobj.CurrentMonthEnding

        Me.Bankaccountnum = bankaccountnum





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
    Friend WithEvents C1Chart3D1 As C1.Win.C1Chart3D.C1Chart3D
    Friend WithEvents btnSetData As System.Windows.Forms.Button
    Friend WithEvents statusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents statusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChartNetMonthTransaction As System.Windows.Forms.MenuItem
    Friend WithEvents mnyChartMonthTransactionsReceipts As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChartsMonthTransactionChecks As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChartMonthlyTransactionAdjustments As System.Windows.Forms.MenuItem
    Friend WithEvents mmnuChartAllMonthTransactions As System.Windows.Forms.MenuItem
    Friend WithEvents cmContour As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem38 As System.Windows.Forms.MenuItem
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents tblChart As System.Windows.Forms.ToolBar
    Friend WithEvents ToolBarButton12 As System.Windows.Forms.ToolBarButton
    Friend WithEvents panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents groupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents rbValueLabels As System.Windows.Forms.RadioButton
    Friend WithEvents rbValues As System.Windows.Forms.RadioButton
    Friend WithEvents groupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents cbLabel As System.Windows.Forms.CheckBox
    Friend WithEvents groupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents upRotZ As System.Windows.Forms.NumericUpDown
    Friend WithEvents groupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents upRotX As System.Windows.Forms.NumericUpDown
    Friend WithEvents checkBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents groupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblDataIndex As System.Windows.Forms.Label
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblMouse As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents groupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblDataCoord As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents groupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents rbSurface As System.Windows.Forms.RadioButton
    Friend WithEvents rbScatter As System.Windows.Forms.RadioButton
    Friend WithEvents rbBar As System.Windows.Forms.RadioButton
    Friend WithEvents btnResetImage As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton6 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton7 As System.Windows.Forms.ToolBarButton
    Friend WithEvents cmView As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton8 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton9 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButton10 As System.Windows.Forms.ToolBarButton
    Friend WithEvents txtChartDetails As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmChartsMain___NOTUSED))
        Me.C1Chart3D1 = New C1.Win.C1Chart3D.C1Chart3D
        Me.btnSetData = New System.Windows.Forms.Button
        Me.statusBar1 = New System.Windows.Forms.StatusBar
        Me.statusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.mmnuChartAllMonthTransactions = New System.Windows.Forms.MenuItem
        Me.mnuChartNetMonthTransaction = New System.Windows.Forms.MenuItem
        Me.mnyChartMonthTransactionsReceipts = New System.Windows.Forms.MenuItem
        Me.mnuChartsMonthTransactionChecks = New System.Windows.Forms.MenuItem
        Me.mnuChartMonthlyTransactionAdjustments = New System.Windows.Forms.MenuItem
        Me.cmContour = New System.Windows.Forms.ContextMenu
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem37 = New System.Windows.Forms.MenuItem
        Me.MenuItem38 = New System.Windows.Forms.MenuItem
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.tblChart = New System.Windows.Forms.ToolBar
        Me.ToolBarButton12 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton6 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton7 = New System.Windows.Forms.ToolBarButton
        Me.cmView = New System.Windows.Forms.ContextMenu
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton8 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton9 = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButton10 = New System.Windows.Forms.ToolBarButton
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.panel1 = New System.Windows.Forms.Panel
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.txtChartDetails = New System.Windows.Forms.TextBox
        Me.groupBox5 = New System.Windows.Forms.GroupBox
        Me.rbValueLabels = New System.Windows.Forms.RadioButton
        Me.rbValues = New System.Windows.Forms.RadioButton
        Me.groupBox8 = New System.Windows.Forms.GroupBox
        Me.cbLabel = New System.Windows.Forms.CheckBox
        Me.groupBox7 = New System.Windows.Forms.GroupBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.upRotZ = New System.Windows.Forms.NumericUpDown
        Me.groupBox6 = New System.Windows.Forms.GroupBox
        Me.upRotX = New System.Windows.Forms.NumericUpDown
        Me.checkBox1 = New System.Windows.Forms.CheckBox
        Me.groupBox3 = New System.Windows.Forms.GroupBox
        Me.lblDataIndex = New System.Windows.Forms.Label
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.lblMouse = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.groupBox2 = New System.Windows.Forms.GroupBox
        Me.lblDataCoord = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.groupBox4 = New System.Windows.Forms.GroupBox
        Me.rbSurface = New System.Windows.Forms.RadioButton
        Me.rbScatter = New System.Windows.Forms.RadioButton
        Me.rbBar = New System.Windows.Forms.RadioButton
        Me.btnResetImage = New System.Windows.Forms.Button
        CType(Me.C1Chart3D1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panel1.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.groupBox5.SuspendLayout()
        Me.groupBox8.SuspendLayout()
        Me.groupBox7.SuspendLayout()
        CType(Me.upRotZ, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupBox6.SuspendLayout()
        CType(Me.upRotX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupBox3.SuspendLayout()
        Me.groupBox1.SuspendLayout()
        Me.groupBox2.SuspendLayout()
        Me.groupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'C1Chart3D1
        '
        Me.C1Chart3D1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.C1Chart3D1.Location = New System.Drawing.Point(208, 40)
        Me.C1Chart3D1.Name = "C1Chart3D1"
        Me.C1Chart3D1.PropBag = "<?xml version=""1.0""?><Chart3DPropBag Version=""""><AxisX><GridStyle Thickness=""5"" C" & _
        "olor=""Black"" Pattern=""Solid"" /></AxisX><View AxisTitleFont=""Microsoft Sans Serif" & _
        ", 12world"" AxisFont=""Microsoft Sans Serif, 10world"" /><ChartGroupsCollection><Ch" & _
        "art3DGroup><Bar ColumnFormat=""Histogram"" Origin=""2"" /><ChartData><Set type=""Char" & _
        "t3DDataSetGrid"" RowCount=""11"" ColumnCount=""11"" RowDelta=""1"" ColumnDelta=""11"" Row" & _
        "Origin=""0"" ColumnOrigin=""0"" Hole=""3.4028234663852886E+38""><Data>4.49999988824129" & _
        "1 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234" & _
        "663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+" & _
        "38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823" & _
        "4663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E" & _
        "+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282" & _
        "34663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886" & _
        "E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028" & _
        "234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823466385288" & _
        "6E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402" & _
        "8234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282346638528" & _
        "86E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40" & _
        "28234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852" & _
        "886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4" & _
        "028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823466385" & _
        "2886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3." & _
        "4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282346638" & _
        "52886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3" & _
        ".4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663" & _
        "852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 " & _
        "3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823466" & _
        "3852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38" & _
        " 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282346" & _
        "63852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+3" & _
        "8 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234" & _
        "663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+" & _
        "38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823" & _
        "4663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E" & _
        "+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282" & _
        "34663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886" & _
        "E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028" & _
        "234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402823466385288" & _
        "6E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.402" & _
        "8234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40282346638528" & _
        "86E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.4028234663852886E+38 3.40" & _
        "28234663852886E+38 3.4028234663852886E+38</Data></Set></ChartData></Chart3DGroup" & _
        "></ChartGroupsCollection><ContourStyles ColorSet=""GreenPurpleYellow"" /><StyleCol" & _
        "lection><NamedStyle Name=""Legend"" ParentName=""Legend.default"" StyleData=""Border=" & _
        "Double,Black,1;BackColor=ControlDark;Opaque=True;Wrap=True;AlignVert=Top;Roundin" & _
        "g=4 4 4 4;"" /><NamedStyle Name=""Footer"" ParentName=""Control"" /><NamedStyle Name=" & _
        """Area"" ParentName=""Area.default"" StyleData=""GradientStyle=None;Border=None,Trans" & _
        "parent,4;AlignHorz=General;BackColor2=Yellow;BackColor=Gainsboro;ForeColor=Black" & _
        ";HatchStyle=None;Rounding=20 20 20 20;"" /><NamedStyle Name=""Control"" ParentName=" & _
        """Control.default"" StyleData=""BackColor=Control;Font=Microsoft Sans Serif, 8.25pt" & _
        ", style=Underline;"" /><NamedStyle Name=""LabelStyleDefault"" ParentName=""Control"" " & _
        "StyleData=""BackColor=Transparent;"" /><NamedStyle Name=""Legend.default"" ParentNam" & _
        "e=""Control"" StyleData=""Wrap=False;AlignVert=Top;"" /><NamedStyle Name=""Header"" Pa" & _
        "rentName=""Control"" /><NamedStyle Name=""Control.default"" ParentName="""" StyleData=" & _
        """ForeColor=ControlText;Border=None,Black,1;BackColor=Control;"" /><NamedStyle Nam" & _
        "e=""Area.default"" ParentName=""Control"" StyleData=""AlignVert=Top;"" /></StyleCollec" & _
        "tion><LegendData Text=""kljhkljh"" Compass=""East"" LocationDefault=""25, 25"" SizeDef" & _
        "ault=""30, 30"" /><FooterData Visible=""True"" Compass=""South"" /><HeaderData Visible" & _
        "=""True"" Compass=""North"" /></Chart3DPropBag>"
        Me.C1Chart3D1.Size = New System.Drawing.Size(336, 312)
        Me.C1Chart3D1.TabIndex = 0
        Me.C1Chart3D1.Visible = False
        '
        'btnSetData
        '
        Me.btnSetData.Location = New System.Drawing.Point(216, 400)
        Me.btnSetData.Name = "btnSetData"
        Me.btnSetData.Size = New System.Drawing.Size(96, 24)
        Me.btnSetData.TabIndex = 10
        Me.btnSetData.Text = "Get Charts ..."
        Me.btnSetData.Visible = False
        '
        'statusBar1
        '
        Me.statusBar1.Location = New System.Drawing.Point(0, 438)
        Me.statusBar1.Name = "statusBar1"
        Me.statusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.statusBarPanel1})
        Me.statusBar1.ShowPanels = True
        Me.statusBar1.Size = New System.Drawing.Size(608, 24)
        Me.statusBar1.TabIndex = 4
        Me.statusBar1.Text = "statusBar1"
        '
        'statusBarPanel1
        '
        Me.statusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.statusBarPanel1.Text = "Move mouse over chart."
        Me.statusBarPanel1.Width = 592
        '
        'Timer1
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2})
        Me.MenuItem1.Text = "File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Exit"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5})
        Me.MenuItem3.Text = "Charts"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mmnuChartAllMonthTransactions, Me.mnuChartNetMonthTransaction, Me.mnyChartMonthTransactionsReceipts, Me.mnuChartsMonthTransactionChecks, Me.mnuChartMonthlyTransactionAdjustments})
        Me.MenuItem5.Text = "Transactions"
        '
        'mmnuChartAllMonthTransactions
        '
        Me.mmnuChartAllMonthTransactions.Index = 0
        Me.mmnuChartAllMonthTransactions.Text = "All Monthly Transactions"
        '
        'mnuChartNetMonthTransaction
        '
        Me.mnuChartNetMonthTransaction.Index = 1
        Me.mnuChartNetMonthTransaction.Text = "Net Monthly Transactions"
        '
        'mnyChartMonthTransactionsReceipts
        '
        Me.mnyChartMonthTransactionsReceipts.Index = 2
        Me.mnyChartMonthTransactionsReceipts.Text = "Monthly Transactions Receipts"
        '
        'mnuChartsMonthTransactionChecks
        '
        Me.mnuChartsMonthTransactionChecks.Index = 3
        Me.mnuChartsMonthTransactionChecks.Text = "Monthly Transactions Checks"
        '
        'mnuChartMonthlyTransactionAdjustments
        '
        Me.mnuChartMonthlyTransactionAdjustments.Index = 4
        Me.mnuChartMonthlyTransactionAdjustments.Text = "Monthly Transactions Adjustments"
        '
        'cmContour
        '
        Me.cmContour.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem19, Me.MenuItem20, Me.MenuItem21, Me.MenuItem22, Me.MenuItem23, Me.MenuItem33, Me.MenuItem34, Me.MenuItem35, Me.MenuItem36, Me.MenuItem37, Me.MenuItem38})
        '
        'MenuItem19
        '
        Me.MenuItem19.Checked = True
        Me.MenuItem19.Index = 0
        Me.MenuItem19.RadioCheck = True
        Me.MenuItem19.Text = "None"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 1
        Me.MenuItem20.RadioCheck = True
        Me.MenuItem20.Text = "Rainbow"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 2
        Me.MenuItem21.RadioCheck = True
        Me.MenuItem21.Text = "RevRainBow"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 3
        Me.MenuItem22.RadioCheck = True
        Me.MenuItem22.Text = "Black - White"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 4
        Me.MenuItem23.RadioCheck = True
        Me.MenuItem23.Text = "White - Black"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 5
        Me.MenuItem33.RadioCheck = True
        Me.MenuItem33.Text = "Green - Blue"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 6
        Me.MenuItem34.RadioCheck = True
        Me.MenuItem34.Text = "Red - White"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 7
        Me.MenuItem35.RadioCheck = True
        Me.MenuItem35.Text = "Blue - Pink"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 8
        Me.MenuItem36.RadioCheck = True
        Me.MenuItem36.Text = "Blue - White - Red"
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 9
        Me.MenuItem37.RadioCheck = True
        Me.MenuItem37.Text = "Black - Red - Yellow"
        '
        'MenuItem38
        '
        Me.MenuItem38.Index = 10
        Me.MenuItem38.RadioCheck = True
        Me.MenuItem38.Text = "Green - Purple - Yellow"
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(0, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(184, 438)
        Me.Splitter1.TabIndex = 11
        Me.Splitter1.TabStop = False
        '
        'tblChart
        '
        Me.tblChart.AllowDrop = True
        Me.tblChart.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.tblChart.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tblChart.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton12, Me.ToolBarButton1, Me.ToolBarButton2, Me.ToolBarButton4, Me.ToolBarButton3, Me.ToolBarButton6, Me.ToolBarButton7, Me.ToolBarButton5, Me.ToolBarButton8, Me.ToolBarButton9, Me.ToolBarButton10})
        Me.tblChart.ButtonSize = New System.Drawing.Size(23, 22)
        Me.tblChart.Divider = False
        Me.tblChart.DropDownArrows = True
        Me.tblChart.ImageList = Me.ImageList1
        Me.tblChart.Location = New System.Drawing.Point(184, 0)
        Me.tblChart.Name = "tblChart"
        Me.tblChart.ShowToolTips = True
        Me.tblChart.Size = New System.Drawing.Size(424, 28)
        Me.tblChart.TabIndex = 13
        '
        'ToolBarButton12
        '
        Me.ToolBarButton12.DropDownMenu = Me.cmContour
        Me.ToolBarButton12.ImageIndex = 2
        Me.ToolBarButton12.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.ToolBarButton12.ToolTipText = "Zone Colors"
        '
        'ToolBarButton1
        '
        Me.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.ToolBarButton1.Text = "|"
        '
        'ToolBarButton2
        '
        Me.ToolBarButton2.ImageIndex = 9
        Me.ToolBarButton2.ToolTipText = "Drop Lines On Scatter Chart"
        '
        'ToolBarButton4
        '
        Me.ToolBarButton4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton3
        '
        Me.ToolBarButton3.ImageIndex = 8
        Me.ToolBarButton3.ToolTipText = "Set Mesh"
        '
        'ToolBarButton6
        '
        Me.ToolBarButton6.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton7
        '
        Me.ToolBarButton7.DropDownMenu = Me.cmView
        Me.ToolBarButton7.ImageIndex = 10
        Me.ToolBarButton7.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.ToolBarButton7.ToolTipText = "Toggle between 3D or 2D Chart Display"
        '
        'cmView
        '
        Me.cmView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MenuItem6, Me.MenuItem7})
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.Text = "3D - Default"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "2D - YZ Plane"
        '
        'ToolBarButton5
        '
        Me.ToolBarButton5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton8
        '
        Me.ToolBarButton8.ImageIndex = 11
        Me.ToolBarButton8.ToolTipText = "Interactive Chart - Toggle With Mouse"
        '
        'ToolBarButton9
        '
        Me.ToolBarButton9.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'ToolBarButton10
        '
        Me.ToolBarButton10.ImageIndex = 12
        Me.ToolBarButton10.ToolTipText = "HoleValue property to set the data value that represents a hole in the data."
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'panel1
        '
        Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panel1.Controls.Add(Me.GroupBox9)
        Me.panel1.Controls.Add(Me.groupBox5)
        Me.panel1.Controls.Add(Me.groupBox8)
        Me.panel1.Controls.Add(Me.groupBox7)
        Me.panel1.Controls.Add(Me.groupBox6)
        Me.panel1.Controls.Add(Me.groupBox3)
        Me.panel1.Controls.Add(Me.groupBox1)
        Me.panel1.Controls.Add(Me.groupBox2)
        Me.panel1.Controls.Add(Me.groupBox4)
        Me.panel1.Controls.Add(Me.btnResetImage)
        Me.panel1.Location = New System.Drawing.Point(1, 8)
        Me.panel1.Name = "panel1"
        Me.panel1.Size = New System.Drawing.Size(176, 416)
        Me.panel1.TabIndex = 14
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.txtChartDetails)
        Me.GroupBox9.Location = New System.Drawing.Point(8, 296)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(152, 64)
        Me.GroupBox9.TabIndex = 11
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Chart Details"
        '
        'txtChartDetails
        '
        Me.txtChartDetails.BackColor = System.Drawing.SystemColors.Control
        Me.txtChartDetails.ForeColor = System.Drawing.Color.Navy
        Me.txtChartDetails.Location = New System.Drawing.Point(3, 14)
        Me.txtChartDetails.Multiline = True
        Me.txtChartDetails.Name = "txtChartDetails"
        Me.txtChartDetails.ReadOnly = True
        Me.txtChartDetails.Size = New System.Drawing.Size(141, 40)
        Me.txtChartDetails.TabIndex = 0
        Me.txtChartDetails.Text = " Selected Chart"
        '
        'groupBox5
        '
        Me.groupBox5.Controls.Add(Me.rbValueLabels)
        Me.groupBox5.Controls.Add(Me.rbValues)
        Me.groupBox5.Location = New System.Drawing.Point(8, 96)
        Me.groupBox5.Name = "groupBox5"
        Me.groupBox5.Size = New System.Drawing.Size(78, 80)
        Me.groupBox5.TabIndex = 5
        Me.groupBox5.TabStop = False
        Me.groupBox5.Text = "Axis anno"
        '
        'rbValueLabels
        '
        Me.rbValueLabels.Location = New System.Drawing.Point(8, 40)
        Me.rbValueLabels.Name = "rbValueLabels"
        Me.rbValueLabels.Size = New System.Drawing.Size(64, 32)
        Me.rbValueLabels.TabIndex = 1
        Me.rbValueLabels.Text = "Value Labels"
        '
        'rbValues
        '
        Me.rbValues.Location = New System.Drawing.Point(8, 20)
        Me.rbValues.Name = "rbValues"
        Me.rbValues.Size = New System.Drawing.Size(56, 16)
        Me.rbValues.TabIndex = 0
        Me.rbValues.Text = "Values"
        '
        'groupBox8
        '
        Me.groupBox8.Controls.Add(Me.cbLabel)
        Me.groupBox8.Location = New System.Drawing.Point(88, 96)
        Me.groupBox8.Name = "groupBox8"
        Me.groupBox8.Size = New System.Drawing.Size(72, 80)
        Me.groupBox8.TabIndex = 9
        Me.groupBox8.TabStop = False
        Me.groupBox8.Text = "Label"
        '
        'cbLabel
        '
        Me.cbLabel.Location = New System.Drawing.Point(4, 16)
        Me.cbLabel.Name = "cbLabel"
        Me.cbLabel.Size = New System.Drawing.Size(52, 16)
        Me.cbLabel.TabIndex = 0
        Me.cbLabel.Text = "Show"
        '
        'groupBox7
        '
        Me.groupBox7.Controls.Add(Me.CheckBox2)
        Me.groupBox7.Controls.Add(Me.upRotZ)
        Me.groupBox7.Location = New System.Drawing.Point(88, 52)
        Me.groupBox7.Name = "groupBox7"
        Me.groupBox7.Size = New System.Drawing.Size(82, 44)
        Me.groupBox7.TabIndex = 8
        Me.groupBox7.TabStop = False
        Me.groupBox7.Text = "RotationZ"
        '
        'CheckBox2
        '
        Me.CheckBox2.Image = CType(resources.GetObject("CheckBox2.Image"), System.Drawing.Image)
        Me.CheckBox2.Location = New System.Drawing.Point(46, 16)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(32, 24)
        Me.CheckBox2.TabIndex = 8
        '
        'upRotZ
        '
        Me.upRotZ.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.upRotZ.Location = New System.Drawing.Point(1, 16)
        Me.upRotZ.Maximum = New Decimal(New Integer() {360, 0, 0, 0})
        Me.upRotZ.Name = "upRotZ"
        Me.upRotZ.Size = New System.Drawing.Size(46, 20)
        Me.upRotZ.TabIndex = 7
        '
        'groupBox6
        '
        Me.groupBox6.Controls.Add(Me.upRotX)
        Me.groupBox6.Controls.Add(Me.checkBox1)
        Me.groupBox6.Location = New System.Drawing.Point(88, 8)
        Me.groupBox6.Name = "groupBox6"
        Me.groupBox6.Size = New System.Drawing.Size(82, 44)
        Me.groupBox6.TabIndex = 6
        Me.groupBox6.TabStop = False
        Me.groupBox6.Text = "RotationX"
        '
        'upRotX
        '
        Me.upRotX.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.upRotX.Location = New System.Drawing.Point(2, 16)
        Me.upRotX.Maximum = New Decimal(New Integer() {360, 0, 0, 0})
        Me.upRotX.Name = "upRotX"
        Me.upRotX.Size = New System.Drawing.Size(46, 20)
        Me.upRotX.TabIndex = 0
        '
        'checkBox1
        '
        Me.checkBox1.Image = CType(resources.GetObject("checkBox1.Image"), System.Drawing.Image)
        Me.checkBox1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.checkBox1.Location = New System.Drawing.Point(47, 17)
        Me.checkBox1.Name = "checkBox1"
        Me.checkBox1.Size = New System.Drawing.Size(32, 24)
        Me.checkBox1.TabIndex = 4
        '
        'groupBox3
        '
        Me.groupBox3.Controls.Add(Me.lblDataIndex)
        Me.groupBox3.Controls.Add(Me.PictureBox3)
        Me.groupBox3.Location = New System.Drawing.Point(8, 256)
        Me.groupBox3.Name = "groupBox3"
        Me.groupBox3.Size = New System.Drawing.Size(152, 40)
        Me.groupBox3.TabIndex = 4
        Me.groupBox3.TabStop = False
        Me.groupBox3.Text = "Data index"
        '
        'lblDataIndex
        '
        Me.lblDataIndex.Location = New System.Drawing.Point(8, 16)
        Me.lblDataIndex.Name = "lblDataIndex"
        Me.lblDataIndex.Size = New System.Drawing.Size(136, 16)
        Me.lblDataIndex.TabIndex = 4
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(64, 0)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 9
        Me.PictureBox3.TabStop = False
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.lblMouse)
        Me.groupBox1.Controls.Add(Me.PictureBox1)
        Me.groupBox1.Location = New System.Drawing.Point(8, 176)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(152, 40)
        Me.groupBox1.TabIndex = 3
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Mouse"
        '
        'lblMouse
        '
        Me.lblMouse.Location = New System.Drawing.Point(8, 16)
        Me.lblMouse.Name = "lblMouse"
        Me.lblMouse.Size = New System.Drawing.Size(136, 16)
        Me.lblMouse.TabIndex = 3
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(48, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'groupBox2
        '
        Me.groupBox2.Controls.Add(Me.lblDataCoord)
        Me.groupBox2.Controls.Add(Me.PictureBox2)
        Me.groupBox2.Location = New System.Drawing.Point(8, 216)
        Me.groupBox2.Name = "groupBox2"
        Me.groupBox2.Size = New System.Drawing.Size(152, 40)
        Me.groupBox2.TabIndex = 3
        Me.groupBox2.TabStop = False
        Me.groupBox2.Text = "Data coordinates"
        '
        'lblDataCoord
        '
        Me.lblDataCoord.Location = New System.Drawing.Point(8, 16)
        Me.lblDataCoord.Name = "lblDataCoord"
        Me.lblDataCoord.Size = New System.Drawing.Size(136, 16)
        Me.lblDataCoord.TabIndex = 4
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(96, 0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 8
        Me.PictureBox2.TabStop = False
        '
        'groupBox4
        '
        Me.groupBox4.Controls.Add(Me.rbSurface)
        Me.groupBox4.Controls.Add(Me.rbScatter)
        Me.groupBox4.Controls.Add(Me.rbBar)
        Me.groupBox4.Location = New System.Drawing.Point(8, 8)
        Me.groupBox4.Name = "groupBox4"
        Me.groupBox4.Size = New System.Drawing.Size(78, 88)
        Me.groupBox4.TabIndex = 3
        Me.groupBox4.TabStop = False
        Me.groupBox4.Text = "Chart Type"
        '
        'rbSurface
        '
        Me.rbSurface.Location = New System.Drawing.Point(8, 64)
        Me.rbSurface.Name = "rbSurface"
        Me.rbSurface.Size = New System.Drawing.Size(64, 16)
        Me.rbSurface.TabIndex = 5
        Me.rbSurface.Text = "Surface"
        '
        'rbScatter
        '
        Me.rbScatter.Location = New System.Drawing.Point(8, 40)
        Me.rbScatter.Name = "rbScatter"
        Me.rbScatter.Size = New System.Drawing.Size(64, 16)
        Me.rbScatter.TabIndex = 4
        Me.rbScatter.Text = "Scatter"
        '
        'rbBar
        '
        Me.rbBar.Location = New System.Drawing.Point(8, 16)
        Me.rbBar.Name = "rbBar"
        Me.rbBar.Size = New System.Drawing.Size(56, 16)
        Me.rbBar.TabIndex = 3
        Me.rbBar.Text = "Bar"
        '
        'btnResetImage
        '
        Me.btnResetImage.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnResetImage.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.btnResetImage.Image = CType(resources.GetObject("btnResetImage.Image"), System.Drawing.Image)
        Me.btnResetImage.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnResetImage.Location = New System.Drawing.Point(8, 392)
        Me.btnResetImage.Name = "btnResetImage"
        Me.btnResetImage.Size = New System.Drawing.Size(96, 20)
        Me.btnResetImage.TabIndex = 7
        Me.btnResetImage.Text = "Reset Chart"
        Me.btnResetImage.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmChartsMain___NOTUSED
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(608, 462)
        Me.Controls.Add(Me.panel1)
        Me.Controls.Add(Me.tblChart)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.statusBar1)
        Me.Controls.Add(Me.C1Chart3D1)
        Me.Controls.Add(Me.btnSetData)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmChartsMain___NOTUSED"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Activity Fund.Net Reports - Charts"
        CType(Me.C1Chart3D1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panel1.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.groupBox5.ResumeLayout(False)
        Me.groupBox8.ResumeLayout(False)
        Me.groupBox7.ResumeLayout(False)
        CType(Me.upRotZ, System.ComponentModel.ISupportInitialize).EndInit()
        Me.groupBox6.ResumeLayout(False)
        CType(Me.upRotX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.groupBox3.ResumeLayout(False)
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox2.ResumeLayout(False)
        Me.groupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  Class Members "


    'property values
    Private _schoolname As String
    Private _schooladdress1 As String
    Private _schoolcity As String
    Private _schoolstate As String
    Private _schoolzip As String
    Private _schoolzipext As String
    Private _schooltelephone As String
    Private _schoolfax As String
    Private _bankaccount As String
    Private currentmonth As String
    Private fiscalyear As Int32
    Private monthbegindate As Date
    Private monthenddate As Date
    Private bankaccountnumber As String



    'Chartdata
    Private totalrows As Int32

    Private mtdtranstot As Single
    Private mtdtransrcpttot As Single
    Private mtdtranscheckstot As Single
    Private mtdtransadjtot As Single



    Private xpixel, ypixel As Integer
    Private bMarker As Boolean = False
    Private pen As New Pen(Color.Red, 2)
    Private x As Single = 0
    Private y As Single = 0
    Private z As Single = 0
    Private nlabel As Integer = 0
    Private old_row As Integer = -1
    Private old_col As Integer = -1
    Private bCapture As Boolean = True
    Private angleIncrement As Integer = 2

    Private setGrid As Chart3DDataSetGrid
    Private setIrGrid As Chart3DDataSetIrGrid

    '>1 COL
    Private flagmulticol As Boolean = False


    'title
    Private MSGTITLE As String = "Activity Fund.Net Reports - Charts"


#End Region

#Region "  Properties "

    Private Property SchoolName() As String
        Get
            Return _schoolname
        End Get
        Set(ByVal Value As String)
            _schoolname = Value
        End Set
    End Property

    Private Property SchoolAddress1() As String
        Get
            Return _schooladdress1
        End Get
        Set(ByVal Value As String)
            _schooladdress1 = Value
        End Set
    End Property

    Private Property SchoolCity() As String
        Get
            Return _schoolcity
        End Get
        Set(ByVal Value As String)
            _schoolcity = Value
        End Set
    End Property

    Private Property SchoolState() As String
        Get
            Return _schoolstate
        End Get
        Set(ByVal Value As String)
            _schoolstate = Value
        End Set
    End Property

    Private Property SchoolZip() As String
        Get
            Return _schoolzip
        End Get
        Set(ByVal Value As String)
            _schoolzip = Value
        End Set
    End Property

    Private Property SchoolFax() As String
        Get
            Return _schoolfax
        End Get
        Set(ByVal Value As String)
            _schoolfax = Value
        End Set
    End Property

    Private Property SchoolTelephone() As String
        Get
            Return _schooltelephone
        End Get
        Set(ByVal Value As String)
            _schooltelephone = Value
        End Set
    End Property
    Private Property Bankaccountnum() As String
        Get
            Return Me._bankaccount
        End Get
        Set(ByVal Value As String)
            _bankaccount = Value
        End Set
    End Property


#End Region

#Region " Start of Application - Load Events"

    Private Sub c1chart3D1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles C1Chart3D1.Load


        rbValues.Checked = True
        'rbSurface.Checked = True
        rbBar.Checked = True

        upRotX.Value = C1Chart3D1.ChartArea.View.RotationX
        upRotZ.Value = C1Chart3D1.ChartArea.View.RotationZ

        C1Chart3D1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        panel1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left

        Me.C1Chart3D1.Size = New Size(404, 384)


    End Sub
#End Region

#Region "  Chkbox Events"
    Private Sub rbChartType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbBar.CheckedChanged, rbSurface.CheckedChanged
        updateChartType()

    End Sub



#End Region

#Region "  Chart Type"
    Private Sub updateChartType()
        ' User Selects Chart Type
        ' Options - Bar,Scatter or Surface
        'bMarker = False
        C1Chart3D1.Refresh()

        If rbBar.Checked Then
            C1Chart3D1.ChartGroups(0).ChartType = Chart3DTypeEnum.Bar
        ElseIf rbScatter.Checked Then
            C1Chart3D1.ChartGroups(0).ChartType = Chart3DTypeEnum.Scatter
        Else
            C1Chart3D1.ChartGroups(0).ChartType = Chart3DTypeEnum.Surface
        End If
    End Sub

    Private Sub ResetCurrChart()
        'remove old graph values/Data
        With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
            Dim i, j As Integer
            For i = 0 To .ColumnCount - 1
                For j = 0 To .RowCount - 1
                    If i <> 0 Or j <> 0 Then
                        .Item(i, j) = .Hole
                    End If
                Next
            Next
            .ColumnCount = 2
            .RowCount = 2
        End With

    End Sub

#Region "  Rotation Events"
    Private Sub upRotX_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles upRotX.ValueChanged

        C1Chart3D1.ChartArea.View.RotationX = CInt(upRotX.Value)
    End Sub


    Private Sub upRotZ_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles upRotZ.ValueChanged

        C1Chart3D1.ChartArea.View.RotationZ = CInt(upRotZ.Value)
    End Sub
#End Region

    Private Sub cbLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'C1Chart3D1.ChartLabels(0).Visible = cbLabel.Checked
    End Sub
#End Region

#Region "  Mouse Events "

    Private Sub C1chart3D1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles C1Chart3D1.MouseMove
        If Not bCapture Then
            Return
        End If
        lblMouse.Text = [String].Format("X={0}; Y={1}", e.X, e.Y)

        Dim _x As Single = 0
        Dim _y As Single = 0
        Dim _z As Single = 0
        If C1Chart3D1.ChartGroups(0).ChartData.CoordToDataCoord(e.X, e.Y, _x, _y, _z) Then
            If _x <> x OrElse _y <> y OrElse _z <> z Then
                x = _x
                y = _y
                z = _z
                lblDataCoord.Text = [String].Format("X={0}; Y={1}; Z={2}", x, y, z)

                'user selects to view values of labels
                If rbValueLabels.Checked Then
                    toggleValueLabelsChart1(True)
                End If


            Else
                Return
            End If
        Else
        End If

        Dim row As Integer = 0
        Dim col As Integer = 0
        If C1Chart3D1.ChartGroups(0).ChartData.CoordToDataIndex(e.X, e.Y, col, row) Then
            lblDataIndex.Text = [String].Format("Col={0}; Row={1}", col, row)

            'Set Labels for acct numbers - visible one @ a time
            If cbLabel.Checked = True Then
                Dim i As Int32
                For i = 1 To totalrows - 1
                    C1Chart3D1.ChartLabels.LabelsCollection(i).Visible = False
                Next
                C1Chart3D1.ChartLabels.LabelsCollection(row).Visible = True
            End If

            C1Chart3D1.ChartLabels(0).AttachMethodData.Column = col
            C1Chart3D1.ChartLabels(0).AttachMethodData.Row = row

            If C1Chart3D1.ChartGroups(0).ChartType = Chart3DTypeEnum.Bar Then
                If old_col <> -1 AndAlso old_row <> -1 Then
                    C1Chart3D1.ChartGroups(0).Bar.SetBarColor(old_col, old_row, Color.LightSteelBlue)
                End If

                C1Chart3D1.ChartGroups(0).Bar.SetBarColor(col, row, Color.Red)

            End If

            old_row = row
            old_col = col
        Else
        End If
    End Sub
    Private Sub chart3D1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles C1Chart3D1.MouseDown
        If rbValueLabels.Checked Then
            If e.Button = MouseButtons.Left Then
                nlabel += 1
            Else
                Dim axis As C1.Win.C1Chart3D.Chart3DAxis
                For Each axis In C1Chart3D1.ChartArea.Axes
                    axis.ValueLabels.Clear()
                Next axis
                nlabel = 0
                C1Chart3D1.Invalidate()
            End If
        Else
            bCapture = Not bCapture
        End If
    End Sub

#End Region

#Region "  Value Labels "

    Private Sub toggleValueLabelsChart1(ByVal show As Boolean)

        If show Then
            Dim axis As C1.Win.C1Chart3D.Chart3DAxis
            For Each axis In C1Chart3D1.ChartArea.Axes
                Dim val As Single
                If axis.Equals(C1Chart3D1.ChartArea.AxisX) Then
                    val = x
                ElseIf axis.Equals(C1Chart3D1.ChartArea.AxisY) Then
                    val = y
                Else
                    val = z
                End If
                If axis.ValueLabels.Count < nlabel + 1 Then
                    axis.ValueLabels.Add(val, val.ToString())

                Else
                    axis.ValueLabels(nlabel).Value = val
                    axis.ValueLabels(nlabel).Text = val.ToString()

                End If

                axis.AnnoMethod = C1.Win.C1Chart3D.AnnotationMethodEnum.ValueLabels
                'Set values to rotate horz
                'axis.AnnoRotated = True
                axis.AnnoPosition = AnnoPositionEnum.Both


                'Change Color of selcted col
                axis.MajorGrid.Style.Color = Color.Red
                axis.AutoMajor = True
                axis.AutoMinor = True
            Next axis

            Me.statusBarPanel1.Text = "Left click to save current axis labels position. Right click to clear all labels"
            bCapture = True
        Else
            Dim axis As C1.Win.C1Chart3D.Chart3DAxis
            For Each axis In C1Chart3D1.ChartArea.Axes
                axis.AnnoMethod = C1.Win.C1Chart3D.AnnotationMethodEnum.Values
                axis.MajorGrid.Style.Color = Color.Black
            Next axis
            statusBarPanel1.Text = "Move mouse over chart to see coordinates mapping results on the left panel.Click on chart to toggle mouse capture"
        End If
    End Sub

#End Region

#Region "  Button Events"


    Private Sub btnResetImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetImage.Click

        Dim axis As C1.Win.C1Chart3D.Chart3DAxis
        For Each axis In C1Chart3D1.ChartArea.Axes
            axis.ValueLabels.Clear()
        Next axis
        nlabel = 0
        '   C1Chart3D1.Invalidate()
        'Set default rotation for graph
        upRotX.Text = CStr(75)
        upRotZ.Text = CStr(115)
        SetChart1Properties()
        C1Chart3D1.ChartArea.View.RotationY = 0 'CInt(upRotX.Value)

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        C1Chart3D1.ChartGroups(0).Elevation.IsMeshed = Not C1Chart3D1.ChartGroups(0).Elevation.IsMeshed
    End Sub
    Private Sub btnDropScatter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If rbScatter.Checked = True Then
            C1Chart3D1.ChartGroups(0).Elevation.DropLines = Not C1Chart3D1.ChartGroups(0).Elevation.DropLines

        End If
    End Sub
#End Region

#Region "  Chart Properties"


    Private Sub SetChart1Properties()
        'Set Properties
        'Me.lblmtdtrans.Visible = True
        'Me.btnLargeImageChart1.Visible = True
        Me.C1Chart3D1.Visible = True
        Me.WindowState = FormWindowState.Maximized


        With C1Chart3D1
            .BackColor = BackColor.LightSteelBlue
            .ChartGroups(0).Elevation.IsShaded = True
            .ChartGroups(0).Elevation.MeshBottomColor = .ChartGroups(0).Elevation.MeshBottomColor.LavenderBlush
            .ChartGroups(0).Elevation.ShadedTopColor = .ChartGroups(0).Elevation.MeshBottomColor.LawnGreen
            .ChartGroups(0).ShouldSerializeBar()




            'Areastyle
            With .AreaStyle
                .Border.BorderStyle = C1Chart3D1.AreaStyle.Border.BorderStyle.Raised
                .ForeColor = ForeColor.Black
                .Border.Color = Color.Green
                .Border.Rounding.All = 20
                .Border.Thickness = 4

            End With



            'ChartArea
            With .ChartArea

                .Axes("Z").AnnoRotated = True
                'sets grid line on plane Y-Z
                .Axes("Y").MajorGrid.IsOnYZPlane = True
                .Axes("X").UnitMajor = 1
                .Axes("X").Max = 1
                .AxisX.AutoMax = False
                .AxisX.Min = 0
                'set max - ie cols
                .AxisX.Max = 0


                If C1Chart3D1.ChartGroups(0).ChartData.SetGrid.ColumnCount > 1 Then
                    flagmulticol = True
                ElseIf C1Chart3D1.ChartGroups(0).ChartData.SetGrid.ColumnCount = 1 Then
                    flagmulticol = False

                End If


                If flagmulticol = True Then
                    'Set X Scale(Col width)
                    .View.XScale = 1
                    .View.YScale = 1
                    .AxisX.AutoMajor = True
                    .AxisX.AutoMajor = True
                    C1Chart3D1.ChartGroups(0).Bar.ColumnWidth = 20
                    C1Chart3D1.ChartGroups(0).Bar.RowWidth = 20


                Else
                    'Set X Scale(Col width)
                    .View.XScale = 0.2
                    .View.YScale = 2
                End If

                With C1Chart3D1.ChartGroups(0)



                End With
                If flagmulticol = True Then
                    C1Chart3D1.ChartGroups(0).Contour.IsZoned = True
                    C1Chart3D1.ChartGroups(0).Bar.ColumnWidth = 80
                    C1Chart3D1.ChartGroups(0).Bar.RowWidth = 80
                    C1Chart3D1.ChartGroups(0).Surface.ColumnMeshFilter = 0


                Else
                    C1Chart3D1.ChartGroups(0).Contour.IsZoned = True
                    C1Chart3D1.ChartGroups(0).Bar.ColumnWidth = 10
                    C1Chart3D1.ChartGroups(0).Bar.RowWidth = 100
                End If

                'C1Chart3D1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateShort

                'C1Chart3D1.ChartArea.AxisX.Visible = False
                Debug.WriteLine(C1Chart3D1.ChartArea.Axes.Count())
            End With
            'Chartarea - style
            With .ChartArea.Style
                .BackColor = BackColor.LightYellow
                .BackColor2 = .BackColor2.Gray
                .GradientStyle = GradientStyleEnum.FromCenter
            End With
            'HeaderStyle
            With .HeaderStyle
                .BackColor = BackColor.Gray()
                .BackColor2 = .BackColor2.Gold()
                .GradientStyle = GradientStyleEnum.HorizontalCenter
                .Border.BorderStyle = C1Chart3D1.AreaStyle.Border.BorderStyle.Raised
                .Border.Rounding.All = 10
            End With



        End With
        'Set default rotation for graph
        upRotX.Text = CStr(75)
        upRotZ.Text = CStr(115)
    End Sub


#End Region

#Region "  Timer Events"



    Private Sub timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Timer for rotation - animation
        If checkBox1.Checked Then
            C1Chart3D1.ChartArea.View.RotationX += angleIncrement
            If C1Chart3D1.ChartArea.View.RotationX >= 360 Then
                C1Chart3D1.ChartArea.View.RotationX = 0
            End If
            upRotX.Value = CDec(C1Chart3D1.ChartArea.View.RotationX)
        End If
        If CheckBox2.Checked Then

            C1Chart3D1.ChartArea.View.RotationZ += angleIncrement
            If C1Chart3D1.ChartArea.View.RotationZ >= 360 Then
                C1Chart3D1.ChartArea.View.RotationZ = 0
            End If
            upRotZ.Value = CDec(C1Chart3D1.ChartArea.View.RotationZ)
        End If
    End Sub 'timer1_Tick


    Private Sub checkBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles checkBox1.CheckedChanged, CheckBox2.CheckedChanged

        Timer1.Enabled = checkBox1.Checked OrElse CheckBox2.Checked
    End Sub 'checkBox1_CheckedChanged

#End Region

#Region "  Retrieval Methods"
    Private Sub GetMonthTransactionsViewAll()
        'Shows receipts,checks,adjustments and net balance
        ResetCurrChart()

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '   0            1           2          3  
        ' banknum      afnum,      afname    status,
        '    4           5           6          7          8
        '_begmonth$  $receipts,   $expend    $adjust    $Total-Net
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim tbl As DataTable
        Dim i As Int32
        Dim mtdtranslbl As String
        Try

            Dim obj As New AF_Reporting.ClassCharts
            tbl = obj.GetMonthlyAcctBalances(Me.Bankaccountnum)
            If tbl.Rows.Count < 1 Then
                MsgBox("No Data", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End If
            'Set tot rows
            totalrows = tbl.Rows.Count

            'Add values to chart
            With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                For i = 0 To tbl.Rows.Count - 1
                    'set values
                    mtdtransrcpttot = CSng(tbl.Rows(i)(5))
                    mtdtranscheckstot = CSng(tbl.Rows(i)(6))
                    mtdtransadjtot = CSng(tbl.Rows(i)(7))
                    mtdtranstot = CSng(tbl.Rows(i)(8))
                    'label - acct numbers
                    mtdtranslbl = CStr(tbl.Rows(i)(1))

                    .ColumnCount = 4

                    flagmulticol = True

                    .RowCount = tbl.Rows.Count

                    'set values of each row in chrt ($)
                    .SetValue(0, i, mtdtransrcpttot)
                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 0
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSalmon
                        .Visible = False
                    End With
                    .SetValue(1, i, mtdtranscheckstot)
                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 1
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightGreen
                        .Visible = False
                    End With
                    .SetValue(2, i, mtdtransadjtot)
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 2
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSkyBlue
                        .Visible = False
                    End With
                    .SetValue(3, i, mtdtranstot)
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 3
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.Honeydew
                        .Visible = False
                    End With

                Next



                'Set zoned chart
                Dim contour(20, 20) As Double
                C1Chart3D1.ChartGroups(0).ChartData.ContourData = contour
                C1Chart3D1.ChartGroups(0).Contour.IsZoned = True
                C1Chart3D1.ChartGroups(0).Contour.NumLevels = 50
                C1Chart3D1.ChartGroups.ContourStyles.ColorSet = ColorSetEnum.RevRainbow
                C1Chart3D1.ChartGroups(0).Contour.IsContoured = True

                'Header Title & titles
                '4d-view
                Dim p, q As Int32
                Dim rnd As Random = New Random
                With C1Chart3D1
                    For p = 0 To 20
                        For q = 0 To 20
                            contour(p, q) = rnd.NextDouble()
                        Next q
                    Next p

                    .Header.Text = "Activity Fund.Net - Reports" & vbCrLf & "Transaction Totals Per Account - Net"
                    .ChartArea.Axes("Y").Title = "Account"
                    .ChartArea.Axes("Z").Title = "Amount"

                    'Titles on X Axis
                    Dim m, n As Int32

                    If C1Chart3D1.ChartGroups.ColumnLabels.Count > 0 Then
                        For m = 0 To C1Chart3D1.ChartGroups.ColumnLabels.Count
                            n = C1Chart3D1.ChartGroups.ColumnLabels.Count
                            If n = 0 Then Exit For
                            C1Chart3D1.ChartGroups.ColumnLabels.RemoveAt(0)
                        Next m


                    End If

                    C1Chart3D1.ChartGroups.ColumnLabels.Add(0, "Rcpts").ToString()
                    C1Chart3D1.ChartGroups.ColumnLabels.Add(1, "Chks").ToString()
                    C1Chart3D1.ChartGroups.ColumnLabels.Add(2, "Adjs").ToString()
                    C1Chart3D1.ChartGroups.ColumnLabels.Add(3, "Net").ToString()

                    With C1Chart3D1.ChartArea.Axes("X")
                        .AnnoMethod = AnnotationMethodEnum.DataLabels
                    End With

                End With
                SetChart1Properties()


            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            flagmulticol = False

            tbl.Dispose()
        End Try
    End Sub

    Private Sub GetMonthTransactions()
        'Shows  net balance
        ResetCurrChart()
        Dim tbl As DataTable
        Dim i As Int32
        Dim mtdtranslbl As String
        Try

            Dim obj As New AF_Reporting.ClassCharts
            tbl = obj.GetMonthlyAcctBalances(Me.Bankaccountnum)
            If tbl.Rows.Count < 1 Then
                MsgBox("No Data", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End If
            'Set tot rows
            totalrows = tbl.Rows.Count

            'Add values to chart
            With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                For i = 0 To tbl.Rows.Count - 1
                    mtdtranstot = CSng(tbl.Rows(i)(8))
                    mtdtranslbl = CStr(tbl.Rows(i)(1))

                    .ColumnCount = 1
                    .RowCount = tbl.Rows.Count
                    'set values of each row in chrt ($)
                    .SetValue(0, i, mtdtranstot)

                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 0
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSalmon
                        .Visible = False

                    End With
                Next
                'Header Title & titles
                With C1Chart3D1
                    .Header.Text = "Activity Fund.Net - Reports" & vbCrLf & "Transaction Totals Per Account - Net"
                    .ChartArea.Axes("Y").Title = "Account"
                    .ChartArea.Axes("Z").Title = "Amount"

                End With
                'Titles on X Axis
                Dim m, n As Int32

                If C1Chart3D1.ChartGroups.ColumnLabels.Count > 0 Then
                    For m = 0 To C1Chart3D1.ChartGroups.ColumnLabels.Count
                        n = C1Chart3D1.ChartGroups.ColumnLabels.Count
                        If n = 0 Then Exit For
                        C1Chart3D1.ChartGroups.ColumnLabels.RemoveAt(0)
                    Next m
                End If
                C1Chart3D1.ChartGroups.ColumnLabels.Add(0, "Net").ToString()

                With C1Chart3D1.ChartArea.Axes("X")
                    .AnnoMethod = AnnotationMethodEnum.DataLabels
                End With

                SetChart1Properties()
                C1Chart3D1.ChartGroups(0).Contour.NumLevels = tbl.Rows.Count
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            tbl.Dispose()
        End Try
    End Sub

    Private Sub GetMonthTransactionsReceipts()
        ResetCurrChart()
        Dim tbl As DataTable
        Dim i As Int32
        Dim mtdtranslbl As String
        Try

            Dim obj As New AF_Reporting.ClassCharts
            tbl = obj.GetMonthlyAcctBalancesReceipts(Me.Bankaccountnum)
            If tbl.Rows.Count < 1 Then
                MsgBox("No Data", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End If
            'Set tot rows
            totalrows = tbl.Rows.Count

            'Add values to chart
            With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                For i = 0 To tbl.Rows.Count - 1
                    mtdtranstot = CSng(tbl.Rows(i)(5))
                    mtdtranslbl = CStr(tbl.Rows(i)(1))

                    .ColumnCount = 1
                    .RowCount = tbl.Rows.Count
                    'set values of each row in chrt ($)
                    .SetValue(0, i, mtdtranstot)

                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 0
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSalmon
                        .Visible = False

                    End With
                Next
                'Header Title & titles
                With C1Chart3D1
                    .Header.Text = "Activity Fund.Net - Reports" & vbCrLf & "Transaction Totals Per Account - Receipts"
                    .ChartArea.Axes("Y").Title = "Account"
                    .ChartArea.Axes("Z").Title = "Receipt Amount"

                End With
                'Titles on X Axis
                Dim m, n As Int32

                If C1Chart3D1.ChartGroups.ColumnLabels.Count > 0 Then
                    For m = 0 To C1Chart3D1.ChartGroups.ColumnLabels.Count
                        n = C1Chart3D1.ChartGroups.ColumnLabels.Count
                        If n = 0 Then Exit For
                        C1Chart3D1.ChartGroups.ColumnLabels.RemoveAt(0)
                    Next m
                End If
                C1Chart3D1.ChartGroups.ColumnLabels.Add(0, "Receipts").ToString()

                With C1Chart3D1.ChartArea.Axes("X")
                    .AnnoMethod = AnnotationMethodEnum.DataLabels
                End With
                SetChart1Properties()
                C1Chart3D1.ChartGroups(0).Contour.NumLevels = tbl.Rows.Count
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            tbl.Dispose()
        End Try
    End Sub
    'Shows receipts

    Private Sub GetMonthTransactionsExpense()
        'Shows checks
        ResetCurrChart()
        Dim tbl As DataTable
        Dim i As Int32
        Dim mtdtranslbl As String
        Try

            Dim obj As New AF_Reporting.ClassCharts
            tbl = obj.GetMonthlyAcctBalancesExpense(Me.Bankaccountnum)
            If tbl.Rows.Count < 1 Then
                MsgBox("No Data", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End If
            'Set tot rows
            totalrows = tbl.Rows.Count

            'Add values to chart
            With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                For i = 0 To tbl.Rows.Count - 1
                    mtdtranstot = CSng(tbl.Rows(i)(6))
                    mtdtranslbl = CStr(tbl.Rows(i)(1))

                    .ColumnCount = 1
                    .RowCount = tbl.Rows.Count
                    'set values of each row in chrt ($)
                    .SetValue(0, i, mtdtranstot)

                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 0
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSalmon
                        .Visible = False

                    End With
                Next
                'Header Title & titles
                With C1Chart3D1
                    .Header.Text = "Activity Fund.Net - Reports" & vbCrLf & "Transaction Totals Per Account - Checks"
                    .ChartArea.Axes("Y").Title = "Account"
                    .ChartArea.Axes("Z").Title = "Check Amount"

                End With
                'Titles on X Axis
                Dim m, n As Int32

                If C1Chart3D1.ChartGroups.ColumnLabels.Count > 0 Then
                    For m = 0 To C1Chart3D1.ChartGroups.ColumnLabels.Count
                        n = C1Chart3D1.ChartGroups.ColumnLabels.Count
                        If n = 0 Then Exit For
                        C1Chart3D1.ChartGroups.ColumnLabels.RemoveAt(0)
                    Next m
                End If
                C1Chart3D1.ChartGroups.ColumnLabels.Add(0, "Checks").ToString()

                With C1Chart3D1.ChartArea.Axes("X")
                    .AnnoMethod = AnnotationMethodEnum.DataLabels
                End With
                SetChart1Properties()
                C1Chart3D1.ChartGroups(0).Contour.NumLevels = tbl.Rows.Count
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            tbl.Dispose()
        End Try
    End Sub

    Private Sub GetMonthTransactionsAdjusments()
        'Shows adjustment transactions for the month

        ResetCurrChart()
        Dim tbl As DataTable
        Dim i As Int32
        Dim mtdtranslbl As String
        Try

            Dim obj As New AF_Reporting.ClassCharts
            tbl = obj.GetMonthlyAcctBalancesAdjustments(Me.Bankaccountnum)
            If tbl.Rows.Count < 1 Then
                MsgBox("No Data", MsgBoxStyle.Information, MSGTITLE)
                Exit Sub
            End If
            'Set tot rows
            totalrows = tbl.Rows.Count

            'Add values to chart
            With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                For i = 0 To tbl.Rows.Count - 1
                    mtdtranstot = CSng(tbl.Rows(i)(7))
                    mtdtranslbl = CStr(tbl.Rows(i)(1))

                    .ColumnCount = 1
                    .RowCount = tbl.Rows.Count
                    'set values of each row in chrt ($)
                    .SetValue(0, i, mtdtranstot)

                    'set Labels of each row in chrt
                    With C1Chart3D1.ChartLabels.LabelsCollection.AddNewLabel
                        .AttachMethod = AttachMethodEnum.DataIndex
                        .AttachMethodData.Column = 0
                        .AttachMethodData.Row = i
                        .Connected = True
                        .Offset = 70
                        .LabelCompass = LabelCompassEnum.NorthEast
                        .Text = "Account " + ControlChars.Lf + mtdtranslbl
                        .Style.BackColor = Color.LightSalmon
                        .Visible = False

                    End With
                Next
                'Header Title & titles
                With C1Chart3D1
                    .Header.Text = "Activity Fund.Net - Reports" & vbCrLf & "Transaction Totals Per Account - Adjustments"
                    .ChartArea.Axes("Y").Title = "Account"
                    .ChartArea.Axes("Z").Title = "Adjustment Amount"

                End With
                'Titles on X Axis
                Dim m, n As Int32

                If C1Chart3D1.ChartGroups.ColumnLabels.Count > 0 Then
                    For m = 0 To C1Chart3D1.ChartGroups.ColumnLabels.Count
                        n = C1Chart3D1.ChartGroups.ColumnLabels.Count
                        If n = 0 Then Exit For
                        C1Chart3D1.ChartGroups.ColumnLabels.RemoveAt(0)
                    Next m
                End If
                C1Chart3D1.ChartGroups.ColumnLabels.Add(0, "Adjustments").ToString()

                With C1Chart3D1.ChartArea.Axes("X")
                    .AnnoMethod = AnnotationMethodEnum.DataLabels
                End With
                SetChart1Properties()
                C1Chart3D1.ChartGroups(0).Contour.NumLevels = tbl.Rows.Count
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            tbl.Dispose()
        End Try
    End Sub

#End Region

#Region "  Menu Events"



    Private Sub mmnuChartAllMonthTransactions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mmnuChartAllMonthTransactions.Click
        Me.C1Chart3D1.Visible = True
        GetMonthTransactionsViewAll()
        Me.txtChartDetails.Text = mmnuChartAllMonthTransactions.Text

    End Sub

    Private Sub mnuChartNetMonthTransaction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChartNetMonthTransaction.Click

        Me.C1Chart3D1.Visible = True
        GetMonthTransactions()
        Me.txtChartDetails.Text = mnuChartNetMonthTransaction.Text
    End Sub
    Private Sub mnyChartMonthTransactionsReceipts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnyChartMonthTransactionsReceipts.Click
        Me.C1Chart3D1.Visible = True
        GetMonthTransactionsReceipts()
        Me.txtChartDetails.Text = mnyChartMonthTransactionsReceipts.Text
    End Sub
    Private Sub mnuChartsMonthTransactionChecks_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChartsMonthTransactionChecks.Click
        Me.C1Chart3D1.Visible = True
        GetMonthTransactionsExpense()
        Me.txtChartDetails.Text = mnuChartsMonthTransactionChecks.Text

    End Sub
    Private Sub mnuChartMonthlyTransactionAdjustments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChartMonthlyTransactionAdjustments.Click
        Me.C1Chart3D1.Visible = True
        GetMonthTransactionsAdjusments()
        Me.txtChartDetails.Text = mnuChartMonthlyTransactionAdjustments.Text
    End Sub



#End Region

#Region "  Color Charts"

    Private Sub miContour_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem19.Click, MenuItem20.Click, MenuItem21.Click, MenuItem22.Click, MenuItem23.Click, MenuItem33.Click, MenuItem34.Click, MenuItem35.Click, MenuItem36.Click, MenuItem37.Click, MenuItem38.Click
        'Set colors on Charts
        Dim mi As MenuItem = CType(sender, MenuItem)
        Dim _mi As MenuItem
        For Each _mi In cmContour.MenuItems
            _mi.Checked = False
        Next _mi
        mi.Checked = True

        Dim i As Integer = Me.cmContour.MenuItems.IndexOf(mi)

        Select Case i
            Case 0
                If C1Chart3D1.ChartGroups(0).Contour.IsZoned Then
                    C1Chart3D1.ChartGroups(0).Contour.IsZoned = False
                End If
            Case Else
                setZoneChart(CType(i, C1.Win.C1Chart3D.ColorSetEnum))
        End Select
    End Sub
    Sub setZoneChart(ByVal clrset As C1.Win.C1Chart3D.ColorSetEnum)
        If Not C1Chart3D1.ChartGroups(0).Contour.IsZoned Then
            C1Chart3D1.ChartGroups(0).Contour.IsZoned = True
        End If
        C1Chart3D1.ChartGroups.ContourStyles.ColorSet = clrset
    End Sub
#End Region

#Region "  Toolbar Events"

    Private Sub tlbChart_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tblChart.ButtonClick
        Select Case tblChart.Buttons.IndexOf(e.Button)
            Case 0
                selectNext(cmContour)
            Case 2
                C1Chart3D1.ChartGroups(0).Elevation.DropLines = Not C1Chart3D1.ChartGroups(0).Elevation.DropLines
            Case 4
                C1Chart3D1.ChartGroups(0).Elevation.IsMeshed = Not C1Chart3D1.ChartGroups(0).Elevation.IsMeshed
            Case 6
                selectNext(cmView)
            Case 8
                C1Chart3D1.ChartArea.View.IsInteractive = Not C1Chart3D1.ChartArea.View.IsInteractive
            Case 10
                setHoles(setGrid, True) ' menuItem9.Checked)
                setHoles(CType(setIrGrid, Chart3DDataSetGrid), True) ', menuItem9.Checked)
            Case Else
        End Select
    End Sub

    Private Sub miView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem4.Click, MenuItem7.Click
        'View 2d or 3d

        Dim mi As MenuItem = CType(sender, MenuItem)
        Dim _mi As MenuItem
        For Each _mi In cmView.MenuItems
            _mi.Checked = False
        Next _mi
        mi.Checked = True

        Select Case Me.cmView.MenuItems.IndexOf(mi)
            Case 0
                C1Chart3D1.ChartArea.View.View3D = C1.Win.C1Chart3D.View3DEnum.Default
                Me.btnResetImage.PerformClick()

                C1Chart3D1.Refresh()

            Case 2
                C1Chart3D1.ChartArea.View.View3D = C1.Win.C1Chart3D.View3DEnum.YZ_2D_Pos
                Me.btnResetImage.PerformClick()
                C1Chart3D1.Refresh()
            Case Else
        End Select
    End Sub

    Sub selectNext(ByVal cm As System.Windows.Forms.ContextMenu)
        'Context Menu Selection
        Dim sel As Integer = 0
        Dim i As Integer
        For i = 0 To cm.MenuItems.Count - 1
            If cm.MenuItems(i).Checked Then
                sel = i
                Exit For
            End If
        Next i

        If sel >= cm.MenuItems.Count - 1 Then
            sel = 0
        Else
            sel += 1
        End If
        If cm.MenuItems(sel).Text.Equals("-") Then
            If sel >= cm.MenuItems.Count - 1 Then
                sel = 0
            Else
                sel += 1
            End If
        End If
        cm.MenuItems(sel).PerformClick()
    End Sub


    Sub setHoles(ByVal grset As Chart3DDataSetGrid, ByVal enable As Boolean)
        Dim i, j As Integer
        If enable Then
            For i = 0 To C1Chart3D1.ChartGroups(0).ChartData.SetGrid.ColumnCount - 1
                For j = 0 To C1Chart3D1.ChartGroups(0).ChartData.SetGrid.RowCount - 1
                    Dim x As Single
                    Dim y As Single

                    With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                        If TypeOf grset Is Chart3DDataSetIrGrid Then

                            Dim s As Chart3DDataSetIrGrid = CType(C1Chart3D1.ChartGroups(0).ChartData.SetGrid, Chart3DDataSetIrGrid)

                            x = CSng(s.GetColumnValue(i))
                            y = CSng(s.GetRowValue(j))
                        Else
                            x = CSng(.MinX + i * C1Chart3D1.ChartGroups(0).ChartData.SetGrid.RowDelta)
                            y = CSng(.MinY + j * C1Chart3D1.ChartGroups(0).ChartData.SetGrid.ColumnDelta)
                        End If
                    End With

                    With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                        Dim val As Single = 0.1F * x * x - 0.4F * y * y
                        If i = j OrElse i = .RowCount - j Then
                            C1Chart3D1.ChartGroups(0).ChartData.SetGrid.SetValue(i, j, val)
                        End If
                    End With
                Next j
            Next i
        Else
            For i = 0 To C1Chart3D1.ChartGroups(0).ChartData.SetGrid.ColumnCount - 1
                For j = 0 To C1Chart3D1.ChartGroups(0).ChartData.SetGrid.RowCount - 1
                    With C1Chart3D1.ChartGroups(0).ChartData.SetGrid
                        If i = j OrElse i = .RowCount - j Then
                            .SetValue(i, j, C1Chart3D1.ChartGroups(0).ChartData.SetGrid.Hole)
                        End If
                    End With
                Next j
            Next i
        End If
    End Sub

#End Region



    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Me.Dispose()

    End Sub


    Private Sub MenuItem5_Select(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem5.Select
        panel1.Enabled = True
        tblChart.Enabled = True
    End Sub

    Private Sub frmChartsMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        panel1.Enabled = False
        tblChart.Enabled = False

    End Sub
End Class


