Imports C1.C1PrintDocument
Imports C1.Win.C1FlexGrid
Imports System.Data
Imports System.Data.SqlClient

Public Class frmFinancialReports
    Inherits System.Windows.Forms.Form

#Region "  Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Dim authobj As AF_Master.Authuser
        Try
            Me.ConnectionString = authobj.ConnectionString
            Me.FiscalYear = authobj.FiscalYear
            Me.FiscalMonthStr = authobj.CurrentMonthString
            Me.SchoolName = authobj.SchoolName
            Me.SchoolAddress1 = authobj.SchoolAddress1
            Me.SchoolAddress2 = authobj.SchoolAddress2
            Me.UseOcas = authobj.UseOCAS
            Me.UserName = authobj.UserFullname
            Dim city, state, zip As String
            city = authobj.SchoolCity
            state = authobj.SchoolState
            zip = authobj.SchoolZipCode
            Me.SchoolCityStateZip = city & ", " & state & " " & zip
            Me.GridDetail.Visible = False
            Me.GridTotals.Visible = False
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
    Friend WithEvents Prev1 As C1.Win.C1PrintPreview.C1PrintPreview
    Friend WithEvents c1pBtnFileOpen1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnFileSave1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnFilePrint1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnPageSetup1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnReflow1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnStop1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnDocInfo1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnShowNavigationBar1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator2 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnMouseHand1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnMouseZoom1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnMouseZoomOut1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnMouseSelect1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnFindText1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator3 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnGoFirst1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnGoPrev1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnGoNext1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnGoLast1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator4 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnHistoryPrev1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnHistoryNext1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator5 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnZoomOut1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnZoomIn1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator6 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnViewActualSize1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnViewFullPage1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnViewPageWidth1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnViewTwoPages1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnViewFourPages1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnSeparator7 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents c1pBtnHelp1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents Doc1 As C1.C1PrintDocument.C1PrintDocument
    Friend WithEvents GridDetail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridTotals As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridWrk As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridWrkTotals As C1.Win.C1FlexGrid.C1FlexGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFinancialReports))
        Me.Prev1 = New C1.Win.C1PrintPreview.C1PrintPreview
        Me.Doc1 = New C1.C1PrintDocument.C1PrintDocument
        Me.c1pBtnFileOpen1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnFileSave1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnFilePrint1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnPageSetup1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnReflow1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnStop1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnDocInfo1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnShowNavigationBar1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator2 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnMouseHand1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnMouseZoom1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnMouseZoomOut1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnMouseSelect1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnFindText1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator3 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnGoFirst1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnGoPrev1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnGoNext1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnGoLast1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator4 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnHistoryPrev1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnHistoryNext1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator5 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnZoomOut1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnZoomIn1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator6 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnViewActualSize1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnViewFullPage1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnViewPageWidth1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnViewTwoPages1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnViewFourPages1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnSeparator7 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.c1pBtnHelp1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.GridDetail = New C1.Win.C1FlexGrid.C1FlexGrid
        Me.GridTotals = New C1.Win.C1FlexGrid.C1FlexGrid
        Me.GridWrk = New C1.Win.C1FlexGrid.C1FlexGrid
        Me.GridWrkTotals = New C1.Win.C1FlexGrid.C1FlexGrid
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridWrk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridWrkTotals, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Prev1
        '
        Me.Prev1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Prev1.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIA"
        Me.Prev1.Document = Me.Doc1
        Me.Prev1.Location = New System.Drawing.Point(0, 16)
        Me.Prev1.Name = "Prev1"
        Me.Prev1.NavigationBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.NavigationBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.NavigationBar.OutlineView.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.NavigationBar.OutlineView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.NavigationBar.OutlineView.Indent = 19
        Me.Prev1.NavigationBar.OutlineView.ItemHeight = 16
        Me.Prev1.NavigationBar.OutlineView.TabIndex = 0
        Me.Prev1.NavigationBar.OutlineView.Visible = False
        Me.Prev1.NavigationBar.Padding = New System.Drawing.Point(6, 3)
        Me.Prev1.NavigationBar.TabIndex = 2
        Me.Prev1.NavigationBar.ThumbnailsView.AutoArrange = True
        Me.Prev1.NavigationBar.ThumbnailsView.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.NavigationBar.ThumbnailsView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.NavigationBar.ThumbnailsView.TabIndex = 0
        Me.Prev1.NavigationBar.ThumbnailsView.Visible = False
        Me.Prev1.NavigationBar.Visible = False
        Me.Prev1.NavigationBar.Width = 160
        Me.Prev1.Size = New System.Drawing.Size(656, 352)
        Me.Prev1.Splitter.Cursor = System.Windows.Forms.Cursors.VSplit
        Me.Prev1.Splitter.Width = 3
        Me.Prev1.StatusBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.StatusBar.TabIndex = 4
        Me.Prev1.TabIndex = 1
        Me.Prev1.ToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.c1pBtnFileOpen1, Me.c1pBtnFileSave1, Me.c1pBtnFilePrint1, Me.c1pBtnPageSetup1, Me.c1pBtnReflow1, Me.c1pBtnStop1, Me.c1pBtnDocInfo1, Me.c1pBtnSeparator1, Me.c1pBtnShowNavigationBar1, Me.c1pBtnSeparator2, Me.c1pBtnMouseHand1, Me.c1pBtnMouseZoom1, Me.c1pBtnMouseZoomOut1, Me.c1pBtnMouseSelect1, Me.c1pBtnFindText1, Me.c1pBtnSeparator3, Me.c1pBtnGoFirst1, Me.c1pBtnGoPrev1, Me.c1pBtnGoNext1, Me.c1pBtnGoLast1, Me.c1pBtnSeparator4, Me.c1pBtnHistoryPrev1, Me.c1pBtnHistoryNext1, Me.c1pBtnSeparator5, Me.c1pBtnZoomOut1, Me.c1pBtnZoomIn1, Me.c1pBtnSeparator6, Me.c1pBtnViewActualSize1, Me.c1pBtnViewFullPage1, Me.c1pBtnViewPageWidth1, Me.c1pBtnViewTwoPages1, Me.c1pBtnViewFourPages1, Me.c1pBtnSeparator7, Me.c1pBtnHelp1})
        Me.Prev1.ToolBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.ToolBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Doc1
        '
        Me.Doc1.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIA"
        Me.Doc1.ColumnSpacingStr = "0.5in"
        Me.Doc1.ColumnSpacingUnit.DefaultType = True
        Me.Doc1.ColumnSpacingUnit.UnitValue = "0.5in"
        Me.Doc1.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Inch
        Me.Doc1.DocumentName = ""
        '
        'c1pBtnFileOpen1
        '
        Me.c1pBtnFileOpen1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileOpen
        Me.c1pBtnFileOpen1.ImageIndex = 0
        Me.c1pBtnFileOpen1.ToolTipText = "File Open"
        '
        'c1pBtnFileSave1
        '
        Me.c1pBtnFileSave1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileSave
        Me.c1pBtnFileSave1.ImageIndex = 1
        Me.c1pBtnFileSave1.ToolTipText = "File Save"
        '
        'c1pBtnFilePrint1
        '
        Me.c1pBtnFilePrint1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FilePrint
        Me.c1pBtnFilePrint1.ImageIndex = 2
        Me.c1pBtnFilePrint1.ToolTipText = "Print"
        '
        'c1pBtnPageSetup1
        '
        Me.c1pBtnPageSetup1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.PageSetup
        Me.c1pBtnPageSetup1.ImageIndex = 3
        Me.c1pBtnPageSetup1.ToolTipText = "Page Setup"
        '
        'c1pBtnReflow1
        '
        Me.c1pBtnReflow1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Reflow
        Me.c1pBtnReflow1.ImageIndex = 4
        Me.c1pBtnReflow1.ToolTipText = "Reflow"
        '
        'c1pBtnStop1
        '
        Me.c1pBtnStop1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Stop
        Me.c1pBtnStop1.ImageIndex = 5
        Me.c1pBtnStop1.ToolTipText = "Stop"
        Me.c1pBtnStop1.Visible = False
        '
        'c1pBtnDocInfo1
        '
        Me.c1pBtnDocInfo1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.DocInfo
        Me.c1pBtnDocInfo1.ImageIndex = 26
        Me.c1pBtnDocInfo1.ToolTipText = "Document information"
        '
        'c1pBtnSeparator1
        '
        Me.c1pBtnSeparator1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'c1pBtnShowNavigationBar1
        '
        Me.c1pBtnShowNavigationBar1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ShowNavigationBar
        Me.c1pBtnShowNavigationBar1.ImageIndex = 6
        Me.c1pBtnShowNavigationBar1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnShowNavigationBar1.ToolTipText = "Show Navigation Bar"
        '
        'c1pBtnSeparator2
        '
        Me.c1pBtnSeparator2.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'c1pBtnMouseHand1
        '
        Me.c1pBtnMouseHand1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseHand
        Me.c1pBtnMouseHand1.ImageIndex = 7
        Me.c1pBtnMouseHand1.Pushed = True
        Me.c1pBtnMouseHand1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnMouseHand1.ToolTipText = "Hand Tool"
        '
        'c1pBtnMouseZoom1
        '
        Me.c1pBtnMouseZoom1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseZoom
        Me.c1pBtnMouseZoom1.ImageIndex = 8
        Me.c1pBtnMouseZoom1.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.c1pBtnMouseZoom1.ToolTipText = "Zoom In Tool"
        '
        'c1pBtnMouseZoomOut1
        '
        Me.c1pBtnMouseZoomOut1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseZoomOut
        Me.c1pBtnMouseZoomOut1.ImageIndex = 25
        Me.c1pBtnMouseZoomOut1.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.c1pBtnMouseZoomOut1.ToolTipText = "Zoom Out Tool"
        Me.c1pBtnMouseZoomOut1.Visible = False
        '
        'c1pBtnMouseSelect1
        '
        Me.c1pBtnMouseSelect1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseSelect
        Me.c1pBtnMouseSelect1.ImageIndex = 9
        Me.c1pBtnMouseSelect1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnMouseSelect1.ToolTipText = "Select Text"
        '
        'c1pBtnFindText1
        '
        Me.c1pBtnFindText1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FindText
        Me.c1pBtnFindText1.ImageIndex = 10
        Me.c1pBtnFindText1.ToolTipText = "Find Text"
        '
        'c1pBtnSeparator3
        '
        Me.c1pBtnSeparator3.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'c1pBtnGoFirst1
        '
        Me.c1pBtnGoFirst1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoFirst
        Me.c1pBtnGoFirst1.Enabled = False
        Me.c1pBtnGoFirst1.ImageIndex = 11
        Me.c1pBtnGoFirst1.ToolTipText = "First Page"
        '
        'c1pBtnGoPrev1
        '
        Me.c1pBtnGoPrev1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoPrev
        Me.c1pBtnGoPrev1.Enabled = False
        Me.c1pBtnGoPrev1.ImageIndex = 12
        Me.c1pBtnGoPrev1.ToolTipText = "Previous Page"
        '
        'c1pBtnGoNext1
        '
        Me.c1pBtnGoNext1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoNext
        Me.c1pBtnGoNext1.ImageIndex = 13
        Me.c1pBtnGoNext1.ToolTipText = "Next Page"
        '
        'c1pBtnGoLast1
        '
        Me.c1pBtnGoLast1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoLast
        Me.c1pBtnGoLast1.ImageIndex = 14
        Me.c1pBtnGoLast1.ToolTipText = "Last Page"
        '
        'c1pBtnSeparator4
        '
        Me.c1pBtnSeparator4.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'c1pBtnHistoryPrev1
        '
        Me.c1pBtnHistoryPrev1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryPrev
        Me.c1pBtnHistoryPrev1.Enabled = False
        Me.c1pBtnHistoryPrev1.ImageIndex = 15
        Me.c1pBtnHistoryPrev1.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.c1pBtnHistoryPrev1.ToolTipText = "Previous View"
        '
        'c1pBtnHistoryNext1
        '
        Me.c1pBtnHistoryNext1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryNext
        Me.c1pBtnHistoryNext1.Enabled = False
        Me.c1pBtnHistoryNext1.ImageIndex = 16
        Me.c1pBtnHistoryNext1.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.c1pBtnHistoryNext1.ToolTipText = "Next View"
        '
        'c1pBtnSeparator5
        '
        Me.c1pBtnSeparator5.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.c1pBtnSeparator5.Visible = False
        '
        'c1pBtnZoomOut1
        '
        Me.c1pBtnZoomOut1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomOut
        Me.c1pBtnZoomOut1.ImageIndex = 17
        Me.c1pBtnZoomOut1.ToolTipText = "Zoom Out"
        Me.c1pBtnZoomOut1.Visible = False
        '
        'c1pBtnZoomIn1
        '
        Me.c1pBtnZoomIn1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomIn
        Me.c1pBtnZoomIn1.ImageIndex = 18
        Me.c1pBtnZoomIn1.ToolTipText = "Zoom In"
        Me.c1pBtnZoomIn1.Visible = False
        '
        'c1pBtnSeparator6
        '
        Me.c1pBtnSeparator6.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator6.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.c1pBtnSeparator6.Visible = False
        '
        'c1pBtnViewActualSize1
        '
        Me.c1pBtnViewActualSize1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewActualSize
        Me.c1pBtnViewActualSize1.ImageIndex = 19
        Me.c1pBtnViewActualSize1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnViewActualSize1.ToolTipText = "Actual Size"
        '
        'c1pBtnViewFullPage1
        '
        Me.c1pBtnViewFullPage1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFullPage
        Me.c1pBtnViewFullPage1.ImageIndex = 20
        Me.c1pBtnViewFullPage1.Pushed = True
        Me.c1pBtnViewFullPage1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnViewFullPage1.ToolTipText = "Full Page"
        '
        'c1pBtnViewPageWidth1
        '
        Me.c1pBtnViewPageWidth1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewPageWidth
        Me.c1pBtnViewPageWidth1.ImageIndex = 21
        Me.c1pBtnViewPageWidth1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnViewPageWidth1.ToolTipText = "Page Width"
        '
        'c1pBtnViewTwoPages1
        '
        Me.c1pBtnViewTwoPages1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewTwoPages
        Me.c1pBtnViewTwoPages1.ImageIndex = 22
        Me.c1pBtnViewTwoPages1.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.c1pBtnViewTwoPages1.ToolTipText = "Two Pages"
        '
        'c1pBtnViewFourPages1
        '
        Me.c1pBtnViewFourPages1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFourPages
        Me.c1pBtnViewFourPages1.ImageIndex = 23
        Me.c1pBtnViewFourPages1.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.c1pBtnViewFourPages1.ToolTipText = "Four Pages"
        '
        'c1pBtnSeparator7
        '
        Me.c1pBtnSeparator7.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.c1pBtnSeparator7.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.c1pBtnSeparator7.Visible = False
        '
        'c1pBtnHelp1
        '
        Me.c1pBtnHelp1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Help
        Me.c1pBtnHelp1.ImageIndex = 24
        Me.c1pBtnHelp1.ToolTipText = "Help"
        Me.c1pBtnHelp1.Visible = False
        '
        'GridDetail
        '
        Me.GridDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridDetail.BackColor = System.Drawing.SystemColors.Window
        Me.GridDetail.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridDetail.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GridDetail.Location = New System.Drawing.Point(0, 0)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.Size = New System.Drawing.Size(656, 368)
        Me.GridDetail.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridDetail.TabIndex = 2
        Me.GridDetail.Visible = False
        '
        'GridTotals
        '
        Me.GridTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridTotals.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GridTotals.Location = New System.Drawing.Point(0, 8)
        Me.GridTotals.Name = "GridTotals"
        Me.GridTotals.Size = New System.Drawing.Size(656, 232)
        Me.GridTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridTotals.TabIndex = 3
        Me.GridTotals.Visible = False
        '
        'GridWrk
        '
        Me.GridWrk.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridWrk.BackColor = System.Drawing.SystemColors.Window
        Me.GridWrk.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridWrk.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GridWrk.Location = New System.Drawing.Point(0, 46)
        Me.GridWrk.Name = "GridWrk"
        Me.GridWrk.Size = New System.Drawing.Size(656, 280)
        Me.GridWrk.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrk.TabIndex = 4
        Me.GridWrk.Visible = False
        '
        'GridWrkTotals
        '
        Me.GridWrkTotals.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridWrkTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridWrkTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridWrkTotals.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GridWrkTotals.Location = New System.Drawing.Point(8, 54)
        Me.GridWrkTotals.Name = "GridWrkTotals"
        Me.GridWrkTotals.Size = New System.Drawing.Size(656, 280)
        Me.GridWrkTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrkTotals.TabIndex = 5
        Me.GridWrkTotals.Visible = False
        '
        'frmFinancialReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 373)
        Me.Controls.Add(Me.Prev1)
        Me.Controls.Add(Me.GridWrkTotals)
        Me.Controls.Add(Me.GridTotals)
        Me.Controls.Add(Me.GridWrk)
        Me.Controls.Add(Me.GridDetail)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(300, 200)
        Me.Name = "frmFinancialReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Financial Reporting"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridWrk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridWrkTotals, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  C1Doc Events "

    Private Sub Doc1_NewPageStarted(ByVal sender As C1.C1PrintDocument.C1PrintDocument, ByVal e As C1.C1PrintDocument.NewPageStartedEventArgs) Handles Doc1.NewPageStarted
        Select Case Me.DocumentName
            Case Else
                PrintHeader()
        End Select
    End Sub

#End Region

#Region "  Class Members "

    'styles
    Private headerstyle As C1DocStyle
    Private footerstyle As C1DocStyle
    Private arialleft8 As C1DocStyle
    Private arialright8 As C1DocStyle
    Private arialleft10 As C1DocStyle
    Private arialleft10bold As C1DocStyle
    Private arialright10 As C1DocStyle
    Private arialright10bold As C1DocStyle
    Private timesleft16 As C1DocStyle
    Private verdanaleft8 As C1DocStyle
    Private verdanaright8 As C1DocStyle
    Private verdanaleft8bold As C1DocStyle
    Private verdanaright8bold As C1DocStyle
    Private verdanaleft10 As C1DocStyle
    Private verdanaright10 As C1DocStyle
    Private verdanaleft10bold As C1DocStyle
    Private verdanaright10bold As C1DocStyle
    'header values
    Private CellMiddleBottom As String = ""
    Private CellMiddleMiddle As String = ""
    Private CellMiddleTop As String = ""
    Private CellRightBottom As String = ""
    Private CellRightMiddle As String = ""
    Private CellRightTop As String = ""

    'styles
    Private docstyle As C1DocStyle
    Private linestyle1 As C1DocStyle
    Private linestyle2 As C1DocStyle
    Private linestyle3 As C1DocStyle
    Private linestyle4 As C1DocStyle
    Private linestyle5 As C1DocStyle
    'property vars
    Private _balanceforwardamount As Double
    Private _balanceforwardcount As Int32
    Private _bankaccountnumber As String
    Private _boldcodefilepath As String
    Private _connectionstring As String
    Private _countyid As String
    Private _districtid As String
    Private _documentname As String
    Private _fiscalyear As Int32
    Private _fiscalmonthstr As String
    Private _haserrors As Boolean = False
    Private MSGTITLE As String = "Activity Fund Reports"
    Private _reportname As String
    Private _schoolname As String
    Private _schooladdress1 As String
    Private _schooladdress2 As String
    Private _schoolcitystatezip As String
    Private _useocas As Boolean
    Private _username As String
    Private cn As SqlConnection

#End Region

#Region "  Delegates & Events "

    Public Event RecordsProcessed(ByVal ecurrentrecord As Int32, ByVal erecordcount As Int32)

    Private Sub EventRecordProcessed(ByVal ecurrentrecord As Int32, ByVal erecordcount As Int32)
        'statusbar update on mainmenu...
        Dim i As Int32
        If i Mod 5 = 0 Then
            RaiseEvent RecordsProcessed(ecurrentrecord, erecordcount)
            Application.DoEvents()
        End If
    End Sub

#End Region

#Region "  Document Settings & Styles "

    Private Sub DefineDocumentSettings(ByVal edocumentname As String)
        With Me.Doc1
            Select Case edocumentname
                Case "DetailOfAccountsSingleAccount", "AdjustmentRegister", "AdjustmentTicket", "TransferRegister", "TransferTicket", "Reconciliation"
                    .DefaultUnit = UnitTypeEnum.Mm
                    .DefaultUnitOfFrames = UnitTypeEnum.Mm
                    .PageSettings.Landscape = False
                    .PageSettings.Margins.Left = 50
                    .PageSettings.Margins.Top = 10
                    .PageSettings.Margins.Right = 50
                    .PageSettings.Margins.Bottom = 20
                    .PageFooter.Height = 10
                Case "EncumbranceDetailOfAccountsSingleAccount"
                    .DefaultUnit = UnitTypeEnum.Mm
                    .DefaultUnitOfFrames = UnitTypeEnum.Mm
                    .PageSettings.Landscape = True
                    .PageSettings.Margins.Left = 50
                    .PageSettings.Margins.Top = 10
                    .PageSettings.Margins.Right = 50
                    .PageSettings.Margins.Bottom = 20
                    .PageFooter.Height = 10
                    '.PageFooter.RenderText.Style.Font = New Font("Verdana", 8, FontStyle.Regular)
                    '.PageFooter.RenderText.Text = "Page [@@PageNo@@] of [@@PageCount@@]"
                Case Else
            End Select
        End With

        DefineFooterStyle()

    End Sub

    Private Sub DefineFooterStyle()
        'style for the footer
        With footerstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .TextColor = Color.Gray
        End With
    End Sub

    Private Sub DefineHeaderStyle(ByVal edocumentname As String)
        Select Case edocumentname
            Case "DetailOfAccountsSingleAccount"
                With headerstyle
                    'style for header and footer of document
                    .BackColor = Color.WhiteSmoke
                    .Font = New Font("Arial", 8, FontStyle.Regular)
                    .LineSpacing = 125
                    .Spacing.TopStr = "2mm"
                    .Spacing.BottomStr = "1mm"
                End With
            Case Else

        End Select
    End Sub

    Private Sub DefineStyles()
        With Me.arialleft8
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With
        With Me.arialright8
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
        With Me.arialleft10
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 10, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With
        With Me.arialleft10bold
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With
        With Me.arialright10
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 10, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
        With Me.timesleft16
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Times", 16, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextColor = Color.Salmon
        End With
        With Me.verdanaleft8
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
        End With
        With Me.verdanaleft8bold
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 8, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
        End With
        With Me.verdanaright8
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
        With Me.verdanaright8bold
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 8, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
        With Me.verdanaleft10
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 10, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
        End With
        With Me.verdanaleft10bold
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 10, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
        End With
        With Me.verdanaright10
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 10, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
        With Me.verdanaright10bold
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 10, FontStyle.Bold)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Right
        End With
    End Sub

#End Region

#Region "  Methods Generation "

    Public Function GenerateAdjustmentRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        'this method retrieves all transactions for a single bank and fiscal year
        'given the specified date range or number range;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4       
        '   bank       fisyr    docnumber  type    amount
        '     5           6         7        8        9 
        '   acct      subacct     xcode   rcode    applied 
        '    10          11        12 
        ' created      descr     remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT bank_acct_num, tran_fisyr, tran_autoinc_key, tran_type," _
            & " tran_amt, af_acct_num, as_acct_num, ocex_code, ocrv_code," _
            & " tran_applied_date, tran_datetime, tran_descr, tran_remarks" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " AND tran_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY tran_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Dim starting, ending As Int32
            Try
                'validate the numbers
                starting = CInt(enumberfrom)
                ending = CInt(enumberto)
            Catch ex As Exception
                Throw New ArgumentException("The beginning or ending number is missing or invalid.")
            End Try

            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT bank_acct_num, tran_fisyr, tran_autoinc_key, tran_type," _
            & " tran_amt, af_acct_num, as_acct_num, ocex_code, ocrv_code, tran_applied_date," _
            & " tran_datetime, tran_descr, tran_remarks" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " AND tran_autoinc_key BETWEEN @p3 AND @p4" _
            & " ORDER BY tran_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", starting)
            cmd.Parameters.Add("@p4", ending)
        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
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
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        Try
            'summarise the totals by transaction type
            Dim amount, bankamt, expamt, intamt, nsfamt, revamt, legchkamt, legrcptamt As Double
            Dim index As Int32
            Dim type As String
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    type = DirectCast(.GetData(index, 3), String).ToUpper
                    amount = CDbl(.GetData(index, 4))
                    Select Case type
                        Case "B"
                            bankamt += amount
                        Case "E"
                            expamt += amount
                        Case "I"
                            intamt += amount
                        Case "J"
                            legrcptamt += amount
                        Case "K"
                            legchkamt += amount
                        Case "N"
                            nsfamt += amount
                        Case "R"
                            revamt += amount
                    End Select
                Next
            End With
            With Me.GridTotals
                .Rows.Count = 0
                .Cols.Count = 8
                .Rows.Add()
                amount = (expamt - (bankamt + legchkamt)) + (intamt + nsfamt + revamt + legrcptamt)
                .SetData(0, 0, bankamt)
                .SetData(0, 1, expamt)
                .SetData(0, 2, intamt)
                .SetData(0, 3, nsfamt)
                .SetData(0, 4, revamt)
                .SetData(0, 5, legchkamt)
                .SetData(0, 6, legrcptamt)
                .SetData(0, 7, amount)
            End With
        Catch ex As Exception
            Throw
        End Try

        ''''Me.Prev1.Visible = False
        ''''Me.GridDetail.Visible = True
        ''''Me.ShowDialog()
        ''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintAdjustmentRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateAdjustmentTicket(ByVal etransactionkey As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5       6  
        '   key        fisyr    trantype   bank     acct    subacct   amt  
        '     7           8         9       10       11       12      13
        ' applied     created    descr   remarks   balance  subname acctname
        '    14          15   
        ' expcode     revcode
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves an adjustment transaction using the adjustment key;
        Dim SSQL As String
        SSQL = "SELECT tran_autoinc_key, tran_fisyr, tran_type, x.bank_acct_num," _
        & " x.af_acct_num, x.as_acct_num, tran_amt, tran_applied_date, tran_datetime," _
        & " tran_descr, tran_remarks," _
        & " as_beg_month_balance + ((as_mtd_receipts + as_mtd_adjust) - as_mtd_expend) AS subcurbal," _
        & " d.as_acct_name, h.af_acct_name, ocex_code, ocrv_code" _
        & " FROM transactions AS x, acct_sub AS d, acct_info AS h" _
        & " WHERE tran_autoinc_key = @p1" _
        & " AND x.af_acct_num = d.af_acct_num" _
        & " AND x.as_acct_num = d.as_acct_num" _
        & " AND d.af_acct_num = h.af_acct_num;"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", etransactionkey)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("adjustment")
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
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table
            PrintAdjustmentTicket()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateDetailOfAccountsMTDAllAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal ebegindate As Date, ByVal eenddate As Date, ByVal ecurrentmonth As String, ByVal etype As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current month detail report for all accounts within a single bank;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'check for current fiscal year;
        If efiscalyear <> Me.FiscalYear Then Throw New ArgumentException("This report is only valid for the current fiscal year...")

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5 As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim cmd As SqlCommand

        Try
            'collection of check items;
            SSQL1 = "SELECT c.bank_acct_num, c.chks_fisyr, c.chks_num, '0' AS doctype, c.chks_status," _
            & " d.ckdt_amount, c.chks_applied_date, c.chks_datetime, c.chks_payee_name," _
            & " d.af_acct_num, d.as_acct_num, d.ckdt_descr, d.ocex_code, '' AS revenuecode," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " c.chks_autoinc_key, p.po_num AS didacts" _
            & " FROM chks_info AS c, chks_detl AS d, purc_detl AS p" _
            & " WHERE c.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND c.bank_acct_num = @p1" _
            & " AND c.chks_fisyr = @p2" _
            & " AND c.chks_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY c.bank_acct_num, c.chks_fisyr, d.af_acct_num, d.as_acct_num, CAST(c.chks_num AS INT), d.ckdt_autoinc_key; "

            'collection of receipt items;
            SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
            & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " h.rcpt_autoinc_key, '' AS didacts" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.rcpt_fisyr = @p2" _
            & " AND h.rcpt_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

            'collection of adjustments;
            SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
            & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
            & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " tran_autoinc_key, '' AS didacts" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " AND tran_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, tran_fisyr, af_acct_num, as_acct_num, tran_autoinc_key; "

            'collection of transfers from;
            SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_from AS account, as_acct_num_from AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND trx_applieddate BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_from, as_acct_num_from, trx_autoinc_key; "

            'collection of transfers to;
            SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_to AS account, as_acct_num_to AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND trx_applieddate BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_to, as_acct_num_to, trx_autoinc_key"

            SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p5", ebegindate)
            cmd.Parameters.Add("@p6", eenddate)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Detail")
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid;
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridWrk
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 21
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    .SetData(i, 20, row.Item(19))
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'sort the grid by account/subaccount
            .Sort(SortFlags.Ascending, 9, 10)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         
        '  mtdrcpts  mtdexpend    mtdadj  acctsubnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SSQL = "SELECT h.bank_acct_num, d.as_beg_month_balance, h.af_acct_num," _
        & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts," _
        & " d.as_mtd_expend, d.as_mtd_adjust, h.af_acct_num + d.as_acct_num" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.bank_acct_num = @p1" _
        & " ORDER BY d.bank_acct_num, h.af_acct_num, d.as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridWrkTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'etype:   1=ytd; 2=mtd; 3=periodical;
            If etype = 2 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = "MTD Detail"
            End If
            If etype = 3 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = ebegindate.ToShortDateString & " To " & eenddate.ToShortDateString
            End If

            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsAllAccounts(etype, efiscalyear)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateDetailOfAccountsMTDSingleAccount(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eaccountnumber As String, ByVal esubaccountnumber As String, ByVal ebegindate As Date, ByVal eenddate As Date, ByVal ecurrentmonth As String, ByVal etype As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current month detail report for a single account
        'within a single bank;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If eaccountnumber Is Nothing OrElse eaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")
        If esubaccountnumber Is Nothing OrElse esubaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")
        If efiscalyear <> Me.FiscalYear Then Throw New ArgumentException("This report is only valid for the current fiscal year...")

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5 As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim cmd As SqlCommand

        Try
            'collection of check items;
            SSQL1 = "SELECT c.bank_acct_num, c.chks_fisyr, c.chks_num, '0' AS doctype, c.chks_status," _
            & " d.ckdt_amount, c.chks_applied_date, c.chks_datetime, c.chks_payee_name," _
            & " d.af_acct_num, d.as_acct_num, d.ckdt_descr, d.ocex_code, '' AS revenuecode," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " c.chks_autoinc_key, p.po_num AS didacts" _
            & " FROM chks_info AS c, chks_detl AS d, purc_detl AS p" _
            & " WHERE c.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND c.bank_acct_num = @p1" _
            & " AND c.chks_fisyr = @p2" _
            & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
            & " AND c.chks_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY c.bank_acct_num, c.chks_fisyr, CAST(c.chks_num AS INT), d.ckdt_autoinc_key; "

            'collection of receipt items;
            SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
            & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " h.rcpt_autoinc_key, '' AS didacts" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.rcpt_fisyr = @p2" _
            & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
            & " AND h.rcpt_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

            'collection of adjustments;
            SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
            & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
            & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " tran_autoinc_key, '' AS didacts" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " AND af_acct_num = @p3 AND as_acct_num = @p4" _
            & " AND tran_applied_date BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, tran_fisyr, tran_autoinc_key; "

            'collection of transfers from;
            SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND af_acct_num_from = @p3 AND as_acct_num_from = @p4" _
            & " AND trx_applieddate BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

            'collection of transfers to;
            SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND af_acct_num_to = @p3 AND as_acct_num_to = @p4" _
            & " AND trx_applieddate BETWEEN @p5 AND @p6" _
            & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key"

            SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", eaccountnumber)
            cmd.Parameters.Add("@p4", esubaccountnumber)
            cmd.Parameters.Add("@p5", ebegindate)
            cmd.Parameters.Add("@p6", eenddate)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Detail")
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid;
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridDetail
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 21
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    .SetData(i, 20, row.Item(19)) 'ponum for check record only;
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 7)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SSQL = "SELECT h.bank_acct_num, d.as_beg_month_balance," _
        & " h.af_acct_num, h.af_acct_name, d.as_acct_num, d.as_acct_name," _
        & " d.as_mtd_receipts, d.as_mtd_expend, d.as_mtd_adjust" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.bank_acct_num = @p1" _
        & " AND d.af_acct_num = @p3" _
        & " AND d.as_acct_num = @p4"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        ''''''''''''''''''
        ''''''test code only ''''''''
        '''''Me.GridTotals.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function
        ''''''''''''''''''
        ''''''''''''''''''

        Try
            'etype:   1=ytd; 2=mtd; 3=periodical;
            If etype = 2 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = "MTD Detail"
            End If
            If etype = 3 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = ebegindate.ToShortDateString & " To " & eenddate.ToShortDateString
            End If
            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsSingleAccount(etype)   '1=ytd; 2=mtd; 3=periodical;
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateDetailOfAccountsYTDAllAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a ytd detail report for all accounts within a single bank;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5 As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim cmd As SqlCommand
 
        Try
            'collection of check items;
            SSQL1 = "SELECT c.bank_acct_num, c.chks_fisyr, c.chks_num, '0' AS doctype, c.chks_status," _
            & " d.ckdt_amount, c.chks_applied_date, c.chks_datetime, c.chks_payee_name," _
            & " d.af_acct_num, d.as_acct_num, d.ckdt_descr, d.ocex_code, '' AS revenuecode," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " c.chks_autoinc_key, p.po_num AS didacts" _
            & " FROM chks_info AS c, chks_detl AS d, purc_detl AS p" _
            & " WHERE c.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND c.bank_acct_num = @p1" _
            & " AND c.chks_fisyr = @p2" _
            & " ORDER BY c.bank_acct_num, c.chks_fisyr, d.af_acct_num, d.as_acct_num, CAST(c.chks_num AS INT), d.ckdt_autoinc_key; "

            'collection of receipt items;
            SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
            & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " h.rcpt_autoinc_key, '' AS didacts" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.rcpt_fisyr = @p2" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

            'collection of adjustments;
            SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
            & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
            & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " tran_autoinc_key, '' AS didacts" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " ORDER BY bank_acct_num, tran_fisyr, af_acct_num, as_acct_num, tran_autoinc_key; "

            'collection of transfers from;
            SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_from AS account, as_acct_num_from AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_from, as_acct_num_from, trx_autoinc_key; "

            'collection of transfers to;
            SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_to AS account, as_acct_num_to AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_to, as_acct_num_to, trx_autoinc_key"

            SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Detail")
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid;
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridWrk
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 21
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    .SetData(i, 20, row.Item(19))
                    i += 1
                Next
            Next

            'check if there are any records;
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 9, 10)
            .AutoSizeCols()
        End With

        'now collect the account information;
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         
        '  mtdrcpts  mtdexpend    mtdadj  acctsubnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If efiscalyear = Me.FiscalYear Then
            SSQL = "SELECT h.bank_acct_num, d.as_beg_year_balance, h.af_acct_num," _
            & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts + d.as_ytd_receipts," _
            & " d.as_mtd_expend + d.as_ytd_expend, d.as_mtd_adjust + d.as_ytd_adjust," _
            & " h.af_acct_num + d.as_acct_num" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND d.bank_acct_num = @p1" _
            & " ORDER BY d.bank_acct_num, h.af_acct_num, d.as_acct_num"
        Else
            SSQL = "SELECT bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name, SUM(ahst_mtd_receipts)," _
            & " SUM(ahst_mtd_expend), SUM(ahst_mtd_adjust), af_acct_num + as_acct_num" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = " & efiscalyear _
            & " GROUP BY bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name, af_acct_num + as_acct_num"
        End If
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridWrkTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'Me.GridWrk.Visible = True
        'Me.Prev1.Visible = False
        'With Me.GridWrkTotals
        '    .Cols(0).Visible = False
        '    .Cols(1).Visible = False
        '    .Cols(2).Visible = False
        '    .Cols(4).Visible = False
        '    .Cols(5).Visible = False
        '    .Cols(6).Visible = False
        '    .Cols(7).Visible = False
        '    .Cols(8).Visible = False
        'End With
        'Me.ShowDialog()
        'Exit Function
        '''''''''''''''''''''''
        '''''''''''''''''''''''

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "YTD Detail"
            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsAllAccounts(1, efiscalyear)  '1=ytd; 2=mtd
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateDetailOfAccountsYTDSingleAccount(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eaccountnumber As String, ByVal esubaccountnumber As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a ytd detail report for a single account  within a single bank;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If eaccountnumber Is Nothing OrElse eaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")
        If esubaccountnumber Is Nothing OrElse esubaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5 As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim cmd As SqlCommand

        Try
            'collection of check items;
            SSQL1 = "SELECT c.bank_acct_num, c.chks_fisyr, c.chks_num, '0' AS doctype, c.chks_status," _
            & " d.ckdt_amount, c.chks_applied_date, c.chks_datetime, c.chks_payee_name," _
            & " d.af_acct_num, d.as_acct_num, d.ckdt_descr, d.ocex_code, '' AS revenuecode," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " c.chks_autoinc_key, p.po_num AS didacts" _
            & " FROM chks_info AS c, chks_detl AS d, purc_detl AS p" _
            & " WHERE c.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND c.bank_acct_num = @p1" _
            & " AND c.chks_fisyr = @p2" _
            & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
            & " ORDER BY c.bank_acct_num, c.chks_fisyr, CAST(c.chks_num AS INT), d.ckdt_autoinc_key; "

            'collection of receipt items;
            SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
            & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " h.rcpt_autoinc_key, '' AS didacts" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.rcpt_fisyr = @p2" _
            & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

            'collection of adjustments;
            SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
            & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
            & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
            & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
            & " tran_autoinc_key, '' AS didacts" _
            & " FROM transactions" _
            & " WHERE bank_acct_num = @p1" _
            & " AND tran_fisyr = @p2" _
            & " AND af_acct_num = @p3 AND as_acct_num = @p4" _
            & " ORDER BY bank_acct_num, tran_fisyr, tran_autoinc_key; "

            'collection of transfers;
            SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND af_acct_num_from = @p3 AND as_acct_num_from = @p4" _
            & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

            'collection of transfers;
            SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
            & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
            & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_autoinc_key, '' AS didacts" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND af_acct_num_to = @p3 AND as_acct_num_to = @p4" _
            & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key"

            SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", eaccountnumber)
            cmd.Parameters.Add("@p4", esubaccountnumber)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Detail")
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid;

        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridDetail
            'initialise the grid;
            .Rows.Count = 0
            .Cols.Count = 21
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    .SetData(i, 20, row.Item(19))
                    i += 1
                Next
            Next

            'check if there are any records;
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 7)
            .AutoSizeCols()
        End With

        'now collect the account information;
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct   begyrbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If efiscalyear = Me.FiscalYear Then
            SSQL = "SELECT h.bank_acct_num, d.as_beg_year_balance, h.af_acct_num," _
            & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts + d.as_ytd_receipts," _
            & " d.as_mtd_expend + d.as_ytd_expend, d.as_mtd_adjust + d.as_ytd_adjust" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND d.bank_acct_num = @p1" _
            & " AND d.af_acct_num = @p3" _
            & " AND d.as_acct_num = @p4"
        Else
            SSQL = "SELECT bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name, SUM(ahst_mtd_receipts)," _
            & " SUM(ahst_mtd_expend), SUM(ahst_mtd_adjust)" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = " & efiscalyear _
            & " AND af_acct_num = @p3" _
            & " AND as_acct_num = @p4" _
            & " GROUP BY bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name"
        End If
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        '''''''''''''''''''''''
        'test code only ''''''''
        '''''Me.GridTotals.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function
        '''''''''''''''''''''''
        '''''''''''''''''''''''

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "YTD Detail"
            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsSingleAccount(1)   '1=ytd; 2=mtd
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateEncumbranceDetailOfAccountsMTDSingleAccount(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eaccountnumber As String, ByVal esubaccountnumber As String, ByVal ebegindate As Date, ByVal eenddate As Date, ByVal ecurrentmonth As String, ByVal etype As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current month detail report for a single account
        'within a single bank;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If eaccountnumber Is Nothing OrElse eaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")
        If esubaccountnumber Is Nothing OrElse esubaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5, SSQL6 As String

        'collection of check items 
        SSQL1 = "SELECT ci.bank_acct_num, ci.chks_fisyr, ci.chks_num, '0' AS doctype, ci.chks_status," _
        & " cd.ckdt_amount, ci.chks_applied_date, ci.chks_datetime, ci.chks_payee_name," _
        & " cd.af_acct_num, cd.as_acct_num, cd.ckdt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " ci.chks_autoinc_key" _
        & " FROM chks_info AS ci, chks_detl AS cd" _
        & " WHERE ci.chks_autoinc_key = cd.chks_autoinc_key" _
        & " AND ci.bank_acct_num = @p1" _
        & " AND ci.chks_fisyr = @p2" _
        & " AND cd.af_acct_num = @p3 AND cd.as_acct_num = @p4" _
        & " AND ci.chks_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY ci.bank_acct_num, ci.chks_fisyr, CAST(ci.chks_num AS INT), cd.ckdt_autoinc_key; "

        'collection of receipt items 
        SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
        & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
        & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.rcpt_autoinc_key" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.rcpt_fisyr = @p2" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
        & " AND h.rcpt_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

        'collection of adjustments
        SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
        & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
        & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " tran_autoinc_key" _
        & " FROM transactions" _
        & " WHERE bank_acct_num = @p1" _
        & " AND tran_fisyr = @p2" _
        & " AND af_acct_num = @p3 AND as_acct_num = @p4" _
        & " AND tran_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, tran_fisyr, tran_autoinc_key; "

        'collection of transfers from
        SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND af_acct_num_from = @p3 AND as_acct_num_from = @p4" _
        & " AND trx_applieddate BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

        'collection of transfers to
        SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND af_acct_num_to = @p3 AND as_acct_num_to = @p4" _
        & " AND trx_applieddate BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

        'collection of purchase order items; 
        SSQL6 = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, '5' AS doctype, d.podt_status," _
        & " d.podt_amount, h.po_applied_date, h.po_datetime, v.vend_name AS 'h.po_descr'," _
        & " d.af_acct_num, d.as_acct_num, d.podt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.po_autoinc_key" _
        & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.vend_number = v.vend_number" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
        & " AND h.po_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY h.bank_acct_num, h.po_fisyr, CAST(h.po_num AS INT), d.podt_autoinc_key"

        SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5 + SSQL6
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        cmd.Parameters.Add("@p5", ebegindate)
        cmd.Parameters.Add("@p6", eenddate)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("Detail")
        Try
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridDetail
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 20
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 7)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    mtdencnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SSQL = "SELECT h.bank_acct_num, d.as_beg_month_balance," _
        & " h.af_acct_num, h.af_acct_name, d.as_acct_num, d.as_acct_name," _
        & " d.as_mtd_receipts, d.as_mtd_expend, d.as_mtd_adjust, d.as_ytd_encumbered" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.bank_acct_num = @p1" _
        & " AND d.af_acct_num = @p3" _
        & " AND d.as_acct_num = @p4"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'collect the encumbered amount for the summary values;
        SSQL = "SELECT ISNULL (SUM(podt_amount), 0)" _
        & " FROM purc_info AS h, purc_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
        & " AND h.po_applied_date BETWEEN @p5 AND @p6"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        cmd.Parameters.Add("@p5", ebegindate)
        cmd.Parameters.Add("@p6", eenddate)
        Dim encnumsum As Double
        Dim value As Object
        Try
            cn.Open()
            encnumsum = CDbl(cmd.ExecuteScalar)
            Me.GridTotals.SetData(1, 9, encnumsum)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        ''''''''''''''''''''''''''''''''''''
        ''''''''' test code only '''''''''''
        ''''Me.GridTotals.Visible = True
        ''''Me.GridDetail.Visible = False
        ''''Me.Prev1.Visible = False
        ''''Me.ShowDialog()
        ''''Exit Function
        ''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''

        Try
            'etype:   1=ytd; 2=mtd; 3=periodical;
            If etype = 2 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = "MTD Detail"
            End If
            If etype = 3 Then
                Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
                Me.CellMiddleBottom = ebegindate.ToShortDateString & " To " & eenddate.ToShortDateString
            End If
            Application.DoEvents()
            'render the report
            PrintEncumbranceDetailOfAccountsSingleAccount(Me.FiscalYear, etype)   '1=ytd; 2=mtd; 3=periodical;
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateEncumbranceDetailOfAccountsMTDAllAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal ebegindate As Date, ByVal eenddate As Date, ByVal ecurrentmonth As String, ByVal etype As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current ytd detail report for all accounts
        'within a single bank;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5, SSQL6 As String

        'collection of check items 
        SSQL1 = "SELECT ci.bank_acct_num, ci.chks_fisyr, ci.chks_num, '0' AS doctype, ci.chks_status," _
        & " cd.ckdt_amount, ci.chks_applied_date, ci.chks_datetime, ci.chks_payee_name," _
        & " cd.af_acct_num, cd.as_acct_num, cd.ckdt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " ci.chks_autoinc_key" _
        & " FROM chks_info AS ci, chks_detl AS cd" _
        & " WHERE ci.chks_autoinc_key = cd.chks_autoinc_key" _
        & " AND ci.bank_acct_num = @p1" _
        & " AND ci.chks_fisyr = @p2" _
        & " AND ci.chks_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY ci.bank_acct_num, ci.chks_fisyr, cd.af_acct_num, cd.as_acct_num, CAST(ci.chks_num AS INT), cd.ckdt_autoinc_key; "

        'collection of receipt items 
        SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
        & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
        & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.rcpt_autoinc_key" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.rcpt_fisyr = @p2" _
        & " AND h.rcpt_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

        'collection of adjustments
        SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
        & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
        & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " tran_autoinc_key" _
        & " FROM transactions" _
        & " WHERE bank_acct_num = @p1" _
        & " AND tran_fisyr = @p2" _
        & " AND tran_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, tran_fisyr, af_acct_num, as_acct_num, tran_autoinc_key; "

        'collection of transfers from
        SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_from AS account, as_acct_num_from AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND trx_applieddate BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_from, as_acct_num_from, trx_autoinc_key; "

        'collection of transfers to
        SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_to AS account, as_acct_num_to AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND trx_applieddate BETWEEN @p5 AND @p6" _
        & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_to, as_acct_num_to, trx_autoinc_key; "

        'collection of purchase order items; 
        SSQL6 = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, '5' AS doctype, h.po_status," _
        & " d.podt_amount, h.po_applied_date, h.po_datetime, v.vend_name AS 'h.po_descr'," _
        & " d.af_acct_num, d.as_acct_num, d.podt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.po_autoinc_key" _
        & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND h.vend_number = v.vend_number" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND h.po_applied_date BETWEEN @p5 AND @p6" _
        & " ORDER BY h.bank_acct_num, h.po_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.po_num AS INT), d.podt_autoinc_key"

        SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5 + SSQL6
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p5", ebegindate)
        cmd.Parameters.Add("@p6", eenddate)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("Detail")
        Try
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridWrk
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 20
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 9, 10)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  ytdrcpts  ytdexpend    ytdadj    ytdencnum  acct-sub
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SSQL = "SELECT h.bank_acct_num, d.as_beg_month_balance, h.af_acct_num," _
        & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts," _
        & " d.as_mtd_expend, d.as_mtd_adjust, d.as_ytd_encumbered," _
        & " h.af_acct_num + d.as_acct_num" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.bank_acct_num = @p1" _
        & " ORDER BY d.bank_acct_num, h.af_acct_num, d.as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridWrkTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        SSQL = "SELECT ISNULL (SUM(podt_amount), 0), af_acct_num, as_acct_num, af_acct_num + as_acct_num" _
        & " FROM purc_info AS h, purc_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND h.po_applied_date BETWEEN @p5 AND @p6" _
        & " GROUP BY af_acct_num, as_acct_num" _
        & " ORDER BY af_acct_num, as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p5", ebegindate)
        cmd.Parameters.Add("@p6", eenddate)
        Dim encnumsum As Double
        Dim dr As SqlDataReader
        Try
            cn.Open()
            dr = cmd.ExecuteReader
            With Me.GridWrkTotals
                Dim acctread, acctsub As String
                Dim amount As Double
                Dim index As Int32
                Do While dr.Read
                    amount = CDbl(dr.Item(0))
                    acctread = DirectCast(dr.Item(3), String)
                    index = .FindRow(acctread, 0, 10, True, True, False)
                    If index >= 0 Then .SetData(index, 9, amount)
                Loop
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            dr.Close()
            cmd.Dispose()
        End Try

        ''''''''''''''''''''''''''''''''''''
        ''''''''' test code only '''''''''''
        '''''Me.GridWrkTotals.Visible = True
        '''''Me.GridWrk.Visible = False
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function
        ''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''

        Try
            Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "MTD Detail"
            Application.DoEvents()
            'render the report
            PrintEncumbranceDetailOfAccountsAllAccounts(2, efiscalyear)      '1=ytd; 2=mtd
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateEncumbranceDetailOfAccountsYTDAllAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current ytd detail report for all accounts
        'within a single bank;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5, SSQL6 As String

        'collection of check items 
        SSQL1 = "SELECT ci.bank_acct_num, ci.chks_fisyr, ci.chks_num, '0' AS doctype, ci.chks_status," _
        & " cd.ckdt_amount, ci.chks_applied_date, ci.chks_datetime, ci.chks_payee_name," _
        & " cd.af_acct_num, cd.as_acct_num, cd.ckdt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " ci.chks_autoinc_key" _
        & " FROM chks_info AS ci, chks_detl AS cd" _
        & " WHERE ci.chks_autoinc_key = cd.chks_autoinc_key" _
        & " AND ci.bank_acct_num = @p1" _
        & " AND ci.chks_fisyr = @p2" _
        & " ORDER BY ci.bank_acct_num, ci.chks_fisyr, cd.af_acct_num, cd.as_acct_num, CAST(ci.chks_num AS INT), cd.ckdt_autoinc_key; "

        'collection of receipt items 
        SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
        & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
        & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.rcpt_autoinc_key" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.rcpt_fisyr = @p2" _
        & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

        'collection of adjustments
        SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
        & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
        & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " tran_autoinc_key" _
        & " FROM transactions" _
        & " WHERE bank_acct_num = @p1" _
        & " AND tran_fisyr = @p2" _
        & " ORDER BY bank_acct_num, tran_fisyr, af_acct_num, as_acct_num, tran_autoinc_key; "

        'collection of transfers from
        SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_from AS account, as_acct_num_from AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_from, as_acct_num_from, trx_autoinc_key; "

        'collection of transfers to
        SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, af_acct_num_to AS account, as_acct_num_to AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " ORDER BY bank_acct_num, trx_fisyr, af_acct_num_to, as_acct_num_to, trx_autoinc_key; "

        'collection of purchase order items; 
        SSQL6 = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, '5' AS doctype, h.po_status," _
        & " d.podt_amount, h.po_applied_date, h.po_datetime, v.vend_name AS 'h.po_descr'," _
        & " d.af_acct_num, d.as_acct_num, d.podt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.po_autoinc_key" _
        & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.vend_number = v.vend_number" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " ORDER BY h.bank_acct_num, h.po_fisyr, d.af_acct_num, d.as_acct_num, CAST(h.po_num AS INT), d.podt_autoinc_key"


        SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5 + SSQL6
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("Detail")
        Try
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridWrk
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 20
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 9, 10)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  ytdrcpts  ytdexpend    ytdadj    ytdencnum  acct-sub
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If efiscalyear = Me.FiscalYear Then
            SSQL = "SELECT h.bank_acct_num, d.as_beg_year_balance, h.af_acct_num," _
            & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts + d.as_ytd_receipts," _
            & " d.as_mtd_expend + d.as_ytd_expend, d.as_mtd_adjust + d.as_ytd_adjust," _
            & " d.as_ytd_encumbered, h.af_acct_num + d.as_acct_num " _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND d.bank_acct_num = @p1" _
            & " ORDER BY d.bank_acct_num, h.af_acct_num, d.as_acct_num"
        Else
            SSQL = "SELECT bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name, SUM(ahst_mtd_receipts)," _
            & " SUM(ahst_mtd_expend), SUM(ahst_mtd_adjust), 0.0 AS Encumbered, af_acct_num + as_acct_num" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = " & efiscalyear _
            & " GROUP BY bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name"
        End If
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridWrkTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        SSQL = "SELECT ISNULL (SUM(podt_amount), 0), af_acct_num, as_acct_num, af_acct_num + as_acct_num" _
        & " FROM purc_info AS h, purc_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " GROUP BY af_acct_num, as_acct_num" _
        & " ORDER BY af_acct_num, as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        Dim encnumsum As Double
        Dim dr As SqlDataReader
        Try
            cn.Open()
            dr = cmd.ExecuteReader
            With Me.GridWrkTotals
                Dim acctread, acctsub As String
                Dim amount As Double
                Dim index As Int32
                Do While dr.Read
                    amount = CDbl(dr.Item(0))
                    acctread = DirectCast(dr.Item(3), String)
                    index = .FindRow(acctread, 0, 10, True, True, False)
                    If index >= 0 Then .SetData(index, 9, amount)
                Loop
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            dr.Close()
            cmd.Dispose()
        End Try

        ''''''''''''''''''''''''''''''''''''
        ''''''''' test code only '''''''''''
        ''''''Me.GridWrkTotals.Visible = True
        ''''''Me.GridWrk.Visible = True
        ''''''Me.Prev1.Visible = True
        ''''''Me.ShowDialog()
        ''''''Exit Function
        ''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "YTD Detail"
            Application.DoEvents()
            'render the report
            PrintEncumbranceDetailOfAccountsAllAccounts(1, efiscalyear)      '1=ytd; 2=mtd
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateEncumbranceDetailOfAccountsYTDSingleAccount(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eaccountnumber As String, ByVal esubaccountnumber As String) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a current ytd detail report for a single account
        'within a single bank;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If eaccountnumber Is Nothing OrElse eaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")
        If esubaccountnumber Is Nothing OrElse esubaccountnumber.Trim.Length = 0 Then Throw New ArgumentException("Account number is missing or invalid...")

        Dim SSQL, SSQL1, SSQL2, SSQL3, SSQL4, SSQL5, SSQL6 As String

        'collection of check items 
        SSQL1 = "SELECT ci.bank_acct_num, ci.chks_fisyr, ci.chks_num, '0' AS doctype, ci.chks_status," _
        & " cd.ckdt_amount, ci.chks_applied_date, ci.chks_datetime, ci.chks_payee_name," _
        & " cd.af_acct_num, cd.as_acct_num, cd.ckdt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " ci.chks_autoinc_key" _
        & " FROM chks_info AS ci, chks_detl AS cd" _
        & " WHERE ci.chks_autoinc_key = cd.chks_autoinc_key" _
        & " AND ci.bank_acct_num = @p1" _
        & " AND ci.chks_fisyr = @p2" _
        & " AND cd.af_acct_num = @p3 AND cd.as_acct_num = @p4" _
        & " ORDER BY ci.bank_acct_num, ci.chks_fisyr, CAST(ci.chks_num AS INT), cd.ckdt_autoinc_key; "

        'collection of receipt items 
        SSQL2 = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, '1' AS doctype, h.rcpt_status," _
        & " d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
        & " d.af_acct_num, d.as_acct_num, d.rcdt_remarks, '' AS expenditurecode, d.ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.rcpt_autoinc_key" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.rcpt_fisyr = @p2" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
        & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, CAST(h.rcpt_num AS INT), d.rcdt_autoinc_key; "

        'collection of adjustments
        SSQL3 = "SELECT bank_acct_num, tran_fisyr, CAST(tran_autoinc_key AS VARCHAR), '2' AS doctype, tran_type," _
        & " tran_amt, tran_applied_date, tran_datetime, tran_descr," _
        & " af_acct_num, as_acct_num, tran_remarks, ocex_code, ocrv_code," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " tran_autoinc_key" _
        & " FROM transactions" _
        & " WHERE bank_acct_num = @p1" _
        & " AND tran_fisyr = @p2" _
        & " AND af_acct_num = @p3 AND as_acct_num = @p4" _
        & " ORDER BY bank_acct_num, tran_fisyr, tran_autoinc_key; "

        'collection of transfers from
        SSQL4 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '3' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND af_acct_num_from = @p3 AND as_acct_num_from = @p4" _
        & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

        'collection of transfers to
        SSQL5 = "SELECT bank_acct_num, trx_fisyr, CAST(trx_autoinc_key AS VARCHAR), '4' AS doctype, '' AS status," _
        & " trx_amt, trx_applieddate, trx_datetime, trx_descr, '' AS account, '' AS subaccount," _
        & " trx_remarks, '' AS expenditurecode, '' AS revenuecode," _
        & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
        & " trx_autoinc_key" _
        & " FROM transfers" _
        & " WHERE bank_acct_num = @p1" _
        & " AND trx_fisyr = @p2" _
        & " AND af_acct_num_to = @p3 AND as_acct_num_to = @p4" _
        & " ORDER BY bank_acct_num, trx_fisyr, trx_autoinc_key; "

        'collection of purchase order items; 
        SSQL6 = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, '5' AS doctype, h.po_status," _
        & " d.podt_amount, h.po_applied_date, h.po_datetime, v.vend_name AS 'h.po_descr'," _
        & " d.af_acct_num, d.as_acct_num, d.podt_descr, ocex_code, '' AS revenuecode," _
        & " '' AS tranacctfrom, '' AS transubfrom, '' AS tranacctto, '' AS transubto," _
        & " h.po_autoinc_key" _
        & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.vend_number = v.vend_number" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4" _
        & " ORDER BY h.bank_acct_num, h.po_fisyr, CAST(h.po_num AS INT), d.podt_autoinc_key"

        SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5 + SSQL6
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("Detail")
        Try
            cn.Open()
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'the data has been collected so it's time to enter them into the detail grid
        Dim tbl As DataTable
        Dim row As DataRow
        Dim i As Int32
        With Me.GridDetail
            'initialise the grid
            .Rows.Count = 0
            .Cols.Count = 20
            For Each tbl In ds.Tables
                For Each row In tbl.Rows
                    .Rows.Add()
                    .SetData(i, 0, row.Item(0))
                    .SetData(i, 1, row.Item(1))
                    .SetData(i, 2, row.Item(2))
                    .SetData(i, 3, row.Item(3))
                    .SetData(i, 4, row.Item(4))
                    .SetData(i, 5, row.Item(5))
                    .SetData(i, 6, row.Item(6))
                    .SetData(i, 7, row.Item(7))
                    .SetData(i, 8, row.Item(8))
                    .SetData(i, 9, row.Item(9))
                    .SetData(i, 10, row.Item(10))
                    .SetData(i, 11, row.Item(11))
                    .SetData(i, 12, row.Item(12))
                    .SetData(i, 13, row.Item(13))
                    .SetData(i, 14, row.Item(14))
                    .SetData(i, 15, row.Item(15))
                    .SetData(i, 16, row.Item(16))
                    .SetData(i, 17, row.Item(17))
                    .SetData(i, 18, row.Item(18))
                    .SetData(i, 19, i)
                    i += 1
                Next
            Next

            'check if there are any records
            If .Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")

            'sort the dog
            .Sort(SortFlags.Ascending, 7)
            .AutoSizeCols()
        End With

        'now collect the account information
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct     begbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    mtdencnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If efiscalyear = Me.FiscalYear Then
            SSQL = "SELECT h.bank_acct_num, d.as_beg_year_balance, h.af_acct_num," _
            & " h.af_acct_name, d.as_acct_num, d.as_acct_name, d.as_mtd_receipts + d.as_ytd_receipts," _
            & " d.as_mtd_expend + d.as_ytd_expend, d.as_mtd_adjust + d.as_ytd_adjust, d.as_ytd_encumbered" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND d.bank_acct_num = @p1" _
            & " AND d.af_acct_num = @p3" _
            & " AND d.as_acct_num = @p4"
        Else
            SSQL = "SELECT bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name, SUM(ahst_mtd_receipts)," _
            & " SUM(ahst_mtd_expend), SUM(ahst_mtd_adjust), 0.0 AS Encumbered" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = " & efiscalyear _
            & " AND af_acct_num = @p3" _
            & " AND as_acct_num = @p4" _
            & " GROUP BY bank_acct_num, ahst_beg_year_balance, af_acct_num," _
            & " af_acct_name, as_acct_num, as_acct_name"
        End If

        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        da = New SqlDataAdapter(cmd)
        Dim dt As New DataTable
        Try
            cn.Open()
            da.Fill(dt)
            Me.GridTotals.DataSource = dt
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        SSQL = "SELECT ISNULL (SUM(podt_amount), 0)" _
        & " FROM purc_info AS h, purc_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.po_fisyr = d.po_fisyr" _
        & " AND h.po_num = d.po_num" _
        & " AND (d.podt_status <> 'D')" _
        & " AND h.bank_acct_num = @p1" _
        & " AND h.po_fisyr = @p2" _
        & " AND d.af_acct_num = @p3 AND d.as_acct_num = @p4"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", eaccountnumber)
        cmd.Parameters.Add("@p4", esubaccountnumber)
        Dim encnumsum As Double
        Try
            cn.Open()
            encnumsum = CDbl(cmd.ExecuteScalar)
            Me.GridTotals.SetData(1, 9, encnumsum)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        '''' test code only '''''''''''
        '''''Me.GridTotals.Visible = True
        '''''Me.GridDetail.Visible = False
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function
        '''''''''''''''''''''''''''''''

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "YTD Detail"
            Application.DoEvents()
            'render the report
            PrintEncumbranceDetailOfAccountsSingleAccount(efiscalyear, 1)   '1=ytd; 2=mtd
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateNoMTDDetailOfAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal ebegindate As Date, ByVal eenddate As Date, ByVal ecurrentmonth As String) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a listing of accounts with NO detail activity for
        'a given fiscal month;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4      
        '   bank        acct    acctname  subacct  subname   
        '     5           6         7
        ' begmonth     cntadj    cnttrx
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand

        'collect those items with no "apparent" activity;
        SSQL = "SELECT d.bank_acct_num, d.af_acct_num, h.af_acct_name," _
        & " d.as_acct_num, d.as_acct_name, d.as_beg_month_balance," _
        & " '' AS adjustments, '' AS transfers" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.as_mtd_receipts = 0.0" _
        & " AND d.as_mtd_expend = 0.0" _
        & " AND d.as_mtd_adjust = 0.0" _
        & " AND d.bank_acct_num = @p1" _
        & " AND h.af_status = 'O'" _
        & " AND d.as_status = 'O'" _
        & " ORDER BY d.bank_acct_num, d.af_acct_num, d.as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("accounts")
        Try
            cn.Open()
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'no accounts were returned so ALL accounts had some activity;
        If tbl.Rows.Count < 1 Then Return False

        'source the data
        Me.GridWrk.DataSource = tbl

        Dim account, subaccount As String
        Dim row, records As Int32
        Try
            With Me.GridWrk
                For row = 1 To .Rows.Count - 1
                    account = DirectCast(.GetData(row, 1), String)
                    subaccount = DirectCast(.GetData(row, 3), String)
                    'check for adjustments
                    SSQL = "SELECT COUNT (*) FROM transactions" _
                    & " WHERE bank_acct_num = @p1" _
                    & " AND tran_fisyr = @p2" _
                    & " AND af_acct_num = @p3 AND as_acct_num = @p4" _
                    & " AND tran_applied_date BETWEEN @p5 AND @p6"
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", ebankaccountnumber)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", account)
                    cmd.Parameters.Add("@p4", subaccount)
                    cmd.Parameters.Add("@p5", ebegindate)
                    cmd.Parameters.Add("@p6", eenddate)
                    Try
                        cn.Open()
                        records = CInt(cmd.ExecuteScalar)
                        'update the grid with the adjustment count
                        .SetData(row, 6, records)
                    Catch ex As Exception
                        Throw
                    Finally
                        cn.Close()
                    End Try
                Next

                For row = 1 To .Rows.Count - 1
                    account = DirectCast(.GetData(row, 1), String)
                    subaccount = DirectCast(.GetData(row, 3), String)
                    records = CInt(.GetData(row, 6))

                    'if adjustment(s) exist, then skip record
                    If records > 0 Then
                        'set the transfer count to some number
                        .SetData(row, 7, 99999)
                        GoTo SkipProcessing
                    End If

                    'check for transfers
                    SSQL = "SELECT COUNT (*) FROM transfers" _
                    & " WHERE bank_acct_num = @p1" _
                    & " AND trx_fisyr = @p2" _
                    & " AND ((af_acct_num_from = @p3 AND as_acct_num_from = @p4)" _
                    & " OR (af_acct_num_to = @p3 AND as_acct_num_to = @p4))" _
                    & " AND trx_applieddate BETWEEN @p5 AND @p6"
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", ebankaccountnumber)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", account)
                    cmd.Parameters.Add("@p4", subaccount)
                    cmd.Parameters.Add("@p5", ebegindate)
                    cmd.Parameters.Add("@p6", eenddate)
                    Try
                        cn.Open()
                        records = CInt(cmd.ExecuteScalar)
                        'update the grid with the adjustment count
                        .SetData(row, 7, records)
                    Catch ex As Exception
                        Throw
                    Finally
                        cn.Close()
                    End Try
SkipProcessing:
                Next

                'load the detail grid with valid accounts (no activity);
                Dim adjrecs, trxrecs, currow As Int32
                Me.GridDetail.Rows.Count = 0
                Me.GridDetail.Cols.Count = 6
                For row = 1 To .Rows.Count - 1
                    adjrecs = CInt(.GetData(row, 6))
                    trxrecs = CInt(.GetData(row, 7))
                    If (adjrecs = 0) And (trxrecs = 0) Then
                        With Me.GridDetail
                            .Rows.Add()
                            .SetData(currow, 0, Me.GridWrk.GetData(row, 0))
                            .SetData(currow, 1, Me.GridWrk.GetData(row, 1))
                            .SetData(currow, 2, Me.GridWrk.GetData(row, 2))
                            .SetData(currow, 3, Me.GridWrk.GetData(row, 3))
                            .SetData(currow, 4, Me.GridWrk.GetData(row, 4))
                            .SetData(currow, 5, Me.GridWrk.GetData(row, 5))
                            currow += 1
                        End With
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''''''''''''''''''''
        '''''Me.GridWrk.Visible = True
        '''''Me.Prev1.Visible = False
        '''''With Me.GridWrkTotals
        '''''    .Cols(0).Visible = False
        '''''    .Cols(1).Visible = False
        '''''End With
        '''''Me.ShowDialog()
        '''''Exit Function
        '''''''''''''''''''''''

        Try
            'etype:   1=ytd; 2=mtd;
            Me.CellMiddleMiddle = ecurrentmonth & ", FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "MTD Detail"
            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsNoActivity(2)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateNoYTDDetailOfAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This generates a listing of accounts with NO detail activity for
        'a given fiscal year;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4      
        '   bank        acct    acctname  subacct  subname   
        '     5           6         7
        ' begmonth     cntadj    cnttrx
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand

        'collect those items with no "apparent" activity;
        SSQL = "SELECT d.bank_acct_num, d.af_acct_num, h.af_acct_name," _
        & " d.as_acct_num, d.as_acct_name, d.as_beg_month_balance," _
        & " '' AS adjustments, '' AS transfers" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " AND d.as_mtd_receipts = 0.0" _
        & " AND d.as_mtd_expend = 0.0" _
        & " AND d.as_mtd_adjust = 0.0" _
        & " AND d.as_ytd_receipts = 0.0" _
        & " AND d.as_ytd_expend = 0.0" _
        & " AND d.as_ytd_adjust = 0.0" _
        & " AND d.bank_acct_num = @p1" _
        & " AND h.af_status = 'O'" _
        & " AND d.as_status = 'O'" _
        & " ORDER BY d.bank_acct_num, d.af_acct_num, d.as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("accounts")
        Try
            cn.Open()
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'no accounts were returned so ALL accounts had some activity;
        If tbl.Rows.Count < 1 Then Return False

        'source the data
        Me.GridWrk.DataSource = tbl

        Dim account, subaccount As String
        Dim row, records As Int32
        Try
            With Me.GridWrk
                For row = 1 To .Rows.Count - 1
                    account = DirectCast(.GetData(row, 1), String)
                    subaccount = DirectCast(.GetData(row, 3), String)
                    'check for adjustments
                    SSQL = "SELECT COUNT (*) FROM transactions" _
                    & " WHERE bank_acct_num = @p1" _
                    & " AND tran_fisyr = @p2" _
                    & " AND af_acct_num = @p3 AND as_acct_num = @p4"
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", ebankaccountnumber)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", account)
                    cmd.Parameters.Add("@p4", subaccount)
                    Try
                        cn.Open()
                        records = CInt(cmd.ExecuteScalar)
                        'update the grid with the adjustment count
                        .SetData(row, 6, records)
                    Catch ex As Exception
                        Throw
                    Finally
                        cn.Close()
                    End Try
                Next

                For row = 1 To .Rows.Count - 1
                    account = DirectCast(.GetData(row, 1), String)
                    subaccount = DirectCast(.GetData(row, 3), String)
                    records = CInt(.GetData(row, 6))

                    'if adjustment(s) exist, then skip record
                    If records > 0 Then
                        'set the transfer count to some number
                        .SetData(row, 7, 99999)
                        GoTo SkipProcessing
                    End If

                    'check for transfers
                    SSQL = "SELECT COUNT (*) FROM transfers" _
                    & " WHERE bank_acct_num = @p1" _
                    & " AND trx_fisyr = @p2" _
                    & " AND ((af_acct_num_from = @p3 AND as_acct_num_from = @p4)" _
                    & " OR (af_acct_num_to = @p3 AND as_acct_num_to = @p4))"
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", ebankaccountnumber)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", account)
                    cmd.Parameters.Add("@p4", subaccount)
                    Try
                        cn.Open()
                        records = CInt(cmd.ExecuteScalar)
                        'update the grid with the adjustment count
                        .SetData(row, 7, records)
                    Catch ex As Exception
                        Throw
                    Finally
                        cn.Close()
                    End Try
SkipProcessing:
                Next

                'load the detail grid with valid accounts (no activity);
                Dim adjrecs, trxrecs, currow As Int32
                Me.GridDetail.Rows.Count = 0
                Me.GridDetail.Cols.Count = 6
                For row = 1 To .Rows.Count - 1
                    adjrecs = CInt(.GetData(row, 6))
                    trxrecs = CInt(.GetData(row, 7))
                    If (adjrecs = 0) And (trxrecs = 0) Then
                        With Me.GridDetail
                            .Rows.Add()
                            .SetData(currow, 0, Me.GridWrk.GetData(row, 0))
                            .SetData(currow, 1, Me.GridWrk.GetData(row, 1))
                            .SetData(currow, 2, Me.GridWrk.GetData(row, 2))
                            .SetData(currow, 3, Me.GridWrk.GetData(row, 3))
                            .SetData(currow, 4, Me.GridWrk.GetData(row, 4))
                            .SetData(currow, 5, Me.GridWrk.GetData(row, 5))
                            currow += 1
                        End With
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''''''''''''''''''''
        '''''Me.GridDetail.Visible = True
        '''''Me.Prev1.Visible = False
        '''''With Me.GridWrk
        '''''    .Cols(0).Visible = False
        '''''    .Cols(1).Visible = False
        '''''End With
        '''''Me.ShowDialog()
        '''''Exit Function
        '''''''''''''''''''''''

        Try
            'etype:   1=ytd; 2=mtd; 3=periodical;
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "YTD Detail"
            Application.DoEvents()
            'render the report
            PrintDetailOfAccountsNoActivity(1)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateTransferRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        'this method retrieves all transfers for a single bank and fiscal year
        'given the specified date range or number range;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0          1          2        3         4         5  
        '   bank       fisyr    docnumber  amount  acctfrom   subfrom
        '     6          7          8        9        10        11 
        '  acctto      subto     applied  created    descr    remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT bank_acct_num, trx_fisyr, trx_autoinc_key, trx_amt," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_applieddate, trx_datetime, trx_descr, trx_remarks" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND trx_applieddate BETWEEN @p3 AND @p4" _
            & " ORDER BY trx_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Dim starting, ending As Int32
            Try
                'validate the numbers
                starting = CInt(enumberfrom)
                ending = CInt(enumberto)
            Catch ex As Exception
                Throw New ArgumentException("The beginning or ending number is missing or invalid.")
            End Try

            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT bank_acct_num, trx_fisyr, trx_autoinc_key, trx_amt," _
            & " af_acct_num_from, as_acct_num_from, af_acct_num_to, as_acct_num_to," _
            & " trx_applieddate, trx_datetime, trx_descr, trx_remarks" _
            & " FROM transfers" _
            & " WHERE bank_acct_num = @p1" _
            & " AND trx_fisyr = @p2" _
            & " AND trx_autoinc_key BETWEEN @p3 AND @p4" _
            & " ORDER BY trx_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", starting)
            cmd.Parameters.Add("@p4", ending)
        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
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
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        Try
            'summarise the transfers
            Dim amount As Double
            Dim index As Int32
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    amount += CDbl(.GetData(index, 3))
                Next
            End With
            With Me.GridTotals
                .Rows.Count = 0
                .Cols.Count = 1
                .Rows.Add()
                .SetData(0, 0, amount)
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintTransferRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateTransferTicket(ByVal etransferkey As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2          3        4        5       6  
        '   key        fisyr      bank     srcacct   srcsub  destacct descsub
        '     7           8         9         10       11       12      13
        '  amount     applied   created     descr   remarks   srcbal srcacctname
        '    14          15        16         17 
        ' srcsubname  destbal  dstacctname destsubname
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves a transfer transaction using the transfer key;
        Dim SSQL As String
        SSQL = "SELECT trx_autoinc_key, trx_fisyr, bank_acct_num, af_acct_num_from," _
        & " as_acct_num_from, af_acct_num_to, as_acct_num_to, trx_amt," _
        & " trx_applieddate, trx_datetime, trx_descr, trx_remarks," _
        & " 0.00 AS Srcbalance, '' AS Srcaccount, '' AS Srcsubaccount," _
        & " 0.00 AS DestBalance, '' AS Destaccount, '' AS Destsubaccount" _
        & " FROM transfers" _
        & " WHERE trx_autoinc_key = @p1;"
        SSQL += "SELECT d.bank_acct_num, d.af_acct_num, as_acct_num," _
        & " as_beg_month_balance + ((as_mtd_receipts + as_mtd_adjust) - as_mtd_expend) AS balance," _
        & " af_acct_name, as_acct_name, (d.bank_acct_num + d.af_acct_num + d.as_acct_num) AS accountkey" _
        & " FROM acct_info AS h, acct_sub AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.af_acct_num = d.af_acct_num" _
        & " ORDER BY d.bank_acct_num, d.af_acct_num, d.as_acct_num"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", etransferkey)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        Try
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria...")
            If ds.Tables(1).Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridWrk.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Dim index, row As Int32
        Dim sourcekey, destkey, acctname, subname As String
        Dim balance As Double
        'find the matching source & destination account balances & names
        With Me.GridDetail
            For index = 1 To .Rows.Count - 1
                'get the keys
                sourcekey = DirectCast(.GetData(index, 2), String)
                sourcekey += DirectCast(.GetData(index, 3), String)
                sourcekey += DirectCast(.GetData(index, 4), String)
                destkey = DirectCast(.GetData(index, 2), String)
                destkey += DirectCast(.GetData(index, 5), String)
                destkey += DirectCast(.GetData(index, 6), String)
                'get information for the source acccount
                row = Me.GridWrk.FindRow(sourcekey, 0, 6, True, True, False)
                If row >= 0 Then
                    balance = CDbl(Me.GridWrk.GetData(row, 3))
                    acctname = DirectCast(Me.GridWrk.GetData(row, 4), String)
                    subname = DirectCast(Me.GridWrk.GetData(row, 5), String)
                    .SetData(index, 12, balance)
                    .SetData(index, 13, acctname)
                    .SetData(index, 14, subname)
                End If
                'get information for the destination acccount
                row = Me.GridWrk.FindRow(destkey, 0, 6, True, True, False)
                If row >= 0 Then
                    balance = CDbl(Me.GridWrk.GetData(row, 3))
                    acctname = DirectCast(Me.GridWrk.GetData(row, 4), String)
                    subname = DirectCast(Me.GridWrk.GetData(row, 5), String)
                    .SetData(index, 15, balance)
                    .SetData(index, 16, acctname)
                    .SetData(index, 17, subname)
                End If
            Next
        End With

        '''''Me.Prev1.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table
            PrintTransferTicket()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateTransferTicket(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal esrcaccountnumber As String, ByVal esrcsubaccountnumber As String, ByVal edstaccountnumber As String, ByVal edstsubaccountnumber As String, ByVal eamount As Double) As Boolean
        'this method retrieves the transfer key and calls the overloaded method;
        Dim SSQL As String
        SSQL = "SELECT trx_autoinc_key FROM transfers" _
        & " WHERE bank_acct_num = @p1 AND trx_fisyr = @p2" _
        & " AND af_acct_num_from = @p3 AND as_acct_num_from = @p4" _
        & " AND af_acct_num_to = @p5 AND as_acct_num_to = @p6" _
        & " AND trx_amt = @p7"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", esrcaccountnumber)
        cmd.Parameters.Add("@p4", esrcsubaccountnumber)
        cmd.Parameters.Add("@p5", edstaccountnumber)
        cmd.Parameters.Add("@p6", edstsubaccountnumber)
        cmd.Parameters.Add("@p7", eamount)
        Dim key As Int32
        Try
            cn.Open()
            key = CInt(cmd.ExecuteScalar)
        Catch ex As Exception
            MsgBox(ex.Message)
            key = 0
        Finally
            cn.Close()
            cmd.Dispose()
        End Try

        Try
            If key <= 0 Then Throw New ArgumentException("No records found for this criteria...")
            GenerateTransferTicket(key)
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateTrialBalance(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal ecurrentmonthbegdate As Date, ByVal ecurrentmonthenddate As Date, ByVal erecondate As Date, ByVal ebankstatement As Double, ByVal einterest As Double, ByVal echarges As Double, ByVal einvestments As Double, ByVal eprintdetail As Boolean) As Boolean

        'pat/fred... use today as the ending date range
        'monthenddate = Now
        'monthenddate = recondate
        'this method retrieves all receipts,checks & adjustments transactions
        'from ceratain bank account number
        'for viewing & returns a dataset cast as a generic object

        Dim SSQL As String
        Dim SSQL1, SSQL2, SSQL3, SSQL4, SSQL5, SSQL6, SSQL7, SSQL8, SSQL9, SSQL10 As String
        Dim SSQL11, SSQL12, SSQL13, SSQL14, SSQL15, SSQL16, SSQL17, SSQL18, SSQL19 As String
        Dim SSQL20, SSQL21 As String

        'left - beginning bank balance [table 0];
        SSQL1 = " SELECT bank_beg_balance FROM bank_info" _
        & " WHERE bank_acct_num = @p1; "

        'left - receipts issued within current month [table 1];
        SSQL2 = "SELECT SUM(rd.rcdt_amount)" _
        & " FROM receipt_detl AS rd, receipt_info AS ri" _
        & " WHERE rd.bank_acct_num = ri.bank_acct_num" _
        & " AND rd.rcpt_fisyr = ri.rcpt_fisyr" _
        & " AND rd.rcpt_num = ri.rcpt_num " _
        & " AND ri.rcpt_status <> 'D'" _
        & " AND ri.bank_acct_num = @p1" _
        & " AND (ri.rcpt_applied_date BETWEEN @p3 AND @p4); "

        'left - less credits (including the voids) [table 2];
        SSQL3 = "SELECT SUM(chks_amount) FROM chks_info" _
        & " WHERE bank_acct_num = @p1" _
        & " AND (chks_status <> 'D')" _
        & " AND (chks_applied_date BETWEEN @p3 AND @p4); "

        'left adjustments [table 3];
        SSQL4 = "SELECT " _
        & " (SELECT ISNULL(SUM(tran_amt), 0.00) FROM transactions" _
        & " WHERE bank_acct_num = @p1 " _
        & " AND tran_type <> 'B' " _
        & " AND tran_applied_date BETWEEN @p3 and @p4) -" _
        & " (SELECT ISNULL(SUM(tran_amt), 0.00) FROM transactions " _
        & " WHERE bank_acct_num = @p1 " _
        & " AND tran_type = 'B' AND tran_applied_date BETWEEN @p3 AND @p4); "

        'right - add deposits in transit [table 4];
        SSQL5 = "SELECT SUM(rd.rcdt_amount)" _
        & " FROM receipt_detl AS rd, receipt_info AS ri" _
        & " WHERE rd.bank_acct_num = ri.bank_acct_num" _
        & " AND rd.rcpt_fisyr = ri.rcpt_fisyr" _
        & " AND rd.rcpt_num = ri.rcpt_num " _
        & " AND ri.rcpt_status <> 'V'" _
        & " AND ri.rcpt_recon_sw = 'N'" _
        & " AND ri.rcpt_posted_sw = 'Y'" _
        & " AND ri.bank_acct_num = @p1; "

        'right - less outstanding checks [table 5];
        SSQL6 = "SELECT SUM(chks_amount)" _
        & " FROM chks_info" _
        & " WHERE chks_recon_sw = 'N'" _
        & " AND chks_status <> 'V' " _
        & " AND bank_acct_num = @p1; "

        'outstanding receipts [table 6];
        SSQL7 = "SELECT ri.rcpt_num, (SELECT SUM(rd.rcdt_amount)" _
        & " FROM receipt_detl AS n" _
        & " WHERE ri.bank_acct_num = n.bank_acct_num" _
        & " AND ri.rcpt_fisyr = n.rcpt_fisyr" _
        & " AND ri.rcpt_num = n.rcpt_num" _
        & " GROUP BY n.bank_acct_num, n.rcpt_fisyr, n.rcpt_num)" _
        & " FROM receipt_info AS ri, receipt_detl AS rd" _
        & " WHERE ri.bank_acct_num = rd.bank_acct_num" _
        & " AND ri.rcpt_fisyr = rd.rcpt_fisyr" _
        & " AND ri.rcpt_num = rd.rcpt_num " _
        & " AND ri.rcpt_status <> 'V'" _
        & " AND ri.rcpt_recon_sw = 'N'" _
        & " AND ri.rcpt_posted_sw = 'Y'" _
        & " AND ri.bank_acct_num = @p1" _
        & " GROUP BY ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num," _
        & " ri.rcpt_applied_date" _
        & " ORDER BY CAST(ri.rcpt_num AS INT); "

        'outstanding checks detail [table 7];
        SSQL8 = "SELECT chks_num, chks_amount" _
        & " FROM chks_info " _
        & " WHERE chks_recon_sw = 'N' " _
        & " AND chks_status <> 'V' " _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST(chks_num AS INT); "

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'pat/fred 2005.03.04... a fix was put in that any check or receipt that
        'is reconciled AND closed (C) is a document that has been cleared AFTER
        'the last closeout.  During a closeout, the status will be changed for 
        'the documents from a (C) to a (F). Since an (F) represents an item that
        'has been cleared and closed-out, this item should never appear again 
        'on a reconciliation.  NOTE: Outstandings will always be shown.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'receipts cleared after last closeout [table 8];
        SSQL9 = "SELECT ri.rcpt_num, (SELECT SUM(rd.rcdt_amount)" _
        & " FROM receipt_detl AS n" _
        & " WHERE ri.bank_acct_num = n.bank_acct_num" _
        & " AND ri.rcpt_fisyr = n.rcpt_fisyr" _
        & " AND ri.rcpt_num = n.rcpt_num" _
        & " GROUP BY n.bank_acct_num, n.rcpt_fisyr, n.rcpt_num)" _
        & " FROM receipt_info AS ri, receipt_detl AS rd" _
        & " WHERE ri.bank_acct_num = rd.bank_acct_num" _
        & " AND ri.rcpt_fisyr = rd.rcpt_fisyr" _
        & " AND ri.rcpt_num = rd.rcpt_num" _
        & " AND ri.rcpt_posted_sw = 'Y' AND ri.rcpt_status = 'C'" _
        & " AND ri.rcpt_recon_sw = 'Y' AND ri.bank_acct_num = @p1" _
        & " GROUP BY ri.bank_acct_num, ri.rcpt_fisyr, ri.rcpt_num, ri.rcpt_applied_date" _
        & " ORDER BY CAST(ri.rcpt_num AS INT); "

        'checks cleared after last closeout [table 9];
        SSQL10 = "SELECT chks_num, chks_amount" _
        & " FROM chks_info" _
        & " WHERE bank_acct_num = @p1" _
        & " AND chks_recon_sw = 'Y' " _
        & " AND chks_posted_sw = 'Y'" _
        & " AND chks_status = 'C'" _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST(chks_num AS INT); "

        'adjustments - all issued within applied month [table 10];
        SSQL11 = "SELECT tran_autoinc_key, tran_amt, tran_type" _
        & " FROM transactions" _
        & " WHERE bank_acct_num = @p1" _
        & " AND tran_fisyr = @p2" _
        & " AND (tran_applied_date BETWEEN @p3 AND @p4)" _
        & " ORDER BY tran_autoinc_key; "

        'void receipts detail [table 11];
        'SSQL12 = "SELECT vr.voidrcpt_num, vr.voidrcpt_amt" _
        '& " FROM voidreceipt AS vr, receipt_info AS ri" _
        '& " WHERE ri.rcpt_num = vr.voidrcpt_num" _
        '& " AND ri.rcpt_status = 'V'" _
        '& " AND ri.bank_acct_num = @p1" _
        '& " AND (vr.voidrcpt_applied_date BETWEEN @p3 AND @p4)" _
        '& " ORDER BY CAST(vr.voidrcpt_num AS INT); "

        SSQL12 = "SELECT DISTINCT voidrcpt_num, voidrcpt_amt" _
        & " FROM voidreceipt" _
        & " WHERE bank_acct_num = @p1 AND voidrcpt_fisyr = @p2" _
        & " AND voidrcpt_applied_date BETWEEN @p3 AND @p4" _
        & " ORDER BY voidrcpt_num; "

        'void checks detail [table 12];
        SSQL13 = "SELECT voidchk_num, SUM(ckdt_amount) FROM voidcheck" _
        & " WHERE bank_acct_num = @p1 AND voidchk_fisyr = @p2" _
        & " AND (voidchk_applied_date BETWEEN @p3 AND @p4)" _
        & " GROUP BY voidchk_num" _
        & " ORDER BY CAST(voidchk_num AS INT); "

        'outstanding legacy checks summary [table 13];
        SSQL14 = "SELECT SUM(outc_chk_amount)" _
        & " FROM outstandingchecks" _
        & " WHERE outc_recon_sw = 'N'" _
        & " AND outc_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1; "

        'outstanding legacy receipts summary [table 14];
        SSQL15 = "SELECT SUM(outr_rcpt_amount)" _
        & " FROM outstandingreceipts" _
        & " WHERE outr_recon_sw = 'N'" _
        & " AND outr_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1; "

        'outstanding legacy checks detail [table 15];
        SSQL16 = "SELECT outc_chk_num, outc_chk_amount" _
        & " FROM outstandingchecks" _
        & " WHERE outc_recon_sw = 'N'" _
        & " AND outc_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST(outc_chk_num AS INT); "

        'outstanding legacy receipts detail [table 16];
        SSQL17 = "SELECT outr_rcpt_num, outr_rcpt_amount" _
        & " FROM outstandingreceipts" _
        & " WHERE outr_recon_sw = 'N'" _
        & " AND outr_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST(outr_rcpt_num AS INT); "

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'for legacy items, the recon switch is used (instead of status) 
        'when an item has been cleared before a closeout (Y);  during the 
        'closeout, the status for the legacy document is changed from a (C) to
        'a (F). Any item with a (F) in the recon switch is an item that has 
        'been cleared and printed on a report before closeout. the cleared item
        'should never appear again on a reconciliation report;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'cleared legacy checks summary [table 17];
        SSQL18 = "SELECT SUM(outc_chk_amount)" _
        & " FROM outstandingchecks" _
        & " WHERE outc_recon_sw = 'Y'" _
        & " AND outc_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1; "

        'cleared legacy receipts summary [table 18];
        SSQL19 = "SELECT SUM(outr_rcpt_amount)" _
        & " FROM outstandingreceipts" _
        & " WHERE outr_recon_sw = 'Y'" _
        & " AND outr_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1; "

        'cleared legacy checks detail [table 19];
        SSQL20 = "SELECT outc_chk_num, outc_chk_amount" _
        & " FROM outstandingchecks" _
        & " WHERE outc_recon_sw = 'Y'" _
        & " AND outc_stale_sw = 'N'" _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST (outc_chk_num AS INT); "

        'cleared legacy receipts detail [table 20];
        SSQL21 = " SELECT outr_rcpt_num, outr_rcpt_amount" _
        & " FROM outstandingreceipts" _
        & " WHERE outr_recon_sw = 'Y' " _
        & " AND outr_stale_sw = 'N' " _
        & " AND bank_acct_num = @p1" _
        & " ORDER BY CAST(outr_rcpt_num AS INT)"

        SSQL = SSQL1 + SSQL2 + SSQL3 + SSQL4 + SSQL5 + SSQL6 + SSQL7 + SSQL8
        SSQL += SSQL9 + SSQL10 + SSQL11 + SSQL12 + SSQL13 + SSQL14 + SSQL15
        SSQL += SSQL16 + SSQL17 + SSQL18 + SSQL19 + SSQL20 + SSQL21

        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", ecurrentmonthbegdate)
        cmd.Parameters.Add("@p4", ecurrentmonthenddate)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("recontables")
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        'load the grid with some values;
        With Me.GridTotals
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '     0            1           2            3            4       
            '   bank        fisyr      begmonth     endmonth     recondate
            '     5            6           7            8
            ' statement    interest     charges    investments
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .Rows.Count = 0
            .Cols.Count = 9
            .Rows.Add()
            .SetData(0, 0, ebankaccountnumber)
            .SetData(0, 1, efiscalyear)
            .SetData(0, 2, ecurrentmonthbegdate)
            .SetData(0, 3, ecurrentmonthenddate)
            .SetData(0, 4, erecondate)
            .SetData(0, 5, ebankstatement)
            .SetData(0, 6, einterest)
            .SetData(0, 7, echarges)
            .SetData(0, 8, einvestments)
        End With

        Try
            Call PrintTrialBalance(ds, eprintdetail)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try


    End Function

#End Region

#Region "  Methods Rendering "

    Private Sub PrintAdjustmentRegister()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4       
        '   bank       fisyr    docnumber  type    amount
        '     5           6         7        8        9 
        '   acct      subacct     xcode   rcode    applied 
        '    10          11        12 
        ' created      descr     remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "AdjustmentRegister"
        Me.ReportName = "Adjustment Register"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, trankey, count As Int32
        Dim totalregister, totalbank, totalexpend, totalinterest, totalnsf, totalrev, totallegcheck, totallegrcpt As Double
        Dim tissuedate, tapplieddate As Date
        Dim tacctnum, tsubacctnum, tdescr, tremarks As String
        Dim ttrantype, ttypedescr As String
        Dim expcode, revcode, prtcode As String
        Dim tamount As Double

        Try
            'get the totals
            With Me.GridTotals
                totalbank = CDbl(.GetData(0, 0))
                totalexpend = CDbl(.GetData(0, 1))
                totalinterest = CDbl(.GetData(0, 2))
                totalnsf = CDbl(.GetData(0, 3))
                totalrev = CDbl(.GetData(0, 4))
                totallegcheck = CDbl(.GetData(0, 5))
                totallegrcpt = CDbl(.GetData(0, 6))
                totalregister = CDbl(.GetData(0, 7))
                totalrev += totallegrcpt
                totalexpend -= totallegcheck
            End With
            'get the bank account number from the first item
            Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 0), String)
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 1 Then
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y + 4, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y + 4, totalregister.ToString.Format("{0:C2}", totalregister), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Description/Remarks", 40, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        trankey = CInt(.GetData(index, 2))
                        ttrantype = DirectCast(.GetData(index, 3), String).ToUpper
                        tamount = CDbl(.GetData(index, 4))
                        tacctnum = DirectCast(.GetData(index, 5), String)
                        tsubacctnum = DirectCast(.GetData(index, 6), String)
                        expcode = DirectCast(.GetData(index, 7), String).Trim
                        revcode = DirectCast(.GetData(index, 8), String).Trim
                        tapplieddate = CDate(.GetData(index, 9))
                        tissuedate = CDate(.GetData(index, 10))
                        tdescr = DirectCast(.GetData(index, 11), String).Trim
                        tremarks = DirectCast(.GetData(index, 12), String).Trim
                        'describe the type of transaction
                        Select Case ttrantype
                            Case "B"
                                ttypedescr = "Bank charge"
                                prtcode = FormatExpenditureCode(expcode)
                            Case "E"
                                ttypedescr = "Expenditure"
                                prtcode = FormatExpenditureCode(expcode)
                            Case "I"
                                ttypedescr = "Interest"
                                prtcode = FormatRevenueCode(revcode)
                            Case "J"
                                ttypedescr = "Leg.Rcpt"
                                prtcode = FormatRevenueCode(revcode)
                            Case "K"
                                ttypedescr = "Leg.Chk"
                                prtcode = FormatExpenditureCode(expcode)
                            Case "N"
                                ttypedescr = "NSF"
                                prtcode = FormatRevenueCode(revcode)
                            Case "R"
                                ttypedescr = "Revenue"
                                prtcode = FormatRevenueCode(revcode)
                            Case Else
                                ttypedescr = "Other"
                                prtcode = ""
                        End Select
                    End With

                    count += 1
                    If currow > 1 Then y += 5
                    .RenderDirectText(2, y, trankey.ToString.Format("{0:D5}", trankey), 15, 5, verdanaleft8)
                    .RenderDirectText(15, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
                    .RenderDirectText(40, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(65, y, ttypedescr, 25, 5, verdanaleft8)
                    .RenderDirectText(90, y, tdescr, 77, 5, verdanaleft8)
                    .RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                    y += 5
                    If Me.UseOcas Then .RenderDirectText(2, y, prtcode, 80, 5, verdanaleft8)
                    .RenderDirectText(90, y, tremarks, 77, 5, verdanaleft8)
                    y += 2

                    If y >= 245 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Description/Remarks", 40, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 10

                If y > 230 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Add Expenditures", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Add Revenue", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 8, "Less Bank Charges", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 12, "Add Interest", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 16, "Add NSF", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 20, "Total Adjustments", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 26, "Number Of Adjustments", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, totalexpend.ToString.Format("{0:F2}", totalexpend), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, totalrev.ToString.Format("{0:F2}", totalrev), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 8, totalbank.ToString.Format("{0:F2}", totalbank), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 12, totalinterest.ToString.Format("{0:F2}", totalinterest), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 16, totalnsf.ToString.Format("{0:F2}", totalnsf), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 20, totalregister.ToString.Format("{0:C2}", totalregister), 25, 5, verdanaright8bold)
                .RenderDirectText(165, y + 26, count.ToString.Format("{0:D2}", count), 25, 5, verdanaright8bold)
                y += 32
                'draw bottom of total box
                .RenderDirectLine(59, y, 190, y, Color.Black, 0.25)
                .RenderDirectLine(59, y + 0.5, 190, y + 0.5, Color.Black, 0.25)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintAdjustmentTicket()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5       6  
        '   key        fisyr    trantype   bank     acct    subacct   amt  
        '     7           8         9       10       11       12      13
        ' applied     created    descr   remarks   balance  subname acctname
        '    14          15   
        ' expcode     revcode
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "AdjustmentTicket"
        Me.ReportName = "Activity Fund Adjustment"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, x, y, trankey, tfisyr As Int32
        Dim tissued, tapplied As Date
        Dim tacctnum, tsubacctnum, tacctname, tsubacctname As String
        Dim tdescr, tremarks, ttrantype, ttypedescr As String
        Dim texpcode, trevcode, prtcode As String
        Dim tamount As Double

        Try
            With Me.GridDetail
                trankey = CInt(Me.GridDetail.GetData(1, 0))
                tfisyr = CInt(Me.GridDetail.GetData(1, 1))
                ttrantype = DirectCast(Me.GridDetail.GetData(1, 2), String).ToUpper
                Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 3), String)
                tacctnum = DirectCast(Me.GridDetail.GetData(1, 4), String)
                tsubacctnum = DirectCast(Me.GridDetail.GetData(1, 5), String)
                tamount = CDbl(Me.GridDetail.GetData(1, 6))
                tapplied = CDate(Me.GridDetail.GetData(1, 7))
                tissued = CDate(Me.GridDetail.GetData(1, 8))
                tdescr = DirectCast(Me.GridDetail.GetData(1, 9), String).Trim
                tremarks = DirectCast(Me.GridDetail.GetData(1, 10), String).Trim
                tsubacctname = DirectCast(Me.GridDetail.GetData(1, 12), String).Trim
                tacctname = DirectCast(Me.GridDetail.GetData(1, 13), String).Trim
                texpcode = DirectCast(Me.GridDetail.GetData(1, 14), String).Trim
                trevcode = DirectCast(Me.GridDetail.GetData(1, 15), String).Trim

                Select Case ttrantype
                    Case "B"
                        ttypedescr = "Bank charge"
                        prtcode = FormatExpenditureCode(texpcode)
                    Case "E"
                        ttypedescr = "Expenditure"
                        prtcode = FormatExpenditureCode(texpcode)
                    Case "I"
                        ttypedescr = "Interest"
                        prtcode = FormatRevenueCode(trevcode)
                    Case "J"
                        ttypedescr = "Legacy Receipt"
                        prtcode = FormatRevenueCode(trevcode)
                    Case "K"
                        ttypedescr = "Legacy Check"
                        prtcode = FormatExpenditureCode(texpcode)
                    Case "N"
                        ttypedescr = "NSF"
                        prtcode = FormatRevenueCode(trevcode)
                    Case "R"
                        ttypedescr = "Revenue"
                        prtcode = FormatRevenueCode(trevcode)
                    Case Else
                        ttypedescr = "Other"
                        prtcode = ""
                End Select
            End With
            Me.CellMiddleBottom = "FY-" + tfisyr.ToString
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                'print the info box left-side
                y = 32
                .RenderDirectText(0, y, "For Bank Account:", 40, 5, verdanaright8bold)
                .RenderDirectText(0, y + 4, Me.BankAccountNumber, 40, 5, verdanaright8)
                .RenderDirectText(0, y + 8, "For Applied Date:", 40, 5, verdanaright8bold)
                .RenderDirectText(0, y + 12, tapplied.ToShortDateString, 40, 5, verdanaright8)
                'print the info box middle
                .RenderDirectText(65, y, "Account:", 20, 5, verdanaright8bold)
                .RenderDirectText(85, y, tacctname, 65, 5, verdanaleft8)
                .RenderDirectText(65, y + 4, tacctnum + "-" + tsubacctnum, 20, 5, verdanaright8)
                .RenderDirectText(85, y + 4, tsubacctname, 65, 5, verdanaleft8)
                'print the info box left
                .RenderDirectText(145, y, "Adjustment number:", 40, 5, verdanaright8bold)
                .RenderDirectText(145, y + 4, trankey.ToString.Format("{0:D5}", trankey), 40, 5, verdanaright8)
                y = 51
                'print line above the column headers
                .RenderDirectLine(0, y, 190, y, Color.Gray, 0.5)
                y = 58
                'print the column headers
                .RenderDirectText(10, y, "Adjustment issued on:", 50, 5, verdanaright8)
                .RenderDirectText(20, y + 5, tissued.ToString.Format("{0:MM/dd/yyyy}", tissued), 40, 5, verdanaright8bold)
                .RenderDirectText(80, y, "Type:", 30, 5, verdanaright8)
                .RenderDirectText(80, y + 5, ttypedescr, 30, 5, verdanaright8bold)
                .RenderDirectText(120, y, "For amount:", 40, 5, verdanaright8)
                .RenderDirectText(120, y + 5, tamount.ToString.Format("{0:C2}", tamount), 40, 5, verdanaright8bold)
                y = 75
                If Me.UseOcas Then
                    .RenderDirectText(40, y, "Coding:", 30, 5, verdanaleft8)
                    .RenderDirectText(40, y + 5, prtcode, 80, 5, verdanaleft8)
                    y += 15
                End If
                .RenderDirectText(40, y, "Description:", 30, 5, verdanaleft8)
                .RenderDirectText(40, y + 5, tdescr, 120, 10, verdanaleft8)
                .RenderDirectText(40, y + 15, "Remarks:", 30, 5, verdanaleft8)
                .RenderDirectText(40, y + 20, tremarks, 120, 10, verdanaleft8)
                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintDetailOfAccountsNoActivity(ByVal etype As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4      
        '   bank        acct    acctname  subacct  subname   
        '     5           6         7
        ' begmonth     cntadj    cnttrx
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "DetailOfAccountsSingleAccount"
        Me.ReportName = "Detail Of Accounts"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        headerstyle = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        'define the styles 
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, x, y As Int32
        Dim begbalance, amount As Double
        Dim tacctnum, tacctname, tsubacctnum, tsubacctname As String
        Dim lbltype As String

        'etype argument determines whether this report is a ytd, mtd, or periodical
        'so that any labels will reflect the report type;
        Select Case etype
            Case 1
                lbltype = "B e g i n n i n g   y e a r l y   b a l a n c e:"
            Case 2
                lbltype = "B e g i n n i n g   m o n t h l y   b a l a n c e:"
            Case Else
                lbltype = "P e r i o d i c a l :"
        End Select

        Try
            With Me.Doc1
                'special font for a special report
                timesleft16.Font = New Font("Arial", 8, FontStyle.Italic)
                timesleft16.TextColor = Color.Black
                .StartDoc()
                x = 0
                y = 50

                For index = 0 To Me.GridDetail.Rows.Count - 1
                    With Me.GridDetail
                        'collect the record
                        Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                        tacctnum = DirectCast(.GetData(index, 1), String)
                        tacctname = DirectCast(.GetData(index, 2), String)
                        tsubacctnum = DirectCast(.GetData(index, 3), String)
                        tsubacctname = DirectCast(.GetData(index, 4), String)
                        begbalance = CDbl(.GetData(index, 5))
                    End With

                    'page break for each account
                    If index > 0 Then .NewPage()

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE HEADER 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'print the total info box left-side
                    .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                    .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                    .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                    .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                    .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                    'print the info box right-side
                    y = 32
                    .RenderDirectText(118, y, "Beginning balance:", 40, 5, verdanaright8bold)
                    .RenderDirectText(118, y + 4, "Receipts:", 40, 5, verdanaright8bold)
                    .RenderDirectText(118, y + 8, "Checks:", 40, 5, verdanaright8bold)
                    .RenderDirectText(118, y + 12, "Adjustments:", 40, 5, verdanaright8bold)
                    .RenderDirectText(118, y + 18, "Ending balance:", 40, 5, verdanaright8bold)
                    'print the money fields
                    .RenderDirectText(160, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, verdanaright8bold)
                    .RenderDirectText(160, y + 4, amount.ToString.Format("{0:F2}", amount), 30, 5, verdanaright8bold)
                    .RenderDirectText(160, y + 8, amount.ToString.Format("{0:F2}", amount), 30, 5, verdanaright8bold)
                    .RenderDirectText(160, y + 12, amount.ToString.Format("{0:F2}", amount), 30, 5, verdanaright8bold)
                    .RenderDirectText(160, y + 18, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, verdanaright8bold)
                    'print line above the column headers
                    .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                    y = 58
                    'print the column headers
                    .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                    .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                    .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                    .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                    .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                    .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                    .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                    y = 65
                    .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                    .RenderDirectText(160, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, arialright8)
                    y = 70


                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE DETAIL 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    y += 5
                    .RenderDirectText(30, y, "N o  a c t i v i t y   r e p o r t e d:", 80, 5, verdanaleft8)
                    .RenderDirectText(100, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                    .RenderDirectText(120, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                    .RenderDirectText(140, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                    'print the running balance
                    .RenderDirectText(165, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, arialright8)
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try

    End Sub

    Private Sub PrintDetailOfAccountsAllAccounts(ByVal etype As Int32, ByVal efiscalyear As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "DetailOfAccountsSingleAccount"
        Me.ReportName = "Detail Of Accounts"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        headerstyle = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        linestyle1 = New C1DocStyle(Me.Doc1)

        'define the styles 
        DefineStyles()

        With Me.linestyle1
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 7, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With

        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, rowindex, numofaccts As Int32
        Dim acct, subacct, nextacct, nextsubacct As String
        Dim haschanged As Boolean
        'for the totals
        Dim tempacct As String
        Dim begbal, mtdrcpt, mtdexp, mtdadj As Double
        Dim acctname, subname As String

        'at this point, GridWrk contains all the records for all the accounts;
        'we will iterate thru GridWrk and load the detail grid with all
        'the records for each new account, then process the report for that 
        'account/sub only.  Then, get the next account in the list.

        With Me.GridWrk
            'initialise the grid and the document;
            Me.GridDetail.Rows.Count = 0
            Me.GridDetail.Cols.Count = 21
            Me.Doc1.StartDoc()

            For index = 0 To .Rows.Count - 1
                acct = DirectCast(.GetData(index, 9), String)
                subacct = DirectCast(.GetData(index, 10), String)
                If index < .Rows.Count - 1 Then
                    nextacct = DirectCast(.GetData(index + 1, 9), String)
                    nextsubacct = DirectCast(.GetData(index + 1, 10), String)
                Else
                    nextacct = ""
                    nextsubacct = ""
                End If

                'map the row
                Me.GridDetail.Rows.Add()
                Me.GridDetail.SetData(currow, 0, Me.GridWrk.GetData(index, 0))
                Me.GridDetail.SetData(currow, 1, Me.GridWrk.GetData(index, 1))
                Me.GridDetail.SetData(currow, 2, Me.GridWrk.GetData(index, 2))
                Me.GridDetail.SetData(currow, 3, Me.GridWrk.GetData(index, 3))
                Me.GridDetail.SetData(currow, 4, Me.GridWrk.GetData(index, 4))
                Me.GridDetail.SetData(currow, 5, Me.GridWrk.GetData(index, 5))
                Me.GridDetail.SetData(currow, 6, Me.GridWrk.GetData(index, 6))
                Me.GridDetail.SetData(currow, 7, Me.GridWrk.GetData(index, 7))
                Me.GridDetail.SetData(currow, 8, Me.GridWrk.GetData(index, 8))
                Me.GridDetail.SetData(currow, 9, Me.GridWrk.GetData(index, 9))
                Me.GridDetail.SetData(currow, 10, Me.GridWrk.GetData(index, 10))
                Me.GridDetail.SetData(currow, 11, Me.GridWrk.GetData(index, 11))
                Me.GridDetail.SetData(currow, 12, Me.GridWrk.GetData(index, 12))
                Me.GridDetail.SetData(currow, 13, Me.GridWrk.GetData(index, 13))
                Me.GridDetail.SetData(currow, 14, Me.GridWrk.GetData(index, 14))
                Me.GridDetail.SetData(currow, 15, Me.GridWrk.GetData(index, 15))
                Me.GridDetail.SetData(currow, 16, Me.GridWrk.GetData(index, 16))
                Me.GridDetail.SetData(currow, 17, Me.GridWrk.GetData(index, 17))
                Me.GridDetail.SetData(currow, 18, Me.GridWrk.GetData(index, 18))
                Me.GridDetail.SetData(currow, 20, Me.GridWrk.GetData(index, 20))

                If acct.Compare(acct, nextacct) <> 0 Then haschanged = True
                If subacct.Compare(subacct, nextsubacct) <> 0 Then haschanged = True

                If haschanged Then
                    'sort the detail grid by issued/created date
                    Me.GridDetail.Sort(SortFlags.Ascending, 7)
                    'find the acct/sub balance in the working totals grid and load into the totalgrid
                    tempacct = acct + subacct
                    rowindex = Me.GridWrkTotals.FindRow(tempacct, 0, 9, True, True, False)
                    If rowindex >= 0 Then
                        With Me.GridWrkTotals
                            Me.BankAccountNumber = DirectCast(.GetData(rowindex, 0), String)
                            begbal = CDbl(.GetData(rowindex, 1))
                            acctname = DirectCast(.GetData(rowindex, 3), String)
                            subname = DirectCast(.GetData(rowindex, 5), String)
                            mtdrcpt = CDbl(.GetData(rowindex, 6))
                            mtdexp = CDbl(.GetData(rowindex, 7))
                            mtdadj = CDbl(.GetData(rowindex, 8))
                            With Me.GridTotals
                                .Rows.Count = 1
                                .Rows.Add()
                                .SetData(1, 0, Me.BankAccountNumber)
                                .SetData(1, 1, begbal)
                                .SetData(1, 2, acct)
                                .SetData(1, 3, acctname)
                                .SetData(1, 4, subacct)
                                .SetData(1, 5, subname)
                                .SetData(1, 6, mtdrcpt)
                                .SetData(1, 7, mtdexp)
                                .SetData(1, 8, mtdadj)
                            End With
                        End With
                    End If

                    'the account/subaccount changed so process the account
                    If numofaccts >= 1 Then Me.Doc1.NewPage()
                    RenderDetailOfAccountsAllAccounts(etype, efiscalyear)
                    currow = 0
                    haschanged = False
                    numofaccts += 1
                    Me.GridDetail.Rows.Count = 0

                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                    Debug.WriteLine("Processing account:  " & numofaccts.ToString)
                Else
                    currow += 1
                End If
            Next

        End With

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try

    End Sub

    Private Sub PrintDetailOfAccountsSingleAccount(ByVal etype As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "DetailOfAccountsSingleAccount"
        Me.ReportName = "Detail Of Accounts"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        headerstyle = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        linestyle1 = New C1DocStyle(Me.Doc1)

        'define the styles 
        DefineStyles()

        With Me.linestyle1
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Arial", 7, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With

        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y As Int32
        Dim begbalance, mtdrcpt, mtdexpend, mtdadjust As Double
        Dim tacctnum, tacctname, tsubacctnum, tsubacctname As String
        Dim lbltype As String

        'etype argument determines whether this report is a ytd, mtd, or periodical
        'so that any labels will reflect the report type;
        Select Case etype
            Case 1
                lbltype = "B e g i n n i n g   y e a r l y   b a l a n c e:"
            Case 2
                lbltype = "B e g i n n i n g   m o n t h l y   b a l a n c e:"
            Case Else
                lbltype = "P e r i o d i c a l :"
        End Select

        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct   begyrbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'collect the summary & descr information
        Try
            With Me.GridTotals
                Me.BankAccountNumber = DirectCast(.GetData(1, 0), String)
                begbalance = CDbl(.GetData(1, 1))
                tacctnum = DirectCast(.GetData(1, 2), String)
                tacctname = DirectCast(.GetData(1, 3), String)
                tsubacctnum = DirectCast(.GetData(1, 4), String)
                tsubacctname = DirectCast(.GetData(1, 5), String)
                mtdrcpt = CDbl(.GetData(1, 6))
                mtdexpend = CDbl(.GetData(1, 7))
                mtdadjust = CDbl(.GetData(1, 8))
                'if this is a periodical report, then all beginning balance is invalid;
                If etype = 3 Then
                    begbalance = 0
                    mtdrcpt = 0
                    mtdexpend = 0
                    mtdadjust = 0
                End If
            End With
        Catch ex As Exception

        End Try

        Try
            Dim tdocnumber, tdoctype, tstatus, tdescr, tremarks, tpurchaseordernumber As String
            Dim texpcode, trevcode, nextbank, nextacct As String
            Dim ttrxacctfrom, ttrxsubacctfrom, ttrxacctto, ttrxsubacctto As String
            Dim tprevtype As String = ""
            Dim tcreated As Date
            Dim tamount, trunningbalance, ttotalrcvd, ttotalpaid, ttotaladj, calctotal As Double
            Dim tempnumber, key, prevkey, encsw As Int32
            Dim doheader, dodetail, isfirstline As Boolean
            encsw = 0

            With Me.Doc1
                'special font for a special report
                timesleft16.Font = New Font("Arial", 8, FontStyle.Italic)
                timesleft16.TextColor = Color.Black

                .StartDoc()
                x = 0
                y = 50

                'if this is a periodical report, then the beginning balance is invalid;
                If etype = 3 Then begbalance = 0
                'initialise the running balance to the beginning balance
                trunningbalance = begbalance

                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                        .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(9.5, 50, "* J & K adjustment document number", 80, 5, timesleft16)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 4, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 12, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 18, "Ending balance:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 12, mtdadjust.ToString.Format("{0:F2}", mtdadjust), 30, 5, verdanaright8bold)
                        calctotal = begbalance + mtdrcpt + mtdadjust - mtdexpend
                        .RenderDirectText(160, y + 18, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                        '''''.RenderDirectLine(0, 49.5, 190, 49.5, Color.Gray, 0.5)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                        .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                        .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                        .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                        y = 63
                        .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                        .RenderDirectText(160, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, arialright8)
                        y = 68
                    End If

                    'collect the data
                    With Me.GridDetail
                        tempnumber = CInt(.GetData(index, 2))
                        tdocnumber = tempnumber.ToString.Format("{0:D8}", tempnumber)
                        tdoctype = DirectCast(.GetData(index, 3), String)
                        tstatus = DirectCast(.GetData(index, 4), String)
                        tamount = CDbl(.GetData(index, 5))
                        'tapplied = CDate(.GetData(index, 6))
                        tcreated = CDate(.GetData(index, 7))
                        tdescr = DirectCast(.GetData(index, 8), String)
                        If tdescr.Trim.Length > 25 Then tdescr = tdescr.Substring(0, 25) & "..."
                        tremarks = DirectCast(.GetData(index, 11), String)
                        texpcode = DirectCast(.GetData(index, 12), String)
                        trevcode = DirectCast(.GetData(index, 13), String)
                        key = CInt(.GetData(index, 18))
                        tpurchaseordernumber = DirectCast(.GetData(index, 20), String)
                    End With

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE HEADER 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If key <> prevkey Then doheader = True
                    If key = prevkey And tdoctype <> tprevtype Then doheader = True

                    If doheader Then
                        If currow > 1 Then y += 10
                        'issued
                        .RenderDirectText(0, y, tcreated.ToShortDateString, 20, 5, arialleft8)
                        'description
                        .RenderDirectText(16, y, tdescr, 60, 5, arialleft8)
                        'docnumber
                        If tdoctype = "0" And encsw = 0 Then
                            .RenderDirectText(73, y, tdocnumber, 45, 5, verdanaleft8bold)
                            .RenderDirectText(93, y, "PO# " & tpurchaseordernumber, 50, 5, linestyle1)
                            encsw = 1
                        Else
                            .RenderDirectText(73, y, tdocnumber, 25, 5, verdanaleft8bold)
                        End If
                        prevkey = key
                        tprevtype = tdoctype
                        doheader = False
                        dodetail = True
                        isfirstline = True
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE DETAIL 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If isfirstline Then
                        y += 2
                        isfirstline = False
                    Else
                        y += 4
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '''''''''''''''''   added 01-13-15     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If tpurchaseordernumber <> "" And encsw = 0 Then
                        .RenderDirectText(93, y, "PO# " & tpurchaseordernumber, 50, 5, linestyle1)
                        y += 3
                    Else
                        .RenderDirectText(93, y, "            ", 50, 5, linestyle1)
                        encsw = 0
                        y += 3
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    .RenderDirectText(16, y, tremarks, 50, 10, arialleft8)

                    Select Case tdoctype
                        Case "0"    'check;
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Check", 20, 5, timesleft16)
                            .RenderDirectText(65, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(120, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a check
                            trunningbalance -= tamount
                            'sum the checks
                            ttotalpaid += tamount
                            dodetail = False
                        Case "1"    'receipt;
                            trevcode = Module1.FormatRevenueCode(trevcode)
                            If dodetail Then .RenderDirectText(0, y, "Receipt", 20, 5, timesleft16)
                            .RenderDirectText(65, y, trevcode, 40, 5, arialleft8)
                            .RenderDirectText(100, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a receipt
                            trunningbalance += tamount
                            ttotalrcvd += tamount
                            dodetail = False
                        Case "2"    'adjustment;
                            Select Case tstatus
                                Case "B"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the bank into the total adjustments
                                    ttotaladj -= tamount
                                Case "E"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an expenditure adj
                                    trunningbalance += tamount
                                    'sum the expenditures into the total adjustments
                                    ttotaladj += tamount
                                Case "I", "N", "R"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(65, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotaladj += tamount
                                Case "J"    'legacy receipts
                                    .RenderDirectText(0, y, "J Adjust *", 25, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(65, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(100, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotalrcvd += tamount
                                Case "K"    'legacy checks
                                    .RenderDirectText(0, y, "K Adjust *", 25, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(120, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the checks
                                    ttotalpaid += tamount
                            End Select
                        Case "3"    'transfer from
                            .RenderDirectText(0, y, "Trx From", 20, 5, timesleft16)
                            .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance -= tamount
                            ttotaladj -= tamount
                        Case "4"    'transfer to
                            .RenderDirectText(0, y, "Trx To", 20, 5, timesleft16)
                            .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance += tamount
                            ttotaladj += tamount
                        Case Else
                            .RenderDirectText(0, y, "Undefined", 20, 5, timesleft16)
                    End Select

                    'print the running balance
                    .RenderDirectText(165, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 25, 5, arialright8)

                    If y >= 243 Then
                        'page break if not the last record
                        If index < (Me.GridDetail.Rows.Count - 1) Then
                            .NewPage()
                            currow = 0
                            .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                            .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                            .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                            .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                            .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                            .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                            'print the column headers
                            y = 58
                            .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                            .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                            .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                            .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                            .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                            .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                            .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                            .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                            'print the continued balance and line
                            y = 63
                            lbltype = "C o n t i n u e d   f r o m   p r e v i o u s   p a g e ..."
                            .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                            .RenderDirectText(160, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)
                            y = 68
                        End If
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                y += 15
                If y >= 243 Then
                    .NewPage()
                    y = 70
                End If

                'draw top of total box
                .RenderDirectLine(59, y - 1, 190, y - 1, Color.Black, 0.25)
                .RenderDirectLine(59, y - 0.5, 190, y - 0.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Totals", 25, 5, verdanaleft8bold)
                .RenderDirectText(75, y, "Beginning", 25, 5, verdanaright8bold)
                .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(75, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, verdanaright8bold)
                .RenderDirectText(100, y, ttotalrcvd.ToString.Format("{0:F2}", ttotalrcvd), 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 25, 5, verdanaright8bold)
                .RenderDirectText(140, y, ttotaladj.ToString.Format("{0:F2}", ttotaladj), 25, 5, verdanaright8bold)
                'calc the total
                calctotal = begbalance + ttotalrcvd + ttotaladj - ttotalpaid
                .RenderDirectText(165, y, calctotal.ToString.Format("{0:F2}", calctotal), 25, 5, verdanaright8bold)
                'draw bottom of total box
                .RenderDirectLine(59, y + 5, 190, y + 5, Color.Black, 0.25)
                .RenderDirectLine(59, y + 5.5, 190, y + 5.5, Color.Black, 0.25)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            'Me.Doc1.RenderBlock(rendertbl)
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintEncumbranceDetailOfAccountsAllAccounts(ByVal etype As Int32, ByVal efiscalyear As Int32)
        ''''''''''''''''''''' GRIDWRK ''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "EncumbranceDetailOfAccountsSingleAccount"
        Me.ReportName = "Encumbrance Detail Of Accounts"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        headerstyle = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        'define the styles 
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, rowindex, numofaccts As Int32
        Dim tbankaccount As String
        Dim acct, subacct, nextacct, nextsubacct As String
        Dim haschanged As Boolean
        'for the totals
        Dim tempacct As String
        Dim begbal, mtdrcpt, mtdexp, mtdadj, mtdencum As Double
        Dim acctname, subname As String

        'at this point, GridWrk contains all the records for all the accounts;
        'we will iterate thru GridWrk and load the detail grid with all
        'the records for each new account, then process the report for that 
        'account/sub only.  Then, get the next account in the list.

        With Me.GridWrk
            'initialise the grid and the document;
            Me.GridDetail.Rows.Count = 0
            Me.GridDetail.Cols.Count = 20
            Me.Doc1.StartDoc()

            For index = 0 To .Rows.Count - 1
                acct = DirectCast(.GetData(index, 9), String)
                subacct = DirectCast(.GetData(index, 10), String)
                If index < .Rows.Count - 1 Then
                    nextacct = DirectCast(.GetData(index + 1, 9), String)
                    nextsubacct = DirectCast(.GetData(index + 1, 10), String)
                Else
                    nextacct = ""
                    nextsubacct = ""
                End If

                'map the row
                Me.GridDetail.Rows.Add()
                Me.GridDetail.SetData(currow, 0, Me.GridWrk.GetData(index, 0))
                Me.GridDetail.SetData(currow, 1, Me.GridWrk.GetData(index, 1))
                Me.GridDetail.SetData(currow, 2, Me.GridWrk.GetData(index, 2))
                Me.GridDetail.SetData(currow, 3, Me.GridWrk.GetData(index, 3))
                Me.GridDetail.SetData(currow, 4, Me.GridWrk.GetData(index, 4))
                Me.GridDetail.SetData(currow, 5, Me.GridWrk.GetData(index, 5))
                Me.GridDetail.SetData(currow, 6, Me.GridWrk.GetData(index, 6))
                Me.GridDetail.SetData(currow, 7, Me.GridWrk.GetData(index, 7))
                Me.GridDetail.SetData(currow, 8, Me.GridWrk.GetData(index, 8))
                Me.GridDetail.SetData(currow, 9, Me.GridWrk.GetData(index, 9))
                Me.GridDetail.SetData(currow, 10, Me.GridWrk.GetData(index, 10))
                Me.GridDetail.SetData(currow, 11, Me.GridWrk.GetData(index, 11))
                Me.GridDetail.SetData(currow, 12, Me.GridWrk.GetData(index, 12))
                Me.GridDetail.SetData(currow, 13, Me.GridWrk.GetData(index, 13))
                Me.GridDetail.SetData(currow, 14, Me.GridWrk.GetData(index, 14))
                Me.GridDetail.SetData(currow, 15, Me.GridWrk.GetData(index, 15))
                Me.GridDetail.SetData(currow, 16, Me.GridWrk.GetData(index, 16))
                Me.GridDetail.SetData(currow, 17, Me.GridWrk.GetData(index, 17))
                Me.GridDetail.SetData(currow, 18, Me.GridWrk.GetData(index, 18))

                If acct.Compare(acct, nextacct) <> 0 Then haschanged = True
                If subacct.Compare(subacct, nextsubacct) <> 0 Then haschanged = True

                If haschanged Then
                    'sort the detail grid by issued/created date
                    Me.GridDetail.Sort(SortFlags.Ascending, 7)
                    'find the acct/sub balance in the working totals grid and load into the totalgrid
                    tempacct = acct + subacct
                    rowindex = Me.GridWrkTotals.FindRow(tempacct, 0, 10, True, True, False)
                    If rowindex >= 0 Then
                        With Me.GridWrkTotals
                            tbankaccount = CType(.GetData(rowindex, 0), String)
                            begbal = CDbl(.GetData(rowindex, 1))
                            acctname = DirectCast(.GetData(rowindex, 3), String)
                            subname = DirectCast(.GetData(rowindex, 5), String)
                            mtdrcpt = CDbl(.GetData(rowindex, 6))
                            mtdexp = CDbl(.GetData(rowindex, 7))
                            mtdadj = CDbl(.GetData(rowindex, 8))
                            mtdencum = CDbl(.GetData(rowindex, 9))

                            With Me.GridTotals
                                .Rows.Count = 1
                                .Rows.Add()
                                .SetData(1, 0, tbankaccount)
                                .SetData(1, 1, begbal)
                                .SetData(1, 2, acct)
                                .SetData(1, 3, acctname)
                                .SetData(1, 4, subacct)
                                .SetData(1, 5, subname)
                                .SetData(1, 6, mtdrcpt)
                                .SetData(1, 7, mtdexp)
                                .SetData(1, 8, mtdadj)
                                .SetData(1, 9, mtdencum)
                            End With
                        End With
                    End If

                    'the account/subaccount changed so process the account
                    If numofaccts >= 1 Then Me.Doc1.NewPage()
                    RenderEncumbranceDetailOfAccountsAllAccounts(etype, efiscalyear)
                    currow = 0
                    haschanged = False
                    numofaccts += 1
                    Me.GridDetail.Rows.Count = 0

                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                    Debug.WriteLine("Processing account:  " & numofaccts.ToString)
                Else
                    currow += 1
                End If
            Next

        End With

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try

    End Sub

    Private Sub PrintEncumbranceDetailOfAccountsSingleAccount(ByVal efiscalyear As Int32, ByVal etype As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "EncumbranceDetailOfAccountsSingleAccount"
        Me.ReportName = "Encumbrance Detail Of Accounts"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        headerstyle = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)

        'define the styles 
        Dim subscript As C1DocStyle = New C1DocStyle(Me.Doc1)
        With subscript
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 7, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With

        DefineStyles()

        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow As Int32
        Dim x, y As Double
        Dim begbalance, mtdrcpt, mtdencumbrance, mtdexpend, mtdadjust As Decimal
        Dim tbankaccount, tacctnum, tacctname, tsubacctnum, tsubacctname As String
        Dim lbltype As String

        'etype argument determines whether this report is a ytd, mtd, or periodical
        'so that any labels will reflect the report type;
        Select Case etype
            Case 1
                lbltype = "B e g i n n i n g   y e a r l y   b a l a n c e:"
            Case 2
                lbltype = "B e g i n n i n g   m o n t h l y   b a l a n c e:"
            Case Else
                lbltype = "P e r i o d i c a l :"
        End Select


        Try
            'collect the summary & descr information
            With Me.GridTotals
                ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
                '     0           1          2         3          4         5   
                '  bankacct     begbal    acctnum   acctname   subnum   subname 
                '     6           7          8         9         10        11     
                '  mtdrcpts  mtdexpend    mtdadj    mtdencnum
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Me.BankAccountNumber = DirectCast(.GetData(1, 0), String)
                tbankaccount = DirectCast(.GetData(1, 0), String)
                begbalance = CType(.GetData(1, 1), Decimal)
                tacctnum = DirectCast(.GetData(1, 2), String)
                tacctname = DirectCast(.GetData(1, 3), String)
                tsubacctnum = DirectCast(.GetData(1, 4), String)
                tsubacctname = DirectCast(.GetData(1, 5), String)
                mtdrcpt = CType(.GetData(1, 6), Decimal)
                mtdexpend = CType(.GetData(1, 7), Decimal)
                mtdadjust = CType(.GetData(1, 8), Decimal)
                mtdencumbrance = CType(.GetData(1, 9), Decimal)
                'if this is a periodical report, then all beginning balance is invalid;
                If etype = 3 Then
                    begbalance = 0
                    mtdrcpt = 0
                    mtdexpend = 0
                    mtdadjust = 0
                End If
            End With
        Catch ex As Exception

        End Try


        Try
            Dim tdocnumber, tdoctype, tstatus, tdescr, tremarks As String
            Dim texpcode, trevcode, nextbank, nextacct As String
            Dim ttrxacctfrom, ttrxsubacctfrom, ttrxacctto, ttrxsubacctto As String
            Dim tprevtype As String = ""
            Dim tcreated As Date
            Dim tamount, trunningbalance, ttotalrcvd, ttotalpaid, ttotaladj, ttotalvoid, calctotal As Decimal
            Dim tendingbalance, ttotalencumbered, tencumberedoutstanding, tunpaid, pcalctotal, penctotal As Decimal
            Dim tempnumber, key, prevkey As Int32
            Dim doheader, dodetail, isfirstline As Boolean

            With Me.Doc1
                'special font for a special report
                timesleft16.Font = New Font("Arial", 8, FontStyle.Italic)
                timesleft16.TextColor = Color.Black

                .StartDoc()
                x = 0
                y = 50

                'if this is a periodical report, then the beginning balance is invalid;
                If etype = 3 Then begbalance = 0
                'initialise the running balance to the beginning balance
                trunningbalance = begbalance

                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                        .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                        'print a legend for the report;
                        .RenderDirectText(0.5, 50, "* J & K adjustment document number", 80, 5, subscript)
                        'print the info box right-side
                        x = 185
                        y = 32
                        .RenderDirectText(x, y, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 8, "Encumbrances:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 12, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 16, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 22, "Ending balance:", 40, 5, verdanaright8bold)
                        x = 225
                        'print the money fields
                        .RenderDirectText(x, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 8, mtdencumbrance.ToString.Format("{0:F2}", mtdencumbrance), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 12, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 16, mtdadjust.ToString.Format("{0:F2}", mtdadjust), 30, 5, verdanaright8bold)
                        calctotal = begbalance + mtdrcpt + mtdadjust - mtdexpend
                        .RenderDirectText(x, y + 22, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                        y = 60
                        'print line above the column headers
                        .RenderDirectLine(0, y, 255, y, Color.Gray, 0.5)
                        x = 120
                        y = 62
                        'print the column headers
                        .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                        .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                        .RenderDirectText(x, y, "Received", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 30, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 50, y, "Paid Out", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 75, y, "Adjusted", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 105, y, "Balance", 25, 5, verdanaright8bold)
                        y = 68
                        .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                        .RenderDirectText(220, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, arialright8)
                        y = 74
                    End If

                    'collect the data
                    With Me.GridDetail
                        tempnumber = CInt(.GetData(index, 2))
                        tdocnumber = tempnumber.ToString.Format("{0:D8}", tempnumber)
                        tdoctype = DirectCast(.GetData(index, 3), String)
                        tstatus = DirectCast(.GetData(index, 4), String)
                        tamount = CType(.GetData(index, 5), Decimal)
                        'tapplied = CDate(.GetData(index, 6))
                        tcreated = CDate(.GetData(index, 7))
                        tdescr = DirectCast(.GetData(index, 8), String)
                        If tdescr.Trim.Length > 25 Then tdescr = tdescr.Substring(0, 25) & "..."
                        tremarks = DirectCast(.GetData(index, 11), String)
                        If tremarks.Trim.Length > 30 Then tremarks = tremarks.Substring(0, 30) & "..."
                        texpcode = DirectCast(.GetData(index, 12), String)
                        trevcode = DirectCast(.GetData(index, 13), String)
                        key = CInt(.GetData(index, 18))
                    End With

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE HEADER 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If key <> prevkey Then doheader = True
                    If key = prevkey And tdoctype <> tprevtype Then doheader = True

                    If doheader Then
                        If currow > 1 Then y += 8
                        'issued
                        .RenderDirectText(0, y, tcreated.ToShortDateString, 20, 5, arialleft8)
                        'description
                        .RenderDirectText(16, y, tdescr, 60, 5, arialleft8)
                        'docnumber
                        .RenderDirectText(75, y, tdocnumber, 25, 5, verdanaleft8bold)
                        prevkey = key
                        tprevtype = tdoctype
                        doheader = False
                        dodetail = True
                        isfirstline = True
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE DETAIL 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If isfirstline Then
                        y += 4
                        isfirstline = False
                    Else
                        y += 4
                    End If

                    .RenderDirectText(16, y, tremarks, 55, 5, arialleft8)

                    Select Case tdoctype
                        Case "0"
                            x = 170
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Check", 20, 5, timesleft16)
                            .RenderDirectText(70, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a check
                            trunningbalance -= tamount
                            'sum the checks
                            ttotalpaid += tamount
                            dodetail = False
                        Case "1"
                            x = 120
                            trevcode = Module1.FormatRevenueCode(trevcode)
                            If dodetail Then .RenderDirectText(0, y, "Receipt", 20, 5, timesleft16)
                            .RenderDirectText(70, y, trevcode, 40, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a receipt
                            trunningbalance += tamount
                            ttotalrcvd += tamount
                            dodetail = False
                        Case "2"
                            x = 195
                            Select Case tstatus
                                Case "B"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the bank into the total adjustments
                                    ttotaladj -= tamount
                                Case "E"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an expenditure adj
                                    trunningbalance += tamount
                                    'sum the expenditures into the total adjustments
                                    ttotaladj += tamount
                                Case "I", "N", "R"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(70, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotaladj += tamount
                                Case "J"    'legacy receipts
                                    x = 120
                                    .RenderDirectText(0, y, "J Adjust", 25, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(-2.5, y + 0.5, "*", 5, 5, subscript)
                                    .RenderDirectText(70, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotalrcvd += tamount
                                Case "K"    'legacy checks
                                    x = 170
                                    .RenderDirectText(0, y, "K Adjust", 25, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(-2.5, y + 0.5, "*", 5, 5, subscript)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the checks
                                    ttotalpaid += tamount
                            End Select
                        Case "3"    'transfer from
                            x = 195
                            .RenderDirectText(0, y, "Trx From", 20, 5, timesleft16)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance -= tamount
                            ttotaladj -= tamount
                        Case "4"    'transfer to
                            x = 195
                            .RenderDirectText(0, y, "Trx To", 20, 5, timesleft16)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance += tamount
                            ttotaladj += tamount
                        Case "5"
                            x = 150
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Encum.", 20, 5, timesleft16)
                            .RenderDirectText(70, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'ttotalencumbered, tencumberedoutstanding, tencumberedpaid
                            Select Case tstatus
                                Case "O", "P"
                                    '.RenderDirectText(-2.5, y, "1", 5, 5, subscript)
                                    tencumberedoutstanding += tamount
                                    ttotalencumbered += tamount
                                Case "C"
                                    ttotalencumbered += tamount
                                    '.RenderDirectText(-2.5, y, "2", 5, 5, subscript)
                                Case Else
                                    ttotalencumbered += tamount
                                    .RenderDirectText(-2.5, y + 0.5, "x", 5, 5, subscript)
                            End Select
                            dodetail = False
                        Case Else
                            .RenderDirectText(0, y, "Undefined", 20, 5, timesleft16)
                    End Select

                    'print the running balance
                    .RenderDirectText(220, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)

                    If y >= 185 Then
                        'page break if not the last record
                        If index < (Me.GridDetail.Rows.Count - 1) Then
                            .NewPage()
                            currow = 0
                            .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                            .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                            .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                            .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                            .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                            .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                            'print a legend for the report;
                            .RenderDirectText(0.5, 50, "* J & K adjustment document number", 80, 5, subscript)
                            y = 60
                            'print line above the column headers
                            .RenderDirectLine(0, y, 255, y, Color.Gray, 0.5)
                            x = 120
                            y = 62
                            'print the column headers;
                            .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                            .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                            .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                            .RenderDirectText(x, y, "Received", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 30, y, "Encumbered", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 53, y, "Paid Out", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 75, y, "Adjusted", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 105, y, "Balance", 25, 5, verdanaright8bold)
                            'print the continued balance and line;
                            y = 68
                            lbltype = "C o n t i n u e d   f r o m   p r e v i o u s   p a g e ..."
                            .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                            .RenderDirectText(220, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)
                            y = 74
                        End If
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'end of detail;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Dim SSQL As String
                Dim cmd As SqlCommand

                Try
                    'collect the total void amount;
                    SSQL = "SELECT ISNULL(SUM(invc_amount), 0.0) FROM invoices" _
                    & " WHERE bank_acct_num = @p1 AND invc_fisyr = @p2 AND invc_status = 'V'" _
                    & " AND af_acct_num = @p3 AND as_acct_num = @p4"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", tbankaccount)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubacctnum)
                    cn.Open()
                    ttotalvoid = CType(cmd.ExecuteScalar, Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                End Try

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'page break for totals, which will be printed on a page by themselves;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                .NewPage()
                y = 60

                'draw top of total box;
                .RenderDirectLine(60, y - 0.5, 250, y - 0.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Totals:", 25, 5, verdanaleft8bold)
                .RenderDirectText(75, y, "Beginning", 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, "Received", 25, 5, verdanaright8bold)
                .RenderDirectText(150, y, "Encumbered", 25, 5, verdanaright8bold)
                .RenderDirectText(173, y, "Paid Out", 25, 5, verdanaright8bold)
                .RenderDirectText(195, y, "Adjusted", 25, 5, verdanaright8bold)
                .RenderDirectText(220, y, "Ending", 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(75, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, ttotalrcvd.ToString.Format("{0:F2}", ttotalrcvd), 25, 5, verdanaright8bold)
                .RenderDirectText(150, y, ttotalencumbered.ToString.Format("{0:F2}", ttotalencumbered), 25, 5, verdanaright8bold)
                .RenderDirectText(173, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 25, 5, verdanaright8bold)
                .RenderDirectText(195, y, ttotaladj.ToString.Format("{0:F2}", ttotaladj), 25, 5, verdanaright8bold)
                tendingbalance = begbalance + ttotalrcvd + ttotaladj - ttotalpaid
                .RenderDirectText(220, y, tendingbalance.ToString.Format("{0:F2}", tendingbalance), 30, 5, verdanaright8bold)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'print the totals on the last page;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                y = 80
                .RenderDirectText(160, y, "Total Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalencumbered.ToString.Format("{0:F2}", ttotalencumbered), 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(160, y, "Less Paid Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 30, 5, verdanaright8bold)
                verdanaright8bold.TextColor = Color.Black
                y += 5
                'calculate unpaid encumbrances;
                tunpaid = ttotalencumbered - ttotalpaid

                .RenderDirectText(160, y, "Unpaid Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tunpaid.ToString.Format("{0:F2}", tunpaid), 30, 5, verdanaright8bold)
                y += 5
                'total voids from query above;
                .RenderDirectText(160, y, "Add Total Voids:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalvoid.ToString.Format("{0:F2}", ttotalvoid), 30, 5, verdanaright8bold)
                y += 5
                'calculate outstanding encumbrance;
                tencumberedoutstanding = tunpaid + ttotalvoid
                .RenderDirectText(160, y, "Total Outstanding Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tencumberedoutstanding.ToString.Format("{0:F2}", tencumberedoutstanding), 30, 5, verdanaright8bold)
                'extra line break;
                y += 10
                .RenderDirectText(160, y, "Ending Balance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(160, y, "Less Outstanding Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tencumberedoutstanding.ToString.Format("{0:F2}", tencumberedoutstanding), 30, 5, verdanaright8bold)
                y += 5
                'calculate projected balance;
                pcalctotal = tendingbalance - tencumberedoutstanding
                .RenderDirectText(160, y, "Projected Balance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, pcalctotal.ToString.Format("{0:F2}", pcalctotal), 30, 5, verdanaright8bold)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom;
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document;
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintTrialBalance(ByVal ds As DataSet, ByVal dodetail As Boolean)
        Me.DocumentName = "Reconciliation"
        Me.ReportName = "Reconciliation"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim currow, count As Int32
        Dim x, y As Single
        Dim tfiscalyear As Int32
        Dim trecondate As Date
        Dim ttrantype, ttypedescr As String
        Dim expcode, revcode, prtcode As String
        Dim calcval, tstatement, tinterest, tcharges, tinvestments As Double
        Dim leftendingbalance, leftbankbegbalance, leftsumreceipts, leftsumchecks, leftsumadjustments As Double
        Dim rightendingbalance, rightsumchecks, rightsumreceipts, rightlegchecks, rightlegreceipts As Double
        'get values from the tables in the dataset
        Try
            If Not IsDBNull(ds.Tables(0).Rows(0).Item(0)) Then leftbankbegbalance = CDbl(ds.Tables(0).Rows(0).Item(0)) Else leftbankbegbalance = 0
            If Not IsDBNull(ds.Tables(1).Rows(0).Item(0)) Then leftsumreceipts = CDbl(ds.Tables(1).Rows(0).Item(0)) Else leftsumreceipts = 0
            If Not IsDBNull(ds.Tables(2).Rows(0).Item(0)) Then leftsumchecks = CDbl(ds.Tables(2).Rows(0).Item(0)) Else leftsumchecks = 0
            If Not IsDBNull(ds.Tables(3).Rows(0).Item(0)) Then leftsumadjustments = CDbl(ds.Tables(3).Rows(0).Item(0)) Else leftsumadjustments = 0
            If Not IsDBNull(ds.Tables(4).Rows(0).Item(0)) Then rightsumreceipts = CDbl(ds.Tables(4).Rows(0).Item(0)) Else rightsumreceipts = 0
            If Not IsDBNull(ds.Tables(5).Rows(0).Item(0)) Then rightsumchecks = CDbl(ds.Tables(5).Rows(0).Item(0)) Else rightsumchecks = 0
            'legacy checks
            If Not IsDBNull(ds.Tables(13).Rows(0).Item(0)) Then rightlegchecks = CDbl(ds.Tables(13).Rows(0).Item(0)) Else rightlegchecks = 0
            'legacy receipts
            If Not IsDBNull(ds.Tables(14).Rows(0).Item(0)) Then rightlegreceipts = CDbl(ds.Tables(14).Rows(0).Item(0)) Else rightlegreceipts = 0
        Catch ex As Exception
            Throw New ArgumentException("An error occurred during render of main page." & vbCrLf & vbCrLf & ex.Message)
        End Try


        Try
            With Me.GridTotals
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
                tfiscalyear = CInt(.GetData(0, 1))
                trecondate = CDate(.GetData(0, 4))
                tstatement = CDbl(.GetData(0, 5))
                tinterest = CDbl(.GetData(0, 6))
                tcharges = CDbl(.GetData(0, 7))
                tinvestments = CDbl(.GetData(0, 8))
            End With
        Catch ex As Exception
            Throw New ArgumentException("An error occurred during header data assignment." & vbCrLf & vbCrLf & ex.Message)
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                currow += 1
                'print the total info box left-side
                y = 44
                .RenderDirectText(0, y, "Bank account:", 30, 5, verdanaright8)
                .RenderDirectText(0, y + 4, Me.BankAccountNumber, 30, 5, verdanaright8)
                .RenderDirectText(40, y, "Reconciliation date: ", 40, 5, verdanaright8)
                .RenderDirectText(40, y + 4, trecondate.ToShortDateString, 40, 5, verdanaright8)
                .RenderDirectText(100, y, "Prepared by: ", 40, 5, verdanaleft8)
                .RenderDirectText(100, y + 4, Me.UserName, 40, 5, verdanaleft8)
                .RenderDirectText(140, y, "For applied period: ", 40, 5, verdanaright8)
                .RenderDirectText(140, y + 4, Me.FiscalMonthStr & ", " & Me.FiscalYear.ToString, 40, 5, verdanaright8bold)
                'print line above the column headers
                .RenderDirectLine(0, 65, 190, 65, Color.Gray, 0.5)
                x = 0
                y = 68
                'print the left side totals
                .RenderDirectText(x, y, "General ledger account balance", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 10, "Add debits", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 15, "Less credits", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 20, "Add adjustments", 50, 5, verdanaleft8)
                x = 45
                .RenderDirectText(x, y, leftbankbegbalance.ToString.Format("{0:C2}", leftbankbegbalance), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 10, leftsumreceipts.ToString.Format("{0:C2}", leftsumreceipts), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 15, leftsumchecks.ToString.Format("{0:C2}", leftsumchecks), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 20, leftsumadjustments.ToString.Format("{0:C2}", leftsumadjustments), 50, 5, verdanaright8)
                'print the right side totals
                x = 100
                .RenderDirectText(x, y, "Balance per bank statement as of reconciliation date", 50, 10, verdanaleft8)
                .RenderDirectText(x, y + 10, "Add receipts in transit", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 15, "Less outstanding checks", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 20, "Interest not yet posted", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 25, "Charges not yet posted", 50, 5, verdanaleft8)
                .RenderDirectText(x, y + 30, "Investments", 50, 5, verdanaleft8)
                x = 140
                .RenderDirectText(x, y, tstatement.ToString.Format("{0:C2}", tstatement), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 10, rightsumreceipts.ToString.Format("{0:C2}", rightsumreceipts), 50, 5, verdanaright8)
                'calc the total outstanding checks 
                calcval = rightsumchecks + rightlegchecks
                .RenderDirectText(x, y + 15, calcval.ToString.Format("{0:C2}", calcval), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 20, tinterest.ToString.Format("{0:C2}", tinterest), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 25, tcharges.ToString.Format("{0:C2}", tcharges), 50, 5, verdanaright8)
                .RenderDirectText(x, y + 30, tinvestments.ToString.Format("{0:C2}", tinvestments), 50, 5, verdanaright8)
                'print the left side balance
                x = 0
                y = 115
                leftendingbalance = leftbankbegbalance + leftsumreceipts + leftsumadjustments - leftsumchecks
                .RenderDirectText(x, y, "Bank Balance Per General Ledger", 40, 10, verdanaleft8bold)
                .RenderDirectText(x + 45, y, leftendingbalance.ToString.Format("{0:C2}", leftendingbalance), 50, 5, verdanaright8bold)
                'print the right side balance
                x = 100
                rightendingbalance = tstatement + (rightsumreceipts + rightlegreceipts) - (rightsumchecks + rightlegchecks) - tinterest + tcharges + tinvestments
                .RenderDirectText(x, y, "Bank Balance Per Statement Reconciliation", 50, 10, verdanaleft8bold)
                .RenderDirectText(x + 40, y, rightendingbalance.ToString.Format("{0:C2}", rightendingbalance), 50, 5, verdanaright8bold)
                'print the variance
                x = 60
                y = 130
                calcval = leftendingbalance - rightendingbalance
                .RenderDirectText(x, y, "Variance:   " & calcval.ToString.Format("{0:C2}", calcval), 50, 5, verdanaright8bold)
                .RenderDirectText(x + 50, y, "***", 10, 5, verdanaleft8bold)
                'draw a line under the variance
                .RenderDirectLine(0, 138, 190, 138, Color.Gray, 0.5)
            End With
            'expose the current record & count to the caller
            'EventRecordProcessed((reccurrent), reccount)
        Catch ex As Exception
            Throw New ArgumentException("An error occurred during render of document header page." & vbCrLf & vbCrLf & ex.Message)
        End Try

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'at this point, the summary page has been completed; if this
        'is a detailed trial balance, then the report will continue, 
        'else we'll exit right here;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not dodetail Then GoTo SummaryOnly
        Dim row As DataRow
        Dim value As String
        Dim amount, sumamount As Decimal
        Dim counter, items, primarykey As Int32

        Try
            With Me.Doc1
                'use a special style
                Dim specialleft8 As New C1DocStyle(Me.Doc1)
                Dim specialright8 As New C1DocStyle(Me.Doc1)

                With specialleft8
                    .Font = New Font("Verdana", 8, FontStyle.Underline)
                    .TextAlignHorz = AlignHorzEnum.Left
                End With
                With specialright8
                    .Font = New Font("Arial", 8, FontStyle.Underline)
                    .TextAlignHorz = AlignHorzEnum.Right
                End With

                'start a new page
                .NewPage()

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do outstanding deposits
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y = 35
                .RenderDirectText(x - 5, y, "Outstanding Receipts", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(6).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(6).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Outstanding Receipts:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do outstanding checks
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Outstanding Checks", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(7).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(7).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Outstanding Checks:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do cleared deposits
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Receipts Cleared This Month", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(8).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(8).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Receipts Cleared:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do cleared checks;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Checks Cleared This Month", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(9).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(9).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Cleared Checks:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do adjustments;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Adjustments This Month", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(10).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    Dim adjtype As String
                    For Each row In ds.Tables(10).Rows
                        primarykey = CInt(row.Item(0))
                        value = primarykey.ToString.Format("{0:D8}", primarykey)
                        amount = CDec(row.Item(1))
                        adjtype = DirectCast(row.Item(2), String).ToUpper
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)

                        If adjtype = "B" Then
                            sumamount -= amount
                        Else
                            sumamount += amount
                        End If

                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Adjustments:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0


                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do void receipts;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                '.RenderDirectText(x - 5, y, "Deposits Voided This Month", 50, 5, verdanaleft8bold)
                .RenderDirectText(x - 5, y, "Receipts Voided This Month", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(11).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(11).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    '.RenderDirectText(x, y, "Total Void Deposits:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x, y, "Total Void Receipts:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do void checks;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Checks Voided This Month", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(12).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    amount = 0D

                    For Each row In ds.Tables(12).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))

                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)

                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next


                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Void Checks:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do outstanding legacy checks;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Legacy Checks Outstanding", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(15).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(15).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Legacy Checks:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do outstanding legacy receipts;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Legacy Receipts Outstanding", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(16).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(16).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Legacy Receipts:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do cleared legacy checks;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Legacy Checks Cleared", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(19).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(19).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Legacy Checks:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0



                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'do cleared legacy receipts;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                x = 10
                y += 10
                'check for end-of-page
                If y > 240 Then .NewPage() : y = 35
                .RenderDirectText(x - 5, y, "Legacy Receipts Cleared", 50, 5, verdanaleft8bold)
                y += 5
                If ds.Tables(20).Rows.Count > 0 Then
                    .RenderDirectText(x + 1, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 20, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 61, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 80, y, "Amount", 25, 5, specialright8)
                    .RenderDirectText(x + 121, y, "Number", 50, 5, specialleft8)
                    .RenderDirectText(x + 140, y, "Amount", 25, 5, specialright8)
                    y += 5
                    For Each row In ds.Tables(20).Rows
                        value = DirectCast(row.Item(0), String)
                        amount = CDec(row.Item(1))
                        If y > 250 Then .NewPage() : y = 35
                        .RenderDirectText(x, y, value, 25, 5, arialleft8)
                        .RenderDirectText(x + 20, y, amount.ToString.Format("{0:F2}", amount), 25, 5, arialright8)
                        sumamount += amount
                        counter += 1
                        items += 1
                        If counter = 3 Then
                            x = 10
                            y += 5
                            counter = 0
                        Else
                            x += 60
                        End If
                    Next
                    'print the totals;
                    If counter = 0 Then y -= 5
                    x = 10
                    y += 7
                    If y > 250 Then .NewPage() : y = 35
                    .RenderDirectText(x, y, "Total Legacy Receipts:", 50, 5, verdanaleft8bold)
                    .RenderDirectText(x + 140, y, "Items:", 25, 5, verdanaright8bold)
                    y += 5
                    .RenderDirectText(x, y, sumamount.ToString.Format("{0:C2}", sumamount), 45, 5, verdanaright8bold)
                    .RenderDirectText(x + 140, y, items.ToString.Format("{0:D1}", items), 25, 5, verdanaright8bold)
                Else
                    'no transactions exist
                    .RenderDirectText(x - 5, y, "No Transactions", 50, 5, verdanaleft8)
                End If
                counter = 0
                sumamount = 0
                items = 0
            End With
        Catch ex As Exception
            Throw New ArgumentException("An error occurred during render of document detail page." & vbCrLf & vbCrLf & ex.Message)
        End Try

SummaryOnly:

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintTransferRegister()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0          1          2        3         4         5  
        '   bank       fisyr    docnumber  amount  acctfrom   subfrom
        '     6          7          8        9        10        11 
        '  acctto      subto     applied  created    descr    remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "TransferRegister"
        Me.ReportName = "Transfer Register"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, trxkey, count As Int32
        Dim totalregister, tamount As Double
        Dim tissuedate, tapplieddate As Date
        Dim tacctfrom, tsubfrom, tacctto, tsubto, tdescr, tremarks As String

        Try
            'get the totals
            totalregister = CDbl(Me.GridTotals.GetData(0, 0))
            'get the bank account number from the first item
            Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 0), String)
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 1 Then
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y + 4, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y + 4, totalregister.ToString.Format("{0:C2}", totalregister), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Source / Destination", 40, 5, verdanaleft8bold)
                        .RenderDirectText(80, y, "Description/Remarks", 40, 5, verdanaleft8bold)
                        .RenderDirectText(145, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        trxkey = CInt(.GetData(index, 2))
                        tamount = CDbl(.GetData(index, 3))
                        tacctfrom = DirectCast(.GetData(index, 4), String)
                        tsubfrom = DirectCast(.GetData(index, 5), String)
                        tacctto = DirectCast(.GetData(index, 6), String)
                        tsubto = DirectCast(.GetData(index, 7), String)
                        tapplieddate = CDate(.GetData(index, 8))
                        tissuedate = CDate(.GetData(index, 9))
                        tdescr = DirectCast(.GetData(index, 10), String).Trim
                        tremarks = DirectCast(.GetData(index, 11), String).Trim
                    End With

                    count += 1
                    If currow > 1 Then y += 5
                    .RenderDirectText(2, y, trxkey.ToString.Format("{0:D5}", trxkey), 15, 5, verdanaleft8)
                    '.RenderDirectText(15, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)

                    .RenderDirectText(15, y, tissuedate.ToString.Format("{0:MM/dd/yyyy}", tissuedate), 20, 5, verdanaright8)

                    .RenderDirectText(40, y, tacctfrom & "-" & tsubfrom, 20, 5, verdanaleft8)
                    .RenderDirectText(80, y, tdescr, 70, 5, verdanaleft8)
                    .RenderDirectText(145, y, tamount.ToString.Format("{0:F2}", (tamount * -1)), 25, 5, verdanaright8)
                    y += 5
                    .RenderDirectText(58, y, tacctto & "-" & tsubto, 20, 5, verdanaleft8)
                    .RenderDirectText(80, y, tremarks, 70, 5, verdanaleft8)
                    .RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                    y += 2

                    If y >= 245 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Source", 25, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Destination", 25, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Description/Remarks", 40, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 10

                If y >= 250 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Number Of Transfers", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, count.ToString.Format("{0:D2}", count), 25, 5, verdanaright8bold)
                y += 6
                'draw bottom of total box
                .RenderDirectLine(59, y, 190, y, Color.Black, 0.25)
                .RenderDirectLine(59, y + 0.5, 190, y + 0.5, Color.Black, 0.25)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintTransferTicket()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2          3        4        5       6  
        '   key        fisyr      bank     srcacct   srcsub  destacct descsub
        '     7           8         9         10       11       12      13
        '  amount     applied   created     descr   remarks   srcbal srcacctname
        '    14          15        16         17 
        ' srcsubname  destbal  dstacctname destsubname
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "TransferTicket"
        Me.ReportName = "Activity Fund Transfer"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, x, y, trxkey, tfisyr As Int32
        Dim tissued, tapplied As Date
        Dim tdescr, tremarks As String
        Dim tsrcacctnum, tsrcsubacctnum, tsrcacctname, tsrcsubacctname As String
        Dim tdstacctnum, tdstsubacctnum, tdstacctname, tdstsubacctname As String
        Dim tamount, tsrcbalance, tdstbalance As Double

        Try
            With Me.GridDetail
                trxkey = CInt(Me.GridDetail.GetData(1, 0))
                tfisyr = CInt(Me.GridDetail.GetData(1, 1))
                Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 2), String)
                tsrcacctnum = DirectCast(Me.GridDetail.GetData(1, 3), String)
                tsrcsubacctnum = DirectCast(Me.GridDetail.GetData(1, 4), String)
                tdstacctnum = DirectCast(Me.GridDetail.GetData(1, 5), String)
                tdstsubacctnum = DirectCast(Me.GridDetail.GetData(1, 6), String)
                tamount = CDbl(Me.GridDetail.GetData(1, 7))
                tapplied = CDate(Me.GridDetail.GetData(1, 8))
                tissued = CDate(Me.GridDetail.GetData(1, 9))
                tdescr = DirectCast(Me.GridDetail.GetData(1, 10), String).Trim
                tremarks = DirectCast(Me.GridDetail.GetData(1, 11), String).Trim
                tsrcbalance = CDbl(Me.GridDetail.GetData(1, 12))
                tsrcacctname = DirectCast(Me.GridDetail.GetData(1, 13), String).Trim
                tsrcsubacctname = DirectCast(Me.GridDetail.GetData(1, 14), String).Trim
                tdstbalance = CDbl(Me.GridDetail.GetData(1, 15))
                tdstacctname = DirectCast(Me.GridDetail.GetData(1, 16), String).Trim
                tdstsubacctname = DirectCast(Me.GridDetail.GetData(1, 17), String).Trim
            End With
            Me.CellMiddleBottom = "FY-" + tfisyr.ToString
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                'print the info box left-side
                y = 32
                .RenderDirectText(0, y, "For Bank Account:", 40, 5, verdanaright8bold)
                .RenderDirectText(0, y + 4, Me.BankAccountNumber, 40, 5, verdanaright8)

                .RenderDirectText(0, y + 8, "For Applied Date:", 40, 5, verdanaright8bold)
                .RenderDirectText(0, y + 12, tapplied.ToShortDateString, 40, 5, verdanaright8)

                'print the info box middle
                '.RenderDirectText(65, y, "Account:", 20, 5, verdanaright8bold)
                '.RenderDirectText(85, y, tacctname, 65, 5, verdanaleft8)
                '.RenderDirectText(65, y + 4, tacctnum + "-" + tsubacctnum, 20, 5, verdanaright8)
                '.RenderDirectText(85, y + 4, tsubacctname, 65, 5, verdanaleft8)
                'print the info box left
                .RenderDirectText(145, y, "Transfer number:", 40, 5, verdanaright8bold)
                .RenderDirectText(145, y + 4, trxkey.ToString.Format("{0:D5}", trxkey), 40, 5, verdanaright8)
                y = 51
                'print line above the column headers
                .RenderDirectLine(0, y, 190, y, Color.Gray, 0.5)
                y = 58
                'print the column headers
                .RenderDirectText(20, y, "Transfer issued on:", 40, 5, verdanaright8)
                .RenderDirectText(20, y + 5, tissued.ToString.Format("{0:MM/dd/yyyy}", tissued), 40, 5, verdanaright8bold)
                .RenderDirectText(115, y, "For amount:", 40, 5, verdanaright8)
                .RenderDirectText(115, y + 5, tamount.ToString.Format("{0:C2}", tamount), 40, 5, verdanaright8bold)
                y = 80
                'left side
                .RenderDirectText(30, y, "Source account:", 30, 5, verdanaleft8bold)
                .RenderDirectText(30, y + 5, tsrcacctnum, 60, 5, verdanaleft8)
                .RenderDirectText(40, y + 5, tsrcacctname, 60, 5, verdanaleft8)
                .RenderDirectText(30, y + 10, tsrcsubacctnum, 60, 5, verdanaleft8)
                .RenderDirectText(40, y + 10, tsrcsubacctname, 60, 5, verdanaleft8)
                .RenderDirectText(30, y + 18, "Debit", 60, 5, verdanaleft8)
                .RenderDirectText(30, y + 18, tamount.ToString.Format("{0:F2}", -tamount), 60, 5, verdanaright8bold)
                .RenderDirectText(30, y + 23, "Account balance", 60, 5, verdanaleft8)
                .RenderDirectText(30, y + 23, tsrcbalance.ToString.Format("{0:C2}", tsrcbalance), 60, 5, verdanaright8bold)
                'right side
                .RenderDirectText(120, y, "Destination account:", 50, 5, verdanaleft8bold)
                .RenderDirectText(120, y + 5, tdstacctnum, 60, 5, verdanaleft8)
                .RenderDirectText(130, y + 5, tdstacctname, 60, 5, verdanaleft8)
                .RenderDirectText(120, y + 10, tdstsubacctnum, 60, 5, verdanaleft8)
                .RenderDirectText(130, y + 10, tdstsubacctname, 60, 5, verdanaleft8)
                .RenderDirectText(120, y + 18, "Credit", 60, 5, verdanaleft8)
                .RenderDirectText(120, y + 18, "+ " & tamount.ToString.Format("{0:F2}", tamount), 60, 5, verdanaright8bold)
                .RenderDirectText(120, y + 23, "Account balance", 60, 5, verdanaleft8)
                .RenderDirectText(120, y + 23, tdstbalance.ToString.Format("{0:C2}", tdstbalance), 60, 5, verdanaright8bold)
                y = 120
                .RenderDirectText(30, y, "Description", 60, 10, verdanaleft8)
                .RenderDirectText(30, y + 5, tdescr, 140, 10, verdanaleft8)
                .RenderDirectText(30, y + 18, "Remarks:", 60, 5, verdanaleft8)
                .RenderDirectText(30, y + 23, tremarks, 140, 10, verdanaleft8)
                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the preview zoom
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RenderDetailOfAccountsAllAccounts(ByVal etype As Int32, ByVal efiscalyear As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17     
        '  expcode     revcode  acctfrom  subfrom  acctto    subto    
        '    18          19        20 
        '   key       ordernum  ponum/na
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim index, currow, x, y As Int32
        Dim begyearbalance, mtdrcpt, mtdexpend, mtdadjust As Double
        Dim tacctnum, tacctname, tsubacctnum, tsubacctname As String
        Dim lbltype As String

        'etype argument determines whether this report is a ytd, mtd, or periodical
        'so that any labels will reflect the report type;
        Select Case etype
            Case 1
                lbltype = "B e g i n n i n g   y e a r l y   b a l a n c e:"
            Case 2
                lbltype = "B e g i n n i n g   m o n t h l y   b a l a n c e:"
            Case Else
                lbltype = "P e r i o d i c a l :"
        End Select
        ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
        '     0           1          2         3          4         5   
        '  bankacct   begyrbal    acctnum   acctname   subnum   subname 
        '     6           7          8         9         10        11     
        '  mtdrcpts  mtdexpend    mtdadj    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'collect the summary & descr information
        Try
            With Me.GridTotals
                Me.BankAccountNumber = DirectCast(.GetData(1, 0), String)
                begyearbalance = CDbl(.GetData(1, 1))
                tacctnum = DirectCast(.GetData(1, 2), String)
                tacctname = DirectCast(.GetData(1, 3), String)
                tsubacctnum = DirectCast(.GetData(1, 4), String)
                tsubacctname = DirectCast(.GetData(1, 5), String)
                mtdrcpt = CDbl(.GetData(1, 6))
                mtdexpend = CDbl(.GetData(1, 7))
                mtdadjust = CDbl(.GetData(1, 8))
                'if this is a periodical report, then all beginning balance is invalid;
                If etype = 3 Then
                    begyearbalance = 0
                    mtdrcpt = 0
                    mtdexpend = 0
                    mtdadjust = 0
                End If
            End With
        Catch ex As Exception

        End Try


        Try
            Dim tdocnumber, tdoctype, tstatus, tdescr, tpurchaseordernumber, tremarks As String
            Dim texpcode, trevcode, nextbank, nextacct As String
            Dim ttrxacctfrom, ttrxsubacctfrom, ttrxacctto, ttrxsubacctto As String
            Dim tprevtype As String = ""
            Dim tcreated As Date
            Dim tamount, trunningbalance, ttotalrcvd, ttotalpaid, ttotaladj, calctotal As Double
            Dim tempnumber, key, prevkey, encsw As Int32
            Dim doheader, dodetail, isfirstline As Boolean

            With Me.Doc1
                'special font for a special report
                timesleft16.Font = New Font("Arial", 8, FontStyle.Italic)
                timesleft16.TextColor = Color.Black

                x = 0
                y = 50

                'initialise the running balance to the beginning balance
                trunningbalance = begyearbalance

                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                        .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(9.5, 50, "* J & K adjustment document number", 80, 5, timesleft16)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 4, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 12, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 18, "Ending balance:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y, begyearbalance.ToString.Format("{0:F2}", begyearbalance), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 12, mtdadjust.ToString.Format("{0:F2}", mtdadjust), 30, 5, verdanaright8bold)
                        calctotal = begyearbalance + mtdrcpt + mtdadjust - mtdexpend
                        .RenderDirectText(160, y + 18, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                        '.RenderDirectLine(0, 49.5, 190, 49.5, Color.Gray, 0.5)

                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                        .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                        .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                        .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                        y = 63
                        .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                        .RenderDirectText(160, y, begyearbalance.ToString.Format("{0:F2}", begyearbalance), 30, 5, arialright8)
                        y = 68
                    End If

                    'collect the data
                    With Me.GridDetail
                        tempnumber = CInt(.GetData(index, 2))
                        tdocnumber = tempnumber.ToString.Format("{0:D8}", tempnumber)
                        tdoctype = DirectCast(.GetData(index, 3), String)
                        tstatus = DirectCast(.GetData(index, 4), String)
                        tamount = CDbl(.GetData(index, 5))
                        'tapplied = CDate(.GetData(index, 6))
                        tcreated = CDate(.GetData(index, 7))
                        tdescr = DirectCast(.GetData(index, 8), String)
                        If tdescr.Trim.Length > 25 Then tdescr = tdescr.Substring(0, 25) & "..."
                        tremarks = DirectCast(.GetData(index, 11), String)
                        If Me.UseOcas Then
                            texpcode = DirectCast(.GetData(index, 12), String)
                            trevcode = DirectCast(.GetData(index, 13), String)
                        Else
                            texpcode = ""
                            trevcode = ""
                        End If
                        key = CInt(.GetData(index, 18))
                        tpurchaseordernumber = DirectCast(.GetData(index, 20), String)
                    End With

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE HEADER 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If key <> prevkey Then doheader = True
                    If key = prevkey And tdoctype <> tprevtype Then doheader = True

                    If doheader Then
                        If currow > 1 Then y += 10
                        'issued
                        .RenderDirectText(0, y, tcreated.ToShortDateString, 20, 5, arialleft8)
                        'description
                        .RenderDirectText(16, y, tdescr, 60, 5, arialleft8)
                        'docnumber
                        If tdoctype = "0" And encsw = 0 Then
                            .RenderDirectText(73, y, tdocnumber, 45, 5, verdanaleft8bold)
                            .RenderDirectText(93, y, "PO# " & tpurchaseordernumber, 50, 5, linestyle1)
                            encsw = 1
                        Else
                            .RenderDirectText(73, y, tdocnumber, 25, 5, verdanaleft8bold)
                        End If
                        prevkey = key
                        tprevtype = tdoctype
                        doheader = False
                        dodetail = True
                        isfirstline = True
                    End If

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE DETAIL 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If isfirstline Then
                        y += 2
                        isfirstline = False
                    Else
                        y += 4
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '''''''''''''''''   added 01-20-15     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If tpurchaseordernumber <> "" And encsw = 0 Then
                        .RenderDirectText(93, y, "PO# " & tpurchaseordernumber, 50, 5, linestyle1)
                        y += 3
                    Else
                        .RenderDirectText(93, y, "            ", 50, 5, linestyle1)
                        encsw = 0
                        y += 3
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                    .RenderDirectText(16, y, tremarks, 50, 10, arialleft8)

                    Select Case tdoctype
                        Case "0"
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Check", 20, 5, timesleft16)
                            .RenderDirectText(65, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(120, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a check
                            trunningbalance -= tamount
                            'sum the checks
                            ttotalpaid += tamount
                            dodetail = False
                        Case "1"
                            trevcode = Module1.FormatRevenueCode(trevcode)
                            If dodetail Then .RenderDirectText(0, y, "Receipt", 20, 5, timesleft16)
                            .RenderDirectText(65, y, trevcode, 40, 5, arialleft8)
                            .RenderDirectText(100, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a receipt
                            trunningbalance += tamount
                            ttotalrcvd += tamount
                            dodetail = False
                        Case "2"
                            Select Case tstatus
                                Case "B"
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the bank into the total adjustments
                                    ttotaladj -= tamount
                                Case "E"
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an expenditure adj
                                    trunningbalance += tamount
                                    'sum the expenditures into the total adjustments
                                    ttotaladj += tamount
                                Case "I", "N", "R"
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    .RenderDirectText(65, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotaladj += tamount
                                Case "J"    'legacy receipts
                                    .RenderDirectText(0, y, "J Adjust *", 25, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(65, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(100, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotalrcvd += tamount
                                Case "K"    'legacy checks
                                    .RenderDirectText(0, y, "K Adjust *", 25, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(65, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(120, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the checks
                                    ttotalpaid += tamount
                            End Select
                        Case "3"    'transfer from
                            .RenderDirectText(0, y, "Trx From", 20, 5, timesleft16)
                            .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance -= tamount
                            ttotaladj -= tamount
                        Case "4"    'transfer to
                            .RenderDirectText(0, y, "Trx To", 20, 5, timesleft16)
                            .RenderDirectText(140, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance += tamount
                            ttotaladj += tamount
                        Case Else
                            .RenderDirectText(0, y, "Undefined", 20, 5, timesleft16)
                    End Select

                    'print the running balance
                    .RenderDirectText(165, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 25, 5, arialright8)

                    'page break if near end-of-page
                    If y >= 243 Then
                        'page break if not the last record
                        If index < (Me.GridDetail.Rows.Count - 1) Then
                            .NewPage()
                            currow = 0
                            .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                            .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                            .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                            .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                            .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                            .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                            'print the column headers
                            y = 58
                            .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                            .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                            .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                            .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                            .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                            .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                            .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                            .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                            'print the continued balance and line
                            y = 63
                            lbltype = "C o n t i n u e d   f r o m   p r e v i o u s   p a g e ..."
                            .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                            .RenderDirectText(160, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)
                            y = 68
                        End If
                    End If
                Next

                y += 15
                If y >= 243 Then
                    .NewPage()
                    y = 70
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 1, 190, y - 1, Color.Black, 0.25)
                .RenderDirectLine(59, y - 0.5, 190, y - 0.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Totals", 25, 5, verdanaleft8bold)
                .RenderDirectText(75, y, "Beginning", 25, 5, verdanaright8bold)
                .RenderDirectText(100, y, "Received", 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, "Paid Out", 25, 5, verdanaright8bold)
                .RenderDirectText(140, y, "Adjusted", 25, 5, verdanaright8bold)
                .RenderDirectText(165, y, "Balance", 25, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(75, y, begyearbalance.ToString.Format("{0:F2}", begyearbalance), 25, 5, verdanaright8bold)
                .RenderDirectText(100, y, ttotalrcvd.ToString.Format("{0:F2}", ttotalrcvd), 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 25, 5, verdanaright8bold)
                .RenderDirectText(140, y, ttotaladj.ToString.Format("{0:F2}", ttotaladj), 25, 5, verdanaright8bold)
                'calc the total
                calctotal = begyearbalance + ttotalrcvd + ttotaladj - ttotalpaid
                .RenderDirectText(165, y, calctotal.ToString.Format("{0:F2}", calctotal), 25, 5, verdanaright8bold)
                'draw bottom of total box
                .RenderDirectLine(59, y + 5, 190, y + 5, Color.Black, 0.25)
                .RenderDirectLine(59, y + 5.5, 190, y + 5.5, Color.Black, 0.25)
            End With
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RenderEncumbranceDetailOfAccountsAllAccounts(ByVal etype As Int32, ByVal efiscalyear As Int32)
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank        fisyr   docnumber doctype  status    amount
        '     6           7         8        9       10        11  
        '  applied     created    descr    acct    subacct   remarks
        '    12          13        14       15       16        17      18   
        '  expcode     revcode  acctfrom  subfrom  acctto    subto     key
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'define the styles 
        Dim subscript As C1DocStyle = New C1DocStyle(Me.Doc1)
        With subscript
            .Borders.AllEmpty = True
            .BorderTableHorz.Empty = True
            .BorderTableVert.Empty = True
            .Font = New Font("Verdana", 7, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
        End With

        Dim index, currow As Int32
        Dim x, y As Double
        Dim begbalance, mtdrcpt, mtdencumbrance, mtdexpend, mtdadjust As Decimal
        Dim tacctnum, tacctname, tbankaccount, tsubacctnum, tsubacctname As String
        Dim lbltype As String

        'etype argument determines whether this report is a ytd, mtd, or periodical
        'so that any labels will reflect the report type;
        Select Case etype
            Case 1
                lbltype = "B e g i n n i n g   y e a r l y   b a l a n c e:"
            Case 2
                lbltype = "B e g i n n i n g   m o n t h l y   b a l a n c e:"
            Case Else
                lbltype = "P e r i o d i c a l :"
        End Select


        Try
            'collect the summary & descr information
            With Me.GridTotals
                ''''''''''''''''''''' GRIDTOTAL ''''''''''''''''''''''''''''''''''
                '     0           1          2         3          4         5   
                '  bankacct     begbal    acctnum   acctname   subnum   subname 
                '     6           7          8         9         10        11     
                '  mtdrcpts  mtdexpend    mtdadj    mtdencnum
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                tbankaccount = CType(.GetData(1, 0), String)
                'for display only (the property may mask the actual bank account number);
                Me.BankAccountNumber = tbankaccount
                begbalance = CType(.GetData(1, 1), Decimal)
                tacctnum = CType(.GetData(1, 2), String)
                tacctname = CType(.GetData(1, 3), String)
                tsubacctnum = CType(.GetData(1, 4), String)
                tsubacctname = CType(.GetData(1, 5), String)
                mtdrcpt = CType(.GetData(1, 6), Decimal)
                mtdexpend = CType(.GetData(1, 7), Decimal)
                mtdadjust = CType(.GetData(1, 8), Decimal)
                mtdencumbrance = CType(.GetData(1, 9), Decimal)
                'if this is a periodical report, then all beginning balance is invalid;
                If etype = 3 Then
                    begbalance = 0
                    mtdrcpt = 0
                    mtdexpend = 0
                    mtdadjust = 0
                    mtdencumbrance = 0
                End If
            End With
        Catch ex As Exception

        End Try


        Try
            Dim tdocnumber, tdoctype, tstatus, tdescr, tremarks As String
            Dim texpcode, trevcode, nextbank, nextacct As String
            Dim ttrxacctfrom, ttrxsubacctfrom, ttrxacctto, ttrxsubacctto As String
            Dim tprevtype As String = ""
            Dim tcreated As Date
            Dim tamount, trunningbalance, ttotaladj, ttotalrcvd, ttotalpaid, ttotalvoid, calctotal, pcalctotal As Decimal
            Dim tendingbalance, ttotalencumbered, tencumberedoutstanding, tencumberedpaid, tunpaid As Decimal
            Dim tempnumber, key, prevkey As Int32
            Dim doheader, dodetail, isfirstline As Boolean

            With Me.Doc1
                'special font for a special report
                timesleft16.Font = New Font("Arial", 8, FontStyle.Italic)
                timesleft16.TextColor = Color.Black

                x = 0
                y = 50

                'if this is a periodical report, then the beginning balance is invalid;
                If etype = 3 Then begbalance = 0
                'initialise the running balance to the beginning balance
                trunningbalance = begbalance

                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, tbankaccount, 40, 5, verdanaright8)
                        .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                        .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                        .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                        'print a legend for the report;
                        .RenderDirectText(0.5, 50, "* J & K adjustment document number", 80, 5, subscript)
                        'print the info box right-side
                        x = 185
                        y = 32
                        .RenderDirectText(x, y, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 8, "Encumbrances:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 12, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 16, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 22, "Ending balance:", 40, 5, verdanaright8bold)
                        x = 225
                        'print the money fields
                        .RenderDirectText(x, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 8, mtdencumbrance.ToString.Format("{0:F2}", mtdencumbrance), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 12, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 30, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 16, mtdadjust.ToString.Format("{0:F2}", mtdadjust), 30, 5, verdanaright8bold)
                        calctotal = begbalance + mtdrcpt + mtdadjust - mtdexpend
                        .RenderDirectText(x, y + 22, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                        y = 60
                        'print line above the column headers
                        .RenderDirectLine(0, y, 255, y, Color.Gray, 0.5)
                        x = 120
                        y = 62
                        'print the column headers
                        .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                        .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                        .RenderDirectText(x, y, "Received", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 30, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 50, y, "Paid Out", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 75, y, "Adjusted", 25, 5, verdanaright8bold)
                        .RenderDirectText(x + 105, y, "Balance", 25, 5, verdanaright8bold)
                        y = 68
                        .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                        .RenderDirectText(220, y, begbalance.ToString.Format("{0:F2}", begbalance), 30, 5, arialright8)
                        y = 74
                    End If

                    'collect the data
                    With Me.GridDetail
                        tempnumber = CInt(.GetData(index, 2))
                        tdocnumber = tempnumber.ToString.Format("{0:D8}", tempnumber)
                        tdoctype = CType(.GetData(index, 3), String)
                        tstatus = CType(.GetData(index, 4), String)
                        tamount = CType(.GetData(index, 5), Decimal)
                        'tapplied = CDate(.GetData(index, 6))
                        tcreated = CDate(.GetData(index, 7))
                        tdescr = CType(.GetData(index, 8), String)
                        If tdescr.Trim.Length > 25 Then tdescr = tdescr.Substring(0, 25) & "..."
                        tremarks = CType(.GetData(index, 11), String)
                        If tremarks.Trim.Length > 30 Then tremarks = tremarks.Substring(0, 30) & "..."
                        texpcode = CType(.GetData(index, 12), String)
                        trevcode = CType(.GetData(index, 13), String)
                        key = CInt(.GetData(index, 18))
                    End With

                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE HEADER 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If key <> prevkey Then doheader = True
                    If key = prevkey And tdoctype <> tprevtype Then doheader = True

                    If doheader Then
                        If currow > 1 Then y += 8
                        'issued
                        .RenderDirectText(0, y, tcreated.ToShortDateString, 20, 5, arialleft8)
                        'description
                        .RenderDirectText(16, y, tdescr, 60, 5, arialleft8)
                        'docnumber
                        .RenderDirectText(75, y, tdocnumber, 25, 5, verdanaleft8bold)
                        prevkey = key
                        tprevtype = tdoctype
                        doheader = False
                        dodetail = True
                        isfirstline = True
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ''''' PRINT THE DETAIL 
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If isfirstline Then
                        y += 4
                        isfirstline = False
                    Else
                        y += 4
                    End If

                    .RenderDirectText(16, y, tremarks, 55, 5, arialleft8)

                    Select Case tdoctype
                        Case "0"
                            x = 170
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Check", 20, 5, timesleft16)
                            .RenderDirectText(70, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a check
                            trunningbalance -= tamount
                            'sum the checks
                            ttotalpaid += tamount
                            dodetail = False
                        Case "1"
                            x = 120
                            trevcode = Module1.FormatRevenueCode(trevcode)
                            If dodetail Then .RenderDirectText(0, y, "Receipt", 20, 5, timesleft16)
                            .RenderDirectText(70, y, trevcode, 40, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'calc the running balance for a receipt
                            trunningbalance += tamount
                            ttotalrcvd += tamount
                            dodetail = False
                        Case "2"
                            x = 195
                            Select Case tstatus
                                Case "B"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the bank into the total adjustments
                                    ttotaladj -= tamount
                                Case "E"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an expenditure adj
                                    trunningbalance += tamount
                                    'sum the expenditures into the total adjustments
                                    ttotaladj += tamount
                                Case "I", "N", "R"
                                    .RenderDirectText(0, y, "Adjust", 20, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(70, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotaladj += tamount
                                Case "J"    'legacy receipts
                                    x = 120
                                    .RenderDirectText(0, y, "J Adjust", 25, 5, timesleft16)
                                    trevcode = Module1.FormatRevenueCode(trevcode)
                                    .RenderDirectText(-2.5, y + 0.5, "*", 5, 5, subscript)
                                    .RenderDirectText(70, y, trevcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for a revenue adjustment
                                    trunningbalance += tamount
                                    'sum the revenues into the total adjustments
                                    ttotalrcvd += tamount
                                Case "K"    'legacy checks
                                    x = 170
                                    .RenderDirectText(0, y, "K Adjust", 25, 5, timesleft16)
                                    texpcode = Module1.FormatExpenditureCode(texpcode)
                                    .RenderDirectText(-2.5, y + 0.5, "*", 5, 5, subscript)
                                    .RenderDirectText(70, y, texpcode, 100, 5, arialleft8)
                                    .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                                    'calc the running balance for an bank adj
                                    trunningbalance -= tamount
                                    'sum the checks
                                    ttotalpaid += tamount
                            End Select
                        Case "3"    'transfer from
                            x = 195
                            .RenderDirectText(0, y, "Trx From", 20, 5, timesleft16)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance -= tamount
                            ttotaladj -= tamount
                        Case "4"    'transfer to
                            x = 195
                            .RenderDirectText(0, y, "Trx To", 20, 5, timesleft16)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            trunningbalance += tamount
                            ttotaladj += tamount
                        Case "5"
                            x = 150
                            texpcode = Module1.FormatExpenditureCode(texpcode)
                            If dodetail Then .RenderDirectText(0, y, "Encum.", 20, 5, timesleft16)
                            .RenderDirectText(70, y, texpcode, 60, 5, arialleft8)
                            .RenderDirectText(x, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, arialright8)
                            'ttotalencumbered, tencumberedoutstanding, tencumberedpaid
                            Select Case tstatus
                                Case "O"
                                    '.RenderDirectText(-2.5, y, "1", 5, 5, subscript)
                                    tencumberedoutstanding += tamount
                                    ttotalencumbered += tamount
                                Case "C"
                                    tencumberedpaid += tamount
                                    ttotalencumbered += tamount
                                    '.RenderDirectText(-2.5, y, "2", 5, 5, subscript)
                                Case Else
                                    ttotalencumbered += tamount
                                    .RenderDirectText(-2.5, y + 0.5, "x", 5, 5, subscript)
                            End Select
                            dodetail = False
                        Case Else
                            .RenderDirectText(0, y, "Undefined", 20, 5, timesleft16)
                    End Select

                    'print the running balance
                    .RenderDirectText(220, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)

                    If y >= 185 Then
                        'page break if not the last record
                        If index < (Me.GridDetail.Rows.Count - 1) Then
                            .NewPage()
                            currow = 0
                            .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                            .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                            .RenderDirectText(45, 36, "Account:", 20, 5, verdanaleft8bold)
                            .RenderDirectText(45, 40, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                            .RenderDirectText(65, 36, tacctname, 60, 5, verdanaleft8)
                            .RenderDirectText(65, 40, tsubacctname, 60, 5, verdanaleft8)
                            'print a legend for the report;
                            .RenderDirectText(0.5, 50, "* J & K adjustment document number", 80, 5, subscript)
                            y = 60
                            'print line above the column headers
                            .RenderDirectLine(0, y, 255, y, Color.Gray, 0.5)
                            x = 120
                            y = 62
                            'print the column headers;
                            .RenderDirectText(0, y, "Issued", 15, 5, verdanaleft8bold)
                            .RenderDirectText(16, y, "Description", 50, 5, verdanaleft8bold)
                            .RenderDirectText(66, y, "Number", 25, 5, verdanaright8bold)
                            .RenderDirectText(x, y, "Received", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 30, y, "Encumbered", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 53, y, "Paid Out", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 75, y, "Adjusted", 25, 5, verdanaright8bold)
                            .RenderDirectText(x + 105, y, "Balance", 25, 5, verdanaright8bold)
                            'print the continued balance and line;
                            y = 68
                            lbltype = "C o n t i n u e d   f r o m   p r e v i o u s   p a g e ..."
                            .RenderDirectText(30, y, lbltype, 80, 5, verdanaleft8)
                            .RenderDirectText(220, y, trunningbalance.ToString.Format("{0:F2}", trunningbalance), 30, 5, arialright8)
                            y = 74
                        End If
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'end of detail;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                Dim SSQL As String
                Dim cmd As SqlCommand

                Try
                    'collect the total void amount;
                    SSQL = "SELECT ISNULL(SUM(invc_amount), 0.0) FROM invoices" _
                    & " WHERE bank_acct_num = @p1 AND invc_fisyr = @p2 AND invc_status = 'V'" _
                    & " AND af_acct_num = @p3 AND as_acct_num = @p4"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", tbankaccount)
                    cmd.Parameters.Add("@p2", efiscalyear)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubacctnum)
                    cn.Open()
                    ttotalvoid = CType(cmd.ExecuteScalar, Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                End Try

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'page break for totals, which will be printed on a page by themselves;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                .NewPage()
                y = 60

                'draw top of total box;
                .RenderDirectLine(60, y - 0.5, 250, y - 0.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Totals:", 25, 5, verdanaleft8bold)
                .RenderDirectText(75, y, "Beginning", 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, "Received", 25, 5, verdanaright8bold)
                .RenderDirectText(150, y, "Encumbered", 25, 5, verdanaright8bold)
                .RenderDirectText(173, y, "Paid Out", 25, 5, verdanaright8bold)
                .RenderDirectText(195, y, "Adjusted", 25, 5, verdanaright8bold)
                .RenderDirectText(220, y, "Ending", 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(75, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, ttotalrcvd.ToString.Format("{0:F2}", ttotalrcvd), 25, 5, verdanaright8bold)
                .RenderDirectText(150, y, ttotalencumbered.ToString.Format("{0:F2}", ttotalencumbered), 25, 5, verdanaright8bold)
                .RenderDirectText(173, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 25, 5, verdanaright8bold)
                .RenderDirectText(195, y, ttotaladj.ToString.Format("{0:F2}", ttotaladj), 25, 5, verdanaright8bold)
                tendingbalance = begbalance + ttotalrcvd + ttotaladj - ttotalpaid
                .RenderDirectText(220, y, tendingbalance.ToString.Format("{0:F2}", tendingbalance), 30, 5, verdanaright8bold)

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'print the totals on the last page;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                y = 80
                .RenderDirectText(160, y, "Total Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalencumbered.ToString.Format("{0:F2}", ttotalencumbered), 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(160, y, "Less Paid Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalpaid.ToString.Format("{0:F2}", ttotalpaid), 30, 5, verdanaright8bold)
                verdanaright8bold.TextColor = Color.Black
                y += 5
                'calculate unpaid encumbrances;
                tunpaid = ttotalencumbered - ttotalpaid
                .RenderDirectText(160, y, "Unpaid Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tunpaid.ToString.Format("{0:F2}", tunpaid), 30, 5, verdanaright8bold)
                y += 5
                'total voids from query above;
                .RenderDirectText(160, y, "Add Total Voids:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, ttotalvoid.ToString.Format("{0:F2}", ttotalvoid), 30, 5, verdanaright8bold)
                y += 5
                'calculate outstanding encumbrance;
                tencumberedoutstanding = tunpaid + ttotalvoid
                .RenderDirectText(160, y, "Total Outstanding Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tencumberedoutstanding.ToString.Format("{0:F2}", tencumberedoutstanding), 30, 5, verdanaright8bold)
                'extra line break;
                y += 10
                .RenderDirectText(160, y, "Ending Balance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, calctotal.ToString.Format("{0:F2}", calctotal), 30, 5, verdanaright8bold)
                y += 5
                .RenderDirectText(160, y, "Less Outstanding Encumbrance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, tencumberedoutstanding.ToString.Format("{0:F2}", tencumberedoutstanding), 30, 5, verdanaright8bold)
                y += 5
                'calculate projected balance;
                pcalctotal = tendingbalance - tencumberedoutstanding
                .RenderDirectText(160, y, "Projected Balance:", 60, 5, verdanaright8bold)
                .RenderDirectText(220, y, pcalctotal.ToString.Format("{0:F2}", pcalctotal), 30, 5, verdanaright8bold)
            End With
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub PrintFooter()
        Try
            'print footer (portrait)
            With Me.Doc1
                If Me.DocumentName = "EncumbranceDetailOfAccountsSingleAccount" Then
                    .RenderDirectLine(0, 205, 255, 205, Color.Black, 0.5)
                    .RenderDirectText(0, 206, "Page [@@PageNo@@] of [@@PageCount@@]", 150, 4, footerstyle)
                    .RenderDirectText(100, 206, "Activity Fund.Net © - A product of ADPC", 75, 4, footerstyle)
                    '.RenderDirectText(231, 206, "1(800)747-2372", 25, 4, footerstyle)
                    'changed by fred 2008.04.10;
                    .RenderDirectText(216, 206, Now.ToString.Format("{0:MM/dd/yyyy HH:mm:ss tt}", Now), 80, 4, footerstyle)
                Else
                    .RenderDirectLine(0, 264, 191.5, 264, Color.Black, 0.5)
                    .RenderDirectText(0, 265, "Page [@@PageNo@@] of [@@PageCount@@]", 150, 4, footerstyle)
                    .RenderDirectText(68, 265, "Activity Fund.Net © - A product of ADPC", 75, 4, footerstyle)
                    '.RenderDirectText(167, 265, "1(800)747-2372", 80, 4, footerstyle)
                    'changed by fred 2008.04.10;
                    .RenderDirectText(152, 265, Now.ToString.Format("{0:MM/dd/yyyy HH:mm:ss tt}", Now), 80, 4, footerstyle)
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PrintHeader()
        Try
            With Me.Doc1
                If Me.DocumentName = "EncumbranceDetailOfAccountsSingleAccount" Then
                    ''''''''''' landscape;
                    'print the top margin line 
                    .RenderDirectLine(0, 14, 255, 14, Color.Black, 0.5)
                    'print the left side of the header
                    .RenderDirectText(2, 15, Me.SchoolName, 80, 5, arialleft10)
                    .RenderDirectText(2, 20, Me.SchoolAddress1, 80, 5, arialleft10)
                    .RenderDirectText(2, 25, Me.SchoolCityStateZip, 80, 5, arialleft10)
                    'print the center of the header
                    verdanaleft8.TextAlignHorz = AlignHorzEnum.Center
                    .RenderDirectText(95, 15, Me.CellMiddleTop, 80, 5, verdanaleft8)
                    .RenderDirectText(95, 20, Me.CellMiddleMiddle, 80, 5, verdanaleft8)
                    .RenderDirectText(95, 25, Me.CellMiddleBottom, 80, 5, verdanaleft8)
                    verdanaleft8.TextAlignHorz = AlignHorzEnum.Left
                    'print the right side of the header
                    .RenderDirectText(185, 15, Me.ReportName, 70, 5, verdanaright10bold)
                    .RenderDirectText(185, 20, Me.CellRightMiddle, 70, 5, verdanaright8)
                    .RenderDirectText(185, 25, Now.ToString.Format("{0:MMMM dd, yyyy}", Now), 70, 5, verdanaright8)
                    'print the bottom header line
                    .RenderDirectLine(0, 31, 255, 31, Color.Gray, 0.5)
                Else
                    ''''''''''' portrait;
                    'print the top margin line 
                    .RenderDirectLine(0, 14, 190, 14, Color.Black, 0.5)
                    'print the left side of the header
                    .RenderDirectText(2, 15, Me.SchoolName, 80, 5, arialleft10)
                    .RenderDirectText(2, 20, Me.SchoolAddress1, 80, 5, arialleft10)
                    .RenderDirectText(2, 25, Me.SchoolCityStateZip, 80, 5, arialleft10)
                    'print the center of the header
                    verdanaleft8.TextAlignHorz = AlignHorzEnum.Center
                    .RenderDirectText(85, 15, Me.CellMiddleTop, 80, 5, verdanaleft8)
                    .RenderDirectText(85, 20, Me.CellMiddleMiddle, 50, 5, verdanaleft8)
                    .RenderDirectText(85, 25, Me.CellMiddleBottom, 50, 5, verdanaleft8)
                    verdanaleft8.TextAlignHorz = AlignHorzEnum.Left
                    'print the right side of the header
                    .RenderDirectText(120, 15, Me.ReportName, 70, 5, verdanaright10bold)
                    .RenderDirectText(150, 20, Me.CellRightMiddle, 40, 5, verdanaright8)
                    .RenderDirectText(150, 25, Now.ToString.Format("{0:MMMM dd, yyyy}", Now), 40, 5, verdanaright8)
                    'print the bottom header line
                    .RenderDirectLine(0, 31, 190, 31, Color.Gray, 0.5)
                End If
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            PrintFooter()
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "  Properties "

    Private Property BankAccountNumber() As String
        Get
            Return _bankaccountnumber
        End Get
        Set(ByVal Value As String)
            Dim index, length As Int32
            Dim s1, s2 As String
            Dim t1, t2 As Int32
            'mask all but the last 4 characters of the bankaccount number;
            t1 = Value.Trim.Length
            If t1 > 4 Then
                t2 = t1 - 4
                For index = 1 To t2
                    s1 &= "* "
                Next
                s2 = s1 & Value.Substring(t1 - 4, 4)
            Else
                s2 = Value
            End If
            _bankaccountnumber = s2
        End Set
    End Property

    Private Property ConnectionString() As String
        Get
            Return _connectionstring
        End Get
        Set(ByVal Value As String)
            _connectionstring = Value.Trim
        End Set
    End Property

    Private Property CountyId() As String
        Get
            Return _countyid
        End Get
        Set(ByVal Value As String)
            If Value.Trim.Length <> 2 Then Throw New ArgumentException("County Id invalid or missing...")
            _countyid = Value.Trim
        End Set
    End Property

    Private Property DistrictId() As String
        Get
            Return _districtid
        End Get
        Set(ByVal Value As String)
            If Value.Trim.Length <> 4 Then Throw New ArgumentException("District Id invalid or missing...")
            _districtid = Value.Trim
        End Set
    End Property

    Private Property DocumentName() As String
        Get
            Return _documentname
        End Get
        Set(ByVal Value As String)
            _documentname = Value.Trim
        End Set
    End Property

    Private Property FiscalMonthStr() As String
        Get
            Return _fiscalmonthstr
        End Get
        Set(ByVal Value As String)
            _fiscalmonthstr = Value.Trim
        End Set
    End Property

    Private Property FiscalYear() As Int32
        Get
            Return _fiscalyear
        End Get
        Set(ByVal Value As Int32)
            _fiscalyear = Value
        End Set
    End Property

    Private Property ReportName() As String
        Get
            Return _reportname
        End Get
        Set(ByVal Value As String)
            _reportname = Value.Trim
        End Set
    End Property

    Private Property SchoolName() As String
        Get
            Return _schoolname
        End Get
        Set(ByVal Value As String)
            _schoolname = Value.Trim
        End Set
    End Property

    Private Property SchoolAddress1() As String
        Get
            Return _schooladdress1

        End Get
        Set(ByVal Value As String)
            _schooladdress1 = Value.Trim
        End Set
    End Property

    Private Property SchoolAddress2() As String
        Get
            Return _schooladdress2
        End Get
        Set(ByVal Value As String)
            _schooladdress2 = Value.Trim
        End Set
    End Property

    Private Property SchoolCityStateZip() As String
        Get
            Return _schoolcitystatezip
        End Get
        Set(ByVal Value As String)
            _schoolcitystatezip = Value.Trim
        End Set
    End Property

    Private Property UseOcas() As Boolean
        Get
            Return _useocas
        End Get
        Set(ByVal Value As Boolean)
            _useocas = Value
        End Set
    End Property

    Private Property UserName() As String
        Get
            Return _username
        End Get
        Set(ByVal Value As String)
            _username = Value.Trim
        End Set
    End Property

#End Region

    Private Sub Prev1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Prev1.Load

    End Sub

    Private Sub Prev1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Prev1.Click
        Debug.WriteLine("Prev1_Click")
    End Sub

    Private Sub Prev1_ButtonClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles Prev1.ButtonClick
        Debug.WriteLine("Prev1_ButtonClick: " + e.Button.ToString())




    End Sub

    Private Sub Prev1_PreviewButtonClick(ByVal sender As Object, ByVal e As C1.Win.C1PrintPreview.PreviewToolBarButtonClickEventArgs) Handles Prev1.PreviewButtonClick
        Debug.WriteLine("Prev1_PreviewButtonClick: " + e.Button.ToString)
    End Sub
End Class
