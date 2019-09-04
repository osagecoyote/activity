Imports C1.C1PrintDocument
Imports C1.Win.C1FlexGrid
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Public Class frmExpenditureReports
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
            Me.CurrentMonthStr = authobj.CurrentMonthString
            Me.SchoolName = authobj.SchoolName
            Me.SchoolNumber = authobj.SchoolNumber
            Me.SchoolAddress1 = authobj.SchoolAddress1
            Me.SchoolAddress2 = authobj.SchoolAddress2
            Me.SchoolTelephone = authobj.SchoolTelephone1
            Me.UseOcas = authobj.UseOCAS
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
    Friend WithEvents GridDetail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridTotals As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Doc1 As C1.C1PrintDocument.C1PrintDocument
    Friend WithEvents GridWrk As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridWrkTotals As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmExpenditureReports))
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
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
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
        Me.Prev1.Location = New System.Drawing.Point(0, 8)
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
        Me.Prev1.Size = New System.Drawing.Size(656, 360)
        Me.Prev1.Splitter.Cursor = System.Windows.Forms.Cursors.VSplit
        Me.Prev1.Splitter.Width = 3
        Me.Prev1.StatusBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.StatusBar.TabIndex = 4
        Me.Prev1.TabIndex = 2
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
        Me.GridDetail.Location = New System.Drawing.Point(0, 8)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.Size = New System.Drawing.Size(656, 336)
        Me.GridDetail.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridDetail.TabIndex = 3
        Me.GridDetail.Visible = False
        '
        'GridTotals
        '
        Me.GridTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridTotals.ForeColor = System.Drawing.SystemColors.WindowText
        Me.GridTotals.Location = New System.Drawing.Point(0, 70)
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
        Me.GridTotals.TabIndex = 4
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
        Me.GridWrk.Location = New System.Drawing.Point(0, 18)
        Me.GridWrk.Name = "GridWrk"
        Me.GridWrk.Size = New System.Drawing.Size(656, 336)
        Me.GridWrk.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrk.TabIndex = 5
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
        Me.GridWrkTotals.Location = New System.Drawing.Point(8, 18)
        Me.GridWrkTotals.Name = "GridWrkTotals"
        Me.GridWrkTotals.Size = New System.Drawing.Size(656, 330)
        Me.GridWrkTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Normal{Font:Arial, 8.25pt;}" & Microsoft.VisualBasic.ChrW(9) & "Fixed{BackColor:Control;ForeColor:ControlText;Border:" & _
        "Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Highlight{BackColor:Highlight;ForeColor:HighlightText;" & _
        "}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & _
        "EmptyArea{BackColor:AppWorkspace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal" & _
        "{BackColor:Black;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackC" & _
        "olor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrkTotals.TabIndex = 6
        Me.GridWrkTotals.Visible = False
        '
        'frmExpenditureReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 373)
        Me.Controls.Add(Me.Prev1)
        Me.Controls.Add(Me.GridWrk)
        Me.Controls.Add(Me.GridDetail)
        Me.Controls.Add(Me.GridWrkTotals)
        Me.Controls.Add(Me.GridTotals)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmExpenditureReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Expenditure Reporting"
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
            Case "PurchaseOrderTicket", "RequisitionTicket"
                'print footer only
                PrintFooter()
            Case Else
                'print header includes the footer
                PrintHeader()
        End Select
    End Sub

#End Region

#Region "  Class Members "

    'styles;
    Private arialleft8 As C1DocStyle
    Private arialright8 As C1DocStyle
    Private arialleft10 As C1DocStyle
    Private arialleft10bold As C1DocStyle
    Private arialright10 As C1DocStyle
    Private arialright10bold As C1DocStyle
    Private docstyle As C1DocStyle
    Private footerstyle As C1DocStyle
    Private specstyle As C1DocStyle
    Private timesleft16 As C1DocStyle
    Private verdanaleft8 As C1DocStyle
    Private verdanaright8 As C1DocStyle
    Private verdanaleft8bold As C1DocStyle
    Private verdanaright8bold As C1DocStyle
    Private verdanaleft10 As C1DocStyle
    Private verdanaright10 As C1DocStyle
    Private verdanaleft10bold As C1DocStyle
    Private verdanaright10bold As C1DocStyle

    'header values;
    Private CellMiddleBottom As String = ""
    Private CellMiddleMiddle As String = ""
    Private CellMiddleTop As String = ""
    Private CellRightBottom As String = ""
    Private CellRightMiddle As String = ""
    Private CellRightTop As String = ""

    'Property Vars;
    Private p_bankaccountnumber As String = ""
    '
    Private CountyId As String = ""
    Private CurrentMonthStr As String = ""
    Private DistrictId As String = ""
    Private DocumentName As String = ""
    Private DoSignatures As Boolean
    Private FiscalYear As Int32
    Private ReportName As String = ""
    Private SchoolName As String = ""
    Private SchoolNumber As String = ""
    Private SchoolAddress1 As String = ""
    Private SchoolAddress2 As String = ""
    Private SchoolCityStateZip As String = ""
    Private SchoolTelephone As String = ""
    Private ShippingName As String = ""
    Private ShippingAddress1 As String = ""
    Private ShippingAddress2 As String = ""
    Private ShippingAddress3 As String = ""
    Private ShippingCity As String = ""
    Private ShippingState As String = ""
    Private ShippingZip As String = ""
    Private SignatureTextLine1 As String = ""
    Private SignatureTextLine2 As String = ""
    Private Signature1 As Image
    Private Signature2 As Image
    Private UseOcas As Boolean
    '
    Private Const MSGTITLE As String = "Activity Fund Reports"
    Private cn As SqlConnection
    Private ConnectionString As String

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
                Case Else
                    .DefaultUnit = UnitTypeEnum.Mm
                    .DefaultUnitOfFrames = UnitTypeEnum.Mm
                    .PageSettings.Landscape = False
                    .PageSettings.Margins.Left = 50
                    .PageSettings.Margins.Top = 10
                    .PageSettings.Margins.Right = 50
                    .PageSettings.Margins.Bottom = 20
                    .PageFooter.Height = 10
                    '.PageFooter.RenderText.Style.Font = New Font("Verdana", 8, FontStyle.Regular)
                    '.PageFooter.RenderText.Text = "Page [@@PageNo@@] of [@@PageCount@@]"
            End Select
        End With

        Call DefineFooterStyle()

    End Sub

    Private Sub DefineFooterStyle()
        'style for the footer
        footerstyle = New C1DocStyle(Me.Doc1)
        With footerstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .TextColor = Color.Gray
        End With
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

#Region "  Methods Helper "

    Private Function FormatOcasCodeX(ByVal ecode As String) As String
        Dim dim1, dim2, dim3, dim4, dim5, dim6, dim7, dim8, dim9 As String
        Dim formattedcode As String
        Try
            'determine the type of code by the length
            Select Case ecode.Length
                Case 16     'revenue code
                    dim1 = ecode.Substring(0, 1)
                    dim2 = ecode.Substring(1, 2)
                    dim3 = ecode.Substring(3, 3)
                    dim4 = ecode.Substring(6, 4)
                    dim5 = ecode.Substring(10, 3)
                    dim6 = ecode.Substring(13, 3)
                    formattedcode = dim1 + "-" + dim2 + "-" + dim3 + "-" + dim4 + "-" + dim5 + "-" + dim6
                Case 26     'expenditure code
                    dim1 = ecode.Substring(0, 1)
                    dim2 = ecode.Substring(1, 2)
                    dim3 = ecode.Substring(3, 3)
                    dim4 = ecode.Substring(6, 4)
                    dim5 = ecode.Substring(10, 3)
                    dim6 = ecode.Substring(13, 3)
                    dim7 = ecode.Substring(16, 4)
                    dim8 = ecode.Substring(20, 3)
                    dim9 = ecode.Substring(23, 3)
                    formattedcode = dim1 + "-" + dim2 + "-" + dim3 + "-" + dim4 + "-" + dim5 + "-" + dim6 + "-" + dim7 + "-" + dim8 + "-" + dim9
                Case Else   'some weird code
            End Select
            Return formattedcode
        Catch ex As Exception
            'in case there's a weird error then return a weird code
            Return "X-XX-XXX-XXXX-XXX-XXX"
        Finally

        End Try
    End Function

#End Region

#Region "  Methods Generation "

    Public Function GenerateChecksOutstanding(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eallfiscalyears As Boolean, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all outstanding checks for a single bank for all fiscal years or a
        'selected fiscal year;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   printed
        '     6           7         8        9       10       11  
        '  payee      ponumber   summary   amount   acct     sub 
        '    12          13        14       15
        '  applied    created     descr  remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet

        If eallfiscalyears Then
            Me.CellMiddleBottom = "All Fiscal Years"
            SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
            & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
            & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
            & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
            & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND h.bank_acct_num = @p1" _
            & " AND (h.chks_recon_sw = 'N')" _
            & " AND (h.chks_status <> 'V')" _
            & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outc_chk_fisyr, outc_chk_num, 'X' AS status," _
            & " outc_recon_sw, '' AS printed, outc_chk_payee_name, '' AS ponum," _
            & " outc_chk_amount, outc_chk_amount, af_acct_num, as_acct_num," _
            & " outc_chk_issue_date, outc_chk_issue_date, outc_chk_descr, '' AS remarks" _
            & " FROM outstandingchecks" _
            & " WHERE bank_acct_num = @p1" _
            & " AND (outc_stale_sw = 'N')" _
            & " AND (outc_recon_sw = 'N')" _
            & " ORDER BY bank_acct_num, outc_chk_fisyr, outc_chk_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
        End If

        If eusedate And Not eallfiscalyears Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
            & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
            & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
            & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
            & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
            & " and h.chks_applied_date between @p3 and @p4" _
            & " AND (h.chks_recon_sw = 'N')" _
            & " AND (h.chks_status <> 'V')" _
            & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outc_chk_fisyr, outc_chk_num, 'X' AS status," _
            & " outc_recon_sw, '' AS printed, outc_chk_payee_name, '' AS ponum," _
            & " outc_chk_amount, outc_chk_amount, af_acct_num, as_acct_num," _
            & " outc_chk_issue_date, outc_chk_issue_date, outc_chk_descr, '' AS remarks" _
            & " FROM outstandingchecks" _
            & " WHERE bank_acct_num = @p1" _
            & " AND outc_chk_fisyr = @p2" _
            & " AND outc_chk_issue_date BETWEEN @p3 AND @p4" _
            & " AND (outc_stale_sw = 'N')" _
            & " AND (outc_recon_sw = 'N')" _
            & " ORDER BY bank_acct_num, outc_chk_fisyr, outc_chk_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If

        If eusenumber And Not eallfiscalyears Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
            & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
            & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
            & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
            & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
            & " AND h.chks_num between @p3 and @p4" _
            & " AND (h.chks_recon_sw = 'N')" _
            & " AND (h.chks_status <> 'V')" _
            & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outc_chk_fisyr, outc_chk_num, 'X' AS status," _
            & " outc_recon_sw, '' AS printed, outc_chk_payee_name, '' AS ponum," _
            & " outc_chk_amount, outc_chk_amount, af_acct_num, as_acct_num," _
            & " outc_chk_issue_date, outc_chk_issue_date, outc_chk_descr, '' AS remarks" _
            & " FROM outstandingchecks" _
            & " WHERE bank_acct_num = @p1" _
            & " AND outc_chk_fisyr = @p2" _
            & " AND outc_chk_num BETWEEN @p3 and @p4" _
            & " AND (outc_stale_sw = 'N')" _
            & " AND (outc_recon_sw = 'N')" _
            & " ORDER BY bank_acct_num, outc_chk_fisyr, outc_chk_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)
        End If

        If (Not eallfiscalyears) And (Not eusedate) And (Not eusenumber) Then
            Me.CellMiddleBottom = "FY-" & efiscalyear.ToString
            SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
            & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
            & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
            & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
            & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
            & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
            & " AND (h.chks_recon_sw = 'N')" _
            & " AND (h.chks_status <> 'V')" _
            & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outc_chk_fisyr, outc_chk_num, 'X' AS status," _
            & " outc_recon_sw, '' AS printed, outc_chk_payee_name, '' AS ponum," _
            & " outc_chk_amount, outc_chk_amount, af_acct_num, as_acct_num," _
            & " outc_chk_issue_date, outc_chk_issue_date, outc_chk_descr, '' AS remarks" _
            & " FROM outstandingchecks" _
            & " WHERE bank_acct_num = @p1" _
            & " AND outc_chk_fisyr = @p2" _
            & " AND (outc_stale_sw = 'N')" _
            & " AND (outc_recon_sw = 'N')" _
            & " ORDER BY bank_acct_num, outc_chk_fisyr, outc_chk_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
        End If

        Try
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("register")
            da.Fill(ds)
            If (ds.Tables(0).Rows.Count < 1) And (ds.Tables(1).Rows.Count < 1) Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridWrk.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'append the legacy recs (gridwrk) into the main grid (griddetail);
            Dim index, currow As Int32
            With Me.GridDetail
                currow = .Rows.Count - 1
                currow += 1
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    .Rows.Add()
                    .SetData(currow, 0, Me.GridWrk.GetData(index, 0))
                    .SetData(currow, 1, Me.GridWrk.GetData(index, 1))
                    .SetData(currow, 2, Me.GridWrk.GetData(index, 2))
                    .SetData(currow, 3, Me.GridWrk.GetData(index, 3))
                    .SetData(currow, 4, Me.GridWrk.GetData(index, 4))
                    .SetData(currow, 5, Me.GridWrk.GetData(index, 5))
                    .SetData(currow, 6, Me.GridWrk.GetData(index, 6))
                    .SetData(currow, 7, Me.GridWrk.GetData(index, 7))
                    .SetData(currow, 8, Me.GridWrk.GetData(index, 8))
                    .SetData(currow, 9, Me.GridWrk.GetData(index, 9))
                    .SetData(currow, 10, Me.GridWrk.GetData(index, 10))
                    .SetData(currow, 11, Me.GridWrk.GetData(index, 11))
                    .SetData(currow, 12, Me.GridWrk.GetData(index, 12))
                    .SetData(currow, 13, Me.GridWrk.GetData(index, 13))
                    .SetData(currow, 14, Me.GridWrk.GetData(index, 14))
                    .SetData(currow, 15, Me.GridWrk.GetData(index, 15))
                    currow += 1
                Next
            End With

            'Me.Prev1.Visible = False
            'Me.GridDetail.Visible = True
            'Me.ShowDialog()
            'Exit Function

            Dim amount As Double
            'sum the checks;
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    amount += CDbl(.GetData(index, 9))
                Next
            End With
            'save total
            With Me.GridTotals
                .Rows.Count = 1
                .Rows.Add()
                .SetData(0, 0, amount)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            Application.DoEvents()
            'render the table
            PrintOutstandingChecksRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateCheckRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all checks for a single bank given the selected filtering criteria;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0        1         2        3        4          5          6        7         8
        '   bank    fisyr    docnumber  status   recon     printed     payee   ponumber  summary
        '    9        10        11       12       13         14         15
        ' amount     acct     subacct  applied  created   hdrdescr    remarks
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable

        Try
            If eusedate Then
                Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
                SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
                & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
                & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
                & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
                & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
                & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
                & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
                & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
                & " AND h.chks_applied_date BETWEEN @p3 AND @p4" _
                & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", ebankaccountnumber)
                cmd.Parameters.Add("@p2", efiscalyear)
                cmd.Parameters.Add("@p3", edatefrom)
                cmd.Parameters.Add("@p4", edateto)
            End If
            If eusenumber Then
                Me.CellMiddleBottom = enumberfrom & " to " & enumberto
                SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
                & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, p.po_num," _
                & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
                & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr" _
                & " FROM chks_info AS h, chks_detl AS d, purc_detl AS p" _
                & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
                & " AND d.podt_autoinc_key = p.podt_autoinc_key" _
                & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
                & " AND h.chks_num BETWEEN @p3 AND @p4" _
                & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", ebankaccountnumber)
                cmd.Parameters.Add("@p2", efiscalyear)
                cmd.Parameters.Add("@p3", enumberfrom)
                cmd.Parameters.Add("@p4", enumberto)
            End If
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("register")
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
            'summarise the total register amount except for voids
            Dim voidamount, totalamount As Double
            Dim index As Int32
            Dim status As String
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    status = DirectCast(.GetData(index, 3), String).ToUpper
                    If status = "V" Then voidamount += CDbl(.GetData(index, 9))
                    totalamount += CDbl(.GetData(index, 9))
                Next
            End With
            With Me.GridTotals
                .Rows.Count = 1
                .Rows.Add()
                .SetData(0, 0, totalamount)
                .SetData(0, 1, voidamount)
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
            PrintCheckRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateInvoicePending(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all open invoices that have not yet been converted into checks;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status  amount   ponum
        '     6           7         8        9       10       11  
        '   acct      subacct    datedue  datepaid  ocas    vendnum
        '    12          13        14       15       16       17  
        ' vendname    applied    issued   invckey  pokey    podtkey
        '    18
        ' ckdtkey
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim cmd As SqlCommand

        Try
            SSQL = "SELECT i.bank_acct_num, invc_fisyr, invc_num, invc_status, invc_amount," _
            & " po_num, af_acct_num, as_acct_num, invc_datedue, invc_datepaid, ocex_code," _
            & " i.vend_number, v.vend_name, invc_applied_date, invc_issued_date," _
            & " invc_autoinc_key, po_autoinc_key, podt_autoinc_key, ckdt_autoinc_key" _
            & " FROM invoices AS i, vend_info AS v" _
            & " WHERE i.vend_number = v.vend_number" _
            & " AND i.bank_acct_num = @p1" _
            & " AND i.invc_fisyr = @p2" _
            & " AND invc_status = 'O'" _
            & " ORDER BY i.bank_acct_num, invc_fisyr, CAST(po_num AS INT), invc_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("pending")
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for this criteria.")
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleBottom = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table;
            Call PrintInvoicePending()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Friend Function GeneratePositivePayFile(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal startnumber As Int32, ByVal endnumber As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Collects information from the checks table by register number, then saves this information
        'to a file on the local (or network) drive; A printout is included in this option;
        'Added on 2016.08.01, Fred;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0        1         2         3         4          5          6         7
        '   bank    fisyr     number    status     payee     amount     issued   register
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim filename As String = ""

        Try
            Me.BankAccountNumber = ebankaccountnumber
            '
            If startnumber > endnumber Then
                MsgBox("The beginning check number must be less than or equal to the ending check number.", MsgBoxStyle.Exclamation, MSGTITLE)
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, MSGTITLE)
            Return False
        End Try

        Try
            'Get checks by beginning and ending number for fiscal year;
            SSQL = "SELECT bank_acct_num, chks_fisyr, chks_num, chks_status, chks_payee_name, chks_amount, chks_datetime, ckrg_autoinc_key" _
            + " FROM chks_info WHERE bank_acct_num = @p1 AND chks_fisyr = @p2 AND CAST(chks_num AS INT) BETWEEN @p3 AND @p4" _
            + " ORDER BY CAST(chks_num AS INT)"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", startnumber)
            cmd.Parameters.Add("@p4", endnumber)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("register")
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then
                MsgBox("No records found for the selected check series.", MsgBoxStyle.Exclamation, MSGTITLE)
                Return True
            End If
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        Try
            'Build the file name for the positive pay file;
            filename = startnumber.ToString + "." + endnumber.ToString + "." + Me.SchoolNumber + ".txt"
            'Save the positive pay file;
            Call SavePositivePay(filename)
        Catch ex As Exception
            Throw
        End Try


        Dim response As DialogResult
        response = MessageBox.Show("Would you like to run a report for this register series?", "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        If response <> DialogResult.Yes Then Return True

        Try
            If startnumber = endnumber Then Me.CellRightMiddle = "Check(s): " + String.Format("{0:D2}", startnumber)
            If startnumber < endnumber Then Me.CellRightMiddle = "Checks: " + String.Format("{0:D2}", startnumber) + " to " + String.Format("{0:D2}", endnumber)
            Application.DoEvents()
            'render the table;
            Call PrintPositivePay()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrderActivity(ByVal epurchaseorderkey As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''''' GridWrk ''''''''''''''''''''''''''''''''''''''''''
        '      0            1           2           3           4           5          6 
        '    bank         fisyr      number      status      amount      account    subacct
        '      7            8           9          10          11          12         13  
        '    code         descr      issued      podtkey     pokey       invckey    ckdtkey
        '     14  
        '  rectype
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''''''''''''
        '      0            1           2           3           4           5           6
        '    pokey      bankacct      fisyr      number      descr       vendor      potype 
        '      7            8           9          10          11          12          13 
        '   issued       reqkey     reqnumber    totenc     totinvc     totspent   outstanding
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim index, holdindex As Int32
        '
        Dim number, status As String
        Dim totalencumbered, totalinvoiced, totaloutstanding, totalspent As Decimal
        Dim recordtype As Int32
        '
        Dim fisyr, requisitionkey As Int32
        Dim bankacct, ponumber, reqnumber As String
        '
        Try
            'collect the header information;
            SSQL = "SELECT po_autoinc_key, bank_acct_num, po_fisyr, po_num, po_descr," _
            & " vend_name, po_type, po_datetime, rqst_autoinc_key, ' ' AS reqnumber," _
            & " 0.0 AS encumbered, 0.0 AS invoiced, 0.0 AS spent, 0.0 AS outstanding" _
            & " FROM purc_info AS p, vend_info AS v" _
            & " WHERE p.vend_number = v.vend_number" _
            & " AND p.po_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", epurchaseorderkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable
            da.Fill(tbl)
            If tbl.Rows.Count > 0 Then
                Me.GridTotals.DataSource = tbl
            Else
                Return False
            End If
            'collect some header information;
            With Me.GridTotals
                bankacct = CType(.GetData(1, 1), String)
                fisyr = CType(.GetData(1, 2), Int32)
                ponumber = CType(.GetData(1, 3), String)
                requisitionkey = CType(.GetData(1, 8), Int32)
            End With
        Catch ex As Exception
            Throw
        End Try

        'if a requisition is tied, retrieve the number & set field;
        If requisitionkey > 0 Then
            Try
                'collect the header information;
                SSQL = "SELECT req_num FROM req_info WHERE req_autoinc_key = @p1"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", requisitionkey)
                If cn.State <> ConnectionState.Open Then cn.Open()
                reqnumber = CType(cmd.ExecuteScalar, String)
                'set the reqnumber in the totals grid;
                Me.GridTotals.SetData(0, 9, reqnumber)
            Catch ex As Exception
                Throw
            Finally
                cn.Close()
            End Try
        End If

        Try
            SSQL = "SELECT p.bank_acct_num, p.po_fisyr, p.po_num, d.podt_status, d.podt_amount," _
            & " d.af_acct_num, d.as_acct_num, d.ocex_code, d.podt_descr, d.podt_datetime," _
            & " d.podt_autoinc_key, p.po_autoinc_key, 0, 0, 1" _
            & " FROM purc_info AS p, purc_detl AS d" _
            & " WHERE p.bank_acct_num = d.bank_acct_num" _
            & " AND p.po_fisyr = d.po_fisyr AND p.po_num = d.po_num" _
            & " AND p.bank_acct_num = @p1 AND p.po_fisyr = @p2 AND p.po_num = @p3"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", bankacct)
            cmd.Parameters.Add("@p2", fisyr)
            cmd.Parameters.Add("@p3", ponumber)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
            'setup the working grid;
            With Me.GridWrk
                .Cols.Count = Me.GridWrkTotals.Cols.Count
                .Rows.Count = 1
                .Cols(0).DataType = GetType(String)
                .Cols(1).DataType = GetType(Int32)
                .Cols(2).DataType = GetType(String)
                .Cols(3).DataType = GetType(String)
                .Cols(4).DataType = GetType(Decimal)
                .Cols(5).DataType = GetType(String)
                .Cols(6).DataType = GetType(String)
                .Cols(7).DataType = GetType(String)
                .Cols(8).DataType = GetType(String)
                .Cols(9).DataType = GetType(Date)
                .Cols(10).DataType = GetType(Int32)
                .Cols(11).DataType = GetType(Int32)
                .Cols(12).DataType = GetType(Int32)
                .Cols(13).DataType = GetType(Int32)
                .Cols(14).DataType = GetType(Int32)
                .SetData(0, 0, "Bank account")
                .SetData(0, 1, "Year")
                .SetData(0, 2, "Number")
                .SetData(0, 3, "Status")
                .SetData(0, 4, "Amount")
                .SetData(0, 5, "Acct")
                .SetData(0, 6, "Sub")
                .SetData(0, 7, "Expenditure code")
                .SetData(0, 8, "Description")
                .SetData(0, 9, "Issued")
                .SetData(0, 10, "PODTKEY")
                .SetData(0, 11, "POKEY")
                .SetData(0, 12, "INVCKEY")
                .SetData(0, 13, "CKDTKEY")
                .SetData(0, 14, "Rectype")
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'load the temporary grid into the working grid;
            With Me.GridWrkTotals
                For index = 1 To .Rows.Count - 1
                    Me.GridWrk.Rows.Add()
                    holdindex = Me.GridWrk.Rows.Count - 1
                    Me.GridWrk.SetData(holdindex, 0, .GetData(index, 0))
                    Me.GridWrk.SetData(holdindex, 1, .GetData(index, 1))
                    Me.GridWrk.SetData(holdindex, 2, .GetData(index, 2))
                    Me.GridWrk.SetData(holdindex, 3, .GetData(index, 3))
                    Me.GridWrk.SetData(holdindex, 4, .GetData(index, 4))
                    Me.GridWrk.SetData(holdindex, 5, .GetData(index, 5))
                    Me.GridWrk.SetData(holdindex, 6, .GetData(index, 6))
                    Me.GridWrk.SetData(holdindex, 7, .GetData(index, 7))
                    Me.GridWrk.SetData(holdindex, 8, .GetData(index, 8))
                    Me.GridWrk.SetData(holdindex, 9, .GetData(index, 9))
                    Me.GridWrk.SetData(holdindex, 10, .GetData(index, 10))
                    Me.GridWrk.SetData(holdindex, 11, .GetData(index, 11))
                    Me.GridWrk.SetData(holdindex, 12, .GetData(index, 12))
                    Me.GridWrk.SetData(holdindex, 13, .GetData(index, 13))
                    Me.GridWrk.SetData(holdindex, 14, .GetData(index, 14))
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'get any invoice tied to the purchase order;
            SSQL = "SELECT bank_acct_num, invc_fisyr, invc_num, invc_status, invc_amount," _
            & " af_acct_num, as_acct_num, ocex_code, ' ', invc_issued_date, podt_autoinc_key," _
            & " po_autoinc_key, invc_autoinc_key, ckdt_autoinc_key, 2" _
            & " FROM invoices" _
            & " WHERE po_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", epurchaseorderkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
            'if any invoices are voided, replace the description with Void tag;
            With Me.GridWrkTotals
                For index = 1 To .Rows.Count - 1
                    number = CType(.GetData(index, 2), String)
                    status = CType(.GetData(index, 3), String)
                    If status = "V" Then
                        number = "[VOID INVOICE] " & number
                        .SetData(index, 2, number)
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'load the temporary grid into the working grid;
            With Me.GridWrkTotals
                For index = 1 To .Rows.Count - 1
                    Me.GridWrk.Rows.Add()
                    holdindex = Me.GridWrk.Rows.Count - 1
                    Me.GridWrk.SetData(holdindex, 0, .GetData(index, 0))
                    Me.GridWrk.SetData(holdindex, 1, .GetData(index, 1))
                    Me.GridWrk.SetData(holdindex, 2, .GetData(index, 2))
                    Me.GridWrk.SetData(holdindex, 3, .GetData(index, 3))
                    Me.GridWrk.SetData(holdindex, 4, .GetData(index, 4))
                    Me.GridWrk.SetData(holdindex, 5, .GetData(index, 5))
                    Me.GridWrk.SetData(holdindex, 6, .GetData(index, 6))
                    Me.GridWrk.SetData(holdindex, 7, .GetData(index, 7))
                    Me.GridWrk.SetData(holdindex, 8, .GetData(index, 2))    'use invoice number;
                    Me.GridWrk.SetData(holdindex, 9, .GetData(index, 9))
                    Me.GridWrk.SetData(holdindex, 10, .GetData(index, 10))
                    Me.GridWrk.SetData(holdindex, 11, .GetData(index, 11))
                    Me.GridWrk.SetData(holdindex, 12, .GetData(index, 12))
                    Me.GridWrk.SetData(holdindex, 13, .GetData(index, 13))
                    Me.GridWrk.SetData(holdindex, 14, .GetData(index, 14))
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'get any check tied to the purchase order via the invoices;
            SSQL = "SELECT k.bank_acct_num, k.chks_fisyr, chks_num, chks_status, ckdt_amount," _
            & " c.af_acct_num, c.as_acct_num, c.ocex_code, c.ckdt_descr, chks_datetime," _
            & " c.podt_autoinc_key, 0, c.invc_autoinc_key, c.ckdt_autoinc_key, 3" _
            & " FROM purc_info AS p, purc_detl AS d, chks_info AS k, chks_detl AS c" _
            & " WHERE p.bank_acct_num = d.bank_acct_num" _
            & " AND p.po_fisyr = d.po_fisyr" _
            & " AND p.po_num = d.po_num" _
            & " AND k.chks_autoinc_key = c.chks_autoinc_key" _
            & " AND d.podt_autoinc_key = c.podt_autoinc_key" _
            & " AND p.po_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", epurchaseorderkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
            'if any checks are voided, replace the description with Void tag;
            With Me.GridWrkTotals
                For index = 1 To .Rows.Count - 1
                    number = CType(.GetData(index, 2), String)
                    status = CType(.GetData(index, 3), String)
                    If status = "V" Then .SetData(index, 8, "[VOID CHECK] " & number)
                Next
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'load the temporary grid into the working grid;
            With Me.GridWrkTotals
                For index = 1 To .Rows.Count - 1
                    Me.GridWrk.Rows.Add()
                    holdindex = Me.GridWrk.Rows.Count - 1
                    Me.GridWrk.SetData(holdindex, 0, .GetData(index, 0))
                    Me.GridWrk.SetData(holdindex, 1, .GetData(index, 1))
                    Me.GridWrk.SetData(holdindex, 2, .GetData(index, 2))
                    Me.GridWrk.SetData(holdindex, 3, .GetData(index, 3))
                    Me.GridWrk.SetData(holdindex, 4, .GetData(index, 4))
                    Me.GridWrk.SetData(holdindex, 5, .GetData(index, 5))
                    Me.GridWrk.SetData(holdindex, 6, .GetData(index, 6))
                    Me.GridWrk.SetData(holdindex, 7, .GetData(index, 7))
                    Me.GridWrk.SetData(holdindex, 8, .GetData(index, 8))
                    Me.GridWrk.SetData(holdindex, 9, .GetData(index, 9))
                    Me.GridWrk.SetData(holdindex, 10, .GetData(index, 10))
                    Me.GridWrk.SetData(holdindex, 11, .GetData(index, 11))
                    Me.GridWrk.SetData(holdindex, 12, .GetData(index, 12))
                    Me.GridWrk.SetData(holdindex, 13, .GetData(index, 13))
                    Me.GridWrk.SetData(holdindex, 14, .GetData(index, 14))
                Next
            End With
            '
            With Me.GridWrk
                'sort the entire working grid by the issued date;
                .Sort(SortFlags.Ascending, 9)
                'summarize the totals for display;
                For index = 1 To .Rows.Count - 1
                    status = CType(.GetData(index, 3), String)
                    recordtype = CType(.GetData(index, 14), Int32)
                    'add up the encumbrance;
                    If recordtype = 1 Then
                        If (status <> "D") And (status <> "V") Then
                            totalencumbered += CType(.GetData(index, 4), Decimal)
                        End If
                    End If
                    'add up the invoiced;
                    If recordtype = 2 Then
                        If status <> "V" Then
                            totalinvoiced += CType(.GetData(index, 4), Decimal)
                        End If
                    End If
                    'add up the spent;
                    If recordtype = 3 Then
                        If status <> "V" Then
                            totalspent += CType(.GetData(index, 4), Decimal)
                        End If
                    End If
                Next
                totaloutstanding = totalencumbered - totalspent
            End With
            'set the total amounts in the header grid;
            With Me.GridTotals
                .SetData(1, 10, totalencumbered)
                .SetData(1, 11, totalinvoiced)
                .SetData(1, 12, totalspent)
                .SetData(1, 13, totaloutstanding)
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridDetail.Visible = False
        '''''Me.GridTotals.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.GridWrkTotals.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            'Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            Call PrintPurchaseOrderActivity()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrder(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal enumber As String) As Boolean
        'this routine is used to fulfill requests from other modules to print a single purchase order;
        Try
            Call GeneratePurchaseOrder(ebankaccountnumber, efiscalyear, enumber, enumber)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrder(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal edate1 As Date, ByVal edate2 As Date) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'prints purchase orders by a range of purchase order applied fiscal dates;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim row As DataRow
        Dim pages As Int32

        Try
            cn = New SqlConnection(Me.ConnectionString)
            SSQL = "SELECT po_num FROM purc_info" _
            & " WHERE bank_acct_num = @p1 AND po_fisyr = @p2" _
            & " AND po_applied_date BETWEEN @p3 AND @p4"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edate1)
            cmd.Parameters.Add("@p4", edate2)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("ticket")
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'define the document;
            Me.DocumentName = "PurchaseOrderTicket"
            Me.ReportName = "Activity Fund - Purchase Order Ticket"
            Call DefineDocumentSettings(Me.DocumentName)
            'start the document;
            Me.Doc1.StartDoc()
            Application.DoEvents()
            '
            Dim number As String
            '
            For Each row In tbl.Rows
                pages += 1
                number = CType(row.Item(0), String)
                Call GeneratePurchaseOrder(ebankaccountnumber, efiscalyear, number, pages)
                Application.DoEvents()
            Next
            'set the preview zoom;
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document;
            Me.Doc1.EndDoc()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrder(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal enumber1 As String, ByVal enumber2 As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'prints purchase orders by a range of purchase order numbers;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim row As DataRow
        Dim pages As Int32

        Try
            cn = New SqlConnection(Me.ConnectionString)
            SSQL = "SELECT po_num FROM purc_info" _
            & " WHERE bank_acct_num = @p1 AND po_fisyr = @p2" _
            & " AND po_num BETWEEN @p3 AND @p4"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumber1)
            cmd.Parameters.Add("@p4", enumber2)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("ticket")
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'define the document;
            Me.DocumentName = "PurchaseOrderTicket"
            Me.ReportName = "Activity Fund - Purchase Order Ticket"
            Call DefineDocumentSettings(Me.DocumentName)
            'start the document;
            Me.Doc1.StartDoc()
            Application.DoEvents()
            '
            Dim number As String
            '
            For Each row In tbl.Rows
                pages += 1
                number = CType(row.Item(0), String)
                Call GeneratePurchaseOrder(ebankaccountnumber, efiscalyear, number, pages)
                Application.DoEvents()
            Next
            'set the preview zoom;
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document;
            Me.Doc1.EndDoc()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function GeneratePurchaseOrder(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal enumber As String, ByVal epage As Int32) As Boolean
        '''''''''''''''''''''''''''''' GRIDTOTALS '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5        6
        '   bank       fisyr     number  amount   status    applied   issued
        '     7           8         9       10       11        12       13
        '  vnumber     vname     vaddr1  vaddr2   vaddr3     vcity    vstate
        '    14          15        16       17       18        19       20 
        '   vzip        vext      vph1     vph2     descr   shipkey  shipattn
        '    21          22        23       24 
        '  vattn      checknum  chkprtd  chkdate
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5        6
        '    qty        cost     amount    acct     name       sub     name
        '     7           8         9       10       11        12       13 
        '   code      chknum   chkstatus  chkdate  status   remarks  invckey
        '    30
        ' podtkey  
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim row As DataRow
        Dim tbl As DataTable
        Dim cmd As SqlCommand
        Dim index As Int32
        '
        cn = New SqlConnection(Me.ConnectionString)
        '
        Try
            'collect header information;
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, 0.00 AS Amount, h.po_status," _
            & " h.po_applied_date, h.po_datetime, h.vend_number, v.vend_name, v.vend_addr1," _
            & " v.vend_addr2, v.vend_addr3, v.vend_city, v.vend_state, v.vend_zip, v.vend_zip_ext," _
            & " v.vend_phone1, v.vend_phone2, h.po_descr, ship_autoinc_key, ship_attn, ship_vendor_attn," _
            & " ' ' AS checknumber, ' ' AS checkprinted, ' ' AS checkdate, vend_fax" _
            & " FROM purc_info AS h, vend_info AS v" _
            & " WHERE h.vend_number = v.vend_number" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_num = @p3"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumber)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("ticket")
            da.Fill(tbl)
            If tbl.Rows.Count > 0 Then
                Me.GridTotals.DataSource = tbl
            End If
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
        End Try

        Try
            'collect purchase detail lines;
            SSQL = "SELECT podt_qty, podt_unitp, podt_amount, af_acct_num, af_acct_name," _
            & " as_acct_num, as_acct_name, ocex_code, ' ' AS chknum, ' ' AS chkstatus, getdate() AS chkdate," _
            & " podt_status, podt_descr, invc_autoinc_key, podt_autoinc_key" _
            & " FROM purc_detl" _
            & " WHERE (podt_status <> 'D')" _
            & " AND (podt_status <> 'V')" _
            & " AND bank_acct_num = @p1 AND po_fisyr = @p2 AND po_num = @p3" _
            & " ORDER BY podt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumber)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("ticket")
            da.Fill(tbl)
            If tbl.Rows.Count > 0 Then
                Me.GridDetail.DataSource = tbl
            End If
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
        End Try

        Try
            'iterate thru the detail grid, summarize the money, and
            'collect the check information for each purchase detail line;
            Dim ckdtkey, checkcount, checkkey, invoicekey As Int32
            Dim purchasetotal As Decimal
            Dim checknumber, checkpaid, checkstatus, prevchecknumber As String
            Dim checkdate As Date
            With Me.GridDetail
                'set default values;
                checknumber = ""
                checkstatus = ""
                checkdate = Now
                checkpaid = ""
                '
                For index = 1 To .Rows.Count - 1
                    purchasetotal += CType(.GetData(index, 2), Decimal)
                    invoicekey = CType(.GetData(index, 13), Int32)
                    If invoicekey > 0 Then
                        SSQL = "SELECT chks_num, chks_status, chks_datetime" _
                        & " FROM chks_info AS c, chks_detl AS d" _
                        & " WHERE c.chks_autoinc_key = d.chks_autoinc_key" _
                        & " AND d.ckdt_autoinc_key IN (" _
                        & " SELECT ckdt_autoinc_key FROM invoices WHERE invc_autoinc_key = @p1)"
                        cmd = New SqlCommand(SSQL, cn)
                        cmd.Parameters.Add("@p1", invoicekey)
                        da = New SqlDataAdapter(cmd)
                        tbl = New DataTable("ticket")
                        da.Fill(tbl)
                        If tbl.Rows.Count > 0 Then
                            checknumber = CType(tbl.Rows(0).Item(0), String)
                            checkstatus = CType(tbl.Rows(0).Item(1), String)
                            checkdate = CType(tbl.Rows(0).Item(2), Date)
                        End If
                    End If
                    'update the grid with the information from the query;
                    .SetData(index, 8, checknumber)
                    .SetData(index, 9, checkstatus)
                    .SetData(index, 10, checkdate)
                    'initialise vars for the next found invoice;
                    checknumber = ""
                    checkstatus = ""
                    checkdate = Now
                    checkpaid = ""
                Next

                'determine the number of non-voided checks tied to this purchase order;
                For index = 1 To .Rows.Count - 1
                    checkstatus = CType(.GetData(index, 9), String).ToUpper
                    Select Case checkstatus
                        Case "C", "O"
                            checknumber = CType(.GetData(index, 8), String).ToUpper
                            checkdate = CType(.GetData(index, 10), Date)
                            checkpaid = "Paid"
                            checkcount += 1
                        Case "F"
                            checknumber = CType(.GetData(index, 8), String).ToUpper
                            checkdate = CType(.GetData(index, 10), Date)
                            checkpaid = "Cleared"
                            checkcount += 1
                        Case "V"
                            '
                    End Select
                    '
                    If index = 1 Then prevchecknumber = checknumber
                    'if more than one non-voided check is tied to the purchase order, set appropriate description;
                    If checkcount > 1 Then
                        If checknumber.Compare(checknumber, prevchecknumber) <> 0 Then
                            checknumber = "Multiple"
                            Exit For
                        End If
                    End If
                Next
            End With
            'update the header record with the check information & total for the purchase order;
            With Me.GridTotals
                .SetData(1, 3, purchasetotal)
                .SetData(1, 22, checknumber)
                .SetData(1, 23, checkpaid)
                .SetData(1, 24, checkdate)
            End With
        Catch ex As Exception
            Throw
        End Try

        'Me.Prev1.Visible = False
        'Me.GridDetail.Visible = True
        'Me.GridWrk.Visible = False
        'Me.GridWrkTotals.Visible = False
        'Me.GridTotals.Visible = False
        'Exit Function

        Try
            'get the signature image(s);
            Call GetSignatureDetails()
            Application.DoEvents()
            'render the purchase order;
            Call PrintPurchaseOrder(epage)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrderRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String, ByVal eopenpurchaseorders As Boolean) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5     6
        '   bank        fisyr    ponumber  vendor  hdrstatus   qty  cost
        '     7           8         9       10       11        12 
        '  amount      hdramt     acct     sub    applied   created
        '    13          14        15       16       17
        ' detlstatus  hdrdescr  remarks  chknumber chkstat
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all purchase orders for a single bank given the selected filtering criteria;
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        Dim filter As String = ""
        If eopenpurchaseorders Then filter = " AND h.po_status = 'O'"
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, '' AS checknumber, '' AS checkstatus, '' AS checkreconsw" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & filter _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.po_num, d.podt_autoinc_key;"
            SSQL += "SELECT po_num, chks_num, chks_status, chks_recon_sw FROM chks_info" _
            & " WHERE bank_acct_num = @p1 AND chks_fisyr = @p2" _
            & " ORDER BY chks_num"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, '' AS checknumber, '' AS checkstatus, '' AS checkreconsw" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & filter _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.po_num, d.podt_autoinc_key;"
            SSQL += "SELECT po_num, chks_num, chks_status, chks_recon_sw FROM chks_info" _
            & " WHERE bank_acct_num = @p1 AND chks_fisyr = @p2" _
            & " ORDER BY chks_num"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)
        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
        Dim ds As New DataSet("register")
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
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Dim chknumber, chkstatus, chkreconsw, ponumber, postatus As String
        Dim index, row As Int32
        Dim amount As Double

        Try
            With Me.GridDetail
                ''''''summarise the detail amounts into the header record
                Dim tcurnum, tnextnum, tholdnum As String
                Dim j, k As Int32
                Dim ttempamt, tamount As Double
                For index = 1 To .Rows.Count - 1
                    tcurnum = DirectCast(.GetData(index, 2), String)
                    For j = index To .Rows.Count - 1
                        tholdnum = DirectCast(.GetData(j, 2), String)
                        tamount = CDbl(.GetData(j, 7))
                        If tcurnum = tholdnum Then
                            ttempamt += tamount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 8, ttempamt)
                    Next
                    index = j - 1
                    ttempamt = 0
                Next

                'get the matching check number for each po
                For index = 1 To .Rows.Count - 1
                    ponumber = DirectCast(.GetData(index, 2), String)
                    postatus = DirectCast(.GetData(index, 4), String).ToUpper
                    If postatus <> "O" Then
                        'get the matching chknumber & status
                        row = Me.GridTotals.FindRow(ponumber, 0, 0, False, True, False)
                        If row >= 0 Then
                            chknumber = DirectCast(Me.GridTotals.GetData(row, 1), String)
                            chkstatus = DirectCast(Me.GridTotals.GetData(row, 2), String).ToUpper
                            chkreconsw = DirectCast(Me.GridTotals.GetData(row, 3), String).ToUpper
                        End If
                        If postatus = "X" Then
                            .SetData(index, 16, "-CLOSED-")
                            .SetData(index, 17, "")
                            .SetData(index, 18, "")
                        Else
                            'update the detail grid with the check information
                            .SetData(index, 16, chknumber)
                            .SetData(index, 17, chkstatus)
                            .SetData(index, 18, chkreconsw)
                        End If
                    End If

                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'we're done with the totals grid so lets clear it out 
            '& put some totals in it;
            Me.GridTotals.DataSource = Nothing
            Me.GridTotals.Cols.Count = 3
            Me.GridTotals.Rows.Count = 1
            Dim totalpoamount As Double
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    chkstatus = DirectCast(.GetData(index, 17), String).ToUpper
                    amount = CDbl(.GetData(index, 7))
                    If chkstatus.Compare(chkstatus, "V") <> 0 Then totalpoamount += amount
                Next
            End With
            Me.GridTotals.SetData(0, 0, totalpoamount)
        Catch ex As Exception
            Throw
        End Try


        '''''Me.Prev1.Visible = True
        '''''Me.GridDetail.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintPurchaseOrderRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrderRegisterByAccount(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String, ByVal eoutstandinginvoices As Boolean) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Note: this is the same report as the GeneratePurchaseOrderRegister routine except it is 
        'defined for invoices;  Eventually, this report will replace the original after FY08 since
        'all schools will be using invoices by that time;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''' GridDetail '''''''''''''''''''''''''''''''''''''''
        '     0             1           2           3           4           5           6
        '   bank          fisyr     ponumber      vendor    hdrstatus      qty        cost
        '     7             8           9          10          11          12          13 
        '  detlamt        hdramt      acct        sub       applied      issued    detlstatus 
        '    14            15          16          17          18          19          20  
        ' hdrdescr       remarks     class       pokey      podtkey     invckey    hdrinvoice
        '    21            22          23          24          25 
        ' detlinvoice    hdrpaid    detlpaid    chknum       payee
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all purchase orders for a single bank given the selected filtering criteria;
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            'collect the purchase orders for the selected criteria;
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, d.ocex_code, h.po_autoinc_key," _
            & " d.podt_autoinc_key, d.invc_autoinc_key," _
            & " 0.00 AS hdrinvoiced, 0.00 AS detlinvoiced, 0.00 AS hdrpaid, 0.00 AS detlpaid," _
            & " '' AS checknumber, '' AS checkpayee" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.af_acct_num, d.as_acct_num, d.po_num, d.podt_autoinc_key; "
            'collect the invoices for the selected criteria;
            SSQL &= "SELECT SUM(invc_amount), 0.00 AS hdramount, i.po_autoinc_key, i.podt_autoinc_key" _
            & " FROM invoices AS i, purc_info AS h, purc_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND i.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND (d.podt_status <> 'D')" _
            & " AND (i.invc_status <> 'V')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " GROUP BY i.po_autoinc_key, i.podt_autoinc_key" _
            & " ORDER BY i.po_autoinc_key, i.podt_autoinc_key; "
            'collect the checks for the selected criteria;
            SSQL &= "SELECT SUM(ckdt_amount), 0.00 AS hdramount, h.po_autoinc_key, c.podt_autoinc_key" _
            & " FROM chks_info AS k, chks_detl AS c, purc_info AS h, purc_detl AS d " _
            & " WHERE k.chks_autoinc_key = c.chks_autoinc_key" _
            & " AND c.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND k.chks_status <> 'V'" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " GROUP BY h.po_autoinc_key, c.podt_autoinc_key" _
            & " ORDER BY h.po_autoinc_key, c.podt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            'collect the purchase orders for the selected criteria;
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, d.ocex_code, h.po_autoinc_key," _
            & " d.podt_autoinc_key, d.invc_autoinc_key," _
            & " 0.00 AS hdrinvoiced, 0.00 AS detlinvoiced, 0.00 AS hdrpaid, 0.00 AS detlpaid," _
            & " '' AS checknumber, '' AS checkpayee" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.af_acct_num, d.as_acct_num, d.po_num, d.podt_autoinc_key; "
            'collect the invoices for the selected criteria;
            SSQL &= "SELECT SUM(invc_amount), 0.00 AS hdramount, i.po_autoinc_key, i.podt_autoinc_key" _
            & " FROM invoices AS i, purc_info AS h, purc_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND i.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND (d.podt_status <> 'D')" _
            & " AND (i.invc_status <> 'V')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " GROUP BY i.po_autoinc_key, i.podt_autoinc_key" _
            & " ORDER BY i.po_autoinc_key, i.podt_autoinc_key; "
            'collect the checks for the selected criteria;
            SSQL &= "SELECT SUM(ckdt_amount), 0.00 AS hdramount, h.po_autoinc_key, c.podt_autoinc_key" _
            & " FROM chks_info AS k, chks_detl AS c, purc_info AS h, purc_detl AS d " _
            & " WHERE k.chks_autoinc_key = c.chks_autoinc_key" _
            & " AND c.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND k.chks_status <> 'V'" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " GROUP BY h.po_autoinc_key, c.podt_autoinc_key" _
            & " ORDER BY h.po_autoinc_key, c.podt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)

            'SSQL += "SELECT af_acct_num, as_acct_num," _
            '& " ISNULL(SUM(podt_amount), 0.0) AS Encumbered," _
            '& " ISNULL(SUM(invoice_total), 0.0) AS Invoiced," _
            '& " ISNULL(SUM(expense_total), 0.0) AS Paid," _
            '& " 0.00 AS HEncumbered, 0.00 AS HInvoiced, 0.00 AS HPaid" _
            '& " FROM purc_detl" _
            '& " WHERE po_fisyr = " & Me.FiscalYear _
            '& " AND (podt_status <> 'D')" _
            '& " AND bank_acct_num = @p1" _
            '& " GROUP BY af_acct_num, as_acct_num" _
            '& " ORDER BY af_acct_num, as_acct_num"

            'select  af_acct_num, as_acct_num,po_fisyr, po_num,podt_status,bank_acct_num,
            'invoice_total(, expense_total)
            'FROM(purc_detl)
            'WHERE bank_acct_num = '3102785'
            'AND po_fisyr = 2009
            'AND podt_status <> 'D'
            'ORDER BY af_acct_num, as_acct_num

        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
        Dim ds As New DataSet("register")
        Try
            da.Fill(ds)
            'verify there are records for the selected criteria;
            If ds.Tables(0).Rows.Count < 1 Then
                MsgBox("No records found for this criteria...", MsgBoxStyle.Information, MSGTITLE)
                Exit Function
            End If
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridWrk.DataSource = ds.Tables(1)
            Me.GridWrkTotals.DataSource = ds.Tables(2)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Dim chknumber, chkstatus, chkreconsw, payee, ponumber, postatus As String
        Dim holdkey, index, pokey, podtkey, row, holdpodtkey As Int32
        Dim hdramount, detlamount, tempamount As Decimal
        Dim balance, encumbered, invoiced, paid As Decimal
        Dim i, j, k As Int32
        Try
            With Me.GridDetail
                'summarise the detail amounts into the header record for the purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 17))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 17))
                        hdramount = CDec(.GetData(j, 7))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 8, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next
            End With

            With Me.GridWrk
                'summarise the detail invoices into the header invoice for each purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 2))
                        hdramount = CDec(.GetData(j, 0))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 1, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next

                'match the invoice amounts for each purchase order line into the detail grid;
                For index = 1 To .Rows.Count - 1
                    detlamount = CDec(.GetData(index, 0))
                    podtkey = CInt(.GetData(index, 3))
                    row = Me.GridDetail.FindRow(podtkey.ToString, 1, 18, True, True, False)
                    If row >= 0 Then
                        Me.GridDetail.SetData(row, 21, detlamount)
                    End If
                Next

                'match the invoice amounts for each purchase order header into the detail grid;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    hdramount = CDec(.GetData(index, 1))
                    row = Me.GridDetail.FindRow(pokey.ToString, 1, 17, True, True, False)
                    If row >= 0 Then
                        With Me.GridDetail
                            For j = row To .Rows.Count - 1
                                holdkey = CInt(.GetData(j, 17))
                                If pokey = holdkey Then
                                    Me.GridDetail.SetData(j, 20, hdramount)
                                Else
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                Next
            End With

            With Me.GridWrkTotals
                'summarise the check amounts into a header amount for each purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 2))
                        hdramount = CDec(.GetData(j, 0))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 1, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next
                'match the check paid amounts for each purchase order line into the detail grid;
                For index = 1 To .Rows.Count - 1
                    detlamount = CDec(.GetData(index, 0))
                    podtkey = CInt(.GetData(index, 3))
                    row = Me.GridDetail.FindRow(podtkey.ToString, 1, 18, True, True, False)
                    If row >= 0 Then
                        Me.GridDetail.SetData(row, 23, detlamount)
                    End If
                Next

                'match the check amounts for each purchase order header into the detail grid;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    hdramount = CDec(.GetData(index, 1))
                    row = Me.GridDetail.FindRow(pokey.ToString, 1, 17, True, True, False)
                    If row >= 0 Then
                        With Me.GridDetail
                            For j = row To .Rows.Count - 1
                                holdkey = CInt(.GetData(j, 17))
                                If pokey = holdkey Then
                                    Me.GridDetail.SetData(j, 22, hdramount)
                                Else
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                Next

            End With
        Catch ex As Exception
            Throw
        End Try

        'calculate totals for the header totals;
        With Me.GridTotals
            .Cols.Count = 4
            .Rows.Count = 2
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    encumbered += CDec(.GetData(index, 7))
                    invoiced += CDec(.GetData(index, 21))
                    paid += CDec(.GetData(index, 23))
                Next
            End With
            balance = encumbered - paid
            .SetData(1, 0, encumbered)
            .SetData(1, 1, invoiced)
            .SetData(1, 2, paid)
            .SetData(1, 3, balance)
        End With


        'Me.Prev1.Visible = False
        'Me.GridDetail.Visible = False
        'Me.GridWrk.Visible = False
        'Me.GridWrkTotals.Visible = False
        'Me.GridTotals.Visible = True
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table;
            Call PrintPurchaseOrderRegisterByAccount(eoutstandinginvoices)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GeneratePurchaseOrderRegisterByInvoice(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String, ByVal eoutstandinginvoices As Boolean) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Note: this is the same report as the GeneratePurchaseOrderRegister routine except it is 
        'defined for invoices;  Eventually, this report will replace the original after FY08 since
        'all schools will be using invoices by that time;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''' GridDetail '''''''''''''''''''''''''''''''''''''''
        '     0             1           2           3           4           5           6
        '   bank          fisyr     ponumber      vendor    hdrstatus      qty        cost
        '     7             8           9          10          11          12          13 
        '  detlamt        hdramt      acct        sub       applied      issued    detlstatus 
        '    14            15          16          17          18          19          20  
        ' hdrdescr       remarks     class       pokey      podtkey     invckey    hdrinvoice
        '    21            22          23          24          25 
        ' detlinvoice    hdrpaid    detlpaid    chknum       payee
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this method retrieves all purchase orders for a single bank given the selected filtering criteria;
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            'collect the purchase orders for the selected criteria;
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, d.ocex_code, h.po_autoinc_key," _
            & " d.podt_autoinc_key, d.invc_autoinc_key," _
            & " 0.00 AS hdrinvoiced, 0.00 AS detlinvoiced, 0.00 AS hdrpaid, 0.00 AS detlpaid," _
            & " '' AS checknumber, '' AS checkpayee" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.po_num, d.podt_autoinc_key; "
            'collect the invoices for the selected criteria;
            SSQL &= "SELECT SUM(invc_amount), 0.00 AS hdramount, i.po_autoinc_key, i.podt_autoinc_key" _
            & " FROM invoices AS i, purc_info AS h, purc_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND i.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND (d.podt_status <> 'D')" _
            & " AND (i.invc_status <> 'V')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " GROUP BY i.po_autoinc_key, i.podt_autoinc_key" _
            & " ORDER BY i.po_autoinc_key, i.podt_autoinc_key; "
            'collect the checks for the selected criteria;
            SSQL &= "SELECT SUM(ckdt_amount), 0.00 AS hdramount, h.po_autoinc_key, c.podt_autoinc_key" _
            & " FROM chks_info AS k, chks_detl AS c, purc_info AS h, purc_detl AS d " _
            & " WHERE k.chks_autoinc_key = c.chks_autoinc_key" _
            & " AND c.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND k.chks_status <> 'V'" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.po_fisyr = @p2" _
            & " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            & " GROUP BY h.po_autoinc_key, c.podt_autoinc_key" _
            & " ORDER BY h.po_autoinc_key, c.podt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            'collect the purchase orders for the selected criteria;
            SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, v.vend_name," _
            & " h.po_status, d.podt_qty, d.podt_unitp, d.podt_amount, 0.00 AS hdramount," _
            & " d.af_acct_num, d.as_acct_num, h.po_applied_date, h.po_datetime," _
            & " d.podt_status, h.po_descr, d.podt_descr, d.ocex_code, h.po_autoinc_key," _
            & " d.podt_autoinc_key, d.invc_autoinc_key," _
            & " 0.00 AS hdrinvoiced, 0.00 AS detlinvoiced, 0.00 AS hdrpaid, 0.00 AS detlpaid," _
            & " '' AS checknumber, '' AS checkpayee" _
            & " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND h.vend_number = v.vend_number" _
            & " AND (d.podt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " ORDER BY d.bank_acct_num, d.po_fisyr, d.po_num, d.podt_autoinc_key; "
            'collect the invoices for the selected criteria;
            SSQL &= "SELECT SUM(invc_amount), 0.00 AS hdramount, i.po_autoinc_key, i.podt_autoinc_key" _
            & " FROM invoices AS i, purc_info AS h, purc_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND i.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND (d.podt_status <> 'D')" _
            & " AND (i.invc_status <> 'V')" _
            & " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " GROUP BY i.po_autoinc_key, i.podt_autoinc_key" _
            & " ORDER BY i.po_autoinc_key, i.podt_autoinc_key; "
            'collect the checks for the selected criteria;
            SSQL &= "SELECT SUM(ckdt_amount), 0.00 AS hdramount, h.po_autoinc_key, c.podt_autoinc_key" _
            & " FROM chks_info AS k, chks_detl AS c, purc_info AS h, purc_detl AS d " _
            & " WHERE k.chks_autoinc_key = c.chks_autoinc_key" _
            & " AND c.podt_autoinc_key = d.podt_autoinc_key" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.po_fisyr = d.po_fisyr" _
            & " AND h.po_num = d.po_num" _
            & " AND k.chks_status <> 'V'" _
            & " AND h.bank_acct_num = @p1" _
            & " AND h.po_fisyr = @p2" _
            & " AND h.po_num BETWEEN @p3 AND @p4" _
            & " GROUP BY h.po_autoinc_key, c.podt_autoinc_key" _
            & " ORDER BY h.po_autoinc_key, c.podt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)
        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
        Dim ds As New DataSet("register")
        Try
            da.Fill(ds)
            'verify there are records for the selected criteria;
            If ds.Tables(0).Rows.Count < 1 Then
                MsgBox("No records found for this criteria...", MsgBoxStyle.Information, MSGTITLE)
                Exit Function
            End If
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridWrk.DataSource = ds.Tables(1)
            Me.GridWrkTotals.DataSource = ds.Tables(2)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Dim chknumber, chkstatus, chkreconsw, payee, ponumber, postatus As String
        Dim holdkey, index, pokey, podtkey, row, holdpodtkey As Int32
        Dim hdramount, detlamount, tempamount As Double
        Dim balance, encumbered, invoiced, paid As Double
        Dim i, j, k As Int32
        Try
            With Me.GridDetail
                'summarise the detail amounts into the header record for the purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 17))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 17))
                        hdramount = CDbl(.GetData(j, 7))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 8, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next
            End With

            With Me.GridWrk
                'summarise the detail invoices into the header invoice for each purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 2))
                        hdramount = CDbl(.GetData(j, 0))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 1, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next

                'match the invoice amounts for each purchase order line into the detail grid;
                For index = 1 To .Rows.Count - 1
                    detlamount = CDbl(.GetData(index, 0))
                    podtkey = CInt(.GetData(index, 3))
                    row = Me.GridDetail.FindRow(podtkey.ToString, 1, 18, True, True, False)
                    If row >= 0 Then
                        Me.GridDetail.SetData(row, 21, detlamount)
                    End If
                Next

                'match the invoice amounts for each purchase order header into the detail grid;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    hdramount = CDbl(.GetData(index, 1))
                    row = Me.GridDetail.FindRow(pokey.ToString, 1, 17, True, True, False)
                    If row >= 0 Then
                        With Me.GridDetail
                            For j = row To .Rows.Count - 1
                                holdkey = CInt(.GetData(j, 17))
                                If pokey = holdkey Then
                                    Me.GridDetail.SetData(j, 20, hdramount)
                                Else
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                Next
            End With

            With Me.GridWrkTotals
                'summarise the check amounts into a header amount for each purchase order;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    For j = index To .Rows.Count - 1
                        holdkey = CInt(.GetData(j, 2))
                        hdramount = CDbl(.GetData(j, 0))
                        If pokey = holdkey Then
                            tempamount += hdramount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 1, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next
                'match the check paid amounts for each purchase order line into the detail grid;
                For index = 1 To .Rows.Count - 1
                    detlamount = CDbl(.GetData(index, 0))
                    podtkey = CInt(.GetData(index, 3))
                    row = Me.GridDetail.FindRow(podtkey.ToString, 1, 18, True, True, False)
                    If row >= 0 Then
                        Me.GridDetail.SetData(row, 23, detlamount)
                    End If
                Next

                'match the check amounts for each purchase order header into the detail grid;
                For index = 1 To .Rows.Count - 1
                    pokey = CInt(.GetData(index, 2))
                    hdramount = CDbl(.GetData(index, 1))
                    row = Me.GridDetail.FindRow(pokey.ToString, 1, 17, True, True, False)
                    If row >= 0 Then
                        With Me.GridDetail
                            For j = row To .Rows.Count - 1
                                holdkey = CInt(.GetData(j, 17))
                                If pokey = holdkey Then
                                    Me.GridDetail.SetData(j, 22, hdramount)
                                Else
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                Next

            End With
        Catch ex As Exception
            Throw
        End Try

        'calculate totals for the header totals;
        With Me.GridTotals
            .Cols.Count = 4
            .Rows.Count = 2
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    encumbered += CDbl(.GetData(index, 7))
                    invoiced += CDbl(.GetData(index, 21))
                    paid += CDbl(.GetData(index, 23))
                Next
            End With
            balance = encumbered - paid
            .SetData(1, 0, encumbered)
            .SetData(1, 1, invoiced)
            .SetData(1, 2, paid)
            .SetData(1, 3, balance)
        End With


        ''''Me.Prev1.Visible = False
        ''''Me.GridDetail.Visible = True
        ''''Me.GridWrk.Visible = True
        ''''Me.GridWrkTotals.Visible = True
        ''''Me.GridTotals.Visible = True
        ''''Me.ShowDialog()
        ''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table;
            Call PrintPurchaseOrderByInvoiceRegister(eoutstandinginvoices)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateRequisition(ByVal erequisitionkey As Int32) As Boolean
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim bankaccount, number As String
        Dim fisyr As Int32
        Try
            cn = New SqlConnection(Me.ConnectionString)
            '
            SSQL = "SELECT bank_acct_num, req_fisyr, req_num FROM req_info WHERE req_autoinc_key = @p1"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", erequisitionkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrequisition")
            da.Fill(tbl)
            With tbl
                If .Rows.Count <> 1 Then Throw New ArgumentException("Invalid record(s) were returned for the requisition.")
                bankaccount = DirectCast(.Rows(0).Item(0), String)
                fisyr = CInt(.Rows(0).Item(1))
                number = DirectCast(.Rows(0).Item(2), String)
                Call GenerateRequisitionTickets(bankaccount, fisyr, False, True, Now, Now, number, number)
            End With
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateRequisitionTickets(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edate1 As Date, ByVal edate2 As Date, ByVal enumber1 As String, ByVal enumber2 As String) As Boolean
        '''''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5         6           7  
        '   bank       fisyr    docnumber  status  reqtype   vendnum   vendname     descr   
        '     8           9        10       11       12        13        14          15       
        '  applied     issued     qty      cost    lineamt   total      acct      acctname
        '    16          17        18       19       20        21        22          23
        '   sub       subname     code   remarks   vaddr1    vaddr2    vaddr3       vcity
        '    24          25        26       27       28        29        30          31   
        '  vstate       vzip    vzipext    vph1    vph1x     reqkey    userkey      pokey     
        '    32          33        34       35     
        '  ponum      poissue    chknum  chkissue
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim cmd As SqlCommand
        Dim index, lpokey, lrequisitionkey As Int32

        If eusenumber Then
            SSQL = "SELECT h.bank_acct_num, h.req_fisyr, h.req_num, h.req_status, h.req_type," _
            & " h.vend_number, v.vend_name, h.req_descr, h.req_applied_date, h.req_datetime," _
            & " d.rqdt_qty, d.rqdt_unitp, d.rqdt_amount, 0.0 AS REQUISITIONTOTAL," _
            & " d.af_acct_num, f.af_acct_name, d.as_acct_num, s.as_acct_name, d.ocex_code, d.rqdt_descr," _
            & " v.vend_addr1, v.vend_addr2, v.vend_addr3, v.vend_city," _
            & " v.vend_state, v.vend_zip, v.vend_zip_ext, v.vend_phone1, v.vend_phone1_ext," _
            & " h.req_autoinc_key, h.user_autoinc_key, po_autoinc_key, ' ' AS ponumber, getdate() AS poissued," _
            & " ' ' AS checknumber, getdate() AS checkissued" _
            & " FROM req_info AS h, req_detl AS d, acct_info AS f, acct_sub AS s, vend_info AS v" _
            & " WHERE h.req_autoinc_key = d.req_autoinc_key" _
            & " AND h.vend_number = v.vend_number" _
            & " AND h.bank_acct_num = f.bank_acct_num" _
            & " AND f.bank_acct_num = s.bank_acct_num" _
            & " AND d.af_acct_num = s.af_acct_num" _
            & " AND d.as_acct_num = s.as_acct_num" _
            & " AND f.af_acct_num = s.af_acct_num" _
            & " AND (d.rqdt_status <> 'D')" _
            & " AND h.bank_acct_num = @p1 AND h.req_fisyr = @p2" _
            & " AND h.req_num BETWEEN @p3 AND @p4" _
            & " ORDER BY h.bank_acct_num, h.req_fisyr, h.req_num, d.rqdt_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumber1)
            cmd.Parameters.Add("@p4", enumber2)
        Else
            '''''SSQL = "SELECT h.bank_acct_num, h.po_fisyr, h.po_num, h.po_status," _
            '''''& " h.vend_number, v.vend_name, h.po_applied_date, h.po_datetime," _
            '''''& " d.podt_qty, d.podt_unitp, d.podt_amount, 0.0 AS poamount," _
            '''''& " d.af_acct_num, d.af_acct_name, d.as_acct_num, d.as_acct_name," _
            '''''& " d.ocex_code, v.vend_addr1, v.vend_addr2, v.vend_addr3, v.vend_city," _
            '''''& " v.vend_state, v.vend_zip, v.vend_zip_ext, v.vend_phone1, v.vend_phone1_ext," _
            '''''& " ' ' AS checknumber, ' ' AS checkprinted, getdate() AS checkdate," _
            '''''& " h.po_descr, d.podt_descr, 0 AS invoices" _
            '''''& " FROM purc_info AS h, purc_detl AS d, vend_info AS v" _
            '''''& " WHERE h.bank_acct_num = d.bank_acct_num" _
            '''''& " AND h.po_fisyr = d.po_fisyr" _
            '''''& " AND h.po_num = d.po_num" _
            '''''& " AND h.vend_number = v.vend_number" _
            '''''& " AND (d.podt_status <> 'D')" _
            '''''& " AND h.bank_acct_num = @p1 AND h.po_fisyr = @p2" _
            '''''& " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            '''''& " ORDER BY h.bank_acct_num, h.po_fisyr, h.po_num, d.podt_autoinc_key; "
            ''''''pull all the checks for the fiscal year;
            '''''SSQL += "SELECT chks_num, chks_printed_sw, chks_datetime, po_num" _
            '''''& " FROM chks_info" _
            '''''& " WHERE bank_acct_num = @p1" _
            '''''& " AND chks_fisyr = @p2; "
            ''''''pull the invoices;
            '''''SSQL += "SELECT i.po_num, COUNT(*) AS itemcount" _
            '''''& " FROM invoices AS i, purc_info AS h, purc_detl AS d" _
            '''''& " WHERE h.bank_acct_num = d.bank_acct_num" _
            '''''& " AND h.po_fisyr = d.po_fisyr" _
            '''''& " AND h.po_num = d.po_num" _
            '''''& " AND d.podt_autoinc_key = i.podt_autoinc_key" _
            '''''& " AND i.bank_acct_num = @p1" _
            '''''& " AND i.invc_fisyr = @p2" _
            '''''& " AND h.po_applied_date BETWEEN @p3 AND @p4" _
            '''''& " GROUP BY i.po_num" _
            '''''& " ORDER BY CAST(i.po_num AS INT)"
            '''''cmd = New SqlCommand(SSQL, cn)
            '''''cmd.Parameters.Add("@p1", ebankaccountnumber)
            '''''cmd.Parameters.Add("@p2", efiscalyear)
            '''''cmd.Parameters.Add("@p3", edate1)
            '''''cmd.Parameters.Add("@p4", edate2)
        End If

        Try

            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrequisition")
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("Missing record(s) were returned for the requisition.")
            Me.GridWrk.DataSource = tbl
            With tbl
                lrequisitionkey = CInt(.Rows(0).Item(29))
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'collect the purchase order information for all requisitions, if exists; 
            SSQL = "SELECT po_autoinc_key, po_num, po_datetime, rqst_autoinc_key" _
            & " FROM purc_info" _
            & " WHERE rqst_autoinc_key > 0 AND po_fisyr = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblpurchaseorder")
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        Try
            'collect the check information for all requisitions, if exists;
            SSQL = "SELECT chks_num, chks_datetime, c.po_num" _
            & " FROM chks_info AS c, purc_info AS p" _
            & " WHERE c.bank_acct_num = p.bank_acct_num" _
            & " AND c.po_fisyr = p.po_fisyr" _
            & " AND c.po_num = p.po_num" _
            & " AND p.po_fisyr = @p1" _
            & " AND p.rqst_autoinc_key > 0"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblchecks")
            da.Fill(tbl)
            Me.GridTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try


        Dim curkey, holdkey As Int32
        Dim lchecknumber, lponumber As String
        Dim lcheckissue, lpoissue As Date
        Dim currow, j, k, invoices, newindex As Int32
        Dim amount, tempamount As Double

        Try
            'summarise the detail amounts into a requisition total;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    'get the requisition key;
                    curkey = CInt(.GetData(index, 29))
                    For j = index To .Rows.Count - 1
                        'iterate the grid & compare keys;
                        holdkey = CInt(.GetData(j, 29))
                        amount = CDbl(.GetData(j, 12))
                        If curkey = holdkey Then
                            tempamount += amount
                        Else
                            Exit For
                        End If
                    Next
                    'backfill the amounts;
                    For k = index To j - 1
                        .SetData(k, 13, tempamount)
                    Next
                    '
                    index = j - 1
                    tempamount = 0
                Next
            End With
        Catch ex As Exception
            Throw
        End Try


        Try
            'set the purchase order information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lpokey = CInt(.GetData(index, 31))
                    If lpokey > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridWrkTotals.FindRow(lpokey.ToString, 1, 0, True, True, False)
                        If newindex > 0 Then
                            lponumber = DirectCast(Me.GridWrkTotals.GetData(newindex, 1), String)
                            lpoissue = CDate(Me.GridWrkTotals.GetData(newindex, 2))
                            .SetData(index, 32, lponumber)
                            .SetData(index, 33, lpoissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the check information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lponumber = DirectCast(.GetData(index, 32), String).Trim
                    If lponumber.Length > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridTotals.FindRow(lponumber, 1, 2, True, True, False)
                        If newindex > 0 Then
                            lchecknumber = DirectCast(Me.GridTotals.GetData(newindex, 0), String)
                            lcheckissue = CDate(Me.GridTotals.GetData(newindex, 1))
                            .SetData(index, 34, lchecknumber)
                            .SetData(index, 35, lcheckissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.GridWrkTotals.Visible = False
        '''''Me.GridTotals.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table;
            Call PrintRequisitionTickets()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateReviewerRequisitionReport(ByVal efiscalyear As Int32, ByVal ereviewerkey As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5         6           7  
        '   bank       fisyr    docnumber  status  reqtype   vendnum   vendname     descr   
        '     8           9        10       11       12        13        14          15       
        '  applied     issued     qty      cost    lineamt   total      acct      acctname
        '    16          17        18       19       20        21        22          23
        '   sub       subname     code   remarks   vaddr1    vaddr2    vaddr3       vcity
        '    24          25        26       27       28        29        30          31   
        '  vstate       vzip    vzipext    vph1    vph1x     reqkey    userkey      pokey     
        '    32          33        34       35     
        '  ponum      poissue    chknum  chkissue
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim cmd As SqlCommand
        Dim index, lpokey, lreviewflag As Int32
        Dim lreviewstr, lreviewername As String

        Try
            'get the user name of the requisitions;
            SSQL = "SELECT rapv_type, user_fullname" _
            & " FROM req_approver AS a, user_info AS u" _
            & " WHERE a.user_autoinc_key = u.user_autoinc_key" _
            & " AND rapv_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ereviewerkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrequisition")
            da.Fill(tbl)
            If tbl.Rows.Count <> 1 Then Throw New ArgumentException("The converter/reviewer information could not be retrieved.")
            With tbl
                lreviewstr = DirectCast(.Rows(0).Item(0), String).ToUpper
                If lreviewstr = "A" Then lreviewflag = 1 Else lreviewflag = 0
                lreviewername = DirectCast(.Rows(0).Item(1), String).Trim
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        Try
            'collect the requisitions;
            SSQL = "SELECT h.bank_acct_num, h.req_fisyr, h.req_num, h.req_status, h.req_type," _
            & " h.vend_number, v.vend_name, h.req_descr, h.req_applied_date, h.req_datetime," _
            & " d.rqdt_qty, d.rqdt_unitp, d.rqdt_amount, 0.0 AS REQUISITIONTOTAL," _
            & " d.af_acct_num, f.af_acct_name, d.as_acct_num, s.as_acct_name, d.ocex_code, d.rqdt_descr," _
            & " v.vend_addr1, v.vend_addr2, v.vend_addr3, v.vend_city," _
            & " v.vend_state, v.vend_zip, v.vend_zip_ext, v.vend_phone1, v.vend_phone1_ext," _
            & " h.req_autoinc_key, h.user_autoinc_key, po_autoinc_key, ' ' AS ponumber, getdate() AS poissued," _
            & " ' ' AS checknumber, getdate() AS checkissued" _
            & " FROM req_info AS h, req_detl AS d, acct_info AS f, acct_sub AS s, vend_info AS v" _
            & " WHERE h.req_autoinc_key = d.req_autoinc_key" _
            & " AND h.vend_number = v.vend_number" _
            & " AND h.bank_acct_num = f.bank_acct_num" _
            & " AND f.bank_acct_num = s.bank_acct_num" _
            & " AND d.af_acct_num = s.af_acct_num" _
            & " AND d.as_acct_num = s.as_acct_num" _
            & " AND f.af_acct_num = s.af_acct_num" _
            & " AND (d.rqdt_status <> 'D')" _
            & " AND h.req_fisyr = @p1" _
            & " AND h.req_autoinc_key IN" _
            & " (SELECT req_autoinc_key FROM req_queue WHERE rapv_autoinc_key = @p2)" _
            & " ORDER BY h.bank_acct_num, h.req_fisyr, h.req_num, d.rqdt_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            cmd.Parameters.Add("@p2", ereviewerkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrequisition")
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No record(s) were found for the requisition converter or reviewer.")
            Me.GridWrk.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'collect the purchase order information for all requisitions, if exists; 
            SSQL = "SELECT po_autoinc_key, po_num, po_datetime, rqst_autoinc_key" _
            & " FROM purc_info" _
            & " WHERE rqst_autoinc_key > 0 AND po_fisyr = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblpurchaseorder")
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'collect the check information for all requisitions, if exists;
            SSQL = "SELECT chks_num, chks_datetime, c.po_num" _
            & " FROM chks_info AS c, purc_info AS p" _
            & " WHERE c.bank_acct_num = p.bank_acct_num" _
            & " AND c.po_fisyr = p.po_fisyr" _
            & " AND c.po_num = p.po_num" _
            & " AND p.po_fisyr = @p1" _
            & " AND p.rqst_autoinc_key > 0"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblchecks")
            da.Fill(tbl)
            Me.GridTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        Dim curkey, holdkey As Int32
        Dim lchecknumber, lponumber As String
        Dim lcheckissue, lpoissue As Date
        Dim currow, j, k, invoices, newindex As Int32
        Dim amount, tempamount As Double

        Try
            'summarise the detail amounts into a requisition total;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    'get the requisition key;
                    curkey = CInt(.GetData(index, 29))
                    For j = index To .Rows.Count - 1
                        'iterate the grid & compare keys;
                        holdkey = CInt(.GetData(j, 29))
                        amount = CDbl(.GetData(j, 12))
                        If curkey = holdkey Then
                            tempamount += amount
                        Else
                            Exit For
                        End If
                    Next
                    'backfill the amounts;
                    For k = index To j - 1
                        .SetData(k, 13, tempamount)
                    Next
                    '
                    index = j - 1
                    tempamount = 0
                Next
            End With
        Catch ex As Exception
            Throw
        End Try


        Try
            'set the purchase order information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lpokey = CInt(.GetData(index, 31))
                    If lpokey > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridWrkTotals.FindRow(lpokey.ToString, 1, 0, True, True, False)
                        If newindex > 0 Then
                            lponumber = DirectCast(Me.GridWrkTotals.GetData(newindex, 1), String)
                            lpoissue = CDate(Me.GridWrkTotals.GetData(newindex, 2))
                            .SetData(index, 32, lponumber)
                            .SetData(index, 33, lpoissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the check information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lponumber = DirectCast(.GetData(index, 32), String).Trim
                    If lponumber.Length > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridTotals.FindRow(lponumber, 1, 2, True, True, False)
                        If newindex > 0 Then
                            lchecknumber = DirectCast(Me.GridTotals.GetData(newindex, 0), String)
                            lcheckissue = CDate(Me.GridTotals.GetData(newindex, 1))
                            .SetData(index, 34, lchecknumber)
                            .SetData(index, 35, lcheckissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.GridWrkTotals.Visible = False
        '''''Me.GridTotals.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table;
            Call PrintReviewerRequisitionReport(lreviewername, lreviewflag)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateUserRequisitionReport(ByVal efiscalyear As Int32, ByVal euserkey As Int32) As Boolean
        '''''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5         6           7  
        '   bank       fisyr    docnumber  status  reqtype   vendnum   vendname     descr   
        '     8           9        10       11       12        13        14          15       
        '  applied     issued     qty      cost    lineamt   total      acct      acctname
        '    16          17        18       19       20        21        22          23
        '   sub       subname     code   remarks   vaddr1    vaddr2    vaddr3       vcity
        '    24          25        26       27       28        29        30          31   
        '  vstate       vzip    vzipext    vph1    vph1x     reqkey    userkey      pokey     
        '    32          33        34       35     
        '  ponum      poissue    chknum  chkissue
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim cmd As SqlCommand
        Dim index, lpokey As Int32
        Dim lusername As String

        Try
            Dim obj As Object
            'get the user name of the requisitions;
            SSQL = "SELECT user_fullname FROM user_info WHERE user_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", euserkey)
            cn.Open()
            obj = cmd.ExecuteScalar
            'test for bad userkey;
            If obj Is Nothing Then Throw New ArgumentException("The user key is invalid or missing.")
            lusername = DirectCast(obj, String)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            cmd.Dispose()
        End Try


        Try
            'collect the requisitions;
            SSQL = "SELECT h.bank_acct_num, h.req_fisyr, h.req_num, h.req_status, h.req_type," _
            & " h.vend_number, v.vend_name, h.req_descr, h.req_applied_date, h.req_datetime," _
            & " d.rqdt_qty, d.rqdt_unitp, d.rqdt_amount, 0.0 AS REQUISITIONTOTAL," _
            & " d.af_acct_num, f.af_acct_name, d.as_acct_num, s.as_acct_name, d.ocex_code, d.rqdt_descr," _
            & " v.vend_addr1, v.vend_addr2, v.vend_addr3, v.vend_city," _
            & " v.vend_state, v.vend_zip, v.vend_zip_ext, v.vend_phone1, v.vend_phone1_ext," _
            & " h.req_autoinc_key, h.user_autoinc_key, po_autoinc_key, ' ' AS ponumber, getdate() AS poissued," _
            & " ' ' AS checknumber, getdate() AS checkissued" _
            & " FROM req_info AS h, req_detl AS d, acct_info AS f, acct_sub AS s, vend_info AS v" _
            & " WHERE h.req_autoinc_key = d.req_autoinc_key" _
            & " AND h.vend_number = v.vend_number" _
            & " AND h.bank_acct_num = f.bank_acct_num" _
            & " AND f.bank_acct_num = s.bank_acct_num" _
            & " AND d.af_acct_num = s.af_acct_num" _
            & " AND d.as_acct_num = s.as_acct_num" _
            & " AND f.af_acct_num = s.af_acct_num" _
            & " AND (d.rqdt_status <> 'D')" _
            & " AND user_autoinc_key = @p1" _
            & " AND h.req_fisyr = @p2" _
            & " ORDER BY h.bank_acct_num, h.req_fisyr, h.req_num, d.rqdt_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", euserkey)
            cmd.Parameters.Add("@p2", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrequisition")
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No record(s) were found for the requisition user.")
            Me.GridWrk.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'collect the purchase order information for all requisitions, if exists; 
            SSQL = "SELECT po_autoinc_key, po_num, po_datetime, rqst_autoinc_key" _
            & " FROM purc_info WHERE rqst_autoinc_key > 0 AND po_fisyr = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblpurchaseorder")
            da.Fill(tbl)
            Me.GridWrkTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'collect the check information for all requisitions, if exists;
            SSQL = "SELECT chks_num, chks_datetime, c.po_num" _
            & " FROM chks_info AS c, purc_info AS p" _
            & " WHERE c.bank_acct_num = p.bank_acct_num" _
            & " AND c.po_fisyr = p.po_fisyr" _
            & " AND c.po_num = p.po_num" _
            & " AND p.po_fisyr = @p1" _
            & " AND p.rqst_autoinc_key > 0"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblchecks")
            da.Fill(tbl)
            Me.GridTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        Dim curkey, holdkey As Int32
        Dim lchecknumber, lponumber As String
        Dim lcheckissue, lpoissue As Date
        Dim currow, j, k, invoices, newindex As Int32
        Dim amount, tempamount As Double

        Try
            'summarise the detail amounts into a requisition total;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    'get the requisition key;
                    curkey = CInt(.GetData(index, 29))
                    For j = index To .Rows.Count - 1
                        'iterate the grid & compare keys;
                        holdkey = CInt(.GetData(j, 29))
                        amount = CDbl(.GetData(j, 12))
                        If curkey = holdkey Then
                            tempamount += amount
                        Else
                            Exit For
                        End If
                    Next
                    'backfill the amounts;
                    For k = index To j - 1
                        .SetData(k, 13, tempamount)
                    Next
                    '
                    index = j - 1
                    tempamount = 0
                Next
            End With
        Catch ex As Exception
            Throw
        End Try


        Try
            'set the purchase order information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lpokey = CInt(.GetData(index, 31))
                    If lpokey > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridWrkTotals.FindRow(lpokey.ToString, 1, 0, True, True, False)
                        If newindex > 0 Then
                            lponumber = DirectCast(Me.GridWrkTotals.GetData(newindex, 1), String)
                            lpoissue = CDate(Me.GridWrkTotals.GetData(newindex, 2))
                            .SetData(index, 32, lponumber)
                            .SetData(index, 33, lpoissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            'set the check information, if exists;
            With Me.GridWrk
                For index = 1 To .Rows.Count - 1
                    lponumber = DirectCast(.GetData(index, 32), String).Trim
                    If lponumber.Length > 0 Then
                        'search the po grid for a matching po key from the requisition grid;
                        newindex = Me.GridTotals.FindRow(lponumber, 1, 2, True, True, False)
                        If newindex > 0 Then
                            lchecknumber = DirectCast(Me.GridTotals.GetData(newindex, 0), String)
                            lcheckissue = CDate(Me.GridTotals.GetData(newindex, 1))
                            .SetData(index, 34, lchecknumber)
                            .SetData(index, 35, lcheckissue)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.GridWrkTotals.Visible = False
        '''''Me.GridTotals.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table;
            Call PrintUserRequisitionReport(lusername)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateUserRequisitionRejectionReport(ByVal efiscalyear As Int32, ByVal euserkey As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Prints a list of requisitions for the user that have been rejected by the converter
        'or reviewer; Requested by Vicki@Western Heights on 2014.05.10;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5         6           7  
        '  reqkey      reqnum   reqissue  amount  reviewer   reqtype  reqlevel    approval
        '     8           9        10       11       12        13        14          15       
        '  comment     comdate  
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL, mUserName As String
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim cmd As SqlCommand
        Dim index, requisitionkey As Int32

        Try
            'get the user name of the requisitions;
            SSQL = "SELECT user_fullname FROM user_info WHERE user_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", euserkey)
            cn.Open()
            mUserName = DirectCast(cmd.ExecuteScalar, String)
            cn.Close()
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
            cmd.Dispose()
        End Try


        Try
            'collect any non-approved req with a rejection flag;
            SSQL = "SELECT q.req_autoinc_key, req_num, req_datetime, 0.00 AS Amount, user_fullname," _
            & " rque_type, rque_level, rque_approved, rcom_comments, rcom_datetime" _
            & " FROM req_approver AS a, req_comments AS c, req_info AS r, req_queue AS q, user_info AS u" _
            & " WHERE a.rapv_autoinc_key = q.rapv_autoinc_key " _
            & " AND q.req_autoinc_key = c.req_autoinc_key" _
            & " AND q.req_autoinc_key = r.req_autoinc_key" _
            & " AND q.rapv_autoinc_key = c.rapv_autoinc_key" _
            & " AND a.user_autoinc_key = u.user_autoinc_key" _
            & " AND rque_approved = 0 " _
            & " AND q.req_autoinc_key IN (SELECT req_autoinc_key FROM req_info WHERE req_fisyr = @p1 AND user_autoinc_key = @p2)" _
            & " AND q.req_autoinc_key IN (SELECT req_autoinc_key FROM req_queue WHERE rque_level < 0)" _
            & " ORDER BY q.req_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            cmd.Parameters.Add("@p2", euserkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("tblrejections")
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then Return False
            Me.GridWrk.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'iterate thru the gridwrk and set the amount field;
            With Me.GridWrk
                Dim prequisitionkey As Int32
                Dim amount As Decimal
                cn = New SqlConnection(Me.ConnectionString)
                For index = 1 To .Rows.Count - 1
                    'get the key for the amount lookup;
                    requisitionkey = DirectCast(.GetData(index, 0), Int32)
                    'if new key, then perform lookup;
                    If requisitionkey <> prequisitionkey Then
                        'calculate the amount;
                        SSQL = "SELECT ISNULL(SUM(rqdt_amount), 0) AS Amount FROM req_detl" _
                        & " WHERE rqdt_status = 'O' and req_autoinc_key = @p1"
                        cmd = New SqlCommand(SSQL, cn)
                        cmd.Parameters.Add("@p1", requisitionkey)
                        If cn.State <> ConnectionState.Open Then cn.Open()
                        amount = DirectCast(cmd.ExecuteScalar, Decimal)
                    End If
                    'set the amount in the grid;
                    .SetData(index, 3, amount)
                    'remember the key;
                    prequisitionkey = requisitionkey
                Next
                cn.Close()
            End With
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
            cmd.Dispose()
        End Try

        '''''Me.Prev1.Visible = False
        '''''Me.GridWrk.Visible = True
        '''''Me.GridWrkTotals.Visible = False
        '''''Me.GridTotals.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table;
            Call PrintUserRequisitionRejectionReport(mUserName)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateVoidCheckRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'retrieves voided checks for a single bank given the selected filtering criteria;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   printed
        '     6           7         8        9       10       11  
        '  payee      ponumber   hdramt   lineamt   acct     sub 
        '    12          13        14       15       16       17  
        '  applied    created   hdrdescr  remarks voidappl voidissue
        '    18
        ' voidremark
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable

        Try
            If eusedate Then
                Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
                SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
                & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, h.po_num," _
                & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
                & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr," _
                & " v.voidchk_applied_date, v.voidchk_datetime, v.voidchk_remarks" _
                & " FROM chks_info AS h, chks_detl AS d, voidcheck AS v" _
                & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
                & " AND h.chks_autoinc_key = v.chks_autoinc_key" _
                & " AND d.ckdt_autoinc_key = v.ckdt_autoinc_key" _
                & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
                & " AND v.voidchk_applied_date BETWEEN @p3 AND @p4" _
                & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", ebankaccountnumber)
                cmd.Parameters.Add("@p2", efiscalyear)
                cmd.Parameters.Add("@p3", edatefrom)
                cmd.Parameters.Add("@p4", edateto)
            End If
            If eusenumber Then
                Me.CellMiddleBottom = enumberfrom & " to " & enumberto
                SSQL = "SELECT h.bank_acct_num, h.chks_fisyr, h.chks_num, h.chks_status," _
                & " h.chks_recon_sw, h.chks_printed_sw, h.chks_payee_name, h.po_num," _
                & " h.chks_amount, d.ckdt_amount, d.af_acct_num, d.as_acct_num," _
                & " h.chks_applied_date, h.chks_datetime, h.chks_descr, d.ckdt_descr," _
                & " v.voidchk_applied_date, v.voidchk_datetime, v.voidchk_remarks" _
                & " FROM chks_info AS h, chks_detl AS d, voidcheck AS v" _
                & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
                & " AND h.chks_autoinc_key = v.chks_autoinc_key" _
                & " AND d.ckdt_autoinc_key = v.ckdt_autoinc_key" _
                & " AND h.bank_acct_num = @p1 AND h.chks_fisyr = @p2" _
                & " AND h.chks_num BETWEEN @p3 AND @p4" _
                & " ORDER BY h.bank_acct_num, h.chks_fisyr, h.chks_num, d.ckdt_autoinc_key"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", ebankaccountnumber)
                cmd.Parameters.Add("@p2", efiscalyear)
                cmd.Parameters.Add("@p3", enumberfrom)
                cmd.Parameters.Add("@p4", enumberto)
            End If
            tbl = New DataTable("register")
            da = New SqlDataAdapter(cmd)
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
            'summarise the total register amount
            Dim amount As Double
            Dim index As Int32
            With Me.GridDetail
                For index = 1 To .Rows.Count - 1
                    amount += CDbl(.GetData(index, 9))
                Next
            End With
            With Me.GridTotals
                .Rows.Count = 1
                .Rows.Add()
                .SetData(0, 0, amount)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintVoidCheckRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region "  Methods Rendering "

    Private Sub PrintCheckRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   printed
        '     6           7         8        9       10       11  
        '  payee      ponumber   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   hdrdescr  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "CheckRegister"
        Me.ReportName = "Check Register"
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

        'this style is only used by this report
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            .TextColor = Color.Gray
        End With

        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, count As Int32
        Dim totalregister, voidregister, checkregister As Double
        Dim tissuedate, tapplieddate As Date
        Dim tacctnum, tsubacctnum, tchknum, tpayee, tdescr, tremarks, tponumber As String
        Dim tprinted, tstatus, trecon, prevchknum, prtstatus As String
        Dim tchkamt, tlineamt, sumamount As Double

        Try
            'get the amount of the register
            checkregister = CDbl(Me.GridTotals.GetData(0, 0))
            'get the voids
            voidregister = CDbl(Me.GridTotals.GetData(0, 1))
            'calc the register amount
            totalregister = checkregister - voidregister
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
                        .RenderDirectText(1, 36, "Status Key:", 40, 5, verdanaleft8bold)
                        .RenderDirectText(3, 41, "1 - Cleared", 40, 4, specstyle)
                        .RenderDirectText(3, 44, "2 - Outstanding", 40, 4, specstyle)
                        .RenderDirectText(3, 47, "3 - Open", 40, 4, specstyle)
                        .RenderDirectText(3, 50, "4 - Void", 40, 4, specstyle)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y + 4, "Check register:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Less voids:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 15, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y + 4, checkregister.ToString.Format("{0:C2}", checkregister), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, voidregister.ToString.Format("{0:C2}", voidregister), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 15, totalregister.ToString.Format("{0:C2}", totalregister), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(123, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        tchknum = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tprinted = DirectCast(.GetData(index, 5), String)
                        tpayee = DirectCast(.GetData(index, 6), String)
                        tponumber = DirectCast(.GetData(index, 7), String)
                        tchkamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tacctnum = DirectCast(.GetData(index, 10), String)
                        tsubacctnum = DirectCast(.GetData(index, 11), String)
                        tapplieddate = CDate(.GetData(index, 12))
                        tissuedate = CDate(.GetData(index, 13))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        If tstatus = "V" Then
                            tdescr = "VOID"
                            tchkamt = 0.0
                        Else
                            sumamount += tlineamt
                        End If
                    End With

                    Select Case tstatus
                        Case "O"
                            prtstatus = "3"
                        Case "C"    'closed, now check if recon is on or off
                            If trecon = "Y" Then
                                prtstatus = "1"
                            Else
                                prtstatus = "2"
                            End If
                        Case "F"    'cleared & closedout
                            prtstatus = "1"
                        Case "V"
                            prtstatus = "4"
                            tpayee = "** VOID **"
                        Case Else
                            prtstatus = "0"
                    End Select

                    If tchknum <> prevchknum Then
                        count += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(-1, y, prtstatus, 5, 5, specstyle)
                        .RenderDirectText(1, y, tchknum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(38, y, tpayee, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, tchkamt.ToString.Format("{0:F2}", tchkamt), 25, 5, verdanaright8)
                    End If
                    .RenderDirectText(82, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(96, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(118, y, tponumber, 18, 5, verdanaleft8)
                    .RenderDirectText(136, y, tremarks, 34, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevchknum = tchknum

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(123, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Expenditures", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Checks", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, count.ToString.Format("{0:D2}", count), 25, 5, verdanaright8bold)
                y += 10
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

    Private Sub PrintInvoicePending()
        Me.DocumentName = "Invoice"
        Me.ReportName = "Outstanding Invoices"
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

        'this style is only used by this report
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            .TextColor = Color.Gray
        End With

        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, linecount, records, x, y As Int32
        Dim ckdtkey, invckey, pokey, podtkey, nextpokey, prevpokey As Int32
        Dim totalamount, tamount, tpoamount As Double
        Dim hdramount As Double
        Dim hdrcount As Int32
        Dim tissuedate, tapplieddate As Date
        Dim tacctnum, tsubacctnum, tinvoicenumber, tponumber As String
        Dim taccount, tclassification, tstatus, tvendornumber, hvendorname, tvendorname As String

        With Me.GridDetail
            'collect header page information;
            For index = 1 To .Rows.Count - 1
                Me.BankAccountNumber = DirectCast(.GetData(1, 0), String)
                hdramount += CDbl(.GetData(index, 4))
                hdrcount += 1
            Next
        End With

        Try
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 1 Then
                        'print the total info box left-side
                        .RenderDirectText(0, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(0, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y + 4, "Total outstanding:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 9, "Total invoices:", 40, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, hdramount.ToString.Format("{0:C2}", hdramount), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 9, hdrcount.ToString.Format("{0:D2}", hdrcount), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Vendor name", 30, 5, verdanaleft8bold)
                        .RenderDirectText(89, y, "Account", 20, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "PO#", 10, 5, verdanaleft8bold)
                        .RenderDirectText(134, y, "Invoice#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        tinvoicenumber = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        tamount = CDbl(.GetData(index, 4))
                        tponumber = DirectCast(.GetData(index, 5), String)
                        tacctnum = DirectCast(.GetData(index, 6), String)
                        tsubacctnum = DirectCast(.GetData(index, 7), String)
                        taccount = tacctnum & "-" & tsubacctnum
                        tclassification = DirectCast(.GetData(index, 10), String)
                        tvendornumber = DirectCast(.GetData(index, 11), String)
                        tvendorname = DirectCast(.GetData(index, 12), String)
                        tapplieddate = CDate(.GetData(index, 13))
                        tissuedate = CDate(.GetData(index, 14))
                        invckey = CInt(.GetData(index, 15))
                        pokey = CInt(.GetData(index, 16))
                        podtkey = CInt(.GetData(index, 17))
                        ckdtkey = CInt(.GetData(index, 18))
                        'verify this is an outstanding invoice, again;
                        If tstatus <> "O" Then Throw New ArgumentException("Invalid status for invoice pending.")
                        'sum the amounts;
                        totalamount += tamount
                        'get the next pokey;
                        If index < (.Rows.Count - 1) Then nextpokey = CInt(.GetData(index + 1, 16)) Else nextpokey = 0
                        records += 1
                    End With

                    linecount += 1

                    'get a running total of the po;
                    tpoamount += tamount

                    If pokey = nextpokey Then
                        If tvendorname <> hvendorname Then
                            .RenderDirectText(20, y, tvendorname, 50, 5, verdanaleft8)
                            hvendorname = tvendorname
                            'Else

                        End If

                        '.RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        .RenderDirectText(-2, y, tissuedate.ToString.Format("{0:MM/dd/yyyy}", tissuedate), 20, 5, verdanaright8)
                        If tvendorname = hvendorname Then
                            .RenderDirectText(90, y, taccount, 20, 5, verdanaright8)
                            .RenderDirectText(110, y, tponumber, 20, 5, verdanaright8)
                            .RenderDirectText(134, y, tinvoicenumber, 30, 10, verdanaleft8)
                            .RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        Else
                            .RenderDirectText(20, y, tvendorname, 50, 5, verdanaleft8)
                        End If

                        y += 2
                    Else
                        'the next purchase order is different so handle the line if single or multiline po;
                        'If linecount = 1 Then
                        If tvendorname = hvendorname Then
                            .RenderDirectText(-2, y, tissuedate.ToString.Format("{0:MM/dd/yyyy}", tissuedate), 20, 5, verdanaright8)
                            .RenderDirectText(90, y, taccount, 20, 5, verdanaright8)
                            .RenderDirectText(110, y, tponumber, 20, 5, verdanaright8)
                            .RenderDirectText(134, y, tinvoicenumber, 30, 10, verdanaleft8)
                            .RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        Else
                            .RenderDirectText(20, y, tvendorname, 50, 5, verdanaleft8)
                            .RenderDirectText(-2, y, tissuedate.ToString.Format("{0:MM/dd/yyyy}", tissuedate), 20, 5, verdanaright8)
                            .RenderDirectText(90, y, taccount, 20, 5, verdanaright8)
                            .RenderDirectText(110, y, tponumber, 20, 5, verdanaright8)
                            .RenderDirectText(134, y, tinvoicenumber, 30, 10, verdanaleft8)
                            .RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        End If

                        y += 2
                        'Else
                        '.RenderDirectText(165, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        y += 5
                        .RenderDirectText(145, y, "Total:", 20, 5, verdanaright8bold)
                        .RenderDirectText(165, y, tpoamount.ToString.Format("{0:F2}", tpoamount), 25, 5, verdanaright8bold)
                        y += 5
                        'End If
                        linecount = 0
                        tpoamount = 0
                        hvendorname = tvendorname
                    End If

                    y += 5

                    If y >= 240 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(0, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(0, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Vendor name", 30, 5, verdanaleft8bold)
                        .RenderDirectText(89, y, "Account", 20, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "PO#", 10, 5, verdanaleft8bold)
                        .RenderDirectText(134, y, "Invoice#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                    prevpokey = pokey
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Outstanding", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, totalamount.ToString.Format("{0:C2}", totalamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 5, "Total Invoices", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 5, records.ToString.Format("{0:D2}", records), 25, 5, verdanaright8bold)
                y += 11
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

    Private Sub PrintOutstandingChecksRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   printed
        '     6           7         8        9       10       11  
        '  payee      ponumber   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   hdrdescr  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "OutstandingChecks"
        Me.ReportName = "Outstanding Checks"
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

        Dim index, currow, x, y, fisyr, holdfisyr, count As Int32
        Dim totalregister As Double
        Dim tissuedate, tapplieddate As Date
        Dim tacctnum, tsubacctnum, tchknum, tpayee, tdescr, tremarks, tponumber As String
        Dim tprinted, tstatus, trecon, prevchknum, prtstatus As String
        Dim tchkamt, tlineamt, sumamount As Double

        Try
            'get the total amount of the register
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
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(122, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        fisyr = CInt(.GetData(index, 1))
                        tchknum = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tprinted = DirectCast(.GetData(index, 5), String)
                        tpayee = DirectCast(.GetData(index, 6), String)
                        tponumber = DirectCast(.GetData(index, 7), String)
                        tchkamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tacctnum = DirectCast(.GetData(index, 10), String)
                        tsubacctnum = DirectCast(.GetData(index, 11), String)
                        tapplieddate = CDate(.GetData(index, 12))
                        tissuedate = CDate(.GetData(index, 13))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        sumamount += tlineamt
                    End With
                    'check for check number break
                    If tchknum <> prevchknum Then
                        count += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(1, y, tchknum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissuedate.ToString.Format("{0:MM/dd/yyyy}", tissuedate), 20, 5, verdanaright8)
                        .RenderDirectText(38, y, tpayee, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, tchkamt.ToString.Format("{0:F2}", tchkamt), 25, 5, verdanaright8)
                    End If
                    If tstatus = "X" Then
                        'legacy check
                        .RenderDirectText(82, y, "", 20, 5, verdanaleft8)
                        tremarks = "Legacy check"
                    Else
                        .RenderDirectText(82, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    End If
                    .RenderDirectText(96, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(118, y, tponumber, 18, 5, verdanaleft8)
                    .RenderDirectText(136, y, tremarks, 34, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevchknum = tchknum
                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(122, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next
                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Expenditures", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Checks", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, count.ToString.Format("{0:D2}", count), 25, 5, verdanaright8bold)
                y += 10
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

    Private Sub PrintPositivePay()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Prints a positive pay file, which is a list of checks by register number;
        'Added on 2016.08.01, Fred;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0        1         2         3         4          5          6         7
        '   bank    fisyr     number    status     payee     amount     issued   register
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "PositivePay"
        Me.ReportName = "Positive Pay File"
        'define styles;
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
        'define the styles;
        DefineStyles()

        'this style is only used by this report;
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            .TextColor = Color.Gray
        End With

        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, countccc As Int32
        Dim check, pregister, register As Int32
        Dim totalchecks, totalregisters As Int32
        Dim payee, status As String
        Dim issued As Date
        Dim amount As Decimal
        Dim dovoid As Boolean

        Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Collect some totals from the detail to print at top of page;
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For index = 1 To Me.GridDetail.Rows.Count - 1
                With Me.GridDetail
                    register = CType(.GetData(index, 7), Int32)
                    If register <> pregister Then totalregisters += 1
                    totalchecks += 1
                    pregister = register
                End With
            Next
            'get the bank account number from the first line;
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
                        'print the bank info left-side;
                        .RenderDirectText(0, 34, "For Bank Account:", 37, 5, verdanaright8bold)
                        .RenderDirectText(0, 38, Me.BankAccountNumber, 37, 5, verdanaright8)
                        'Print the summary info on right side (Only on first page);
                        .RenderDirectText(140, 34, "Total registers:", 30, 5, verdanaright8bold)
                        .RenderDirectText(168, 34, String.Format("{0:D2}", totalregisters), 20, 5, verdanaright8bold)
                        .RenderDirectText(140, 39, "Total checks:", 30, 5, verdanaright8bold)
                        .RenderDirectText(168, 39, String.Format("{0:D2}", totalchecks), 20, 5, verdanaright8bold)
                        'print line above the column headers;
                        .RenderDirectLine(0, 46, 190, 46, Color.Gray, 0.5)
                        y = 50
                        'print the column headers;
                        .RenderDirectText(0, y, "Register", 18, 5, verdanaright8bold)
                        .RenderDirectText(17, y, "Number", 20, 5, verdanaright8bold)
                        .RenderDirectText(44, y, "Issued", 20, 5, verdanaright8bold)
                        .RenderDirectText(70, y, "Amount", 20, 5, verdanaright8bold)
                        .RenderDirectText(96, y, "Payee", 25, 5, verdanaleft8bold)
                        y = 57
                    End If

                    'Collect the detail;
                    With Me.GridDetail
                        check = CType(.GetData(index, 2), Int32)
                        status = CType(.GetData(index, 3), String)
                        payee = CType(.GetData(index, 4), String).Trim
                        amount = CType(.GetData(index, 5), Decimal)
                        issued = CType(.GetData(index, 6), Date)
                        register = CType(.GetData(index, 7), Int32)
                    End With

                    'Check for void switch;
                    If status = "V" Then dovoid = True

                    'Render the detail;
                    If dovoid = False Then
                        .RenderDirectText(0, y, String.Format("{0:D2}", register), 15, 5, verdanaright8)
                        .RenderDirectText(16, y, String.Format("{0:D2}", check), 20, 5, verdanaright8)
                        .RenderDirectText(44, y, issued.ToShortDateString, 25, 10, verdanaright8)
                        .RenderDirectText(70, y, String.Format("{0:F2}", amount), 20, 5, verdanaright8)
                        .RenderDirectText(96, y, payee, 94, 5, verdanaleft8)
                    Else
                        'Highlight void check;
                        .RenderDirectText(0, y, String.Format("{0:D2}", register), 15, 5, verdanaright8)
                        .RenderDirectText(16, y, String.Format("{0:D2}", check), 20, 5, verdanaright8bold)
                        .RenderDirectText(36, y, "******", 20, 5, verdanaleft8bold)
                        .RenderDirectText(44, y, issued.ToShortDateString, 25, 10, verdanaright8)
                        .RenderDirectText(70, y, String.Format("{0:F2}", amount), 20, 5, verdanaright8bold)
                        .RenderDirectText(96, y, "***VOID*** " + payee, 94, 5, verdanaleft8bold)
                        'Turn off void switch;
                        dovoid = False
                    End If

                    'CRLF;
                    y += 7

                    If y >= 256 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the bank info left-side;
                        .RenderDirectText(0, 34, "For Bank Account:", 37, 5, verdanaright8bold)
                        .RenderDirectText(0, 38, Me.BankAccountNumber, 37, 5, verdanaright8)
                        'print line above the column headers;
                        .RenderDirectLine(0, 46, 190, 46, Color.Gray, 0.5)
                        y = 50
                        'print the column headers;
                        .RenderDirectText(0, y, "Register", 18, 5, verdanaright8bold)
                        .RenderDirectText(17, y, "Number", 20, 5, verdanaright8bold)
                        .RenderDirectText(44, y, "Issued", 20, 5, verdanaright8bold)
                        .RenderDirectText(70, y, "Amount", 20, 5, verdanaright8bold)
                        .RenderDirectText(96, y, "Payee", 25, 5, verdanaleft8bold)
                        y = 57
                        currow = 0
                    End If
                Next
                'Only print this footer line if enough room on page, do not page break for this line only;
                If y < 256 Then .RenderDirectText(96, y + 2, "***** E N D   O F   F I L E *****", 95, 5, verdanaleft8bold)
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

    Private Sub PrintPurchaseOrder(ByVal epage As Int32)
        'define styles;
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        specstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        With timesleft16
            'used for the continuation pages;
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Center
            .TextColor = Color.Salmon
        End With

        With specstyle
            'this style is used for the Conditions box printed on the purchase order;
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 6, FontStyle.Regular)
            .TextAlignHorz = AlignHorzEnum.Justify
            .TextColor = Color.Gray
        End With

        Dim cond1 As String = "1. Invoices to be rendered in duplicate."
        Dim cond2 As String = "2. No payment to be made until order is complete."
        Dim cond3 As String = "3. Goods to be delivered F.O.B. as per address in upper left."
        Dim cond4 As String = "4. Exempt from sales tax per state statute."
        Dim cond5 As String = "5. Deliveries acknowledge subject to Purchaser's inspection."

        'header vars;
        Dim happlied, hcheckissued, hissued As Date
        Dim hfisyr, hqty, hshippingkey As Int32
        Dim htotal As Decimal
        Dim hvendname, hvendaddr1, hvendaddr2, hvendaddr3, hvendcity, hvendstate, hvendzip, hvendzipx, hvendfull As String
        Dim hvendfax, hvendph1, hvendph2, hvendphone As String
        Dim hchecknumber, hcheckpaid, hdescr, hnumber, hshipattn, hshipvendorattn As String
        'detail line vars;
        Dim daccount, dexpenditure, dremarks, dstatus, dsubaccount As String
        Dim dchecknumber As String
        Dim dcost, damount As Decimal
        Dim dqty As Int32
        'function vars;
        Dim x, y, index As Int32

        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'render the header portion of the purchase order;
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            index = 1
            With Me.Doc1
                With Me.GridTotals
                    'collect the header information from the first row;
                    Me.BankAccountNumber = CType(.GetData(index, 0), String)
                    hfisyr = CType(.GetData(index, 1), Int32)
                    hnumber = CType(.GetData(index, 2), String)
                    htotal = CType(.GetData(index, 3), Decimal)
                    happlied = CType(.GetData(index, 5), Date)
                    hissued = CType(.GetData(index, 6), Date)
                    hvendname = CType(.GetData(index, 8), String)
                    hvendaddr1 = CType(.GetData(index, 9), String)
                    hvendaddr2 = CType(.GetData(index, 10), String)
                    hvendaddr3 = CType(.GetData(index, 11), String)
                    hvendcity = CType(.GetData(index, 12), String)
                    hvendstate = CType(.GetData(index, 13), String)
                    hvendzip = CType(.GetData(index, 14), String)
                    hvendzipx = CType(.GetData(index, 15), String)
                    hvendph1 = CType(.GetData(index, 16), String)
                    hvendph2 = CType(.GetData(index, 17), String)
                    hdescr = CType(.GetData(index, 18), String)
                    hshippingkey = CType(.GetData(index, 19), Int32)
                    hshipattn = CType(.GetData(index, 20), String)
                    hshipvendorattn = CType(.GetData(index, 21), String)
                    hchecknumber = CType(.GetData(index, 22), String)
                    hcheckpaid = CType(.GetData(index, 23), String)
                    hcheckissued = CType(.GetData(index, 24), Date)
                    hvendfax = CType(.GetData(index, 25), String)
                    'build the vendor city line;
                    hvendfull = hvendcity & ", " & hvendstate & " " & hvendzip
                    If hvendfull.Trim = "," Then hvendfull = ""
                    'get the current shipping information;
                    Call GetShippingInformation(hshippingkey)
                End With
                'create a new page if current page greater than one;
                If epage > 1 Then .NewPage()
                'set the beginning y coordinate;
                y = 20
                'left side;
                .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                'right side;
                .RenderDirectText(100, y, "Activity Fund Purchase Order", 90, 5, verdanaright10)
                .RenderDirectText(145, y + 5, "PO Number: ", 30, 5, verdanaleft10)
                .RenderDirectText(160, y + 5, hnumber, 30, 5, verdanaright10bold)
                .RenderDirectText(130, y + 10, "Issued: " & hissued.ToShortDateString, 60, 5, verdanaright10)
                .RenderDirectText(130, y + 15, "Total:  " & htotal.ToString.Format("{0:C2}", htotal), 60, 5, verdanaright10bold)
                'draw line under the title information;
                .RenderDirectLine(0, y + 22, 190, y + 22, Color.LightGray, 0.5)
                'draw the vendor;
                x = 21
                y = 44
                .RenderDirectText(2, y, "To:", 60, 5, arialleft8)
                .RenderDirectText(x, y, hvendname, 74, 5, arialleft8) : y += 4
                'render the vendor lines, and wrap accordingly if available;
                If hvendaddr1.Length > 0 Then .RenderDirectText(x, y, hvendaddr1, 60, 5, arialleft8) : y += 4
                If hvendaddr2.Length > 0 Then .RenderDirectText(x, y, hvendaddr2, 60, 5, arialleft8) : y += 4
                If hvendaddr3.Length > 0 Then .RenderDirectText(x, y, hvendaddr3, 60, 5, arialleft8) : y += 4
                If hvendfull.Length > 0 Then .RenderDirectText(x, y, hvendfull, 60, 5, arialleft8) : y += 4

                'build the vendor phone line;
                hvendphone = ""
                If hvendph1.Length > 0 Then hvendphone = hvendph1 + Space(2)
                If hvendph2.Length > 0 Then hvendphone &= hvendph2 + Space(2)
                If hvendfax.Length > 0 Then hvendphone &= hvendfax

                'render the phone line if space available;
                If hvendphone.Length > 0 And y < 64 Then
                    'hvendph1 &= Space(3) & hvendph2
                    .RenderDirectText(x, y, hvendphone, 60, 5, arialleft8)
                    y += 4
                End If
                'rendor the Attn: for the shipping line;
                If hshipvendorattn.Length > 0 Then
                    hshipvendorattn = "ATTN: " & hshipvendorattn
                    .RenderDirectText(x, y, hshipvendorattn, 60, 5, arialleft8)
                End If

                ''''''''''''''''''''''''''''''
                'draw the ship to;
                ''''''''''''''''''''''''''''''
                x = 21
                y = 69
                Dim shipcitystatezip As String = Me.ShippingCity & ", " & Me.ShippingState & " " & Me.ShippingZip
                If shipcitystatezip.Trim = "," Then shipcitystatezip = ""
                .RenderDirectText(2, y, "Ship To:", 60, 5, arialleft8)
                .RenderDirectText(x, y, Me.ShippingName, 60, 5, arialleft8) : y += 4
                If Me.ShippingAddress1.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress1, 60, 5, arialleft8) : y += 4
                If Me.ShippingAddress2.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress2, 60, 5, arialleft8) : y += 4
                If Me.ShippingAddress3.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress3, 60, 5, arialleft8) : y += 4
                If shipcitystatezip.Length > 0 Then .RenderDirectText(x, y, shipcitystatezip, 60, 5, arialleft8) : y += 4
                If hshipattn.Length > 0 Then
                    hshipattn = "ATTN: " & hshipattn
                    .RenderDirectText(x, y, hshipattn, 60, 5, arialleft8)
                End If

                ''''''''''''''''''''''''''''''
                'draw condition box;
                ''''''''''''''''''''''''''''''
                x = 128
                y = 48
                specstyle.TextColor = Color.Black
                .RenderDirectText(x, 45, "CONDITIONS", 60, 3, specstyle)
                specstyle.TextColor = Color.Gray
                .RenderDirectText(x, 48, cond1, 60, 3, specstyle)
                .RenderDirectText(x, 51, cond2, 60, 3, specstyle)
                .RenderDirectText(x, 54, cond3, 60, 3, specstyle)
                .RenderDirectText(x, 57, cond4, 60, 3, specstyle)
                .RenderDirectText(x, 60, cond5, 60, 3, specstyle)
                'draw box around the conditions;
                .RenderDirectRectangle(125.5, y - 4, 189.5, y + 15.5, Color.LightGray, 0.5, Color.Transparent)
                'draw the signature line;
                x = 128
                y = 88

                ''''''''''''''''''''''''''''''
                'Draw the signature image;
                ''''''''''''''''''''''''''''''
                If Me.DoSignatures Then
                    Dim imgalign As New C1.C1PrintDocument.ImageAlignDef
                    imgalign.AlignHorz = ImageAlignHorzEnum.Left
                    imgalign.StretchHorz = True
                    imgalign.StretchVert = True
                    imgalign.KeepAspectRatio = True
                    'if primary signature image is available, then print the image;
                    If Not Me.Signature1 Is Nothing Then Doc1.RenderDirectImage(128, 76, Me.Signature1, 300, 14, imgalign)
                    'print the name of the primary signer under the first line;
                    specstyle.TextAlignHorz = AlignHorzEnum.Right
                    .RenderDirectText(150, y + 1, Me.SignatureTextLine1, 39, 3, specstyle)
                    'if secondary signature image is available, then print the image;
                    'If Not Me.Signature2 Is Nothing Then Doc1.RenderDirectImage(101, 40, Me.Signature2, 300, 14, imgalign)
                    'print the name of the secondary signer under the second line;
                    '.RenderDirectText(101, 51, Me.SignatureTextLine2, 80, 5, arialleft8)
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Else
                    specstyle.TextAlignHorz = AlignHorzEnum.Right
                    .RenderDirectText(x + 23, y + 1, "Activity Fund Custodian", 38, 5, specstyle)
                End If
                ''''''''''''''''''''''''''''''
                'End of signatures;
                ''''''''''''''''''''''''''''''

                ''''''''''''''''''''''''''''''
                'draw the signature line;
                ''''''''''''''''''''''''''''''
                x = 128
                y = 88
                .RenderDirectLine(x, y, 189, y, Color.Gray, 1.0)
                specstyle.TextAlignHorz = AlignHorzEnum.Left
                .RenderDirectText(x, y + 1, "Purchase approved by", 60, 4, specstyle)
                .RenderDirectText(x, y + 4, "FY-" & hfisyr.ToString, 60, 4, specstyle)
                specstyle.TextAlignHorz = AlignHorzEnum.Right
                .RenderDirectText(x + 10, y + 4, "SCHOOL ACTIVITY FNDS - 60", 51, 4, specstyle)
                'draw the description;
                x = 21
                y = 95
                .RenderDirectText(2, y, "Description:", 20, 5, arialleft8)
                .RenderDirectText(x, y, hdescr, 100, 15, arialleft8)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'if a check is tied to this purchase order, then draw the check information;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If hchecknumber.Length > 0 Then
                    y = 101
                    .RenderDirectRectangle(127, y - 1, 189.5, y + 8.5, Color.LightGray, 0.5, Color.WhiteSmoke)
                    .RenderDirectText(128, y, "Status:", 20, 5, verdanaleft8)
                    .RenderDirectText(142, y, "Check number:", 25, 5, verdanaleft8)
                    .RenderDirectText(157, y, "Issued:", 30, 5, verdanaright8)
                    .RenderDirectText(128, y + 3, "Paid", 20, 5, verdanaleft8bold)
                    .RenderDirectText(136, y + 3, hchecknumber, 30, 5, verdanaright8bold)
                    .RenderDirectText(159, y + 3, hcheckissued.ToShortDateString, 30, 5, verdanaright8bold)
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                y = 105
                specstyle.TextAlignHorz = AlignHorzEnum.Left
                .RenderDirectText(2, y, "For applied period " & happlied.ToString.Format("{0:MMMM, yyyy}", happlied), 75, 5, specstyle)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                y = 111
                'draw line under the po header information;
                .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                'draw the column headers;
                y = 111
                .RenderDirectText(2, y, "Account", 20, 5, arialleft8)
                .RenderDirectText(18, y, "Expenditure coding", 30, 5, arialleft8)
                .RenderDirectText(122, y, "Check", 20, 5, arialleft8)
                .RenderDirectText(140, y, "Qty", 10, 5, arialright8)
                .RenderDirectText(150, y, "Cost", 20, 5, arialright8)
                .RenderDirectText(170, y, "Amount", 20, 5, arialright8)
                'draw line under the column headers;
                y = 116
                .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                y = 119



                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'print the detail lines for the purchase order;
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    With Me.GridDetail
                        dqty = CType(.GetData(index, 0), Int32)
                        dcost = CType(.GetData(index, 1), Decimal)
                        damount = CType(.GetData(index, 2), Decimal)
                        daccount = CType(.GetData(index, 3), String)
                        dsubaccount = CType(.GetData(index, 5), String)
                        dexpenditure = FormatExpenditureCode(CType(.GetData(index, 7), String))
                        dchecknumber = CType(.GetData(index, 8), String)
                        dstatus = CType(.GetData(index, 11), String).ToUpper
                        dremarks = CType(.GetData(index, 12), String)
                        If dstatus = "V" Then
                            If damount >= 0 Then
                                dremarks = "***VOID CHECK " & dchecknumber & "***"
                            End If
                            If damount < 0 Then
                                dremarks = "***VOID REVERSE ENTRY " & dchecknumber & "***"
                                dchecknumber = ""
                            End If
                        End If
                    End With
                    '
                    .RenderDirectText(1, y, daccount & "-" & dsubaccount, 20, 5, verdanaleft8)
                    .RenderDirectText(18, y, dexpenditure, 70, 5, verdanaleft8)
                    .RenderDirectText(110, y, dchecknumber, 25, 5, arialright8)
                    .RenderDirectText(140, y, dqty.ToString.Format("{0:D2}", dqty), 10, 5, verdanaright8)
                    .RenderDirectText(150, y, dcost.ToString.Format("{0:F2}", dcost), 20, 5, verdanaright8)
                    .RenderDirectText(165, y, damount.ToString.Format("{0:F2}", damount), 25, 5, verdanaright8)
                    y += 5
                    .RenderDirectText(18, y, dremarks, 172, 5, verdanaleft8)
                    'check if it's a page break;
                    If y > 245 Then
                        .NewPage()
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Purchase Order", 90, 5, verdanaright10)
                        .RenderDirectText(145, y + 5, "PO Number: ", 30, 5, verdanaleft10)
                        .RenderDirectText(160, y + 5, hnumber, 30, 5, verdanaright10bold)
                        .RenderDirectText(130, y + 10, "Issued: " & hissued.ToShortDateString, 60, 5, verdanaright10)
                        .RenderDirectText(130, y + 15, "Total:  " & htotal.ToString.Format("{0:C2}", htotal), 60, 5, verdanaright10bold)
                        'draw continuation information;
                        .RenderDirectText(0, y + 15, "Continued from previous page...", 190, 10, timesleft16)
                        'draw line under the title information;
                        .RenderDirectLine(0, y + 22, 190, y + 22, Color.LightGray, 0.5)
                        'draw the column headers;
                        y = 42
                        .RenderDirectText(2, y, "Account", 20, 5, arialleft8)
                        .RenderDirectText(18, y, "Expenditure coding", 30, 5, arialleft8)
                        .RenderDirectText(122, y, "Check", 20, 5, arialleft8)
                        .RenderDirectText(140, y, "Qty", 10, 5, arialright8)
                        .RenderDirectText(150, y, "Cost", 20, 5, arialright8)
                        .RenderDirectText(170, y, "Amount", 20, 5, arialright8)
                        'draw line under the column headers;
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        'subtract here since you're about to add 8 again;
                        y = 42
                    End If
                    y += 8
                Next
            End With
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''''''''Private Sub PrintPurchaseOrderKick(ByVal epage As Int32)
    ''''''''    'define styles;
    ''''''''    arialleft8 = New C1DocStyle(Me.Doc1)
    ''''''''    arialright8 = New C1DocStyle(Me.Doc1)
    ''''''''    arialleft10 = New C1DocStyle(Me.Doc1)
    ''''''''    arialleft10bold = New C1DocStyle(Me.Doc1)
    ''''''''    arialright10 = New C1DocStyle(Me.Doc1)
    ''''''''    footerstyle = New C1DocStyle(Me.Doc1)
    ''''''''    specstyle = New C1DocStyle(Me.Doc1)
    ''''''''    timesleft16 = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaleft8 = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaleft8bold = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaright8 = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaright8bold = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaleft10 = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaleft10bold = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaright10 = New C1DocStyle(Me.Doc1)
    ''''''''    verdanaright10bold = New C1DocStyle(Me.Doc1)
    ''''''''    'define the styles;
    ''''''''    Call DefineStyles()
    ''''''''    'define the document;
    ''''''''    Call DefineDocumentSettings(Me.DocumentName)

    ''''''''    With timesleft16
    ''''''''        'used for the continuation pages;
    ''''''''        .Font = New Font("Arial", 10, FontStyle.Bold)
    ''''''''        .TextAlignHorz = AlignHorzEnum.Center
    ''''''''        .TextColor = Color.Salmon
    ''''''''    End With

    ''''''''    With specstyle
    ''''''''        'this style is used for the Conditions box printed on the purchase order;
    ''''''''        .Borders.AllEmpty = True
    ''''''''        .Font = New Font("Arial", 6, FontStyle.Regular)
    ''''''''        .TextAlignHorz = AlignHorzEnum.Justify
    ''''''''        .TextColor = Color.Gray
    ''''''''    End With

    ''''''''    Dim condhrd1 As String = "NON-KICKBACK AFFIDAVIT"
    ''''''''    Dim condhrd2 As String = "STATE OF"
    ''''''''    Dim condhrd3 As String = "SS"
    ''''''''    Dim condhrd4 As String = " COUNTY OF"

    ''''''''    Dim cond1 As String = "The undersigned(architect, contractor, supplier, or engineer), of lawful age,"
    ''''''''    Dim cond2 As String = "being first duly sworn, on oath says that this (invoice, claim, or contract) is true and"
    ''''''''    Dim cond3 As String = "correct. Affiant further states that the (work, services, or materials) as shown by this"
    ''''''''    Dim cond4 As String = "invoice or claim will be (completed or supplied) in accordance with the plans,"
    ''''''''    Dim cond5 As String = "specifications, orders, or request furnished the affiant Affiant further states that"
    ''''''''    Dim cond6 As String = "(s)he has made no payment, given or donated or agreed to pay, give or donate,"
    ''''''''    Dim cond7 As String = "either directly or indirectly, to any elected official, officer, or employee of the State of"
    ''''''''    Dim cond8 As String = "Oklahoma, any county or local subdivision of the state, of money or any other thing"
    ''''''''    Dim cond9 As String = "of value to obtain payment or the award of this contract."

    ''''''''    Dim condend1 As String = "Oklahoma law requires school districts to obtain a properly signed and witnessed NON-KICKBACK AFFIDAVIT from any"
    ''''''''    Dim condend2 As String = "vendor submitting an invoice/purchase order for $25,000.00 or more"



    ''''''''    'header vars;
    ''''''''    Dim hvendname As String
    ''''''''    hvendname = "SEAN ALEXANDER"




    ''''''''    ''Dim happlied, hcheckissued, hissued As Date
    ''''''''    ''Dim hfisyr, hqty, hshippingkey As Int32
    ''''''''    ''Dim htotal As Decimal
    ''''''''    ''Dim hvendname, hvendaddr1, hvendaddr2, hvendaddr3, hvendcity, hvendstate, hvendzip, hvendzipx, hvendfull As String
    ''''''''    ''Dim hvendfax, hvendph1, hvendph2, hvendphone As String
    ''''''''    ''Dim hchecknumber, hcheckpaid, hdescr, hnumber, hshipattn, hshipvendorattn As String
    ''''''''    'detail line vars;
    ''''''''    ''Dim daccount, dexpenditure, dremarks, dstatus, dsubaccount As String
    ''''''''    ''Dim dchecknumber As String
    ''''''''    ''Dim dcost, damount As Decimal
    ''''''''    ''Dim dqty As Int32
    ''''''''    'function vars;
    ''''''''    Dim x, y, index As Int32

    ''''''''    Try
    ''''''''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''        'render the header portion of the purchase order;
    ''''''''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''        index = 1
    ''''''''        With Me.Doc1
    ''''''''            ''With Me.GridTotals
    ''''''''            ''    'collect the header information from the first row;
    ''''''''            ''    Me.BankAccountNumber = CType(.GetData(index, 0), String)
    ''''''''            ''    hfisyr = CType(.GetData(index, 1), Int32)
    ''''''''            ''    hnumber = CType(.GetData(index, 2), String)
    ''''''''            ''    htotal = CType(.GetData(index, 3), Decimal)
    ''''''''            ''    happlied = CType(.GetData(index, 5), Date)
    ''''''''            ''    hissued = CType(.GetData(index, 6), Date)
    ''''''''            ''    hvendname = CType(.GetData(index, 8), String)
    ''''''''            ''    hvendaddr1 = CType(.GetData(index, 9), String)
    ''''''''            ''    hvendaddr2 = CType(.GetData(index, 10), String)
    ''''''''            ''    hvendaddr3 = CType(.GetData(index, 11), String)
    ''''''''            ''    hvendcity = CType(.GetData(index, 12), String)
    ''''''''            ''    hvendstate = CType(.GetData(index, 13), String)
    ''''''''            ''    hvendzip = CType(.GetData(index, 14), String)
    ''''''''            ''    hvendzipx = CType(.GetData(index, 15), String)
    ''''''''            ''    hvendph1 = CType(.GetData(index, 16), String)
    ''''''''            ''    hvendph2 = CType(.GetData(index, 17), String)
    ''''''''            ''    hdescr = CType(.GetData(index, 18), String)
    ''''''''            ''    hshippingkey = CType(.GetData(index, 19), Int32)
    ''''''''            ''    hshipattn = CType(.GetData(index, 20), String)
    ''''''''            ''    hshipvendorattn = CType(.GetData(index, 21), String)
    ''''''''            ''    hchecknumber = CType(.GetData(index, 22), String)
    ''''''''            ''    hcheckpaid = CType(.GetData(index, 23), String)
    ''''''''            ''    hcheckissued = CType(.GetData(index, 24), Date)
    ''''''''            ''    hvendfax = CType(.GetData(index, 25), String)
    ''''''''            ''    'build the vendor city line;
    ''''''''            ''    hvendfull = hvendcity & ", " & hvendstate & " " & hvendzip
    ''''''''            ''    If hvendfull.Trim = "," Then hvendfull = ""
    ''''''''            ''    'get the current shipping information;
    ''''''''            ''    Call GetShippingInformation(hshippingkey)
    ''''''''            ''End With
    ''''''''            'create a new page if current page greater than one;
    ''''''''            If epage > 1 Then .NewPage()
    ''''''''            'set the beginning y coordinate;
    ''''''''            y = 20
    ''''''''            'left side;
    ''''''''            .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
    ''''''''            .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
    ''''''''            .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
    ''''''''            .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
    ''''''''            'right side;
    ''''''''            .RenderDirectText(100, y, condhrd1, 90, 5, verdanaright10)
    ''''''''            .RenderDirectText(145, y + 5, condhrd2, 30, 5, verdanaleft10)
    ''''''''            .RenderDirectText(160, y + 5, condhrd3, 30, 5, verdanaright10bold)
    ''''''''            .RenderDirectText(130, y + 10, condhrd4, 60, 5, verdanaright10)

    ''''''''            'draw line under the title information;
    ''''''''            .RenderDirectLine(0, y + 22, 190, y + 22, Color.LightGray, 0.5)

    ''''''''            'draw the vendor;
    ''''''''            x = 21
    ''''''''            y = 44
    ''''''''            .RenderDirectText(2, y, "To:", 60, 5, arialleft8)
    ''''''''            .RenderDirectText(x, y, hvendname, 74, 5, arialleft8) : y += 4
    ''''''''            'render the vendor lines, and wrap accordingly if available;
    ''''''''            ''If hvendaddr1.Length > 0 Then .RenderDirectText(x, y, hvendaddr1, 60, 5, arialleft8) : y += 4
    ''''''''            ''If hvendaddr2.Length > 0 Then .RenderDirectText(x, y, hvendaddr2, 60, 5, arialleft8) : y += 4
    ''''''''            ''If hvendaddr3.Length > 0 Then .RenderDirectText(x, y, hvendaddr3, 60, 5, arialleft8) : y += 4
    ''''''''            ''If hvendfull.Length > 0 Then .RenderDirectText(x, y, hvendfull, 60, 5, arialleft8) : y += 4

    ''''''''            'build the vendor phone line;
    ''''''''            ''hvendphone = ""
    ''''''''            ''If hvendph1.Length > 0 Then hvendphone = hvendph1 + Space(2)
    ''''''''            ''If hvendph2.Length > 0 Then hvendphone &= hvendph2 + Space(2)
    ''''''''            ''If hvendfax.Length > 0 Then hvendphone &= hvendfax

    ''''''''            'render the phone line if space available;
    ''''''''            ''If hvendphone.Length > 0 And y < 64 Then
    ''''''''            ''    'hvendph1 &= Space(3) & hvendph2
    ''''''''            ''    .RenderDirectText(x, y, hvendphone, 60, 5, arialleft8)
    ''''''''            ''    y += 4
    ''''''''            ''End If
    ''''''''            'rendor the Attn: for the shipping line;
    ''''''''            ''If hshipvendorattn.Length > 0 Then
    ''''''''            ''    hshipvendorattn = "ATTN: " & hshipvendorattn
    ''''''''            ''    .RenderDirectText(x, y, hshipvendorattn, 60, 5, arialleft8)
    ''''''''            ''End If

    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            'draw the ship to;
    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            x = 21
    ''''''''            y = 69
    ''''''''            ''Dim shipcitystatezip As String = Me.ShippingCity & ", " & Me.ShippingState & " " & Me.ShippingZip
    ''''''''            ''If shipcitystatezip.Trim = "," Then shipcitystatezip = ""
    ''''''''            ''.RenderDirectText(2, y, "Ship To:", 60, 5, arialleft8)
    ''''''''            ''.RenderDirectText(x, y, Me.ShippingName, 60, 5, arialleft8) : y += 4
    ''''''''            ''If Me.ShippingAddress1.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress1, 60, 5, arialleft8) : y += 4
    ''''''''            ''If Me.ShippingAddress2.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress2, 60, 5, arialleft8) : y += 4
    ''''''''            ''If Me.ShippingAddress3.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress3, 60, 5, arialleft8) : y += 4
    ''''''''            ''If shipcitystatezip.Length > 0 Then .RenderDirectText(x, y, shipcitystatezip, 60, 5, arialleft8) : y += 4
    ''''''''            ''If hshipattn.Length > 0 Then
    ''''''''            ''    hshipattn = "ATTN: " & hshipattn
    ''''''''            ''    .RenderDirectText(x, y, hshipattn, 60, 5, arialleft8)
    ''''''''            ''End If

    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            'draw condition box;
    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            x = 128
    ''''''''            y = 48
    ''''''''            specstyle.TextColor = Color.Black
    ''''''''            ''.RenderDirectText(x, 45, "CONDITIONS", 60, 3, specstyle)
    ''''''''            'specstyle.TextColor = Color.Gray
    ''''''''            .RenderDirectText(x, 48, cond1, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 51, cond2, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 54, cond3, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 57, cond4, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 60, cond5, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 60, cond6, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 60, cond7, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 60, cond8, 60, 3, specstyle)
    ''''''''            .RenderDirectText(x, 60, cond9, 60, 3, specstyle)

    ''''''''            'draw box around the conditions;
    ''''''''            .RenderDirectRectangle(125.5, y - 4, 189.5, y + 15.5, Color.LightGray, 0.5, Color.Transparent)
    ''''''''            'draw the signature line;
    ''''''''            x = 128
    ''''''''            y = 88


    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            'draw the signature line;
    ''''''''            ''''''''''''''''''''''''''''''
    ''''''''            x = 128
    ''''''''            y = 88
    ''''''''            .RenderDirectLine(x, y, 189, y, Color.Gray, 1.0)
    ''''''''            specstyle.TextAlignHorz = AlignHorzEnum.Left
    ''''''''            .RenderDirectText(x, y + 1, condend1, 60, 4, specstyle)
    ''''''''            .RenderDirectText(x, y + 4, condend2, 60, 4, specstyle)

    ''''''''            'draw the description;
    ''''''''            x = 21
    ''''''''            y = 95

    ''''''''            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''            y = 105
    ''''''''            ''specstyle.TextAlignHorz = AlignHorzEnum.Left
    ''''''''            ''.RenderDirectText(2, y, "For applied period " & happlied.ToString.Format("{0:MMMM, yyyy}", happlied), 75, 5, specstyle)
    ''''''''            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''            y = 111
    ''''''''            'draw line under the po header information;
    ''''''''            '.RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
    ''''''''            'draw the column headers;
    ''''''''            y = 111
    ''''''''            ''.RenderDirectText(2, y, "Account", 20, 5, arialleft8)
    ''''''''            ''.RenderDirectText(18, y, "Expenditure coding", 30, 5, arialleft8)
    ''''''''            ''.RenderDirectText(122, y, "Check", 20, 5, arialleft8)
    ''''''''            ''.RenderDirectText(140, y, "Qty", 10, 5, arialright8)
    ''''''''            ''.RenderDirectText(150, y, "Cost", 20, 5, arialright8)
    ''''''''            ''.RenderDirectText(170, y, "Amount", 20, 5, arialright8)
    ''''''''            'draw line under the column headers;
    ''''''''            y = 116
    ''''''''            '.RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
    ''''''''            y = 119

    ''''''''            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''        End With
    ''''''''    Catch ex As Exception
    ''''''''        Throw
    ''''''''    End Try
    ''''''''End Sub

    Private Sub PrintPurchaseOrderActivity()
        '''''''''''''''''''''''''''''''''''''' GridWrk ''''''''''''''''''''''''''''''''''''''''''
        '      0            1           2           3           4           5          6 
        '    bank         fisyr      number      status      amount      account    subacct
        '      7            8           9          10          11          12         13  
        '    code         descr      issued      podtkey     pokey       invckey    ckdtkey
        '     14  
        '  rectype
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''''''''''''
        '      0            1           2           3           4           5           6
        '    pokey      bankacct      fisyr      number      descr       vendor      potype 
        '      7            8           9          10          11          12          13 
        '   issued       reqkey     reqnumber    totenc     totinvc     totspent   outstanding
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "PurchaseOrderHistory"
        Me.ReportName = "Purchase Order Activity"
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
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)
        '
        Dim index, x, y As Int32
        '
        Dim hdescr, hnumber, hreqnumber, htype, hvendor As String
        Dim hfisyr As Int32
        Dim hissued As Date
        Dim hencumbered, hinvoiced, hspent, houtstanding As Decimal
        '
        Dim daccount, dchecknumber, dexpenditure, dremarks, dstatus As String
        Dim dissued As Date
        Dim damount As Decimal
        Dim drectype As Int32


        'collect the header information;
        With Me.GridTotals
            index = 1
            Me.BankAccountNumber = CType(.GetData(index, 1), String)
            hfisyr = CType(.GetData(index, 2), Int32)
            hnumber = CType(.GetData(index, 3), String)
            hdescr = CType(.GetData(index, 4), String)
            hvendor = CType(.GetData(index, 5), String)
            htype = CType(.GetData(index, 6), String).ToUpper
            If htype = "B" Then htype = "(B)lanket" Else htype = "(R)egular"
            hissued = CType(.GetData(index, 7), Date)
            hreqnumber = CType(.GetData(index, 9), String)
            hencumbered = CType(.GetData(index, 10), Decimal)
            hinvoiced = CType(.GetData(index, 11), Decimal)
            hspent = CType(.GetData(index, 12), Decimal)
            houtstanding = CType(.GetData(index, 13), Decimal)
        End With

        Try
            With Me.Doc1
                .StartDoc()
                'print the total info box left-side;
                y = 30
                .RenderDirectText(0, y + 4, "For Bank Account:", 35, 5, verdanaright8bold)
                .RenderDirectText(30, y + 4, "Year:", 20, 5, verdanaright8bold)
                .RenderDirectText(50, y + 4, "Purchase order:", 30, 5, verdanaright8bold)
                .RenderDirectText(80, y + 4, "Issued:", 20, 5, verdanaright8bold)
                .RenderDirectText(-5, y + 8, Me.BankAccountNumber, 40, 5, verdanaright8)
                .RenderDirectText(30, y + 8, hfisyr.ToString, 20, 5, verdanaright8)
                .RenderDirectText(50, y + 8, hnumber, 25, 5, verdanaright8)
                .RenderDirectText(81, y + 8, hissued.ToShortDateString, 20, 5, verdanaright8)
                .RenderDirectText(0, y + 16, "Vendor:", 35, 5, verdanaright8bold)
                .RenderDirectText(37, y + 16, hvendor, 90, 10, verdanaleft8)
                'print the info box right-side;
                .RenderDirectText(118, y + 2, "Encumbered:", 40, 5, verdanaright8bold)
                .RenderDirectText(118, y + 7, "Invoiced:", 40, 5, verdanaright8bold)
                .RenderDirectText(118, y + 12, "Paid:", 40, 5, verdanaright8bold)
                .RenderDirectText(118, y + 20, "Outstanding:", 40, 5, verdanaright8bold)
                'print the money fields;
                .RenderDirectText(160, y + 2, hencumbered.ToString.Format("{0:F2}", hencumbered), 30, 5, verdanaright8bold)
                .RenderDirectText(160, y + 7, hinvoiced.ToString.Format("{0:F2}", hinvoiced), 30, 5, verdanaright8bold)
                .RenderDirectText(160, y + 12, hspent.ToString.Format("{0:F2}", hspent), 30, 5, verdanaright8bold)
                '.RenderDirectText(160, y + 14, hvoided.ToString.Format("{0:F2}", totalvoided), 30, 5, verdanaright8bold)
                .RenderDirectText(160, y + 20, houtstanding.ToString.Format("{0:C2}", houtstanding), 30, 5, verdanaright8bold)
                'print line above the column headers;
                .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                y = 58
                'print the column headers;
                .RenderDirectText(2, y, "Issued", 15, 5, verdanaleft8bold)
                .RenderDirectText(20, y, "Account", 20, 5, verdanaleft8bold)
                .RenderDirectText(38, y, "Description/Invoice", 50, 5, verdanaleft8bold)
                .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                .RenderDirectText(120, y, "Invoiced", 25, 5, verdanaright8bold)
                .RenderDirectText(145, y, "Spent", 25, 5, verdanaright8bold)
                .RenderDirectText(172, y, "Check", 25, 5, verdanaleft8bold)
                y = 65
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    With Me.GridWrk
                        dstatus = DirectCast(.GetData(index, 3), String)
                        damount = CType(.GetData(index, 4), Decimal)
                        daccount = CType(.GetData(index, 5), String)
                        daccount &= "-" & CType(.GetData(index, 6), String)
                        dexpenditure = CType(.GetData(index, 7), String)
                        dremarks = CType(.GetData(index, 8), String)
                        dissued = CType(.GetData(index, 9), Date)
                        drectype = CType(.GetData(index, 14), Int32)
                        Select Case drectype
                            Case 1
                            Case 2
                            Case 3
                                dchecknumber = CType(.GetData(index, 2), String)
                        End Select
                    End With

                    'bypass voided encumbrances;
                    If drectype = 1 And dstatus = "V" Then GoTo skip

                    'do not print void encumbrances;
                    If drectype = 1 Then
                        .RenderDirectText(95, y, damount.ToString.Format("{0:F2}", damount), 25, 5, verdanaright8)
                    End If

                    If drectype = 2 Then
                        .RenderDirectText(120, y, damount.ToString.Format("{0:F2}", damount), 25, 5, verdanaright8)
                    End If

                    If drectype = 3 Then
                        .RenderDirectText(145, y, damount.ToString.Format("{0:F2}", damount), 25, 5, verdanaright8)
                        .RenderDirectText(172, y, dchecknumber, 25, 10, verdanaleft8)
                    End If
                    '
                    .RenderDirectText(0, y, dissued.ToShortDateString, 20, 5, verdanaleft8)
                    .RenderDirectText(20, y, daccount, 20, 5, verdanaleft8)
                    .RenderDirectText(38, y, dremarks, 60, 10, verdanaleft8)
                    '
                    y += 8

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        y = 34
                        'print the column headers;
                        .RenderDirectText(2, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Account", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Description/Invoice", 50, 5, verdanaleft8bold)
                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(120, y, "Invoiced", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Spent", 25, 5, verdanaright8bold)
                        .RenderDirectText(172, y, "Check", 25, 5, verdanaleft8bold)
                        y = 41
                    End If
Skip:
                    Application.DoEvents()
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

    Private Sub PrintPurchaseOrderRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5     6
        '   bank        fisyr    ponumber  vendor  hdrstatus   qty  cost
        '     7           8         9       10       11        12 
        '  amount      hdramt     acct     sub    applied   created
        '    13          14        15       16       17
        ' detlstatus  hdrdescr  remarks  chknumber chkstat
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "PurchaseOrderRegister"
        Me.ReportName = "Purchase Order Register"
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

        'this style is only used by this report
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            .TextColor = Color.Gray
        End With

        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim topofpage As Boolean
        Dim index, x, y, tqty, topen, tclosed, tvoided, tcount As Int32
        Dim totalregister As Double
        Dim tissuedate, tapplieddate As Date
        Dim tvendor, tacctnum, tsubacctnum, tdescr, tremarks, tponumber As String
        Dim tstatus, prevponum, prtstatus, tchknum, tchkstatus, tchkreconsw As String
        Dim thdramt, tlineamt, tcost, tsumamt As Double

        Try
            'get the total amount of the register
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
                    If index = 1 Then
                        'print the total info box left-side
                        .RenderDirectText(25, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(1, 34, "Status Key:", 40, 5, verdanaleft8bold)
                        .RenderDirectText(3, 39, "1 - Cleared (Check)", 40, 4, specstyle)
                        .RenderDirectText(3, 42, "2 - Outstanding (Check)", 40, 4, specstyle)
                        .RenderDirectText(3, 45, "3 - Open", 40, 4, specstyle)
                        .RenderDirectText(3, 48, "4 - Void (Check)", 40, 4, specstyle)
                        'print the info box right-side
                        y = 30
                        .RenderDirectText(118, y + 4, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y + 4, totalregister.ToString.Format("{0:C2}", totalregister), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, y, "Vendor", 30, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Description/Remarks", 60, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Check", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        '2nd line of column headers
                        y = 62
                        .RenderDirectText(12, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(37, y, "Qty", 25, 5, verdanaleft8bold)
                        .RenderDirectText(55, y, "Cost", 22, 5, verdanaleft8bold)
                        .RenderDirectText(77, y, "Line", 22, 5, verdanaleft8bold)
                        '.RenderDirectText(90, y, "Remarks", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If

                    With Me.GridDetail
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '     0           1         2        3        4         5     6
                        '   bank        fisyr    ponumber  vendor  hdrstatus   qty  cost
                        '     7           8         9       10       11        12 
                        '  amount      hdramt     acct     sub    applied   created
                        '    13          14        15       16       17
                        ' detlstatus  hdrdescr  remarks  chknumber chkstat
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''                        
                        '''''tchknum = DirectCast(.GetData(index, 2), String)
                        tponumber = DirectCast(.GetData(index, 2), String).ToUpper
                        tvendor = DirectCast(.GetData(index, 3), String).ToUpper
                        tstatus = DirectCast(.GetData(index, 4), String).ToUpper
                        tqty = CInt(.GetData(index, 5))
                        tcost = CDbl(.GetData(index, 6))
                        tlineamt = CDbl(.GetData(index, 7))
                        thdramt = CDbl(.GetData(index, 8))
                        tacctnum = DirectCast(.GetData(index, 9), String)
                        tsubacctnum = DirectCast(.GetData(index, 10), String)
                        tapplieddate = CDate(.GetData(index, 11))
                        tissuedate = CDate(.GetData(index, 12))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        tchknum = DirectCast(.GetData(index, 16), String)
                        tchkstatus = DirectCast(.GetData(index, 17), String).ToUpper
                        tchkreconsw = DirectCast(.GetData(index, 18), String).ToUpper
                    End With

                    Select Case tchkstatus
                        Case "O"    'check is unprinted
                            prtstatus = "2"
                            tclosed += 1
                        Case "C"    'check is printed
                            If tchkreconsw = "Y" Then
                                prtstatus = "1"
                            Else
                                prtstatus = "2"
                            End If
                            tclosed += 1
                        Case "F"    'check is cleared
                            prtstatus = "1"
                            tclosed += 1
                        Case "V"    'check is voided
                            prtstatus = "4"
                            tvoided += 1
                        Case Else   'po not converted to check
                            topen += 1
                            prtstatus = "3"
                    End Select

                    If tponumber <> prevponum Then
                        'linefeed extra space between po's if not top of page
                        If Not topofpage Then y += 3
                        .RenderDirectText(-1, y, prtstatus, 5, 5, specstyle)
                        .RenderDirectText(1, y, tponumber, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(45, y, tvendor, 45, 5, verdanaleft8)
                        .RenderDirectText(90, y, tdescr, 58, 5, verdanaleft8)
                        If prtstatus = "4" Then
                            'handle voided check;
                            .RenderDirectText(164, y, "[V]", 10, 5, verdanaleft8bold)
                            thdramt = 0
                        End If
                        .RenderDirectText(148, y, tchknum, 20, 5, verdanaleft8)
                        .RenderDirectText(165, y, thdramt.ToString.Format("{0:F2}", thdramt), 25, 5, verdanaright8)
                        topofpage = False
                        tsumamt += thdramt
                        y += 5
                    End If
                    .RenderDirectText(12, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(32, y, tqty.ToString.Format("{0:D2}", tqty), 12, 5, verdanaright8)
                    .RenderDirectText(45, y, tcost.ToString.Format("{0:F2}", tcost), 20, 5, verdanaright8)
                    .RenderDirectText(65, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(90, y, tremarks, 78, 10, verdanaleft8)
                    y += 7
                    'get the current ponumber
                    prevponum = tponumber

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(45, y, "Vendor", 25, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Description/Remarks", 60, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Check", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        '2nd line of column headers
                        y = 62
                        .RenderDirectText(12, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(37, y, "Qty", 25, 5, verdanaleft8bold)
                        .RenderDirectText(55, y, "Cost", 22, 5, verdanaleft8bold)
                        .RenderDirectText(77, y, "Line", 22, 5, verdanaleft8bold)
                        '.RenderDirectText(90, y, "Remarks", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                tcount = topen + tclosed + tvoided
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Encumbered", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, tsumamt.ToString.Format("{0:C2}", tsumamt), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Purchase Orders", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, tcount.ToString.Format("{0:D2}", tcount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 8, "Open", 50, 5, verdanaright8)
                .RenderDirectText(165, y + 8, topen.ToString.Format("{0:D1}", topen), 25, 5, verdanaright8)
                .RenderDirectText(60, y + 12, "Closed", 50, 5, verdanaright8)
                .RenderDirectText(165, y + 12, tclosed.ToString.Format("{0:D1}", tclosed), 25, 5, verdanaright8)
                .RenderDirectText(60, y + 16, "Voided", 50, 5, verdanaright8)
                .RenderDirectText(165, y + 16, tvoided.ToString.Format("{0:D1}", tvoided), 25, 5, verdanaright8)
                y += 22
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

    Private Sub PrintPurchaseOrderRegisterByAccount(ByVal eoutstandinginvoices As Boolean)
        Me.DocumentName = "PurchaseOrderRegister"
        Me.ReportName = "Purchase Order Register"
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

        'this style is only used by this report;
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            '.TextColor = Color.Gray
        End With

        'define the document;
        DefineDocumentSettings(Me.DocumentName)

        Dim topofpage As Boolean
        Dim index, invoicekey, x, y As Int32
        Dim totalbalance, totalencumbered, totalinvoiced, totalpaid As Decimal
        Dim rptbalance, rptencumbered, rptinvoiced, rptpaid As Decimal
        Dim paidbyinvoice As Boolean
        Dim tissuedate, tapplieddate As Date
        Dim tvendor, tacctnum, tsubacctnum, tdescr, tremarks, tponumber, tclass As String
        Dim tstatus, prevponum, prtstatus As String
        'Dim thdramt, tlineamt, tcost, tsumamt As Double
        Dim encumbered, invoiced, paid, balance As Decimal
        Dim lencumbered, linvoiced, lpaid, lbalance As Decimal

        Try
            'get the total amount of the register;
            With Me.GridTotals
                totalencumbered = CDec(Me.GridTotals.GetData(1, 0))
                totalinvoiced = CDec(Me.GridTotals.GetData(1, 1))
                totalpaid = CDec(Me.GridTotals.GetData(1, 2))
                totalbalance = CDec(Me.GridTotals.GetData(1, 3))
            End With
            'get the bank account number from the first item;
            Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 0), String)
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    If index = 1 Then
                        'print the total info box left-side;
                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
                        'print the info box right-side;
                        y = 30
                        'do not print totals for outstanding invoices;
                        If Not eoutstandinginvoices = True Then
                            .RenderDirectText(118, y + 4, "Encumbered:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 8, "Invoiced:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 12, "Paid:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 18, "Balance:", 40, 5, verdanaright8bold)
                            'print the money fields
                            .RenderDirectText(160, y + 4, totalencumbered.ToString.Format("{0:C2}", totalencumbered), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 8, totalinvoiced.ToString.Format("{0:C2}", totalinvoiced), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 12, totalpaid.ToString.Format("{0:C2}", totalpaid), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 18, totalbalance.ToString.Format("{0:C2}", totalbalance), 30, 5, verdanaright8bold)
                        End If
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
                        '2nd line of column headers;
                        y = 62
                        .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If

                    With Me.GridDetail
                        tponumber = DirectCast(.GetData(index, 2), String).ToUpper
                        tvendor = DirectCast(.GetData(index, 3), String).ToUpper
                        tstatus = DirectCast(.GetData(index, 4), String).ToUpper
                        lencumbered = CDec(.GetData(index, 7))
                        encumbered = CDec(.GetData(index, 8))
                        tacctnum = DirectCast(.GetData(index, 9), String)
                        tsubacctnum = DirectCast(.GetData(index, 10), String)
                        tapplieddate = CDate(.GetData(index, 11))
                        tissuedate = CDate(.GetData(index, 12))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        tclass = DirectCast(.GetData(index, 16), String)
                        invoicekey = CInt(.GetData(index, 19))
                        invoiced = CDec(.GetData(index, 20))
                        linvoiced = CDec(.GetData(index, 21))
                        paid = CDec(.GetData(index, 22))
                        lpaid = CDec(.GetData(index, 23))
                    End With

                    'added for outstanding balances on 04.03.2008;
                    If eoutstandinginvoices = True Then
                        If encumbered = paid Then
                            GoTo BYPASS
                        End If
                    End If

                    rptencumbered += lencumbered
                    rptinvoiced += linvoiced
                    rptpaid += lpaid

                    If tponumber <> prevponum Then
                        'linefeed extra space between po's if not top of page;
                        If Not topofpage Then y += 3

                        'If tponumber = "00001316" Then Stop

                        If invoiced = 0 And paid > 0 Then paidbyinvoice = False
                        If invoiced = 0 And paid = 0 Then paidbyinvoice = True
                        If invoiced > 0 Then paidbyinvoice = True
                        If Not paidbyinvoice Then .RenderDirectText(-1, y + 0.5, "*", 5, 5, specstyle)

                        'calculate the balance;
                        balance = Math.Round(encumbered - paid, 2)
                        'If balance <> 0 Then
                        .RenderDirectText(1, y, tponumber, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(38, y, tvendor, 45, 5, verdanaleft8)
                        'render the header balances;
                        .RenderDirectText(95, y, encumbered.ToString.Format("{0:F2}", encumbered), 20, 5, verdanaright8)
                        If paidbyinvoice Then .RenderDirectText(120, y, invoiced.ToString.Format("{0:F2}", invoiced), 20, 5, verdanaright8)
                        .RenderDirectText(145, y, paid.ToString.Format("{0:F2}", paid), 20, 5, verdanaright8)
                        'calculate the balance;
                        'balance = Math.Round(encumbered - paid, 2)
                        .RenderDirectText(165, y, balance.ToString.Format("{0:F2}", balance), 25, 5, verdanaright8bold)
                        topofpage = False
                        'tsumamt += thdramt
                        y += 5
                    End If
                    'End If
                    .RenderDirectText(8, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    'If Me.UseOcas Then .RenderDirectText(30, y, tclass, 60, 5, verdanaleft8)
                    'render the detail balances;
                    .RenderDirectText(95, y, lencumbered.ToString.Format("{0:F2}", lencumbered), 20, 5, verdanaright8)
                    If paidbyinvoice Then .RenderDirectText(120, y, linvoiced.ToString.Format("{0:F2}", linvoiced), 20, 5, verdanaright8)
                    .RenderDirectText(145, y, lpaid.ToString.Format("{0:F2}", lpaid), 20, 5, verdanaright8)
                    '''''.RenderDirectText(170, y, balance.ToString.Format("{0:F2}", balance), 20, 5, verdanaright8)

                    y += 5
                    'get the current ponumber;
                    prevponum = tponumber

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side;
                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
                        '2nd line of column headers;
                        y = 62
                        .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If
                    'End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
BYPASS:
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Encumbered", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, rptencumbered.ToString.Format("{0:C2}", rptencumbered), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Invoiced", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, rptinvoiced.ToString.Format("{0:C2}", rptinvoiced), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 8, "Total Paid", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 8, rptpaid.ToString.Format("{0:C2}", rptpaid), 25, 5, verdanaright8bold)
                rptbalance = Math.Round(rptencumbered - rptpaid, 2)
                .RenderDirectText(60, y + 14, "Balance", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 14, rptbalance.ToString.Format("{0:C2}", rptbalance), 25, 5, verdanaright8bold)
                y += 22
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

    Private Sub PrintPurchaseOrderByInvoiceRegister(ByVal eoutstandinginvoices As Boolean)
        Me.DocumentName = "PurchaseOrderRegister"
        Me.ReportName = "Purchase Order Register"
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

        'this style is only used by this report;
        Dim specstyle As New C1DocStyle(Me.Doc1)
        With specstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .Spacing.LeftStr = "1mm"
            .Spacing.RightStr = "1mm"
            .Spacing.TopStr = ".5mm"
            .TextAlignHorz = AlignHorzEnum.Left
            .TextPosition = TextPositionEnum.Superscript
            '.TextColor = Color.Gray
        End With

        'define the document;
        DefineDocumentSettings(Me.DocumentName)

        Dim topofpage As Boolean
        Dim index, invoicekey, x, y As Int32
        Dim totalbalance, totalencumbered, totalinvoiced, totalpaid As Decimal
        Dim rptbalance, rptencumbered, rptinvoiced, rptpaid As Decimal
        Dim paidbyinvoice As Boolean
        Dim tissuedate, tapplieddate As Date
        Dim tvendor, tacctnum, tsubacctnum, tdescr, tremarks, tponumber, tclass As String
        Dim tstatus, prevponum, prtstatus As String
        'Dim thdramt, tlineamt, tcost, tsumamt As Double
        Dim encumbered, invoiced, paid, balance As Decimal
        Dim lencumbered, linvoiced, lpaid, lbalance As Decimal

        Try
            'get the total amount of the register;
            With Me.GridTotals
                totalencumbered = CDec(Me.GridTotals.GetData(1, 0))
                totalinvoiced = CDec(Me.GridTotals.GetData(1, 1))
                totalpaid = CDec(Me.GridTotals.GetData(1, 2))
                totalbalance = CDec(Me.GridTotals.GetData(1, 3))
            End With
            'get the bank account number from the first item;
            Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 0), String)
        Catch ex As Exception
            Throw
        End Try

        Try
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridDetail.Rows.Count - 1
                    If index = 1 Then
                        'print the total info box left-side;
                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
                        'print the info box right-side;
                        y = 30
                        'do not print totals for outstanding invoices;
                        If Not eoutstandinginvoices = True Then
                            .RenderDirectText(118, y + 4, "Encumbered:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 8, "Invoiced:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 12, "Paid:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, y + 18, "Balance:", 40, 5, verdanaright8bold)
                            'print the money fields
                            .RenderDirectText(160, y + 4, totalencumbered.ToString.Format("{0:C2}", totalencumbered), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 8, totalinvoiced.ToString.Format("{0:C2}", totalinvoiced), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 12, totalpaid.ToString.Format("{0:C2}", totalpaid), 30, 5, verdanaright8bold)
                            .RenderDirectText(160, y + 18, totalbalance.ToString.Format("{0:C2}", totalbalance), 30, 5, verdanaright8bold)
                        End If
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
                        '2nd line of column headers;
                        y = 62
                        .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If

                    With Me.GridDetail
                        tponumber = DirectCast(.GetData(index, 2), String).ToUpper
                        tvendor = DirectCast(.GetData(index, 3), String).ToUpper
                        tstatus = DirectCast(.GetData(index, 4), String).ToUpper
                        lencumbered = CDec(.GetData(index, 7))
                        encumbered = CDec(.GetData(index, 8))
                        tacctnum = DirectCast(.GetData(index, 9), String)
                        tsubacctnum = DirectCast(.GetData(index, 10), String)
                        tapplieddate = CDate(.GetData(index, 11))
                        tissuedate = CDate(.GetData(index, 12))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        tclass = DirectCast(.GetData(index, 16), String)
                        invoicekey = CInt(.GetData(index, 19))
                        invoiced = CDec(.GetData(index, 20))
                        linvoiced = CDec(.GetData(index, 21))
                        paid = CDec(.GetData(index, 22))
                        lpaid = CDec(.GetData(index, 23))
                    End With

                    'added for outstanding balances on 04.03.2008;
                    If eoutstandinginvoices = True Then
                        If encumbered = paid Then
                            GoTo BYPASS
                        End If
                    End If

                    rptencumbered += lencumbered
                    rptinvoiced += linvoiced
                    rptpaid += lpaid

                    If tponumber <> prevponum Then
                        'linefeed extra space between po's if not top of page;
                        If Not topofpage Then y += 3

                        'If tponumber = "00001316" Then Stop

                        If invoiced = 0 And paid > 0 Then paidbyinvoice = False
                        If invoiced = 0 And paid = 0 Then paidbyinvoice = True
                        If invoiced > 0 Then paidbyinvoice = True
                        If Not paidbyinvoice Then .RenderDirectText(-1, y + 0.5, "*", 5, 5, specstyle)

                        .RenderDirectText(1, y, tponumber, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(38, y, tvendor, 45, 5, verdanaleft8)
                        'render the header balances;
                        .RenderDirectText(95, y, encumbered.ToString.Format("{0:F2}", encumbered), 20, 5, verdanaright8)
                        If paidbyinvoice Then .RenderDirectText(120, y, invoiced.ToString.Format("{0:F2}", invoiced), 20, 5, verdanaright8)
                        .RenderDirectText(145, y, paid.ToString.Format("{0:F2}", paid), 20, 5, verdanaright8)
                        'calculate the balance;
                        balance = Math.Round(encumbered - paid, 2)
                        .RenderDirectText(165, y, balance.ToString.Format("{0:F2}", balance), 25, 5, verdanaright8bold)
                        topofpage = False
                        'tsumamt += thdramt
                        y += 5
                    End If
                    .RenderDirectText(8, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    'If Me.UseOcas Then .RenderDirectText(30, y, tclass, 60, 5, verdanaleft8)
                    'render the detail balances;
                    .RenderDirectText(95, y, lencumbered.ToString.Format("{0:F2}", lencumbered), 20, 5, verdanaright8)
                    If paidbyinvoice Then .RenderDirectText(120, y, linvoiced.ToString.Format("{0:F2}", linvoiced), 20, 5, verdanaright8)
                    .RenderDirectText(145, y, lpaid.ToString.Format("{0:F2}", lpaid), 20, 5, verdanaright8)
                    '''''.RenderDirectText(170, y, balance.ToString.Format("{0:F2}", balance), 20, 5, verdanaright8)

                    y += 5

                    'get the current ponumber;
                    prevponum = tponumber

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side;
                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
                        '2nd line of column headers;
                        y = 62
                        .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
                        topofpage = True
                        y = 70
                    End If
                    'End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
BYPASS:
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Encumbered", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, rptencumbered.ToString.Format("{0:C2}", rptencumbered), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Invoiced", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, rptinvoiced.ToString.Format("{0:C2}", rptinvoiced), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 8, "Total Paid", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 8, rptpaid.ToString.Format("{0:C2}", rptpaid), 25, 5, verdanaright8bold)
                rptbalance = Math.Round(rptencumbered - rptpaid, 2)
                .RenderDirectText(60, y + 14, "Balance", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 14, rptbalance.ToString.Format("{0:C2}", rptbalance), 25, 5, verdanaright8bold)
                y += 22
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

    '''''Private Sub PrintPurchaseOrderOutstandingRegister()
    '        Me.DocumentName = "PurchaseOrderOutstandingRegister"
    '        Me.ReportName = "Purchase Order Outstanding Register"
    '        'define styles
    '        arialleft8 = New C1DocStyle(Me.Doc1)
    '        arialright8 = New C1DocStyle(Me.Doc1)
    '        arialleft10 = New C1DocStyle(Me.Doc1)
    '        arialleft10bold = New C1DocStyle(Me.Doc1)
    '        arialright10 = New C1DocStyle(Me.Doc1)
    '        footerstyle = New C1DocStyle(Me.Doc1)
    '        timesleft16 = New C1DocStyle(Me.Doc1)
    '        verdanaleft8 = New C1DocStyle(Me.Doc1)
    '        verdanaleft8bold = New C1DocStyle(Me.Doc1)
    '        verdanaright8 = New C1DocStyle(Me.Doc1)
    '        verdanaright8bold = New C1DocStyle(Me.Doc1)
    '        verdanaleft10 = New C1DocStyle(Me.Doc1)
    '        verdanaleft10bold = New C1DocStyle(Me.Doc1)
    '        verdanaright10 = New C1DocStyle(Me.Doc1)
    '        verdanaright10bold = New C1DocStyle(Me.Doc1)
    '        'define the styles
    '        DefineStyles()

    '        'this style is only used by this report;
    '        Dim specstyle As New C1DocStyle(Me.Doc1)
    '        With specstyle
    '            .Font = New Font("Verdana", 8, FontStyle.Regular)
    '            .Spacing.LeftStr = "1mm"
    '            .Spacing.RightStr = "1mm"
    '            .Spacing.TopStr = ".5mm"
    '            .TextAlignHorz = AlignHorzEnum.Left
    '            .TextPosition = TextPositionEnum.Superscript
    '            '.TextColor = Color.Gray
    '        End With

    '        'define the document;
    '        DefineDocumentSettings(Me.DocumentName)

    '        Dim topofpage As Boolean
    '        Dim index, invoicekey, x, y As Int32
    '        Dim totalbalance, totalencumbered, totalinvoiced, totalpaid As Decimal
    '        Dim rptbalance, rptencumbered, rptinvoiced, rptpaid As Decimal
    '        Dim paidbyinvoice As Boolean
    '        Dim tissuedate, tapplieddate As Date
    '        Dim tvendor, tacctnum, tsubacctnum, tdescr, tremarks, tponumber, tclass As String
    '        Dim tstatus, prevponum, prtstatus As String
    '        'Dim thdramt, tlineamt, tcost, tsumamt As Double
    '        Dim encumbered, invoiced, paid, balance As Decimal
    '        Dim lencumbered, linvoiced, lpaid, lbalance As Decimal

    '        Try
    '            'get the total amount of the register;
    '            With Me.GridTotals
    '                totalencumbered = CDec(Me.GridTotals.GetData(1, 0))
    '                totalinvoiced = CDec(Me.GridTotals.GetData(1, 1))
    '                totalpaid = CDec(Me.GridTotals.GetData(1, 2))
    '                totalbalance = CDec(Me.GridTotals.GetData(1, 3))
    '            End With
    '            'get the bank account number from the first item;
    '            Me.BankAccountNumber = DirectCast(Me.GridDetail.GetData(1, 0), String)
    '        Catch ex As Exception
    '            Throw
    '        End Try

    '        Try
    '            With Me.Doc1
    '                .StartDoc()
    '                For index = 1 To Me.GridDetail.Rows.Count - 1
    '                    If index = 1 Then
    '                        'print the total info box left-side;
    '                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
    '                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
    '                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
    '                        'print the info box right-side;
    '                        y = 30
    '                        'do not print totals for outstanding invoices;
    '                        'If Not eoutstandinginvoices = True Then
    '                        .RenderDirectText(118, y + 4, "Encumbered:", 40, 5, verdanaright8bold)
    '                        .RenderDirectText(118, y + 8, "Invoiced:", 40, 5, verdanaright8bold)
    '                        .RenderDirectText(118, y + 12, "Paid:", 40, 5, verdanaright8bold)
    '                        .RenderDirectText(118, y + 18, "Balance:", 40, 5, verdanaright8bold)
    '                        'print the money fields
    '                        .RenderDirectText(160, y + 4, totalencumbered.ToString.Format("{0:C2}", totalencumbered), 30, 5, verdanaright8bold)
    '                        .RenderDirectText(160, y + 8, totalinvoiced.ToString.Format("{0:C2}", totalinvoiced), 30, 5, verdanaright8bold)
    '                        .RenderDirectText(160, y + 12, totalpaid.ToString.Format("{0:C2}", totalpaid), 30, 5, verdanaright8bold)
    '                        .RenderDirectText(160, y + 18, totalbalance.ToString.Format("{0:C2}", totalbalance), 30, 5, verdanaright8bold)
    '                        'End If
    '                    'print line above the column headers;
    '                    .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
    '                    y = 58
    '                    'print the column headers;
    '                    .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
    '                    .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
    '                    .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
    '                    .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
    '                    .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
    '                    .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
    '                    .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
    '                    '2nd line of column headers;
    '                    y = 62
    '                    .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
    '                    topofpage = True
    '                    y = 70
    '                    End If

    '                    With Me.GridDetail
    '                        tponumber = DirectCast(.GetData(index, 2), String).ToUpper
    '                        tvendor = DirectCast(.GetData(index, 3), String).ToUpper
    '                        tstatus = DirectCast(.GetData(index, 4), String).ToUpper
    '                        lencumbered = CDec(.GetData(index, 7))
    '                        encumbered = CDec(.GetData(index, 8))
    '                        tacctnum = DirectCast(.GetData(index, 9), String)
    '                        tsubacctnum = DirectCast(.GetData(index, 10), String)
    '                        tapplieddate = CDate(.GetData(index, 11))
    '                        tissuedate = CDate(.GetData(index, 12))
    '                        tdescr = DirectCast(.GetData(index, 14), String)
    '                        tremarks = DirectCast(.GetData(index, 15), String)
    '                        tclass = DirectCast(.GetData(index, 16), String)
    '                                invoicekey = CInt(.GetData(index, 19))     ------
    '                                invoiced = CDec(.GetData(index, 20))       ------
    '                                linvoiced = CDec(.GetData(index, 21))      ------
    '                        paid = CDec(.GetData(index, 22))
    '                        lpaid = CDec(.GetData(index, 23))
    '                    End With

    '                    'added for outstanding balances on 04.03.2008;
    '                    'If eoutstandinginvoices = True Then
    '                    '    If encumbered = paid Then
    '                    '        GoTo BYPASS
    '                    '    End If
    '                    'End If

    '                    rptencumbered += lencumbered
    '                    rptinvoiced += linvoiced
    '                    rptpaid += lpaid

    '                    If tponumber <> prevponum Then
    '                        'linefeed extra space between po's if not top of page;
    '                        If Not topofpage Then y += 3

    '                        'If tponumber = "00001316" Then Stop       --------

    '                        If invoiced = 0 And paid > 0 Then paidbyinvoice = False
    '                        If invoiced = 0 And paid = 0 Then paidbyinvoice = True
    '                        If invoiced > 0 Then paidbyinvoice = True
    '                        If Not paidbyinvoice Then .RenderDirectText(-1, y + 0.5, "*", 5, 5, specstyle)

    '                        'calculate the balance;
    '                        balance = Math.Round(encumbered - paid, 2)
    '                        'If balance <> 0 Then
    '                        .RenderDirectText(1, y, tponumber, 20, 5, verdanaleft8)
    '                        .RenderDirectText(18, y, tissuedate.ToShortDateString, 20, 5, verdanaright8)
    '                        .RenderDirectText(38, y, tvendor, 45, 5, verdanaleft8)
    '                        'render the header balances;
    '                        .RenderDirectText(95, y, encumbered.ToString.Format("{0:F2}", encumbered), 20, 5, verdanaright8)
    '                        If paidbyinvoice Then .RenderDirectText(120, y, invoiced.ToString.Format("{0:F2}", invoiced), 20, 5, verdanaright8)
    '                        .RenderDirectText(145, y, paid.ToString.Format("{0:F2}", paid), 20, 5, verdanaright8)
    '                        'calculate the balance;
    '                        'balance = Math.Round(encumbered - paid, 2)
    '                        .RenderDirectText(165, y, balance.ToString.Format("{0:F2}", balance), 25, 5, verdanaright8bold)
    '                        topofpage = False
    '                        'tsumamt += thdramt
    '                        y += 5
    '                    End If
    '                    'End If
    '                    .RenderDirectText(8, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
    '                    'If Me.UseOcas Then .RenderDirectText(30, y, tclass, 60, 5, verdanaleft8)
    '                    'render the detail balances;
    '                    .RenderDirectText(95, y, lencumbered.ToString.Format("{0:F2}", lencumbered), 20, 5, verdanaright8)
    '                    If paidbyinvoice Then .RenderDirectText(120, y, linvoiced.ToString.Format("{0:F2}", linvoiced), 20, 5, verdanaright8)
    '                    .RenderDirectText(145, y, lpaid.ToString.Format("{0:F2}", lpaid), 20, 5, verdanaright8)
    '                    '''''.RenderDirectText(170, y, balance.ToString.Format("{0:F2}", balance), 20, 5, verdanaright8)

    '                    y += 5
    '                    'get the current ponumber;
    '                    prevponum = tponumber

    '                    If y >= 250 Then
    '                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
    '                        .NewPage()
    '                        'print the total info box left-side;
    '                        .RenderDirectText(5, 34, "For Bank Account:", 40, 5, verdanaright8bold)
    '                        .RenderDirectText(5, 38, Me.BankAccountNumber, 40, 5, verdanaright8)
    '                        .RenderDirectText(3, 51, "(*) - Denotes a non-invoiced payment", 80, 5, specstyle)
    '                        'print line above the column headers;
    '                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
    '                        y = 58
    '                        'print the column headers;
    '                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
    '                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
    '                        .RenderDirectText(38, y, "Vendor", 30, 5, verdanaleft8bold)
    '                        .RenderDirectText(95, y, "Encumbered", 25, 5, verdanaright8bold)
    '                        .RenderDirectText(118, y, "Invoiced", 25, 5, verdanaright8bold)
    '                        .RenderDirectText(145, y, "Paid", 20, 5, verdanaright8bold)
    '                        .RenderDirectText(170, y, "Balance", 20, 5, verdanaright8bold)
    '                        '2nd line of column headers;
    '                        y = 62
    '                        .RenderDirectText(9, y, "Account", 25, 5, verdanaleft8bold)
    '                        topofpage = True
    '                        y = 70
    '                    End If
    '                    'End If
    '                    'expose the current record & count to the caller
    '                    'EventRecordProcessed((reccurrent), reccount)
    'BYPASS:
    '                Next

    '                'print totals
    '                y += 10
    '                If y > 240 Then
    '                    .NewPage()
    '                    y = 65
    '                End If
    '                'draw top of total box
    '                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
    '                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
    '                .RenderDirectText(60, y, "Total Encumbered", 50, 5, verdanaright8bold)
    '                .RenderDirectText(165, y, rptencumbered.ToString.Format("{0:C2}", rptencumbered), 25, 5, verdanaright8bold)
    '                .RenderDirectText(60, y + 4, "Total Invoiced", 50, 5, verdanaright8bold)
    '                .RenderDirectText(165, y + 4, rptinvoiced.ToString.Format("{0:C2}", rptinvoiced), 25, 5, verdanaright8bold)
    '                .RenderDirectText(60, y + 8, "Total Paid", 50, 5, verdanaright8bold)
    '                .RenderDirectText(165, y + 8, rptpaid.ToString.Format("{0:C2}", rptpaid), 25, 5, verdanaright8bold)
    '                rptbalance = Math.Round(rptencumbered - rptpaid, 2)
    '                .RenderDirectText(60, y + 14, "Balance", 50, 5, verdanaright8bold)
    '                .RenderDirectText(165, y + 14, rptbalance.ToString.Format("{0:C2}", rptbalance), 25, 5, verdanaright8bold)
    '                y += 22
    '                'draw bottom of total box
    '                .RenderDirectLine(59, y, 190, y, Color.Black, 0.25)
    '                .RenderDirectLine(59, y + 0.5, 190, y + 0.5, Color.Black, 0.25)
    '            End With
    '        Catch ex As Exception
    '            Throw
    '        End Try

    '        Try
    '            'set the preview zoom
    '            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
    '            'finish the document
    '            Me.Doc1.EndDoc()
    '        Catch ex As Exception
    '            Throw
    '        End Try
    '    End Sub

    Private Sub PrintRequisitionTickets()
        'this routine prints one or more requisitions;
        Me.DocumentName = "RequisitionTicket"
        Me.ReportName = "Activity Fund - Requisition Ticket"
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
        footerstyle = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim count, currow, index, rowindex As Int32
        Dim nextrequisitionkey, requisitionkey As Int32
        Dim haschanged As Boolean

        'At this point, GridWrk contains records for one or more requisitions;
        'We will iterate thru GridWrk and load the detail grid with a single
        'requisition, then process the report for that requisisition only.
        'Then, get the next requisition in the grid until eof;

        With Me.GridWrk
            'initialise the grid and the document;
            Me.GridDetail.DataSource = Nothing
            Me.GridDetail.Rows.Count = 0
            Me.GridDetail.Cols.Count = .Cols.Count
            '
            Me.Doc1.StartDoc()
            '
            For index = 1 To .Rows.Count - 1
                requisitionkey = CInt(.GetData(index, 29))
                If index < .Rows.Count - 1 Then
                    nextrequisitionkey = CInt(.GetData(index + 1, 29))
                Else
                    nextrequisitionkey = 0
                End If

                'map the row;
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
                Me.GridDetail.SetData(currow, 19, Me.GridWrk.GetData(index, 19))
                Me.GridDetail.SetData(currow, 20, Me.GridWrk.GetData(index, 20))
                Me.GridDetail.SetData(currow, 21, Me.GridWrk.GetData(index, 21))
                Me.GridDetail.SetData(currow, 22, Me.GridWrk.GetData(index, 22))
                Me.GridDetail.SetData(currow, 23, Me.GridWrk.GetData(index, 23))
                Me.GridDetail.SetData(currow, 24, Me.GridWrk.GetData(index, 24))
                Me.GridDetail.SetData(currow, 25, Me.GridWrk.GetData(index, 25))
                Me.GridDetail.SetData(currow, 26, Me.GridWrk.GetData(index, 26))
                Me.GridDetail.SetData(currow, 27, Me.GridWrk.GetData(index, 27))
                Me.GridDetail.SetData(currow, 28, Me.GridWrk.GetData(index, 28))
                Me.GridDetail.SetData(currow, 29, Me.GridWrk.GetData(index, 29))
                Me.GridDetail.SetData(currow, 30, Me.GridWrk.GetData(index, 30))
                Me.GridDetail.SetData(currow, 31, Me.GridWrk.GetData(index, 31))
                Me.GridDetail.SetData(currow, 32, Me.GridWrk.GetData(index, 32))
                Me.GridDetail.SetData(currow, 33, Me.GridWrk.GetData(index, 33))
                Me.GridDetail.SetData(currow, 34, Me.GridWrk.GetData(index, 34))
                Me.GridDetail.SetData(currow, 35, Me.GridWrk.GetData(index, 35))

                If requisitionkey <> nextrequisitionkey Then haschanged = True

                If haschanged Then
                    If count >= 1 Then Me.Doc1.NewPage()
                    Call RenderRequisitionTickets()
                    currow = 0
                    haschanged = False
                    Me.GridDetail.Rows.Count = 0
                    count += 1
                Else
                    currow += 1
                End If
            Next
        End With

        Try
            'set the preview zoom;
            Me.Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
            'finish the document;
            Me.Doc1.EndDoc()
        Catch ex As Exception
            Throw
        End Try

    End Sub

    Private Sub PrintReviewerRequisitionReport(ByVal ereviewername As String, ByVal econverter As Int32)
        Me.DocumentName = "RequisitionTicket"
        Me.ReportName = "Activity Fund - Requisition Report"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        specstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        With timesleft16
            'used for the continuation pages;
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Center
            .TextColor = Color.Salmon
        End With

        With specstyle
            'this style is used for the Conditions box printed on the purchase order;
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 6, FontStyle.Regular)
            .TextAlignHorz = AlignHorzEnum.Justify
            .TextColor = Color.Gray
        End With

        Dim tcheckissued, tpoissued, treqapplied, treqissued As Date
        Dim tchecknumber, tponumber, treqnumber As String
        Dim tacctnum, tsubacctnum, tdescription, tremarks, texpcode As String
        Dim tvendname, tvendaddr1, tvendaddr2, tvendcity, tvendstate, tvendzip, tvendfull, tvendphone As String
        Dim treqtype, tstatus, prtcode As String
        Dim tcost, tlineamount, tamount As Decimal
        Dim tfiscalyear, tqty, x, y, index As Int32
        Dim tpurchaseorderkey, tprevkey, treqkey As Int32
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim encumberedcount, paidcount, requisitioncount As Int32
        Dim requisitiontotal As Decimal

        Try
            'fred 2008.04.23;
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    With Me.GridWrk
                        'collect the header information from the first row
                        Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                        tfiscalyear = CInt(.GetData(index, 1))
                        treqnumber = CType(.GetData(index, 2), String)
                        tstatus = CType(.GetData(index, 3), String).ToUpper
                        treqtype = CType(.GetData(index, 4), String).ToUpper
                        tvendname = CType(.GetData(index, 6), String)
                        tdescription = CType(.GetData(index, 7), String)
                        treqapplied = CDate(.GetData(index, 8))
                        treqissued = CDate(.GetData(index, 9))
                        tqty = CInt(.GetData(index, 10))
                        tcost = CDec(.GetData(index, 11))
                        tlineamount = CDec(.GetData(index, 12))
                        tamount = CDec(.GetData(index, 13))
                        tacctnum = CType(.GetData(index, 14), String)
                        tsubacctnum = CType(.GetData(index, 16), String)
                        texpcode = CType(.GetData(index, 18), String)
                        prtcode = FormatExpenditureCode(texpcode)
                        tremarks = CType(.GetData(index, 19), String)
                        tvendaddr1 = CType(.GetData(index, 20), String).Trim
                        tvendaddr2 = CType(.GetData(index, 21), String).Trim
                        tvendcity = CType(.GetData(index, 23), String).Trim
                        tvendstate = CType(.GetData(index, 24), String).Trim
                        tvendzip = CType(.GetData(index, 25), String).Trim
                        tvendfull = tvendcity & ", " & tvendstate & " " & tvendzip
                        If tvendfull.Trim = "," Then tvendfull = ""
                        tvendphone = CType(.GetData(index, 27), String)
                        treqkey = CInt(.GetData(index, 29))
                        tpurchaseorderkey = CInt(.GetData(index, 31))
                        tponumber = CType(.GetData(index, 32), String).Trim
                        tpoissued = CDate(.GetData(index, 33))
                        tchecknumber = CType(.GetData(index, 34), String).Trim
                        tcheckissued = CDate(.GetData(index, 35))
                        requisitiontotal += tlineamount
                    End With

                    If index = 1 Then
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Requisition Report", 90, 5, verdanaright10bold)
                        If econverter = 0 Then
                            .RenderDirectText(100, y + 5, "Requisitions for reviewer:", 90, 5, verdanaright10)
                        End If
                        If econverter = 1 Then
                            .RenderDirectText(100, y + 5, "Requisitions for converter:", 90, 5, verdanaright10)
                        End If
                        .RenderDirectText(100, y + 10, ereviewername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Description", 25, 5, verdanaleft8bold)
                        .RenderDirectText(145, y, "Encumbrance", 25, 5, verdanaright8bold)
                        .RenderDirectText(172, y, "Check", 15, 5, verdanaright8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If

                    'print line1 if different requisition;
                    If treqkey <> tprevkey Then
                        If y > 50 Then y += 8
                        .RenderDirectRectangle(0, y, 66, y + 4.25, Color.LightGray, 0.25, Color.LightGray)
                        .RenderDirectText(0, y, treqnumber, 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, treqissued.ToString.Format("{0:MM/dd/yyyy}", treqissued), 24, 5, verdanaleft8)
                        .RenderDirectText(40, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        .RenderDirectText(67, y, tdescription, 65, 10, verdanaleft8)
                        '.RenderDirectText(67, y, tvendname, 65, 10, verdanaleft8)
                        If tpurchaseorderkey > -1 Then
                            .RenderDirectText(145, y, tponumber, 23, 5, verdanaright8bold)
                            .RenderDirectText(165, y, tchecknumber, 25, 5, verdanaright8bold)
                        End If
                        'if po has been deleted;
                        If tpurchaseorderkey = -1 Then
                            verdanaright8bold.TextColor = Color.Maroon
                            .RenderDirectText(120, y, "*** Purchase order deleted ***", 70, 5, verdanaright8bold)
                            verdanaright8bold.TextColor = Color.Black
                        End If
                        'tally the requisition count;
                        requisitioncount += 1
                        'tally the approved requisitions;
                        If tponumber.Length > 0 Then encumberedcount += 1
                        'tally the paid requisitions;
                        If tchecknumber.Length > 0 Then paidcount += 1
                        tprevkey = treqkey
                        y += 5
                    End If

                    'print line2 (detail line);
                    .RenderDirectText(2, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(40, y, tlineamount.ToString.Format("{0:F2}", tlineamount), 25, 5, verdanaright8)
                    .RenderDirectText(120, y, prtcode, 70, 5, verdanaright8)
                    y += 4

                    'check if it's a page break;
                    If y >= 248 Then
                        .NewPage()
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Requisition Report", 90, 5, verdanaright10bold)
                        If econverter = 0 Then
                            .RenderDirectText(100, y + 5, "Requisitions for reviewer:", 90, 5, verdanaright10)
                        End If
                        If econverter = 1 Then
                            .RenderDirectText(100, y + 5, "Requisitions for converter:", 90, 5, verdanaright10)
                        End If
                        .RenderDirectText(100, y + 10, ereviewername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Description", 25, 5, verdanaleft8bold)
                        .RenderDirectText(145, y, "Encumbrance", 25, 5, verdanaright8bold)
                        .RenderDirectText(172, y, "Check", 15, 5, verdanaright8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If
                Next

                'print totals page;
                .NewPage()
                y = 20
                'draw the shadowbox;
                .RenderDirectRectangle(0, y + 14, 190, y + 34.5, Color.LightGray, 0.25, Color.LightGray)
                'render the totals;
                If econverter = 0 Then
                    .RenderDirectText(0, y, "Requisition summary " & "FY-" & tfiscalyear & " for reviewer:", 200, 5, verdanaleft10bold)
                End If
                If econverter = 1 Then
                    .RenderDirectText(0, y, "Requisition summary " & "FY-" & tfiscalyear & " for converter:", 200, 5, verdanaleft10bold)
                End If
                .RenderDirectText(0, y, ereviewername, 190, 5, verdanaright10bold)
                y += 16
                .RenderDirectText(10, y, "Total amount:", 25, 5, verdanaleft8)
                .RenderDirectText(35, y, requisitiontotal.ToString.Format("{0:C2}", requisitiontotal), 30, 5, verdanaright8bold)
                'item counts;
                .RenderDirectText(100, y, "Total items:", 25, 5, verdanaleft8)
                .RenderDirectText(125, y, requisitioncount.ToString.Format("{0:D2}", requisitioncount), 35, 5, verdanaright8bold)
                y += 6
                .RenderDirectText(105, y, "Encumbered items:", 35, 5, verdanaleft8)
                .RenderDirectText(125, y, encumberedcount.ToString.Format("{0:D2}", encumberedcount), 35, 5, verdanaright8bold)
                y += 6
                .RenderDirectText(105, y, "Paid items:", 35, 5, verdanaleft8)
                .RenderDirectText(125, y, paidcount.ToString.Format("{0:D2}", paidcount), 35, 5, verdanaright8bold)
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

    Private Sub PrintUserRequisitionReport(ByVal eusername As String)
        Me.DocumentName = "RequisitionTicket"
        Me.ReportName = "Activity Fund - Requisition Report"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        specstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        With timesleft16
            'used for the continuation pages;
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Center
            .TextColor = Color.Salmon
        End With

        With specstyle
            'this style is used for the Conditions box printed on the purchase order;
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 6, FontStyle.Regular)
            .TextAlignHorz = AlignHorzEnum.Justify
            .TextColor = Color.Gray
        End With

        Dim tcheckissued, tpoissued, treqapplied, treqissued As Date
        Dim tchecknumber, tponumber, treqnumber As String
        Dim tacctnum, tsubacctnum, tdescription, tremarks, texpcode As String
        Dim tvendname, tvendaddr1, tvendaddr2, tvendcity, tvendstate, tvendzip, tvendfull, tvendphone As String
        Dim treqtype, tstatus, prtcode As String
        Dim tcost, tlineamount, tamount As Decimal
        Dim tfiscalyear, tqty, x, y, index As Int32
        Dim tpurchaseorderkey, tprevkey, treqkey As Int32
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim encumberedcount, paidcount, requisitioncount As Int32
        Dim requisitiontotal As Decimal

        Try
            'fred 2008.04.23;
            With Me.Doc1
                .StartDoc()
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    With Me.GridWrk
                        'collect the header information from the first row
                        Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                        tfiscalyear = CInt(.GetData(index, 1))
                        treqnumber = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        treqtype = DirectCast(.GetData(index, 4), String).ToUpper
                        tvendname = DirectCast(.GetData(index, 6), String)
                        tdescription = DirectCast(.GetData(index, 7), String)
                        treqapplied = CDate(.GetData(index, 8))
                        treqissued = CDate(.GetData(index, 9))
                        tqty = CInt(.GetData(index, 10))
                        tcost = CDec(.GetData(index, 11))
                        tlineamount = CDec(.GetData(index, 12))
                        tamount = CDec(.GetData(index, 13))
                        tacctnum = DirectCast(.GetData(index, 14), String)
                        tsubacctnum = DirectCast(.GetData(index, 16), String)
                        texpcode = DirectCast(.GetData(index, 18), String)
                        prtcode = FormatExpenditureCode(texpcode)
                        tremarks = DirectCast(.GetData(index, 19), String)
                        tvendaddr1 = DirectCast(.GetData(index, 20), String).Trim
                        tvendaddr2 = DirectCast(.GetData(index, 21), String).Trim
                        tvendcity = DirectCast(.GetData(index, 23), String).Trim
                        tvendstate = DirectCast(.GetData(index, 24), String).Trim
                        tvendzip = DirectCast(.GetData(index, 25), String).Trim
                        tvendfull = tvendcity & ", " & tvendstate & " " & tvendzip
                        If tvendfull.Trim = "," Then tvendfull = ""
                        tvendphone = DirectCast(.GetData(index, 27), String)
                        treqkey = CInt(.GetData(index, 29))
                        tpurchaseorderkey = CInt(.GetData(index, 31))
                        tponumber = DirectCast(.GetData(index, 32), String).Trim
                        tpoissued = CDate(.GetData(index, 33))
                        tchecknumber = DirectCast(.GetData(index, 34), String).Trim
                        tcheckissued = CDate(.GetData(index, 35))
                        requisitiontotal += tlineamount
                    End With

                    If index = 1 Then
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Requisition Report", 90, 5, verdanaright10bold)
                        .RenderDirectText(100, y + 5, "Requisitions for submitter:", 90, 5, verdanaright10)
                        .RenderDirectText(100, y + 10, eusername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Description", 25, 5, verdanaleft8bold)
                        .RenderDirectText(145, y, "Encumbrance", 25, 5, verdanaright8bold)
                        .RenderDirectText(172, y, "Check", 15, 5, verdanaright8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If

                    'print line1 if different requisition;
                    If treqkey <> tprevkey Then
                        'If y > 46 Then y += 8
                        If y > 50 Then y += 8
                        .RenderDirectRectangle(0, y, 66, y + 4.25, Color.LightGray, 0.25, Color.LightGray)
                        .RenderDirectText(0, y, treqnumber, 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, treqissued.ToString.Format("{0:MM/dd/yyyy}", treqissued), 24, 5, verdanaleft8)
                        .RenderDirectText(40, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        'if description is too long to print, then truncate;
                        If tdescription.Length > 60 Then tdescription = tdescription.Substring(0, 60) & "..."
                        .RenderDirectText(67, y, tdescription, 65, 10, verdanaleft8)
                        If tpurchaseorderkey > -1 Then
                            .RenderDirectText(145, y, tponumber, 23, 5, verdanaright8bold)
                            .RenderDirectText(165, y, tchecknumber, 25, 5, verdanaright8bold)
                        End If
                        'if po has been deleted;
                        If tpurchaseorderkey = -1 Then
                            verdanaright8bold.TextColor = Color.Maroon
                            .RenderDirectText(120, y, "*** Purchase order deleted ***", 70, 5, verdanaright8bold)
                            verdanaright8bold.TextColor = Color.Black
                        End If
                        tprevkey = treqkey
                        'tally the requisition count;
                        requisitioncount += 1
                        'tally the approved requisitions;
                        If tponumber.Length > 0 Then encumberedcount += 1
                        'tally the paid requisitions;
                        If tchecknumber.Length > 0 Then paidcount += 1
                        y += 5
                    End If

                    'print line2 (detail line);
                    .RenderDirectText(2, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(40, y, tlineamount.ToString.Format("{0:F2}", tlineamount), 25, 5, verdanaright8)
                    .RenderDirectText(120, y, prtcode, 70, 5, verdanaright8)
                    y += 4

                    'check if it's a page break;
                    If y >= 250 Then
                        .NewPage()
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Requisition Report", 90, 5, verdanaright10bold)
                        .RenderDirectText(100, y + 5, "Requisitions for submitter:", 90, 5, verdanaright10)
                        .RenderDirectText(100, y + 10, eusername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Description", 25, 5, verdanaleft8bold)
                        .RenderDirectText(145, y, "Encumbrance", 25, 5, verdanaright8bold)
                        .RenderDirectText(172, y, "Check", 15, 5, verdanaright8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If
                Next

                'print totals page;
                .NewPage()
                y = 20
                'draw the shadowbox;
                .RenderDirectRectangle(0, y + 14, 190, y + 34.5, Color.LightGray, 0.25, Color.LightGray)
                'render the totals;
                .RenderDirectText(0, y, "Requisition summary " & "FY-" & tfiscalyear & " for submitter:", 200, 5, verdanaleft10)
                .RenderDirectText(0, y, eusername, 190, 5, verdanaright10bold)
                y += 16
                .RenderDirectText(10, y, "Total amount:", 25, 5, verdanaleft8)
                .RenderDirectText(35, y, requisitiontotal.ToString.Format("{0:C2}", requisitiontotal), 30, 5, verdanaright8bold)
                'item counts;
                .RenderDirectText(100, y, "Total items:", 25, 5, verdanaleft8)
                .RenderDirectText(125, y, requisitioncount.ToString.Format("{0:D2}", requisitioncount), 35, 5, verdanaright8bold)
                y += 6
                .RenderDirectText(105, y, "Encumbered items:", 35, 5, verdanaleft8)
                .RenderDirectText(125, y, encumberedcount.ToString.Format("{0:D2}", encumberedcount), 35, 5, verdanaright8bold)
                y += 6
                .RenderDirectText(105, y, "Paid items:", 35, 5, verdanaleft8)
                .RenderDirectText(125, y, paidcount.ToString.Format("{0:D2}", paidcount), 35, 5, verdanaright8bold)
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

    Private Sub PrintUserRequisitionRejectionReport(ByVal eusername As String)
        Me.DocumentName = "RequisitionTicket"
        Me.ReportName = "Activity Fund - Requisition Report"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        specstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        With timesleft16
            'used for the continuation pages;
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Center
            .TextColor = Color.Salmon
        End With

        With specstyle
            'this style is used for the Conditions box printed on the purchase order;
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 6, FontStyle.Regular)
            .TextAlignHorz = AlignHorzEnum.Justify
            .TextColor = Color.Gray
        End With

        Dim tcommentdate, treqissued As Date
        Dim tcomment, treqnumber, treviewer As String
        Dim treqtype, tstatus, prtcode As String
        Dim tamount As Decimal
        Dim prequisitionkey, requisitionkey, tapproval, tfiscalyear, treqlevel, x, y, index As Int32
        Dim requisitioncount As Int32
        Dim SSQL As String
        Dim cmd As SqlCommand

        Try
            With Me.Doc1
                .StartDoc()
                'get a count of the number of requisitions;
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    With Me.GridWrk
                        requisitionkey = CType(.GetData(index, 0), Int32)
                        If requisitionkey <> prequisitionkey Then
                            requisitioncount += 1
                        End If
                        prequisitionkey = requisitionkey
                    End With
                Next

                'reinitialize the prior key;
                prequisitionkey = -1
                For index = 1 To Me.GridWrk.Rows.Count - 1
                    With Me.GridWrk
                        'collect the header information from the first row
                        requisitionkey = CType(.GetData(index, 0), Int32)
                        treqnumber = DirectCast(.GetData(index, 1), String)
                        treqissued = CDate(.GetData(index, 2))
                        tamount = DirectCast(.GetData(index, 3), Decimal)
                        treviewer = DirectCast(.GetData(index, 4), String)
                        treqtype = DirectCast(.GetData(index, 5), String).ToUpper
                        treqlevel = CType(.GetData(index, 6), Int32)
                        tapproval = CType(.GetData(index, 7), Int32)
                        tcomment = DirectCast(.GetData(index, 8), String).Trim
                        If tcomment.Length < 1 Then tcomment = "** No comment provided **"
                        tcommentdate = CDate(.GetData(index, 9))
                    End With

                    If index = 1 Then
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Rejected Requisition Report", 90, 5, verdanaright10bold)
                        .RenderDirectText(100, y + 5, "Rejected for submitter:", 90, 5, verdanaright10)
                        .RenderDirectText(100, y + 10, eusername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(37, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Reviewer", 25, 5, verdanaleft8bold)
                        .RenderDirectText(93, y, "Rejected", 25, 5, verdanaright8bold)
                        .RenderDirectText(127, y, "Comment", 25, 5, verdanaleft8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If

                    'only print lines if it's a rejected line;
                    If treqlevel = -1 Then
                        'only print req header once;
                        If requisitionkey <> prequisitionkey Then
                            .RenderDirectRectangle(0, y, 63, y + 4.25, Color.LightGray, 0.25, Color.LightGray)
                            .RenderDirectText(0, y, treqnumber, 20, 5, verdanaleft8bold)
                            .RenderDirectText(20, y, treqissued.ToString.Format("{0:MM/dd/yyyy}", treqissued), 24, 5, verdanaleft8)
                            .RenderDirectText(37, y, tamount.ToString.Format("{0:F2}", tamount), 25, 5, verdanaright8)
                        End If
                        .RenderDirectText(67, y, treviewer, 65, 5, verdanaleft8)
                        .RenderDirectText(95, y, tcommentdate.ToString.Format("{0:MM/dd/yyyy}", tcommentdate), 25, 5, verdanaright8)
                        .RenderDirectText(125, y, tcomment, 70, 15, verdanaleft8)
                        prequisitionkey = requisitionkey
                        'tally the requisition count;
                        y += 15
                    End If

                    'check if it's a page break;
                    If y >= 250 Then
                        .NewPage()
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Rejected Requisition Report", 90, 5, verdanaright10bold)
                        .RenderDirectText(100, y + 5, "Rejected for submitter:", 90, 5, verdanaright10)
                        .RenderDirectText(100, y + 10, eusername, 90, 5, verdanaright10)
                        'draw line under the title information;
                        y = 42
                        .RenderDirectText(2, y, "Number", 20, 5, verdanaleft8bold)
                        .RenderDirectText(24, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(37, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(67, y, "Reviewer", 25, 5, verdanaleft8bold)
                        .RenderDirectText(93, y, "Rejected", 25, 5, verdanaright8bold)
                        .RenderDirectText(127, y, "Comment", 25, 5, verdanaleft8bold)
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 50
                    End If
                Next

                If y < 250 Then
                    'only print this if there is enough room; no need to create new page;
                    y += 4
                    Dim str As String = String.Format("{0:D1}", requisitioncount) + " total requisitions rejected"
                    .RenderDirectText(2, y, str, 50, 5, verdanaleft10)
                    .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                End If
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

    Private Sub PrintVoidCheckRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   printed
        '     6           7         8        9       10       11  
        '  payee      ponumber   hdramt   lineamt   acct     sub 
        '    12          13        14       15       16       17  
        '  applied    created   hdrdescr  remarks voidappl voidissue
        '    18
        ' voidremark
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VoidCheckRegister"
        Me.ReportName = "Void Check Register"
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

        Dim index, currow, x, y, count As Int32
        Dim totalregister As Double
        Dim tissuedate, tapplieddate As Date
        Dim tacctnum, tsubacctnum, tchknum, tpayee, tdescr, tremarks, tponumber As String
        Dim tprinted, tstatus, trecon, prevchknum, prtstatus As String
        Dim tchkamt, tlineamt, sumamount As Double
        Dim tvoidissued, tvoidapplied As Date
        Dim tvoidremarks As String

        Try
            'get the total amount of the register
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
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(123, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        tchknum = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tprinted = DirectCast(.GetData(index, 5), String)
                        tpayee = DirectCast(.GetData(index, 6), String)
                        tponumber = DirectCast(.GetData(index, 7), String)
                        tchkamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tacctnum = DirectCast(.GetData(index, 10), String)
                        tsubacctnum = DirectCast(.GetData(index, 11), String)
                        tapplieddate = CDate(.GetData(index, 12))
                        tissuedate = CDate(.GetData(index, 13))
                        tdescr = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        tvoidapplied = CDate(.GetData(index, 16))
                        tvoidissued = CDate(.GetData(index, 17))
                        tvoidremarks = DirectCast(.GetData(index, 18), String)
                        sumamount += tlineamt
                    End With

                    If tchknum <> prevchknum Then
                        count += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(1, y, "Voided on", 20, 5, verdanaleft8bold)
                        .RenderDirectText(15, y, tvoidissued.ToShortDateString, 23, 5, verdanaright8)
                        .RenderDirectText(38, y, tvoidremarks, 152, 5, verdanaleft8)
                        y += 5
                        .RenderDirectText(1, y, tchknum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tapplieddate.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(38, y, tpayee, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, tchkamt.ToString.Format("{0:F2}", tchkamt), 25, 5, verdanaright8)
                    End If
                    .RenderDirectText(82, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(96, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(118, y, tponumber, 18, 5, verdanaleft8)
                    .RenderDirectText(136, y, tremarks, 34, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevchknum = tchknum

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(23, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(38, y, "Payee", 25, 5, verdanaleft8bold)
                        .RenderDirectText(83, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(96, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(123, y, "PO#", 25, 5, verdanaleft8bold)
                        .RenderDirectText(136, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        currow = 0
                        y = 65
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 10
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Expenditures", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Checks", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, count.ToString.Format("{0:D2}", count), 25, 5, verdanaright8bold)
                y += 10
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

    Private Sub PrintFooter()
        Try
            'print footer (portrait)
            With Me.Doc1
                .RenderDirectLine(0, 264, 191.5, 264, Color.Black, 0.5)
                .RenderDirectText(0, 265, "Page [@@PageNo@@] of [@@PageCount@@]", 150, 4, footerstyle)
                .RenderDirectText(68, 265, "Activity Fund.Net © - A product of ADPC", 75, 4, footerstyle)
                '.RenderDirectText(167, 265, "1(800)747-2372", 80, 4, footerstyle)
                'changed by fred 2008.04.10;
                .RenderDirectText(152, 265, Now.ToString.Format("{0:MM/dd/yyyy HH:mm:ss tt}", Now), 80, 4, footerstyle)
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PrintHeader()
        Try
            With Me.Doc1
                'print the top margin line 
                .RenderDirectLine(0, 14, 190, 14, Color.Black, 0.5)
                'print the left side of the header
                .RenderDirectText(2, 15, Me.SchoolName, 80, 5, arialleft10)
                .RenderDirectText(2, 20, Me.SchoolAddress1, 80, 5, arialleft10)
                .RenderDirectText(2, 25, Me.SchoolCityStateZip, 80, 5, arialleft10)
                'print the center of the header
                verdanaleft8.TextAlignHorz = AlignHorzEnum.Center
                .RenderDirectText(85, 15, Me.CellMiddleTop, 80, 5, verdanaleft8)
                .RenderDirectText(85, 20, Me.CellMiddleMiddle, 40, 5, verdanaleft8)
                .RenderDirectText(85, 25, Me.CellMiddleBottom, 40, 5, verdanaleft8)
                verdanaleft8.TextAlignHorz = AlignHorzEnum.Left
                'print the right side of the header
                .RenderDirectText(120, 15, Me.ReportName, 70, 5, verdanaright10bold)
                .RenderDirectText(150, 20, Me.CellRightMiddle, 40, 5, verdanaright8)
                .RenderDirectText(150, 25, Now.ToString.Format("{0:MMMM dd, yyyy}", Now), 40, 5, verdanaright8)
                'print the bottom header line
                .RenderDirectLine(0, 31, 190, 31, Color.Gray, 0.5)
            End With
            Call PrintFooter()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RenderRequisitionTickets()
        Me.DocumentName = "RequisitionTicket"
        Me.ReportName = "Activity Fund - Requisition Ticket"
        'define styles
        arialleft8 = New C1DocStyle(Me.Doc1)
        arialright8 = New C1DocStyle(Me.Doc1)
        arialleft10 = New C1DocStyle(Me.Doc1)
        arialleft10bold = New C1DocStyle(Me.Doc1)
        arialright10 = New C1DocStyle(Me.Doc1)
        footerstyle = New C1DocStyle(Me.Doc1)
        specstyle = New C1DocStyle(Me.Doc1)
        timesleft16 = New C1DocStyle(Me.Doc1)
        verdanaleft8 = New C1DocStyle(Me.Doc1)
        verdanaleft8bold = New C1DocStyle(Me.Doc1)
        verdanaright8 = New C1DocStyle(Me.Doc1)
        verdanaright8bold = New C1DocStyle(Me.Doc1)
        verdanaleft10 = New C1DocStyle(Me.Doc1)
        verdanaleft10bold = New C1DocStyle(Me.Doc1)
        verdanaright10 = New C1DocStyle(Me.Doc1)
        verdanaright10bold = New C1DocStyle(Me.Doc1)
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        With timesleft16
            'used for the continuation pages;
            .Font = New Font("Arial", 10, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Center
            .TextColor = Color.Salmon
        End With

        With specstyle
            'this style is used for the Conditions box printed on the purchase order;
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 6, FontStyle.Regular)
            .TextAlignHorz = AlignHorzEnum.Justify
            .TextColor = Color.Gray
        End With

        Dim cond1 As String = "1. Invoices to be rendered in duplicate."
        Dim cond2 As String = "2. No payment to be made until order is complete."
        Dim cond3 As String = "3. Goods to be delivered F.O.B. as per address in upper left."
        Dim cond4 As String = "4. Exempt from sales tax per state statute."
        Dim cond5 As String = "5. Deliveries acknowledge subject to Purchaser's inspection."

        Dim tcheckissued, tpoissued, treqapplied, treqissued As Date
        Dim tchecknumber, tponumber, treqnumber As String
        Dim tacctnum, tsubacctnum, tdescription, tremarks, texpcode As String
        Dim tvendname, tvendaddr1, tvendaddr2, tvendcity, tvendstate, tvendzip, tvendfull, tvendphone As String
        Dim treqtype, tstatus, prtcode As String
        Dim tcost, tlineamount, tamount As Decimal
        Dim tfiscalyear, tqty, x, y, index As Int32
        Dim tpurchaseorderkey As Int32
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim multi As Int32

        Try
            'fred 2008.04.17;
            For index = 0 To Me.GridDetail.Rows.Count - 1
                With Me.GridDetail

                    'collect the header information from the first row
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tfiscalyear = CInt(.GetData(index, 1))
                    treqnumber = DirectCast(.GetData(index, 2), String)
                    tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                    treqtype = DirectCast(.GetData(index, 4), String).ToUpper
                    tvendname = DirectCast(.GetData(index, 6), String)
                    tdescription = DirectCast(.GetData(index, 7), String)
                    treqapplied = CDate(.GetData(index, 8))
                    treqissued = CDate(.GetData(index, 9))
                    tqty = CInt(.GetData(index, 10))
                    tcost = CDec(.GetData(index, 11))
                    tlineamount = CDec(.GetData(index, 12))
                    tamount = CDec(.GetData(index, 13))
                    tacctnum = DirectCast(.GetData(index, 14), String)
                    tsubacctnum = DirectCast(.GetData(index, 16), String)
                    texpcode = DirectCast(.GetData(index, 18), String)
                    prtcode = FormatExpenditureCode(texpcode)
                    tremarks = DirectCast(.GetData(index, 19), String)
                    tvendaddr1 = DirectCast(.GetData(index, 20), String).Trim
                    tvendaddr2 = DirectCast(.GetData(index, 21), String).Trim
                    tvendcity = DirectCast(.GetData(index, 23), String).Trim
                    tvendstate = DirectCast(.GetData(index, 24), String).Trim
                    tvendzip = DirectCast(.GetData(index, 25), String).Trim
                    tvendfull = tvendcity & ", " & tvendstate & " " & tvendzip
                    If tvendfull.Trim = "," Then tvendfull = ""
                    tvendphone = DirectCast(.GetData(index, 27), String)
                    tpurchaseorderkey = CInt(.GetData(index, 31))
                    tponumber = DirectCast(.GetData(index, 32), String).Trim
                    tpoissued = CDate(.GetData(index, 33))
                    tchecknumber = DirectCast(.GetData(index, 34), String).Trim
                    tcheckissued = CDate(.GetData(index, 35))
                    Call GetShippingInformation(0)
                End With

                With Me.Doc1
                    If index = 0 Then
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Requisition", 90, 5, verdanaright10bold)
                        .RenderDirectText(143, y + 5, "Req. Number: ", 30, 5, verdanaleft10)
                        .RenderDirectText(160, y + 5, treqnumber, 30, 5, verdanaright10)
                        .RenderDirectText(130, y + 10, "Issued: " & treqissued.ToShortDateString, 60, 5, verdanaright10)
                        .RenderDirectText(130, y + 15, "Total:  " & tamount.ToString.Format("{0:C2}", tamount), 60, 5, verdanaright10bold)
                        'draw line under the title information;
                        .RenderDirectLine(0, y + 22, 190, y + 22, Color.LightGray, 0.5)
                        'draw the vendor;
                        x = 21
                        y = 44
                        arialleft8.Font = New Font("Arial", 9, FontStyle.Regular)
                        .RenderDirectText(2, y, "To:", 60, 5, arialleft8)
                        .RenderDirectText(x, y, tvendname, 60, 5, arialleft8) : y += 4
                        If tvendaddr1.Length > 0 Then .RenderDirectText(x, y, tvendaddr1, 60, 5, arialleft8) : y += 4
                        If tvendaddr2.Length > 0 Then .RenderDirectText(x, y, tvendaddr2, 60, 5, arialleft8) : y += 4
                        If tvendfull.Length > 0 Then .RenderDirectText(x, y, tvendfull, 60, 5, arialleft8) : y += 4
                        If tvendphone.Length > 0 Then .RenderDirectText(x, y, tvendphone, 60, 5, arialleft8)
                        'draw the ship to;
                        x = 21
                        y = 69
                        Dim shipfull As String = Me.ShippingCity & ", " & Me.ShippingState & " " & Me.ShippingZip
                        If shipfull.Trim = "," Then shipfull = ""
                        .RenderDirectText(2, y, "Ship To:", 60, 5, arialleft8)
                        .RenderDirectText(x, y, Me.ShippingName, 60, 5, arialleft8) : y += 4
                        If Me.ShippingAddress1.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress1, 60, 5, arialleft8) : y += 4
                        If Me.ShippingAddress2.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress2, 60, 5, arialleft8) : y += 4
                        If Me.ShippingAddress3.Length > 0 Then .RenderDirectText(x, y, Me.ShippingAddress3, 60, 5, arialleft8) : y += 4
                        If shipfull.Length > 0 Then .RenderDirectText(x, y, shipfull, 60, 5, arialleft8)
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'draw condition box stuff;
                        x = 128
                        y = 48
                        specstyle.TextColor = Color.Black
                        .RenderDirectText(x, 45, "CONDITIONS", 60, 3, specstyle)
                        specstyle.TextColor = Color.Gray
                        .RenderDirectText(x, 48, cond1, 60, 3, specstyle)
                        .RenderDirectText(x, 51, cond2, 60, 3, specstyle)
                        .RenderDirectText(x, 54, cond3, 60, 3, specstyle)
                        .RenderDirectText(x, 57, cond4, 60, 3, specstyle)
                        .RenderDirectText(x, 60, cond5, 60, 3, specstyle)
                        'draw box around the conditions
                        .RenderDirectRectangle(125.5, y - 4, 189.5, y + 15.5, Color.LightGray, 0.5, Color.Transparent)
                        'draw the signature line 
                        x = 128
                        y = 88

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '''''''''''''''''''''Draw the signature image ''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '''''If Me.DoSignatures Then
                        '''''    Dim imgalign As New C1.C1PrintDocument.ImageAlignDef
                        '''''    imgalign.AlignHorz = ImageAlignHorzEnum.Left
                        '''''    imgalign.StretchHorz = True
                        '''''    imgalign.StretchVert = True
                        '''''    imgalign.KeepAspectRatio = True
                        '''''    'if primary signature image is available, then print the image
                        '''''    If Not Me.Signature1 Is Nothing Then Doc1.RenderDirectImage(128, 76, Me.Signature1, 300, 14, imgalign)
                        '''''    'print the name of the primary signer under the first line
                        '''''    specstyle.TextAlignHorz = AlignHorzEnum.Right
                        '''''    .RenderDirectText(150, y + 1, Me.SignatureTextLine1, 39, 5, specstyle)
                        '''''    'if secondary signature image is available, then print the image
                        '''''    'If Not Me.Signature2 Is Nothing Then Doc1.RenderDirectImage(101, 40, Me.Signature2, 300, 14, imgalign)
                        '''''    'print the name of the secondary signer under the second line
                        '''''    '.RenderDirectText(101, 51, Me.SignatureTextLine2, 80, 5, arialleft8)
                        '''''    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '''''End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '''''''''''''''''''''End of signature image ''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        x = 128
                        y = 101

                        If tpurchaseorderkey = 0 Then
                            .RenderDirectText(x, y - 5, "*** NOT AN APPROVED PURCHASE ***", 80, 5, arialleft8)
                        End If
                        'draw the signature line;
                        .RenderDirectLine(x, y, 189, y, Color.Gray, 1.0)
                        specstyle.TextAlignHorz = AlignHorzEnum.Left
                        .RenderDirectText(x, y + 1, "Purchase approved by", 60, 4, specstyle)
                        .RenderDirectText(x, y + 4, "FY-" & tfiscalyear.ToString, 60, 4, specstyle)
                        specstyle.TextAlignHorz = AlignHorzEnum.Right
                        .RenderDirectText(x + 10, y + 4, "SCHOOL ACTIVITY FNDS - 60", 51, 4, specstyle)
                        .RenderDirectText(x + 23, y + 1, "Activity Fund Custodian", 38, 5, specstyle)
                        'draw the description;
                        x = 21
                        y = 95
                        .RenderDirectText(2, y, "Description:", 20, 5, arialleft8)
                        .RenderDirectText(x, y, tdescription, 100, 15, arialleft8)

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''' Purchase order & check information '''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        'render the purchase order & check information, if available;
                        If tpurchaseorderkey > 0 Then
                            y = 69
                            .RenderDirectText(128, y, "Purchase order issued:", 50, 5, verdanaleft8bold)
                            .RenderDirectText(128, y + 4, tponumber, 20, 5, verdanaleft8)
                            .RenderDirectText(147, y + 4, tpoissued.ToString.Format("{0:MMM dd, yyyy}", tpoissued), 25, 5, verdanaright8)
                            'check available;
                            If tchecknumber.Length > 0 Then
                                y = 78
                                .RenderDirectText(128, y, "Check issued:", 50, 5, verdanaleft8bold)
                                .RenderDirectText(128, y + 4, tchecknumber, 20, 5, verdanaleft8)
                                .RenderDirectText(147, y + 4, tcheckissued.ToString.Format("{0:MMM dd, yyyy}", tcheckissued), 25, 5, verdanaright8)
                            End If
                            'no check available;
                            If tchecknumber.Length = 0 Then
                                y = 78
                                .RenderDirectText(128, y, "No check issued", 50, 5, verdanaleft8bold)
                            End If
                        End If

                        If tpurchaseorderkey = -1 Then
                            'purchase order was issued but later deleted;
                            y = 69
                            .RenderDirectText(128, y, "Purchase order deleted", 50, 5, verdanaleft8bold)
                            .RenderDirectText(128, y + 9, "The requisition is closed and locked", 100, 5, verdanaleft8bold)
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        y = 105
                        specstyle.TextAlignHorz = AlignHorzEnum.Left
                        .RenderDirectText(2, y, "For applied period " & treqapplied.ToString.Format("{0:MMMM, yyyy}", treqapplied), 75, 5, specstyle)
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        y = 111
                        'draw line under the po header information;
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        'draw the column headers;
                        .RenderDirectText(2, y, "Account", 20, 5, arialleft8)
                        .RenderDirectText(18, y, "Expenditure coding", 30, 5, arialleft8)
                        .RenderDirectText(80, y, "Remarks", 20, 5, arialleft8)
                        arialleft8.TextAlignHorz = AlignHorzEnum.Right
                        .RenderDirectText(140, y, "Qty", 10, 5, arialleft8)
                        .RenderDirectText(150, y, "Cost", 20, 5, arialleft8)
                        .RenderDirectText(170, y, "Amount", 20, 5, arialleft8)
                        arialleft8.TextAlignHorz = AlignHorzEnum.Left
                        'draw line under the column headers;
                        y = 116
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        y = 119
                    End If
                    'print the detail information;
                    .RenderDirectText(1, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(18, y, prtcode, 70, 5, verdanaleft8)
                    .RenderDirectText(80, y, tremarks, 62, 10, verdanaleft8)
                    .RenderDirectText(140, y, tqty.ToString.Format("{0:D2}", tqty), 10, 5, verdanaright8)
                    .RenderDirectText(150, y, tcost.ToString.Format("{0:F2}", tcost), 20, 5, verdanaright8)
                    .RenderDirectText(165, y, tlineamount.ToString.Format("{0:F2}", tlineamount), 25, 5, verdanaright8)
                    'check if it's a page break;
                    If y > 245 Then
                        .NewPage()
                        'do the header stuff;
                        y = 20
                        'left side;
                        .RenderDirectText(2, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(2, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        .RenderDirectText(2, y + 15, Me.SchoolTelephone, 100, 5, verdanaleft10)
                        'right side;
                        .RenderDirectText(100, y, "Activity Fund Purchase Order", 90, 5, verdanaright10)
                        .RenderDirectText(143, y + 5, "Req Number :", 30, 5, verdanaleft10)
                        .RenderDirectText(160, y + 5, treqnumber, 30, 5, verdanaright10bold)
                        .RenderDirectText(130, y + 10, "Issued: " & treqissued.ToShortDateString, 60, 5, verdanaright10)
                        .RenderDirectText(130, y + 15, "Total:  " & tamount.ToString.Format("{0:C2}", tamount), 60, 5, verdanaright10bold)
                        'draw continuation information;
                        .RenderDirectText(0, y + 15, "Continued from previous page...", 190, 10, timesleft16)
                        'draw line under the title information;
                        .RenderDirectLine(0, y + 22, 190, y + 22, Color.LightGray, 0.5)
                        'draw the column headers;
                        y = 42
                        .RenderDirectText(2, y, "Account", 20, 5, arialleft8)
                        .RenderDirectText(18, y, "Expenditure coding", 30, 5, arialleft8)
                        .RenderDirectText(80, y, "Remarks", 20, 5, arialleft8)
                        arialleft8.TextAlignHorz = AlignHorzEnum.Right
                        .RenderDirectText(140, y, "Qty", 10, 5, arialleft8)
                        .RenderDirectText(150, y, "Cost", 20, 5, arialleft8)
                        .RenderDirectText(170, y, "Amount", 20, 5, arialleft8)
                        arialleft8.TextAlignHorz = AlignHorzEnum.Left
                        'draw line under the column headers;
                        y = 47
                        .RenderDirectLine(0, y, 190, y, Color.LightGray, 0.5)
                        'subtract here since you're about to add 8 again;
                        y = 42
                    End If
                    y += 8
                End With
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function SavePositivePay(ByVal efilename As String) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Writes a positive pay file, which is a list of checks by register number. 
        'Added on 2016.08.01, Fred for request by Gayle@Western Heights;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'THIS FILE USES THE BANK STANDARD FORMAT PROVIDED BY BOK;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '  0-04	    Filler
        '  5-14	    Bank account (Right justified, zero filled)
        ' 25-34     Check number (Right justified, zero filled)
        ' 40-51     Amount (Right justified, no decimal, zero filled)
        ' 53        Void indicator (V only)
        ' 55-62     Issued MMddyyyy
        ' 75-134    Payee (Left justifed, blank filled)
        '136-195    Payee Line 2 (Optional, blank filled)
        '197-256    Payee Line 3 (Optional, blank filled)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'THIS FILE USES THE FORMAT FOR THE TREASURER'S WARRFLE.TXT FILE (NOT IN USE);
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '01	    3 (Hardcoded symbol)
        '02-26	Payee
        '27-31 	Date	(7/13/2010 = 07130) MMDDY or 2 digit month + 2 digit day + 1 digit ending year;
        '32-36	Encumbrance
        '37-41	Check number
        '42-50	Check amount (zero filled)	$14517.75 = 001451775
        '51	    + (Hardcoded symbol)
        '52-77	Account code (26 chars)
        '78-80	School ID (3 chars) : 338 = Hugo
        '81	    Filler (Hardcoded space)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0        1         2         3         4          5          6         7
        '   bank    fisyr     number    status     payee     amount     issued   register
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim response As DialogResult
        Dim currentpath As String = ""

        Try
            'Call to registry to get paths;
            Call GetRegistry()
            'Display save dialog window;
            With Me.SaveFileDialog1
                .Filter = "Dat files (*.txt)|*.txt|All files (*.*)|*.*"
                .FileName = efilename
                .InitialDirectory = Module1.SaveFilePath
                'Display dialog;
                response = .ShowDialog()
                If response <> DialogResult.OK Then Return False
                'Set the returned file path;
                currentpath = .FileName
                'Store the new path;
                Module1.SaveFilePath = Path.GetDirectoryName(currentpath)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, MSGTITLE)
            Return False
        End Try



        'Write the positive pay file;
        Dim fs As FileStream
        Dim sw As StreamWriter
        Try
            'create a reference to a filestream;
            fs = New FileStream(currentpath, FileMode.Create)
            'open the filestream;
            sw = New StreamWriter(fs)



            Dim index, holdindex, number As Int32
            Dim issued As Date
            Dim bankaccount, payee, record, stramount, voidsw As String
            Dim mm, dd, yyyy As String
            Dim cents, dollars As String


            For index = 1 To Me.GridDetail.Rows.Count - 1
                With Me.GridDetail
                    bankaccount = CType(.GetData(index, 0), String)
                    number = CType(.GetData(index, 2), Int32)
                    voidsw = CType(.GetData(index, 3), String)
                    payee = CType(.GetData(index, 4), String).Trim
                    stramount = CType(.GetData(index, 5), String)
                    issued = CType(.GetData(index, 6), Date)
                End With
                'Test bank account number (10 character limit)
                If bankaccount.Length > 10 Then bankaccount.Substring(0, 10)
                'Test void switch;
                If voidsw <> "V" Then voidsw = Space(1)
                'Truncate payee if too long;
                If payee.Length > 25 Then payee = payee.Substring(0, 25)
                'Parse the date field;
                mm = String.Format("{0:D2}", issued.Month)
                dd = String.Format("{0:D2}", issued.Day)
                yyyy = issued.Year.ToString
                'Parse the decimal field;
                holdindex = stramount.LastIndexOf(".")
                dollars = stramount.Substring(0, holdindex)
                cents = stramount.Substring(holdindex + 1, stramount.Length - (holdindex + 1))
                stramount = String.Concat(dollars, cents).PadLeft(12, "0"c)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'THIS FILE USES THE BANK STANDARD FORMAT PROVIDED BY BOK;
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '  0-04	    Filler
                '  5-14	    Bank account (Right justified, blank or zero filled)
                ' 25-34     Check number (Right justified, zero filled)
                ' 40-51     Amount (Right justified, no decimal, zero filled)
                ' 53        Void indicator (V only)
                ' 55-62     Issued MMddyyyy
                ' 75-134    Payee (Left justifed, blank filled)
                '136-195    Payee Line 2 (Optional, blank filled)
                '197-256    Payee Line 3 (Optional, blank filled)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Build the record;
                record = Space(4)
                record += bankaccount.PadLeft(10)
                record += Space(10)
                record += String.Format("{0:D10}", number)
                record += Space(5)
                record += stramount
                record += Space(1)
                record += voidsw
                record += Space(1)
                record += mm + dd + yyyy
                record += Space(12)
                record += payee.PadRight(60)
                record += Space(1)
                record += Space(60)
                record += Space(1)
                record += Space(60)

                'Zero fill example;
                'record += String.Format("{0:D5}", 0)        '32-36 PO (N/A)

                'Write the record;
                sw.WriteLine(record)
            Next
        Catch ex As Exception
            Return False
        Finally
            sw.Close()
            fs.Close()
        End Try


        Try
            'Save the filepath settings to the registry;
            Module1.SetRegistry()
        Catch ex As Exception
            Return False
        End Try

        '''''''''''''''''''''''''''''''''''''
        'Success
        '''''''''''''''''''''''''''''''''''''
        Return True
    End Function

#End Region

#Region "  Methods Retrieval "

    Private Sub GetShippingInformation(ByVal eshippingkey As Int32)
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Try
            If eshippingkey = 0 Then
                SSQL = "SELECT TOP 1 ship_name, ship_addr1, ship_addr2, ship_addr3," _
                & " ship_city, ship_state, ship_zip, ship_zip_ext" _
                & " FROM shipping"
            Else
                SSQL = "SELECT TOP 1 ship_name, ship_addr1, ship_addr2, ship_addr3," _
                & " ship_city, ship_state, ship_zip, ship_zip_ext" _
                & " FROM shipping" _
                & " WHERE ship_autoinc_key = " & eshippingkey
            End If
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("signatures")
            cn.Open()
            da.Fill(tbl)
            With tbl
                If .Rows.Count = 1 Then
                    Me.ShippingName = CStr(.Rows(0).Item(0))
                    Me.ShippingAddress1 = CStr(.Rows(0).Item(1))
                    Me.ShippingAddress2 = CStr(.Rows(0).Item(2))
                    Me.ShippingAddress3 = CStr(.Rows(0).Item(3))
                    Me.ShippingCity = CStr(.Rows(0).Item(4))
                    Me.ShippingState = CStr(.Rows(0).Item(5))
                    Me.ShippingZip = CStr(.Rows(0).Item(6))
                End If
            End With
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try
    End Sub

    Private Sub GetSignatureDetails()
        'this method retrieves all signature details as saved in school info
        '& returns a datatable cast as a generic object
        Dim SSQL As String
        SSQL = "SELECT sign_title, sign_fname, sign_mi, sign_lname," _
        & " sign_signature, sign_po_sw FROM signatures" _
        & " WHERE sign_po_sw = 'Y'" _
        & " ORDER BY sign_autoinc_key"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("signatures")
        Try
            cn.Open()
            da.Fill(tbl)
        Catch ex As Exception
            'if sigs not found, then exit
            Exit Sub
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'load the signature information into the vars for later user by the 
        'print mechanism
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1          2          3          4        5
        ' sign_title  sign_fname  sign_mi  sign_lname  signatures  posw
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If tbl.Rows.Count < 1 Then Exit Sub
        Dim currow As Int32
        Dim arrayimage1 As Byte()
        Dim title, fname, mi, lname, sw As String
        Dim ms As MemoryStream

        With tbl
            Me.DoSignatures = False

            For currow = 0 To .Rows.Count - 1
                sw = DirectCast(.Rows(currow).Item(5), String).Trim
                If sw.ToUpper = "Y" Then Me.DoSignatures = True
                If Not Me.DoSignatures Then Exit Sub

                title = DirectCast(.Rows(currow).Item(0), String)
                fname = DirectCast(.Rows(currow).Item(1), String).Trim
                mi = DirectCast(.Rows(currow).Item(2), String).Trim
                lname = DirectCast(.Rows(currow).Item(3), String).Trim

                Try
                    'get signature 1
                    If currow = 0 Then
                        SignatureTextLine1 = fname + " " + mi + " " + lname
                        'if no image is available, then continue
                        arrayimage1 = CType(.Rows(currow).Item(4), Byte())
                        ms = New MemoryStream(arrayimage1)
                        If ms.Length > 0 Then Me.Signature1 = Image.FromStream(ms)
                    End If
                Catch ex As Exception

                End Try

                Try
                    'get signature 2
                    If currow = 1 Then
                        SignatureTextLine2 = fname + " " + mi + " " + lname
                        'if no image is available, then continue
                        arrayimage1 = CType(.Rows(currow).Item(4), Byte())
                        ms = New MemoryStream(arrayimage1)
                        If ms.Length > 0 Then Me.Signature2 = Image.FromStream(ms)
                    End If
                Catch ex As Exception

                End Try
            Next
        End With
    End Sub

#End Region

#Region "  Properties "

    Private Property BankAccountNumber() As String
        Get
            Return p_bankaccountnumber
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
            p_bankaccountnumber = s2
        End Set
    End Property

#End Region

End Class
