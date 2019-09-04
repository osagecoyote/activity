Imports C1.C1PrintDocument
Imports C1.Win.C1FlexGrid
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Public Class frmRevenueReports
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
            Me.UseOcas = authobj.UseOCAS
            Me.SchoolName = authobj.SchoolName
            Me.SchoolAddress1 = authobj.SchoolAddress1
            Me.SchoolAddress2 = authobj.SchoolAddress2
            Dim city, state, zip As String
            city = authobj.SchoolCity
            state = authobj.SchoolState
            zip = authobj.SchoolZipCode
            Me.SchoolCityStateZip = city & ", " & state & " " & zip
            Me.GridDetail.Visible = False
            Me.GridTotals.Visible = False
            footerstyle = New C1DocStyle(Me.Doc1)
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
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRevenueReports))
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
        Me.Prev1.TabIndex = 0
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
        Me.GridDetail.Location = New System.Drawing.Point(0, 8)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.Rows.Fixed = 0
        Me.GridDetail.Size = New System.Drawing.Size(656, 336)
        Me.GridDetail.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridDetail.TabIndex = 1
        Me.GridDetail.Visible = False
        '
        'GridTotals
        '
        Me.GridTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridTotals.Location = New System.Drawing.Point(32, 8)
        Me.GridTotals.Name = "GridTotals"
        Me.GridTotals.Rows.Fixed = 0
        Me.GridTotals.Size = New System.Drawing.Size(400, 160)
        Me.GridTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridTotals.TabIndex = 2
        Me.GridTotals.Visible = False
        '
        'GridWrk
        '
        Me.GridWrk.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridWrk.BackColor = System.Drawing.SystemColors.Window
        Me.GridWrk.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridWrk.Location = New System.Drawing.Point(0, 8)
        Me.GridWrk.Name = "GridWrk"
        Me.GridWrk.Rows.Fixed = 0
        Me.GridWrk.Size = New System.Drawing.Size(656, 336)
        Me.GridWrk.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrk.TabIndex = 3
        Me.GridWrk.Visible = False
        '
        'GridWrkTotals
        '
        Me.GridWrkTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridWrkTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridWrkTotals.Location = New System.Drawing.Point(32, 24)
        Me.GridWrkTotals.Name = "GridWrkTotals"
        Me.GridWrkTotals.Rows.Fixed = 0
        Me.GridWrkTotals.Size = New System.Drawing.Size(400, 160)
        Me.GridWrkTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridWrkTotals.TabIndex = 4
        Me.GridWrkTotals.Visible = False
        '
        'frmRevenueReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 373)
        Me.Controls.Add(Me.Prev1)
        Me.Controls.Add(Me.GridWrk)
        Me.Controls.Add(Me.GridDetail)
        Me.Controls.Add(Me.GridTotals)
        Me.Controls.Add(Me.GridWrkTotals)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(300, 200)
        Me.Name = "frmRevenueReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Revenue Reporting"
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
            Case "DepositTicket"
                'deposit ticket does not use the PrintHeader or PrintFooter routine
            Case "ReceiptTicket"
                'receipt ticket does not use the PrintHeader routine
                PrintFooter()
            Case Else
                PrintHeader()
        End Select
    End Sub

#End Region

#Region "  Class Members "

    'styles
    Private docstyle As C1DocStyle
    'Private footerstyle As C1DocStyle
    Private arialleft8 As C1DocStyle
    Private arialright8 As C1DocStyle
    Private arialleft10 As C1DocStyle
    Private arialleft10bold As C1DocStyle
    Private arialright10 As C1DocStyle
    Private arialright10bold As C1DocStyle
    Private footerstyle As C1DocStyle
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
    'property vars
    Private _balanceforwardamount As Double
    Private _balanceforwardcount As Int32
    Private _bankaccountnumber As String
    Private _boldcodefilepath As String
    Private _connectionstring As String
    Private _countyid As String
    Private _districtid As String
    Private _documentname As String
    Private _dosignatures As Boolean
    Private _fiscalyear As Int32
    Private _haserrors As Boolean = False
    Private MSGTITLE As String = "Activity Fund Reporting"
    Private _reportname As String
    Private _schoolname As String
    Private _schooladdress1 As String
    Private _schooladdress2 As String
    Private _schoolcitystatezip As String
    Private _useocas As Boolean
    Private FilePath As String = ""                 'used with the 1098-T revenue report;
    Private SignatureTextLine1 As String = ""
    Private SignatureTextLine2 As String = ""
    Private Signature1 As Image
    Private Signature2 As Image
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
                Case "DepositReport", "DepositSummaryReport", "ReceiptRegister", "ReceiptTicket", "DepositTicket", "VoidReceiptRegister", "OutstandingReceipts"
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
                Case Else
            End Select
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
        With footerstyle
            .Font = New Font("Verdana", 8, FontStyle.Regular)
            .TextColor = Color.Gray
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

    Public Function GenerateDeposit(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal edepositnum As String, ByVal eusecreditcard As Boolean, ByVal eprintdepositticket As Boolean) As Boolean
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank       fisyr    docnumber  status   recon     posted
        '     6           7         8        9       10        11  
        '  applied    created   rcvdfrom  lineamt  paytype   paydescr
        '    12          13        14       15       16        17 
        '  rcptchk      acct      sub    remarks  totalamt depositkey
        '    18          19
        '  depdate    depnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        'get the non-voided receipts for the given deposit number, bank, & year;
        SSQL = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, h.rcpt_status," _
        & " h.rcpt_recon_sw, h.rcpt_posted_sw, h.rcpt_applied_date, h.rcpt_datetime," _
        & " h.rcpt_rcvd_from, d.rcdt_amount, d.rcdt_pymt_type, '' AS paydescr," _
        & " d.rcdt_pymt_chknum, d.af_acct_num, d.as_acct_num, d.rcdt_remarks," _
        & " 0.0 AS totalamount, h.dpst_autoinc_key, p.dpst_transdate, p.dpst_num" _
        & " FROM receipt_info AS h, receipt_detl AS d, deposit_info AS p" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND h.dpst_autoinc_key = p.dpst_autoinc_key" _
        & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2" _
        & " AND p.dpst_num = @p3" _
        & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, CAST(h.rcpt_num AS INT)," _
        & " d.af_acct_num, d.as_acct_num;"
        'get the sum of each payment type for the collected receipts;
        SSQL += "SELECT rcdt_pymt_type, SUM(d.rcdt_amount) AS Amount" _
        & " FROM receipt_info AS h, receipt_detl AS d, deposit_info AS p " _
        & " WHERE h.bank_acct_num = d.bank_acct_num AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num AND h.dpst_autoinc_key = p.dpst_autoinc_key " _
        & " AND (h.rcpt_status <> 'V')" _
        & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2 AND p.dpst_num = @p3" _
        & " GROUP BY rcdt_pymt_type"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", edepositnum)
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
            'throw error if no records returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            If ds.Tables(1).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            'datasource the sums
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Dim index, curtbl, rowindex As Int32
        Dim tpaytype, tpaydescr As String
        Dim tfisyr, tcount As Int32
        Dim tamount, tcashamt, tcheckamt, tcoinamt, tcreditamt, tgrandamt As Double
        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tpaytype = DirectCast(.GetData(index, 10), String)
                tamount = CDbl(.GetData(index, 9))
                tgrandamt += tamount
                Select Case tpaytype
                    Case "1"
                        tpaydescr = "Cash"
                        tcashamt += tamount
                    Case "2"
                        tpaydescr = "Check"
                        tcheckamt += tamount
                    Case "3"
                        tpaydescr = "Coin"
                        tcoinamt += tamount
                    Case "4"
                        tpaydescr = "Credit"
                        tcreditamt += tamount
                End Select
                .SetData(index, 11, tpaydescr)
            Next

            'summarise the detail amounts into the header record
            Dim tcurnum, tnextnum, tholdnum As String
            Dim j, k As Int32
            Dim ttempamt As Double
            'summarise the receipt lines into a header amt
            For index = 0 To .Rows.Count - 1
                tcurnum = DirectCast(.GetData(index, 2), String)
                For j = index To .Rows.Count - 1
                    tholdnum = DirectCast(.GetData(j, 2), String)
                    tamount = CDbl(.GetData(j, 9))
                    If tcurnum = tholdnum Then
                        ttempamt += tamount
                    Else
                        Exit For
                    End If
                Next
                For k = index To j - 1
                    .SetData(k, 16, ttempamt)
                Next
                index = j - 1
                ttempamt = 0
            Next
        End With

        'test only
        '''''With Me.GridDetail
        '''''    .Cols(0).Visible = False
        '''''    .Cols(1).Visible = False
        '''''    .Cols(3).Visible = False
        '''''    .Cols(4).Visible = False
        '''''    .Cols(5).Visible = False
        '''''    .Cols(6).Visible = False
        '''''    .Cols(7).Visible = False
        '''''    .Cols(8).Visible = False
        '''''    .Cols(10).Visible = False
        '''''    .Cols(11).Visible = False
        '''''    .Cols(12).Visible = False
        '''''    .Cols(16).Width = 100
        '''''End With
        Try
            Application.DoEvents()
            'render the table
            If eprintdepositticket Then
                'prints the micr-encoded deposit ticket (currently for Kiamichi only)
                PrintDepositTicket()
            Else
                'prints the daily report
                Me.CellMiddleBottom = "FY-" & efiscalyear
                PrintDeposit()
            End If
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateDepositSummary(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal edepositnumberfrom As String, ByVal edepositnumberto As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '    bank       fisyr    number   amount   remarks   depdate
        '     6           7
        ' rcptcount  depositsum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String

        'changed from p.dpst_datetime to p.dspt_transdate 01-03-2011 pec.

        'SSQL = "SELECT p.bank_acct_num, p.dpst_fisyr, p.dpst_num, p.dpst_amount," _
        '& " p.dpst_remarks, p.dpst_datetime, COUNT(*) AS receiptcount, 0.00 AS TOTAL" _
        '& " FROM deposit_info AS p, receipt_info AS h" _
        '& " WHERE p.dpst_autoinc_key = h.dpst_autoinc_key" _
        '& " AND p.bank_acct_num = @p1" _
        '& " AND p.dpst_fisyr = @p2" _
        '& " AND p.dpst_num BETWEEN @p3 AND @p4" _
        '& " GROUP BY p.bank_acct_num, p.dpst_fisyr, p.dpst_num, p.dpst_amount," _
        '& " p.dpst_remarks, p.dpst_datetime" _
        '& " ORDER BY p.bank_acct_num, p.dpst_fisyr, CAST(p.dpst_num AS INT)"

        SSQL = "SELECT p.bank_acct_num, p.dpst_fisyr, p.dpst_num, p.dpst_amount," _
        & " p.dpst_remarks, p.dpst_transdate, COUNT(*) AS receiptcount, 0.00 AS TOTAL" _
        & " FROM deposit_info AS p, receipt_info AS h" _
        & " WHERE p.dpst_autoinc_key = h.dpst_autoinc_key" _
        & " AND p.bank_acct_num = @p1" _
        & " AND p.dpst_fisyr = @p2" _
        & " AND p.dpst_num BETWEEN @p3 AND @p4" _
        & " GROUP BY p.bank_acct_num, p.dpst_fisyr, p.dpst_num, p.dpst_amount," _
        & " p.dpst_remarks, p.dpst_transdate" _
        & " ORDER BY p.bank_acct_num, p.dpst_fisyr, CAST(p.dpst_num AS INT)"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", edepositnumberfrom)
        cmd.Parameters.Add("@p4", edepositnumberto)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable
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
            'throw error if no records returned
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        'calc the total
        With Me.GridDetail
            Dim currow As Int32
            Dim calcval, totalval As Double
            'sum the amount
            For currow = 0 To .Rows.Count - 1
                calcval = CDbl(.GetData(currow, 3))
                totalval += calcval
            Next
            'set the amount
            For currow = 0 To .Rows.Count - 1
                .SetData(currow, 7, totalval)
            Next
        End With


        ''''''test only
        '''''Me.Prev1.Visible = False
        '''''With Me.GridDetail
        '''''    .Visible = True
        '''''    '.Cols(0).Visible = False
        '''''    '.Cols(1).Visible = False
        '''''End With
        '''''Me.Show()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear
            Me.CellMiddleBottom = edepositnumberfrom & " To " & edepositnumberto
            Application.DoEvents()
            'render the table
            PrintDepositSummary()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateOutstandingReceiptsRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eallfiscalyears As Boolean) As Boolean
        'this method retrieves all outstanding receipts for a single bank
        'for all fiscal years or a selected fiscal year;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   posted
        '     6           7         8        9       10       11  
        '  paytype    paydescr   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   rcvdfrom  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eallfiscalyears Then
            Me.CellMiddleBottom = "All Fiscal Years"
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.bank_acct_num = @p1" _
            & " AND (r.rcpt_posted_sw = 'Y')" _
            & " AND (r.rcpt_recon_sw = 'N')" _
            & " AND (r.rcpt_status <> 'V')" _
            & " ORDER BY r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, d.rcdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outr_rcpt_fisyr, outr_rcpt_num, 'X' AS status," _
            & " outr_recon_sw, '' AS posted, '' AS pymttype, '' AS pymtdescr," _
            & " outr_rcpt_amount, outr_rcpt_amount, af_acct_num, as_acct_num," _
            & " outr_rcpt_issue_date, outr_rcpt_issue_date, outr_rcpt_rcvd_from, outr_rcpt_descr" _
            & " FROM outstandingreceipts" _
            & " WHERE bank_acct_num = @p1" _
            & " AND outr_stale_sw = 'N'" _
            & " AND outr_recon_sw = 'N'" _
            & " ORDER BY bank_acct_num, outr_rcpt_fisyr, outr_rcpt_num"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
        Else
            Me.CellMiddleBottom = "FY-" & efiscalyear.ToString
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.bank_acct_num = @p1 AND r.rcpt_fisyr = @p2" _
            & " AND (r.rcpt_posted_sw = 'Y')" _
            & " AND (r.rcpt_recon_sw = 'N')" _
            & " AND (r.rcpt_status <> 'V')" _
            & " ORDER BY r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, d.rcdt_autoinc_key;"
            SSQL += "SELECT bank_acct_num, outr_rcpt_fisyr, outr_rcpt_num, 'X' AS status," _
            & " outr_recon_sw, '' AS posted, '' AS pymttype, '' AS pymtdescr," _
            & " outr_rcpt_amount, outr_rcpt_amount, af_acct_num, as_acct_num," _
            & " outr_rcpt_issue_date, outr_rcpt_issue_date, outr_rcpt_rcvd_from, outr_rcpt_descr" _
            & " FROM outstandingreceipts" _
            & " WHERE bank_acct_num = @p1" _
            & " AND outr_stale_sw = 'N'" _
            & " AND outr_recon_sw = 'N'" _
            & " ORDER BY bank_acct_num, outr_rcpt_fisyr, outr_rcpt_num"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
        End If
        Dim da As New SqlDataAdapter(cmd)
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
            If (ds.Tables(0).Rows.Count < 1) And (ds.Tables(1).Rows.Count < 1) Then Throw New ArgumentException("No records found for this criteria...")
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridWrk.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Dim index, currow As Int32
        Try
            With Me.GridDetail
                currow = .Rows.Count - 1
                currow += 1
                For index = 0 To Me.GridWrk.Rows.Count - 1
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
        Catch ex As Exception

        End Try

        Dim curtbl, rowindex As Int32
        Dim tstatus, tpaytype, tpaydescr As String
        Dim tfisyr, tcount As Int32
        Dim tamount, tcashamt, tcheckamt, tcoinamt, tcreditamt, tlegacyamt, tgrandamt As Double
        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tpaytype = DirectCast(.GetData(index, 6), String)
                tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                tamount = CDbl(.GetData(index, 9))
                tgrandamt += tamount
                Select Case tpaytype
                    Case "1"
                        tpaydescr = "Cash"
                        tcashamt += tamount
                    Case "2"
                        tpaydescr = "Check"
                        tcheckamt += tamount
                    Case "3"
                        tpaydescr = "Coin"
                        tcoinamt += tamount
                    Case "4"
                        tpaydescr = "Credit"
                        tcreditamt += tamount
                    Case Else
                        tpaydescr = "Other"
                        tlegacyamt += tamount
                End Select
                .SetData(index, 7, tpaydescr)
            Next

            'summarise the detail amounts into the header record
            Dim tcurnum, tnextnum, tholdnum As String
            Dim j, k As Int32
            Dim ttempamt As Double
            For index = 0 To .Rows.Count - 1
                tcurnum = DirectCast(.GetData(index, 2), String)
                For j = index To .Rows.Count - 1
                    tholdnum = DirectCast(.GetData(j, 2), String)
                    tamount = CDbl(.GetData(j, 9))
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
        End With

        With Me.GridTotals
            .Cols.Count = 6
            .Rows.Count = 1
            .Rows.Add()
            .SetData(0, 0, tcashamt)
            .SetData(0, 1, tcheckamt)
            .SetData(0, 2, tcoinamt)
            .SetData(0, 3, tcreditamt)
            .SetData(0, 4, tlegacyamt)
            .SetData(0, 5, tgrandamt)
        End With

        '''''Me.Prev1.Visible = False
        '''''Me.GridDetail.Visible = True
        '''''Me.GridWrk.Visible = True
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Application.DoEvents()
            'render the table
            PrintOutstandingReceiptsRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateReceiptTicket(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal enumber As String) As Boolean
        'this routine is used to fulfill requests from other modules to print a single receipt ticket;
        Try
            GenerateReceiptTickets(ebankaccountnumber, efiscalyear, False, True, Now, Now, enumber, enumber)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateReceiptTickets(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edate1 As Date, ByVal edate2 As Date, ByVal enumber1 As String, ByVal enumber2 As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank       fisyr    docnumber  status   recon     posted
        '     6           7         8        9       10        11  
        '  applied    created   rcvdfrom  paytype paydescr  rcptchk
        '    12          13        14       15       16        17 
        '  lineamt    totalamt    acct     sub    remarks revenuecode
        '    18          19        20       21 
        '  hdrkey     detlkey   acctname subname
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusenumber Then
            'order the receipt by the created date of the detail lines 
            SSQL = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num," _
            & " h.rcpt_status, h.rcpt_recon_sw, h.rcpt_posted_sw," _
            & " h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.rcdt_pymt_type, '' AS paydescr, d.rcdt_pymt_chknum," _
            & " d.rcdt_amount, 0.00 AS Total, d.af_acct_num, d.as_acct_num," _
            & " d.rcdt_remarks, d.ocrv_code, h.rcpt_autoinc_key, d.rcdt_autoinc_key," _
            & " d.af_acct_name, d.as_acct_name" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2" _
            & " AND h.rcpt_num BETWEEN @p3 AND @p4" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, rcdt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumber1)
            cmd.Parameters.Add("@p4", enumber2)
        Else
            'order the receipt by the created date of the detail lines 
            SSQL = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num," _
            & " h.rcpt_status, h.rcpt_recon_sw, h.rcpt_posted_sw," _
            & " h.rcpt_applied_date, h.rcpt_datetime, h.rcpt_rcvd_from," _
            & " d.rcdt_pymt_type, '' AS paydescr, d.rcdt_pymt_chknum," _
            & " d.rcdt_amount, 0.00 AS Total, d.af_acct_num, d.as_acct_num," _
            & " d.rcdt_remarks, d.ocrv_code, h.rcpt_autoinc_key, d.rcdt_autoinc_key," _
            & " d.af_acct_name, d.as_acct_name" _
            & " FROM receipt_info AS h, receipt_detl AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2" _
            & " AND h.rcpt_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, rcdt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edate1)
            cmd.Parameters.Add("@p4", edate2)
        End If
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("receipts")
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
            Me.GridWrk.DataSource = ds.Tables(0)
            'Me.GridWrkTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Dim curnum, holdnum, paytype, paydescr As String
        Dim currow, index, j, k As Int32
        Dim amount, tempamount As Double
        Try
            With Me.GridWrk
                'summarise the detail amounts into the header record
                For index = 0 To .Rows.Count - 1
                    curnum = DirectCast(.GetData(index, 2), String)
                    For j = index To .Rows.Count - 1
                        holdnum = DirectCast(.GetData(j, 2), String)
                        amount = CDbl(.GetData(j, 12))
                        If curnum = holdnum Then
                            tempamount += amount
                        Else
                            Exit For
                        End If
                    Next
                    For k = index To j - 1
                        .SetData(k, 13, tempamount)
                    Next
                    index = j - 1
                    tempamount = 0
                Next

                'set the payment description based on the paycode
                For index = 0 To .Rows.Count - 1
                    paytype = DirectCast(.GetData(index, 9), String)
                    Select Case paytype
                        Case "1"
                            paydescr = "Cash"
                        Case "2"
                            paydescr = "Check"
                        Case "3"
                            paydescr = "Coin"
                        Case "4"
                            paydescr = "Credit"
                    End Select
                    .SetData(index, 10, paydescr)
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        'Me.Prev1.Visible = False
        'Me.GridWrk.Cols(12).Width = 50
        'Me.GridWrk.Cols(13).Width = 50
        'Me.GridWrk.Visible = True
        'Me.ShowDialog()
        'Exit Function

        Try
            Application.DoEvents()
            GetSignatureDetails()
            Application.DoEvents()
            'render the table
            PrintReceiptTickets()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateReceiptRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String, ByVal eusesearch As Boolean, ByVal esearchstring As String) As Boolean
        'this method retrieves all receipts for a single bank given the selected filtering criteria;
        Dim SSQL, filter As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)

        If eusesearch Then filter = " AND rcpt_rcvd_from LIKE '%" & esearchstring & "%'" Else filter = ""

        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.bank_acct_num = @p1 AND r.rcpt_fisyr = @p2" _
            & " AND r.rcpt_applied_date BETWEEN @p3 AND @p4" _
            & filter _
            & " ORDER BY r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, d.rcdt_datetime"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.bank_acct_num = @p1 AND r.rcpt_fisyr = @p2" _
            & " AND r.rcpt_num BETWEEN @p3 AND @p4" _
            & filter _
            & " ORDER BY r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, d.rcdt_datetime"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)
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

        Dim index, curtbl, rowindex As Int32
        Dim tstatus, tpaytype, tpaydescr As String
        Dim tfisyr, tcount As Int32
        Dim tamount, tcashamt, tcheckamt, tcoinamt, tcreditamt, tgrandamt As Double
        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tpaytype = DirectCast(.GetData(index, 6), String)
                tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                tamount = CDbl(.GetData(index, 9))
                If tstatus = "V" Then tamount = 0.0
                tgrandamt += tamount
                Select Case tpaytype
                    Case "1"
                        tpaydescr = "Cash"
                        tcashamt += tamount
                    Case "2"
                        tpaydescr = "Check"
                        tcheckamt += tamount
                    Case "3"
                        tpaydescr = "Coin"
                        tcoinamt += tamount
                    Case "4"
                        tpaydescr = "Credit"
                        tcreditamt += tamount
                End Select
                .SetData(index, 7, tpaydescr)
            Next

            'summarise the detail amounts into the header record
            Dim tcurnum, tnextnum, tholdnum As String
            Dim j, k As Int32
            Dim ttempamt As Double
            For index = 0 To .Rows.Count - 1
                tcurnum = DirectCast(.GetData(index, 2), String)
                For j = index To .Rows.Count - 1
                    tholdnum = DirectCast(.GetData(j, 2), String)
                    tamount = CDbl(.GetData(j, 9))
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
        End With

        With Me.GridTotals
            .Rows.Count = 1
            .Rows.Add()
            .SetData(0, 0, tcashamt)
            .SetData(0, 1, tcheckamt)
            .SetData(0, 2, tcoinamt)
            .SetData(0, 3, tcreditamt)
            .SetData(0, 4, tgrandamt)
        End With

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintReceiptRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Friend Function GenerateReceiptRegisterAllBanks(ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        'this method retrieves all receipts for all banks given the selected filtering criteria;
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellRightMiddle = "Register by date"
            Me.CellRightBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.rcpt_fisyr = @p1" _
            & " AND r.rcpt_applied_date BETWEEN @p2 AND @p3" _
            & " ORDER BY r.bank_acct_num, r.rcpt_datetime"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            cmd.Parameters.Add("@p2", edatefrom)
            cmd.Parameters.Add("@p3", edateto)
        End If
        If eusenumber Then
            Me.CellRightMiddle = "Register by number"
            Me.CellRightBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT r.bank_acct_num, r.rcpt_fisyr, r.rcpt_num, r.rcpt_status," _
            & " r.rcpt_recon_sw, r.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, d.af_acct_num, d.as_acct_num," _
            & " r.rcpt_applied_date, r.rcpt_datetime, r.rcpt_rcvd_from, d.rcdt_remarks" _
            & " FROM receipt_info AS r, receipt_detl AS d" _
            & " WHERE r.bank_acct_num = d.bank_acct_num" _
            & " AND r.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND r.rcpt_num = d.rcpt_num" _
            & " AND r.rcpt_fisyr = @p1" _
            & " AND r.rcpt_num BETWEEN @p2 AND @p3" _
            & " ORDER BY r.bank_acct_num, r.rcpt_datetime"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", efiscalyear)
            cmd.Parameters.Add("@p2", enumberfrom)
            cmd.Parameters.Add("@p3", enumberto)
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

        Dim index, curtbl, rowindex As Int32
        Dim tpaytype, tpaydescr As String
        Dim tfisyr, tcount As Int32
        Dim tamount, tcashamt, tcheckamt, tcoinamt, tcreditamt, tgrandamt As Double
        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tpaytype = DirectCast(.GetData(index, 6), String)
                tamount = CDbl(.GetData(index, 9))
                tgrandamt += tamount
                Select Case tpaytype
                    Case "1"
                        tpaydescr = "Cash"
                        tcashamt += tamount
                    Case "2"
                        tpaydescr = "Check"
                        tcheckamt += tamount
                    Case "3"
                        tpaydescr = "Coin"
                        tcoinamt += tamount
                    Case "4"
                        tpaydescr = "Credit"
                        tcreditamt += tamount
                End Select
                .SetData(index, 7, tpaydescr)
            Next

            'summarise the detail amounts into the header record
            Dim tbanknum, tcurnum, tnextnum, tholdrcpt, tholdbank As String
            Dim j, k As Int32
            Dim ttempamt As Double
            For index = 0 To .Rows.Count - 1
                tbanknum = DirectCast(.GetData(index, 0), String)
                tcurnum = DirectCast(.GetData(index, 2), String)
                For j = index To .Rows.Count - 1
                    tholdbank = DirectCast(.GetData(j, 0), String)
                    tholdrcpt = DirectCast(.GetData(j, 2), String)
                    tamount = CDbl(.GetData(j, 9))
                    If (tbanknum = tholdbank) And (tcurnum = tholdrcpt) Then
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
        End With

        With Me.GridTotals
            .Rows.Count = 1
            .Rows.Add()
            .SetData(0, 0, tcashamt)
            .SetData(0, 1, tcheckamt)
            .SetData(0, 2, tcoinamt)
            .SetData(0, 3, tcreditamt)
            .SetData(0, 4, tgrandamt)
        End With

        Try

            Return True

            Application.DoEvents()
            'render the table
            PrintReceiptRegisterAllBanks()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateVoidReceiptRegister(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eusedate As Boolean, ByVal eusenumber As Boolean, ByVal edatefrom As Date, ByVal edateto As Date, ByVal enumberfrom As String, ByVal enumberto As String) As Boolean
        'this method retrieves all voided receipts for a single bank given the selected filtering criteria;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5          6         7 
        '   bank        fisyr   docnumber  status   recon    posted     paycode  paydescr
        '     8           9        10        11      12        13         14        15  
        '  hdramt     lineamt   applied   created   acct      sub       descr     remarks
        '    16          17        18        19      20        21 
        ' revcode    voidappl  voidissue  vremarks hdrkey   detlkey
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        If eusedate Then
            Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
            SSQL = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, h.rcpt_status," _
            & " h.rcpt_recon_sw, h.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime," _
            & " d.af_acct_num, d.as_acct_num, h.rcpt_rcvd_from, d.rcdt_remarks," _
            & " d.ocrv_code, v.voidrcpt_applied_date, v.voidrcpt_datetime, v.voidrcpt_remarks," _
            & " h.rcpt_autoinc_key, d.rcdt_autoinc_key" _
            & " FROM receipt_info AS h, receipt_detl AS d, voidreceipt AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = v.bank_acct_num" _
            & " AND h.rcpt_fisyr = v.voidrcpt_fisyr" _
            & " AND h.rcpt_num = v.voidrcpt_num" _
            & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2" _
            & " AND v.voidrcpt_applied_date BETWEEN @p3 AND @p4" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, d.rcdt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", edatefrom)
            cmd.Parameters.Add("@p4", edateto)
        End If
        If eusenumber Then
            Me.CellMiddleBottom = enumberfrom & " to " & enumberto
            SSQL = "SELECT h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, h.rcpt_status," _
            & " h.rcpt_recon_sw, h.rcpt_posted_sw, d.rcdt_pymt_type, '' AS PAYTYPE," _
            & " 0.00 AS RcptAmount, d.rcdt_amount, h.rcpt_applied_date, h.rcpt_datetime," _
            & " d.af_acct_num, d.as_acct_num, h.rcpt_rcvd_from, d.rcdt_remarks," _
            & " d.ocrv_code, v.voidrcpt_applied_date, v.voidrcpt_datetime, v.voidrcpt_remarks," _
            & " h.rcpt_autoinc_key, d.rcdt_autoinc_key" _
            & " FROM receipt_info AS h, receipt_detl AS d, voidreceipt AS v" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
            & " AND h.rcpt_num = d.rcpt_num" _
            & " AND h.bank_acct_num = v.bank_acct_num" _
            & " AND h.rcpt_fisyr = v.voidrcpt_fisyr" _
            & " AND h.rcpt_num = v.voidrcpt_num" _
            & " AND h.bank_acct_num = @p1 AND h.rcpt_fisyr = @p2" _
            & " AND h.rcpt_num BETWEEN @p3 AND @p4" _
            & " ORDER BY h.bank_acct_num, h.rcpt_fisyr, h.rcpt_num, d.rcdt_autoinc_key"
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", enumberfrom)
            cmd.Parameters.Add("@p4", enumberto)
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

        Dim index, curtbl, rowindex As Int32
        Dim tstatus, tpaytype, tpaydescr As String
        Dim tfisyr, tcount As Int32
        Dim tamount, tcashamt, tcheckamt, tcoinamt, tcreditamt, tgrandamt As Double
        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tpaytype = DirectCast(.GetData(index, 6), String)
                tamount = CDbl(.GetData(index, 9))
                tgrandamt += tamount
                Select Case tpaytype
                    Case "1"
                        tpaydescr = "Cash"
                        tcashamt += tamount
                    Case "2"
                        tpaydescr = "Check"
                        tcheckamt += tamount
                    Case "3"
                        tpaydescr = "Coin"
                        tcoinamt += tamount
                    Case "4"
                        tpaydescr = "Credit"
                        tcreditamt += tamount
                End Select
                .SetData(index, 7, tpaydescr)
            Next

            'summarise the detail amounts into the header record
            Dim tcurnum, tnextnum, tholdnum As String
            Dim j, k As Int32
            Dim ttempamt As Double
            For index = 0 To .Rows.Count - 1
                tcurnum = DirectCast(.GetData(index, 2), String)
                For j = index To .Rows.Count - 1
                    tholdnum = DirectCast(.GetData(j, 2), String)
                    tamount = CDbl(.GetData(j, 9))
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
        End With

        '''''With Me.GridDetail
        '''''    .Cols(8).Width = 75
        '''''    .Visible = True
        '''''End With
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        With Me.GridTotals
            .Rows.Count = 1
            .Rows.Add()
            .SetData(0, 0, tcashamt)
            .SetData(0, 1, tcheckamt)
            .SetData(0, 2, tcoinamt)
            .SetData(0, 3, tcreditamt)
            .SetData(0, 4, tgrandamt)
        End With

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table
            PrintVoidReceiptRegister()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Friend Function Generate1098TReport(ByVal ecalendaryear As Int32, ByVal eaccountnumber As String, ByVal esubaccountnumber As String) As Boolean
        'this method retrieves receipts for a given calendar year;
        Dim SSQL, filter As String
        Dim response As DialogResult
        Dim cmd As SqlCommand
        Dim date1, date2 As Date
        cn = New SqlConnection(Me.ConnectionString)

        date1 = CDate("01/01/" & ecalendaryear.ToString)
        date2 = CDate("12/31/" & ecalendaryear.ToString)
        date2 = date1.AddYears(1)
        date2 = date2.AddSeconds(-1)

        '''''Me.CellMiddleBottom = edatefrom.ToShortDateString & " to " & edateto.ToShortDateString
        SSQL = "SELECT " & ecalendaryear & ", af_acct_num + '-' + as_acct_num, h.rcpt_rcvd_from, SUM(rcdt_amount) AS total" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND af_acct_num = @p1" _
        & " AND as_acct_num = @p2" _
        & " AND rcpt_applied_date BETWEEN @p3 AND @p4" _
        & " GROUP BY af_acct_num + '-' + as_acct_num, h.rcpt_rcvd_from" _
        & " ORDER BY af_acct_num + '-' + as_acct_num, h.rcpt_rcvd_from"
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", eaccountnumber)
        cmd.Parameters.Add("@p2", esubaccountnumber)
        cmd.Parameters.Add("@p3", date1)
        cmd.Parameters.Add("@p4", date2)
        Dim da As New SqlDataAdapter(cmd)
        Dim tbl As New DataTable("register")
        Try
            da.Fill(tbl)
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected calendar year " & ecalendaryear.ToString & ".")
            Me.GridTotals.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            With Me.SaveFileDialog1
                .Filter = "Comma delimited files (*.csv)|*.csv|All files (*.*)|*.*"
                .FileName = "received.csv"
                response = .ShowDialog()
                If response <> DialogResult.OK Then Exit Function
                'set the file path
                Me.FilePath = .FileName
                Me.GridTotals.SaveGrid(Me.FilePath, FileFormatEnum.TextComma)
                Application.DoEvents()
            End With
            'prompt if a detail report is needed?
            response = MessageBox.Show("Would you like to view a detailed report of the summarized file?", MSGTITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
            If response <> DialogResult.Yes Then Exit Function
        Catch ex As Exception
            Throw
        End Try

        'do the detail report;
        SSQL = "SELECT " & ecalendaryear & ", h.rcpt_num, af_acct_num + '-' + as_acct_num, h.rcpt_applied_date, h.rcpt_rcvd_from, rcdt_amount AS total" _
        & " FROM receipt_info AS h, receipt_detl AS d" _
        & " WHERE h.bank_acct_num = d.bank_acct_num" _
        & " AND h.rcpt_fisyr = d.rcpt_fisyr" _
        & " AND h.rcpt_num = d.rcpt_num" _
        & " AND af_acct_num = @p1" _
        & " AND as_acct_num = @p2" _
        & " AND rcpt_applied_date BETWEEN @p3 AND @p4" _
        & " ORDER BY h.rcpt_rcvd_from, h.rcpt_num"
        cn = New SqlConnection(Me.ConnectionString)
        cmd = New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", eaccountnumber)
        cmd.Parameters.Add("@p2", esubaccountnumber)
        cmd.Parameters.Add("@p3", date1)
        cmd.Parameters.Add("@p4", date2)
        da = New SqlDataAdapter(cmd)
        tbl = New DataTable("register")
        Try
            da.Fill(tbl)
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            Me.CellMiddleMiddle = "Account: " & eaccountnumber & "-" & esubaccountnumber
            Me.CellMiddleBottom = "Calendar- " & ecalendaryear.ToString
            Application.DoEvents()
            'render the table;
            Print1098TReport()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

#End Region

#Region "  Methods Rendering "

    Private Sub PrintDeposit()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NOTE:  this routine prints the daily deposit report and is 
        'similar to the daily report
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank       fisyr    docnumber  status   recon     posted
        '     6           7         8        9       10        11  
        '  applied    created   rcvdfrom  lineamt  paytype   paydescr
        '    12          13        14       15       16        17 
        '  rcptchk      acct      sub    remarks  totalamt depositkey
        '    18          19
        '  depdate    depnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim currow As Int32
        Dim tdepositnum As String
        Dim tdepositdate As Date

        Me.DocumentName = "DepositReport"
        Me.ReportName = "Daily Deposit Report"
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
        'define the styles 
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Try
            With Me.GridDetail
                'get the bank account number from the first item
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
                tdepositnum = DirectCast(.GetData(0, 19), String)
                tdepositdate = CDate(.GetData(0, 18))
            End With
        Catch ex As Exception

        End Try

        ''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''
        '      0            1    
        '   pymttype     totalamt
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim paytype As String
        Dim cashamt, chkamt, coinamt, creditamt, totalamt As Double

        With Me.GridTotals
            For currow = 0 To .Rows.Count - 1
                paytype = DirectCast(.GetData(currow, 0), String)
                Select Case paytype
                    Case "1"    'cash
                        cashamt = CDbl(.GetData(currow, 1))
                    Case "2"    'check  
                        chkamt = CDbl(.GetData(currow, 1))
                    Case "3"    'coin
                        coinamt = CDbl(.GetData(currow, 1))
                    Case "4"    'credit
                        creditamt = CDbl(.GetData(currow, 1))
                End Select
            Next
            totalamt = cashamt + chkamt + coinamt + creditamt
        End With

        Try
            Dim tcreated, tapplied As Date
            Dim tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks As String
            Dim tpaytype, tstatus, trecon, prevtrcptnum As String
            Dim trcptamt, tlineamt As Double
            Dim x, y, index As Int32
            Dim dopagebreak, isvoid As Boolean

            'start the document
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    trcptnum = DirectCast(.GetData(index, 2), String)
                    tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                    trecon = DirectCast(.GetData(index, 4), String).ToUpper
                    tapplied = CDate(.GetData(index, 6))
                    tcreated = CDate(.GetData(index, 7))
                    trcvdfrom = DirectCast(.GetData(index, 8), String)
                    tlineamt = CDbl(.GetData(index, 9))
                    tpaytype = DirectCast(.GetData(index, 11), String)
                    tacctnum = DirectCast(.GetData(index, 13), String)
                    tsubacctnum = DirectCast(.GetData(index, 14), String)
                    tremarks = DirectCast(.GetData(index, 15), String)
                    trcptamt = CDbl(.GetData(index, 16))
                End With

                Select Case tstatus
                    Case "V"
                        isvoid = True
                    Case Else
                        isvoid = False
                End Select

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the total info box left-side
                        .RenderDirectText(x, y, "For Bank Account:", 50, 5, verdanaleft8bold)
                        .RenderDirectText(x, y + 4, Me.BankAccountNumber, 50, 5, verdanaleft8)
                        .RenderDirectText(x, y + 10, "For Deposit Number:", 50, 5, verdanaleft8bold)
                        .RenderDirectText(x, y + 14, tdepositnum, 50, 5, verdanaleft8)
                        .RenderDirectText(x + 20, y + 14, tdepositdate.ToShortDateString, 50, 5, verdanaleft8)
                        'print the total info box right-side
                        x = 138
                        .RenderDirectText(x, y, "Cash:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 4, "Checks:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 8, "Coin:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 12, "Credit Card:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 18, "Total:", 25, 5, verdanaright8)
                        'print the money fields
                        x = 165
                        .RenderDirectText(x, y, cashamt.ToString.Format("{0:F2}", cashamt), 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 4, chkamt.ToString.Format("{0:F2}", chkamt), 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 8, coinamt.ToString.Format("{0:F2}", coinamt), 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 12, creditamt.ToString.Format("{0:F2}", creditamt), 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 18, totalamt.ToString.Format("{0:C2}", totalamt), 25, 5, verdanaright8bold)
                        'print the lines under the info box
                        y = 57
                        .RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        '.RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        'print the column headers
                        y = 58
                        .RenderDirectText(3, y, "Issued", 25, 5, verdanaleft8bold)
                        .RenderDirectText(19, y, "Number", 25, 5, verdanaleft8bold)
                        .RenderDirectText(34, y, "Acct-Sub", 25, 5, verdanaleft8bold)
                        .RenderDirectText(51, y, "Rcvd From", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(141, y, "Amount", 25, 5, verdanaleft8bold)
                        .RenderDirectText(156, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Total", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    'print the detail
                    If trcptnum = prevtrcptnum Then
                        .RenderDirectText(35, y, tacctnum & "-" & tsubacctnum, 50, 5, arialleft8)
                        .RenderDirectText(85, y, tremarks, 50, 5, arialleft8)
                        .RenderDirectText(125, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 30, 5, arialright8)
                        .RenderDirectText(156, y, tpaytype, 15, 5, arialleft8)
                    Else
                        '.RenderDirectText(2, y, tapplied.ToShortDateString, 50, 5, arialleft8)
                        .RenderDirectText(2, y, tcreated.ToShortDateString, 50, 5, arialleft8)
                        .RenderDirectText(19, y, trcptnum, 50, 5, arialleft8)
                        .RenderDirectText(35, y, tacctnum & "-" & tsubacctnum, 50, 5, arialleft8)
                        If isvoid Then
                            verdanaleft8bold.TextColor = Color.Red
                            verdanaright8.TextColor = Color.Red
                            .RenderDirectText(51, y, "VOID", 35, 5, verdanaleft8bold)
                            .RenderDirectText(125, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 30, 5, verdanaright8)
                            .RenderDirectText(170, y, "0.00", 20, 5, verdanaright8)
                            verdanaleft8bold.TextColor = Color.Black
                            verdanaright8.TextColor = Color.Black
                        Else
                            .RenderDirectText(51, y, trcvdfrom, 35, 5, arialleft8)
                            .RenderDirectText(125, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 30, 5, arialright8)
                            .RenderDirectText(170, y, trcptamt.ToString.Format("{0:C2}", trcptamt), 20, 5, arialright8)
                        End If
                        .RenderDirectText(85, y, tremarks, 50, 5, arialleft8)
                        .RenderDirectText(156, y, tpaytype, 15, 5, arialleft8)
                    End If
                    y += 5
                    currow += 1
                    'get the current rcptnumber
                    prevtrcptnum = trcptnum

                    'check for page break & print new column headers if true
                    If .CurrentPage = 1 And currow = 38 Then dopagebreak = True
                    If .CurrentPage > 1 And currow Mod 41 = 0 Then dopagebreak = True
                    If dopagebreak Then
                        .NewPage()
                        y = 32
                        '.RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        '.RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        y = 33
                        .RenderDirectText(3, y, "Issued", 25, 5, verdanaleft8bold)
                        .RenderDirectText(19, y, "Number", 25, 5, verdanaleft8bold)
                        .RenderDirectText(34, y, "Acct-Sub", 25, 5, verdanaleft8bold)
                        .RenderDirectText(51, y, "Rcvd From", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(141, y, "Amount", 25, 5, verdanaleft8bold)
                        .RenderDirectText(156, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Total", 25, 5, verdanaright8bold)
                        y = 39
                        dopagebreak = False
                        currow = 1
                    End If
                End With
                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            Next
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

    Private Sub PrintDepositTicket()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NOTE:  this routine conforms to a micr-encoded pre-printed form
        'supplied by the Kiamichi district;  this form may later be used
        'by other schools if the form is purchased;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank       fisyr    docnumber  status   recon     posted
        '     6           7         8        9       10        11  
        '  applied    created   rcvdfrom  lineamt  paytype   paydescr
        '    12          13        14       15       16        17 
        '  rcptchk      acct      sub    remarks  totalamt depositkey
        '    18          19
        '  depdate    depnum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim currow, rows As Int32
        Dim val0, val1, calcval As Double
        Dim tdepositnum As String
        Dim tdepositdate As Date

        Me.DocumentName = "DepositTicket"
        Me.ReportName = "Activity Fund - Deposit Ticket"
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
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Try
            With Me.GridDetail
                'get the bank account number from the first item
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
                tdepositnum = DirectCast(.GetData(0, 19), String)
                tdepositdate = CDate(.GetData(0, 18))
            End With
        Catch ex As Exception

        End Try

        ''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''
        '      0            1    
        '   pymttype     totalamt
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim paytype As String
        Dim cashamt, chkamt, coinamt, creditamt, otheramt, totalamt As Double
        With Me.GridTotals
            For currow = 0 To .Rows.Count - 1
                paytype = DirectCast(.GetData(currow, 0), String)
                Select Case paytype
                    Case "1"    'cash
                        cashamt = CDbl(.GetData(currow, 1))
                    Case "2"    'check  
                        chkamt = CDbl(.GetData(currow, 1))
                    Case "3"    'coin
                        coinamt = CDbl(.GetData(currow, 1))
                    Case "4"    'credit
                        creditamt = CDbl(.GetData(currow, 1))
                End Select
            Next
            otheramt = coinamt + creditamt
            totalamt = cashamt + chkamt + otheramt
        End With

        Try
            Dim trcptdate As Date
            Dim trcptamt, tlineamt As Double
            Dim x, y, itemcount As Int32
            Dim trcptnum, prevtrcptnum, trcvdfrom, trcptcheck As String
            Dim tpaytype As String

            With Me.Doc1
                'begin the document here
                .StartDoc()
                x = 125
                y = 20
                'get the item count
                itemcount = Me.GridDetail.Rows.Count
                'print the header information on the ticket
                .RenderDirectText(x, y, cashamt.ToString.Format("{0:F2}", cashamt), 40, 5, arialright10)
                .RenderDirectText(x, y + 7, chkamt.ToString.Format("{0:F2}", chkamt), 40, 5, arialright10)
                .RenderDirectText(x, y + 14, otheramt.ToString.Format("{0:F2}", otheramt), 40, 5, arialright10)
                .RenderDirectText(127, y + 21, itemcount.ToString, 40, 5, arialleft10)
                .RenderDirectText(x, y + 27, totalamt.ToString.Format("{0:F2}", totalamt), 40, 5, arialright10)
                'print the date on the header
                .RenderDirectText(22, 47, tdepositdate.ToShortDateString, 40, 5, arialleft10)
                y = 102
                'print the horizontal lines on the stub
                .RenderDirectLine(3, 102, 188, 102, arialleft10)         'top 1
                .RenderDirectLine(3, 107, 188, 107, arialleft10)         'top 2
                .RenderDirectLine(3, 238, 188, 238, arialleft10)         'bottom
                'print the vertical lines on the stub
                .RenderDirectLine(3, 102, 3, 238, arialleft10)           'left1
                .RenderDirectLine(41, 102, 41, 238, arialleft10)         'left2
                .RenderDirectLine(92, 102, 92, 238, arialleft10)         'left3
                .RenderDirectLine(143, 102, 143, 238, arialleft10)       'left4
                .RenderDirectLine(188, 102, 188, 238, arialleft10)       'left5
                'print the column headers
                .RenderDirectText(5, 102.5, "Payment Method", 40, 5, arialleft10)
                .RenderDirectText(43, 102.5, "Received From", 40, 5, arialleft10)
                .RenderDirectText(94, 102.5, "Check/Ref No.", 40, 5, arialleft10)
                .RenderDirectText(173, 102.5, "Amount", 40, 5, arialleft10)

                'now print the lines
                y = 108
                For currow = 0 To Me.GridDetail.Rows.Count - 1
                    rows += 1
                    With Me.GridDetail
                        trcptnum = DirectCast(.GetData(currow, 2), String)
                        trcvdfrom = DirectCast(.GetData(currow, 8), String)
                        tlineamt = CDbl(.GetData(currow, 9))
                        tpaytype = DirectCast(.GetData(currow, 11), String)
                        trcptcheck = DirectCast(.GetData(currow, 12), String)
                        trcptamt = CDbl(.GetData(currow, 16))
                    End With

                    'render the detail line
                    .RenderDirectText(5, y, tpaytype, 40, 5, verdanaleft8)
                    .RenderDirectText(43, y, trcvdfrom, 52, 5, verdanaleft8)
                    .RenderDirectText(94, y, trcptcheck, 40, 5, verdanaleft8)
                    .RenderDirectText(146, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 40, 5, verdanaright8)
                    y += 5

                    If currow = (Me.GridDetail.Rows.Count - 1) Then
                        'last record
                        .RenderDirectLine(156, y - 0.5, 186, y - 0.5, verdanaright8)
                        .RenderDirectText(146, y, totalamt.ToString.Format("{0:F2}", totalamt), 40, 5, verdanaright8bold)
                    End If

                    'check for page break
                    If rows Mod 26 = 0 Then
                        .NewPage()
                        'print special text for a continued page on the header
                        .RenderDirectText(128, 10, "C", 80, 10, timesleft16)
                        .RenderDirectText(131, 15, "O", 80, 10, timesleft16)
                        .RenderDirectText(134, 20, "N", 80, 10, timesleft16)
                        .RenderDirectText(137, 25, "T", 80, 10, timesleft16)
                        .RenderDirectText(140, 30, "I", 80, 10, timesleft16)
                        .RenderDirectText(143, 35, "N", 80, 10, timesleft16)
                        .RenderDirectText(146, 40, "U", 80, 10, timesleft16)
                        .RenderDirectText(149, 45, "E", 80, 10, timesleft16)
                        .RenderDirectText(152, 50, "D", 80, 10, timesleft16)
                        'print the horizontal lines on the stub
                        .RenderDirectLine(3, 102, 188, 102, arialleft10)         'top 1
                        .RenderDirectLine(3, 107, 188, 107, arialleft10)         'top 2
                        .RenderDirectLine(3, 238, 188, 238, arialleft10)         'bottom
                        'print the vertical lines on the stub
                        .RenderDirectLine(3, 102, 3, 238, arialleft10)           'left1
                        .RenderDirectLine(41, 102, 41, 238, arialleft10)         'left2
                        .RenderDirectLine(92, 102, 92, 238, arialleft10)         'left3
                        .RenderDirectLine(143, 102, 143, 238, arialleft10)       'left4
                        .RenderDirectLine(188, 102, 188, 238, arialleft10)       'left5
                        'print the column headers
                        .RenderDirectText(5, 102.5, "Payment Method", 40, 5, verdanaleft8)
                        .RenderDirectText(43, 102.5, "Received From", 40, 5, verdanaleft8)
                        .RenderDirectText(94, 102.5, "Check/Ref No.", 40, 5, verdanaleft8)
                        .RenderDirectText(173, 102.5, "Amount", 40, 5, verdanaleft8)
                        'print the date on the header
                        .RenderDirectText(22, 47, tdepositdate.ToShortDateString, 40, 5, arialleft10)
                        y = 108
                    End If
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

    Private Sub PrintDepositSummary()
        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '    bank       fisyr    number   amount   remarks   datetime
        '     6           7
        ' rcptcount  depositsum
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "DepositSummaryReport"
        Me.ReportName = "Deposit Summary"
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
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currowx, x, y As Int32
        Dim tamount, totalamount As Double
        Dim tdepositnum, tremarks As String
        Dim tdepositdate As Date
        Dim titemcount, totalitems, totaldeposits As Int32

        Try
            'collect the total amount & total items
            With Me.GridDetail
                totaldeposits = .Rows.Count
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
                'the total amount has already been summarized in the query
                totalamount = CDbl(.GetData(0, 7))
                'get the total items
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    titemcount = CInt(.GetData(index, 6))
                    totalitems += titemcount
                Next
            End With

            'start the doc
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
                '     0           1         2        3        4         5   
                '    bank       fisyr    number   amount   remarks   datetime
                '     6           7
                ' rcptcount  depositsum
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                With Me.GridDetail
                    tdepositnum = DirectCast(.GetData(index, 2), String)
                    tamount = CDbl(.GetData(index, 3))
                    tremarks = DirectCast(.GetData(index, 4), String)
                    tdepositdate = CDate(.GetData(index, 5))
                    titemcount = CInt(.GetData(index, 6))
                End With

                With Me.Doc1
                    If index = 0 Then
                        x = 10
                        y = 37
                        'print the total info box left-side
                        .RenderDirectText(x, y, "For Bank Account:", 50, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, Me.BankAccountNumber, 50, 5, verdanaright8)
                        'print the total info box right-side
                        x = 138
                        .RenderDirectText(x, y, "Total amount:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 4, "Total deposits:", 25, 5, verdanaright8)
                        .RenderDirectText(x, y + 8, "Total items:", 25, 5, verdanaright8)
                        'print the money fields
                        x = 165
                        .RenderDirectText(x, y, totalamount.ToString.Format("{0:C2}", totalamount), 25, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 4, totaldeposits.ToString.Format("{0:D1}", totaldeposits), 25, 5, verdanaright8bold)
                        .RenderDirectText(x, y + 8, totalitems.ToString.Format("{0:D1}", totalitems), 25, 5, verdanaright8bold)
                        'print the lines under the info box
                        .RenderDirectLine(0, 53, 190, 53, Color.Gray, 0.5)
                        'print the column headers
                        y = 55
                        .RenderDirectText(5, y, "Issued", 25, 5, verdanaleft8bold)
                        .RenderDirectText(33, y, "Number", 25, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(132, y, "Items", 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y, "Amount", 30, 5, verdanaright8bold)
                        y = 62
                    End If
                    .RenderDirectText(2, y, tdepositdate.ToString.Format("{0:MM/dd/yyyy}", tdepositdate), 25, 5, verdanaleft8)
                    .RenderDirectText(35, y, tdepositnum, 25, 5, verdanaleft8)
                    .RenderDirectText(60, y, tremarks, 80, 5, verdanaleft8)
                    .RenderDirectText(140, y, titemcount.ToString.Format("{0:D1}", titemcount), 20, 5, verdanaright8)
                    .RenderDirectText(160, y, tamount.ToString.Format("{0:F2}", tamount), 30, 5, verdanaright8)
                    y += 5

                    If y >= 255 Then
                        .NewPage()
                        y = 37
                        'print the total info box left-side
                        .RenderDirectText(10, y, "For Bank Account:", 50, 5, verdanaright8bold)
                        .RenderDirectText(10, y + 4, Me.BankAccountNumber, 50, 5, verdanaright8)
                        .RenderDirectLine(0, 53, 190, 53, Color.Gray, 0.5)
                        y = 55
                        .RenderDirectText(5, y, "Issued", 25, 5, verdanaleft8bold)
                        .RenderDirectText(33, y, "Number", 25, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(132, y, "Items", 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y, "Amount", 30, 5, verdanaright8bold)
                        y = 62
                    End If

                End With

                Debug.WriteLine(index.ToString)

                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            Next
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

    Private Sub PrintOutstandingReceiptsRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   posted
        '     6           7         8        9       10       11  
        '  paytype    paydescr   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   rcvdfrom  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "OutstandingReceipts"
        Me.ReportName = "Outstanding Receipts"
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
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, rcptcount As Int32
        Dim val1, val2, val3, val4, val5, calcval, sumamount As Double
        Dim tapplied, tissued As Date
        Dim tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks As String
        Dim tpaytype, tstatus, trecon, prevtrcptnum, prtstatus As String
        Dim trcptamt, tlineamt As Double

        Try
            ''''''''''''''''''''' GridTotals ''''''''''''''''''''''''''''''''''''''''''
            '      0            1          2          3           4             5
            '   castamt     checkamt    coinamt   creditamt   legacyamt    grandtotal
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Me.GridTotals
                val1 = CDbl(.GetData(0, 0))
                val2 = CDbl(.GetData(0, 1))
                val3 = CDbl(.GetData(0, 2))
                val4 = CDbl(.GetData(0, 3))
                val5 = CDbl(.GetData(0, 4))
                calcval = val1 + val2 + val3 + val4 + val5
            End With
            With Me.GridDetail
                'get the bank account number from the first item
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try

            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y, "Cash:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 4, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Coin:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 12, "Credit Card:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 16, "Legacy/Other:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 22, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y, val1.ToString.Format("{0:F2}", val1), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, val2.ToString.Format("{0:F2}", val2), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, val3.ToString.Format("{0:F2}", val3), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 12, val4.ToString.Format("{0:F2}", val4), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 16, val5.ToString.Format("{0:F2}", val5), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 22, calcval.ToString.Format("{0:C2}", calcval), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        y = 54
                        .RenderDirectLine(0, 59, 190, 59, Color.Gray, 0.5)
                        y = 62
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 69
                    End If

                    With Me.GridDetail
                        trcptnum = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tpaytype = DirectCast(.GetData(index, 7), String)
                        trcptamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tacctnum = DirectCast(.GetData(index, 10), String)
                        tsubacctnum = DirectCast(.GetData(index, 11), String)
                        tapplied = CDate(.GetData(index, 12))
                        tissued = CDate(.GetData(index, 13))
                        trcvdfrom = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        If tstatus = "V" Then
                            trcptamt = 0.0
                        Else
                            sumamount += tlineamt
                        End If

                    End With

                    If trcptnum <> prevtrcptnum Then
                        rcptcount += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(1, y, trcptnum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissued.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(40, y, trcvdfrom, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, trcptamt.ToString.Format("{0:F2}", trcptamt), 25, 5, verdanaright8)
                    End If
                    .RenderDirectText(85, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(103, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(125, y, tpaytype, 15, 5, verdanaleft8)
                    .RenderDirectText(137, y, tremarks, 33, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevtrcptnum = trcptnum

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
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
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
                .RenderDirectText(60, y, "Total Received", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Receipts", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, rcptcount.ToString.Format("{0:D2}", rcptcount), 25, 5, verdanaright8bold)
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

    Private Sub PrintReceiptTickets()
        'this routine collects information from a working grid & 
        'inserts it into the detail grid (the current po) and renders it
        'until no po's are left in the working grid;
        Me.DocumentName = "ReceiptTicket"
        Me.ReportName = "Activity Fund - Receipt Ticket"
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
        'define the styles 
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, rowindex, doccount As Int32
        Dim curnum, nextnum As String
        Dim haschanged As Boolean

        'At this point, GridWrk contains records for multiple purchase orders;
        'We will iterate thru GridWrk and load the detail grid with a single
        'purchase order, then process the report for that purchase order only.
        'Then, get the next purchase order in the grid until eof;

        ''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5   
        '   bank       fisyr    docnumber  status   recon     posted
        '     6           7         8        9       10        11  
        '  applied    created   rcvdfrom  paytype paydescr  rcptchk
        '    12          13        14       15       16        17 
        '  lineamt    totalamt    acct     sub    remarks revenuecode
        '    18          19        20       21 
        '  hdrkey     detlkey   acctname subname
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With Me.GridWrk
            'initialise the grid and the document;
            Me.GridDetail.Rows.Count = 0
            Me.GridDetail.Cols.Count = .Cols.Count
            Me.Doc1.StartDoc()

            For index = 0 To .Rows.Count - 1
                curnum = DirectCast(.GetData(index, 2), String)
                If index < .Rows.Count - 1 Then
                    nextnum = DirectCast(.GetData(index + 1, 2), String)
                Else
                    nextnum = ""
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
                Me.GridDetail.SetData(currow, 19, Me.GridWrk.GetData(index, 19))
                Me.GridDetail.SetData(currow, 20, Me.GridWrk.GetData(index, 20))
                Me.GridDetail.SetData(currow, 21, Me.GridWrk.GetData(index, 21))

                If curnum.Compare(curnum, nextnum, True) <> 0 Then haschanged = True

                If haschanged Then
                    'the account/subaccount changed so process the account
                    If doccount >= 1 Then Me.Doc1.NewPage()
                    RenderReceiptTickets()
                    currow = 0
                    haschanged = False
                    'numofaccts += 1
                    Me.GridDetail.Rows.Count = 0
                    doccount += 1
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                    'Debug.WriteLine("Processing account:  " & numofaccts.ToString)
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

    Private Sub PrintReceiptRegister()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   posted
        '     6           7         8        9       10       11  
        '  paytype    paydescr   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   rcvdfrom  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "ReceiptRegister"
        Me.ReportName = "Receipt Register"
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
        DefineStyles()

        'special style for this report
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

        Dim index, currow, x, y, rcptcount As Int32
        Dim val1, val2, val3, val4, calcval, sumamount As Double
        Dim tapplied, tissued As Date
        Dim tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks As String
        Dim tpaytype, tstatus, trecon, prevtrcptnum, prtstatus As String
        Dim trcptamt, tlineamt As Double

        Try
            ''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''
            '      0            1          2          3           4 
            '   castamt     checkamt    coinamt   creditamt   grandtotal
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Me.GridTotals
                val1 = CDbl(.GetData(0, 0))
                val2 = CDbl(.GetData(0, 1))
                val3 = CDbl(.GetData(0, 2))
                val4 = CDbl(.GetData(0, 3))
                calcval = val1 + val2 + val3 + val4
            End With
            With Me.GridDetail
                'get the bank account number from the first item
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try

            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
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
                        .RenderDirectText(118, y, "Cash:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 4, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Coin:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 12, "Credit Card:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 18, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y, val1.ToString.Format("{0:F2}", val1), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, val2.ToString.Format("{0:F2}", val2), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, val3.ToString.Format("{0:F2}", val3), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 12, val4.ToString.Format("{0:F2}", val4), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 18, calcval.ToString.Format("{0:C2}", calcval), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        trcptnum = DirectCast(.GetData(index, 2), String)
                        tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tpaytype = DirectCast(.GetData(index, 7), String)
                        trcptamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tacctnum = DirectCast(.GetData(index, 10), String)
                        tsubacctnum = DirectCast(.GetData(index, 11), String)
                        tapplied = CDate(.GetData(index, 12))
                        tissued = CDate(.GetData(index, 13))
                        trcvdfrom = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        If tstatus = "V" Then
                            trcptamt = 0.0
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
                            trcvdfrom = "** VOID **"
                        Case Else
                            prtstatus = "0"
                    End Select

                    If trcptnum <> prevtrcptnum Then
                        rcptcount += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(-1, y, prtstatus, 5, 5, specstyle)
                        .RenderDirectText(1, y, trcptnum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissued.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(40, y, trcvdfrom, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, trcptamt.ToString.Format("{0:F2}", trcptamt), 25, 5, verdanaright8)
                    End If
                    .RenderDirectText(85, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(103, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(125, y, tpaytype, 15, 5, verdanaleft8)
                    .RenderDirectText(137, y, tremarks, 33, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevtrcptnum = trcptnum

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
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
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
                .RenderDirectText(60, y, "Total Received", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Total Receipts", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, rcptcount.ToString.Format("{0:D2}", rcptcount), 25, 5, verdanaright8bold)
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

    Private Sub PrintReceiptRegisterAllBanks()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NOTE:  This routine uses rendertables and should be rewritten
        '       in the future;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4        5   
        '   bank       fisyr    docnumber  status   recon   posted
        '     6           7         8        9       10       11  
        '  paytype    paydescr   hdramt   lineamt   acct     sub 
        '    12          13        14       15
        '  applied    created   rcvdfrom  remarks
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim currow, bodyrow As Int32
        Dim val0, val1, val2, val3, calcval As Double
        Dim tbankacctnum As String

        Me.DocumentName = "ReceiptRegister"
        Me.ReportName = "Activity Fund Receipt Register"
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
        'define styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Try
            Me.Doc1.StartDoc()
            'draw rectangle around the header totals
            'Me.Doc1.RenderDirectRectangle(118, 35, 190, 60)
            'define the rendertbl (body) attributes
            'With rendertbl
            '    .BeginUpdate()
            '    .CanSplit = True
            '    .Columns.AddSome(9)
            '    'total = 190
            '    .Columns(0).WidthStr = "16mm"   'number
            '    .Columns(1).WidthStr = "14mm"   'status
            '    .Columns(2).WidthStr = "18mm"   'date
            '    .Columns(3).WidthStr = "30mm"   'rcvdfrom
            '    .Columns(4).WidthStr = "15mm"   'account
            '    .Columns(5).WidthStr = "20mm"   'lineamt
            '    .Columns(6).WidthStr = "15mm"   'paytype
            '    .Columns(7).WidthStr = "40mm"   'remarks
            '    .Columns(8).WidthStr = "20mm"   'amount
            '    .Style = arialleft8
            '    .StyleTableCell = arialleft8
            '    .Columns(0).StyleTableCell.TextAlignHorz = AlignHorzEnum.Left
            '    .Columns(1).StyleTableCell.TextAlignHorz = AlignHorzEnum.Left
            '    .Columns(2).StyleTableCell.TextAlignHorz = AlignHorzEnum.Center
            '    .Columns(3).StyleTableCell.TextAlignHorz = AlignHorzEnum.Left
            '    .Columns(4).StyleTableCell.TextAlignHorz = AlignHorzEnum.Center
            '    .Columns(5).StyleTableCell.TextAlignHorz = AlignHorzEnum.Right
            '    .Columns(8).StyleTableCell.TextAlignHorz = AlignHorzEnum.Right
            '    'always add the rows after other attributes are set (for performance)
            '    .Body.Rows.AddSome(Me.GridDetail.Rows.Count * 2)
            '    .EndUpdate()
            'End With
        Catch ex As Exception
            Throw
        End Try

        Try
            Dim trcptdate As Date
            Dim tholdbank, tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks As String
            Dim tpaytype, tstatus, trecon, prevtrcptnum As String
            Dim prtstatus As String
            Dim trcptamt, tlineamt As Double

            bodyrow = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '     0           1         2        3        4        5   
            '   bank       fisyr    docnumber  status   recon   posted
            '     6           7         8        9       10       11  
            '  paytype    paydescr   hdramt   lineamt   acct     sub 
            '    12          13        14       15
            '  applied    created   rcvdfrom  remarks
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'With rendertbl
            '    .BeginUpdate()
            '    For currow = 0 To Me.GridDetail.Rows.Count - 1
            '        With Me.GridDetail
            '            tbankacctnum = DirectCast(.GetData(currow, 0), String)
            '            trcptnum = DirectCast(.GetData(currow, 2), String)
            '            tstatus = DirectCast(.GetData(currow, 3), String).ToUpper
            '            trecon = DirectCast(.GetData(currow, 4), String).ToUpper
            '            tpaytype = DirectCast(.GetData(currow, 7), String)
            '            trcptamt = CDbl(.GetData(currow, 8))
            '            tlineamt = CDbl(.GetData(currow, 9))
            '            tacctnum = DirectCast(.GetData(currow, 10), String)
            '            tsubacctnum = DirectCast(.GetData(currow, 11), String)
            '            trcptdate = CDate(.GetData(currow, 12))
            '            trcvdfrom = DirectCast(.GetData(currow, 14), String)
            '            tremarks = DirectCast(.GetData(currow, 15), String)

            '            Select Case tstatus
            '                Case "O"
            '                    prtstatus = "Open"
            '                Case "C"    'closed, now check if recon is on or off
            '                    If trecon = "Y" Then
            '                        prtstatus = "Recon"
            '                    Else
            '                        prtstatus = "Closed"
            '                    End If
            '                Case "F"    'cleared & closedout
            '                    prtstatus = "Cleared"
            '                Case Else
            '                    prtstatus = tstatus
            '            End Select
            '        End With


            '        'handle the first record
            '        If currow = 0 Then tholdbank = tbankacctnum

            '        If tholdbank <> tbankacctnum Then
            '            tholdbank = tbankacctnum
            '            Me.Doc1.NewPage()
            '        End If

            '        .Body.Rows.AddSome(2)
            '        If trcptnum = prevtrcptnum Then
            '            .Body.Cell(bodyrow, 0).RenderText.Text = ""
            '            .Body.Cell(bodyrow, 1).RenderText.Text = ""
            '            .Body.Cell(bodyrow, 2).RenderText.Text = ""
            '            .Body.Cell(bodyrow, 3).RenderText.Text = ""
            '        Else
            '            .Body.Cell(bodyrow, 0).RenderText.Text = trcptnum
            '            .Body.Cell(bodyrow, 1).RenderText.Text = prtstatus
            '            .Body.Cell(bodyrow, 2).RenderText.Text = trcptdate.ToShortDateString
            '            .Body.Cell(bodyrow, 3).RenderText.Text = trcvdfrom
            '            .Body.Cell(bodyrow, 8).RenderText.Text = trcptamt.ToString.Format("{0:F2}", trcptamt)
            '        End If
            '        .Body.Cell(bodyrow, 4).RenderText.Text = tacctnum & "-" & tsubacctnum
            '        .Body.Cell(bodyrow, 5).RenderText.Text = tlineamt.ToString.Format("{0:F2}", tlineamt)
            '        .Body.Cell(bodyrow, 6).RenderText.Text = tpaytype
            '        .Body.Cell(bodyrow, 7).RenderText.Text = tremarks

            '        'get the current rcptnumber
            '        prevtrcptnum = trcptnum
            '        'add some space before the next record
            '        .Body.Rows(bodyrow).Height = (.Body.Rows(bodyrow).Height * 1.25)
            '        bodyrow += 2
            '        'expose the current record & count to the caller
            '        'EventRecordProcessed((reccurrent), reccount)
            '    Next
            '    .EndUpdate()
            'End With

            'Me.Doc1.RenderBlock(rendertbl)
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

    Private Sub PrintVoidReceiptRegister()
        ''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5          6         7 
        '   bank        fisyr   docnumber  status   recon    posted     paycode  paydescr
        '     8           9        10        11      12        13         14        15  
        '  hdramt     lineamt   applied   created   acct      sub       descr     remarks
        '    16          17        18        19      20        21 
        ' revcode    voidappl  voidissue  vremarks hdrkey   detlkey
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VoidReceiptRegister"
        Me.ReportName = "Void Receipt Register"
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
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currow, x, y, rcptcount As Int32
        Dim val1, val2, val3, val4, calcval, sumamount As Double
        Dim tissued, tapplied As Date
        Dim tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks As String
        Dim tpaytype, tstatus, trecon, prevtrcptnum, prtstatus As String
        Dim tvoidapplied, tvoidissued As Date
        Dim tvoidremarks As String
        Dim trcptamt, tlineamt As Double

        Try
            ''''''''''''''''''''' GridTotals '''''''''''''''''''''''''''''''
            '      0            1          2          3           4 
            '   castamt     checkamt    coinamt   creditamt   grandtotal
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Me.GridTotals
                val1 = CDbl(.GetData(0, 0))
                val2 = CDbl(.GetData(0, 1))
                val3 = CDbl(.GetData(0, 2))
                val4 = CDbl(.GetData(0, 3))
                calcval = val1 + val2 + val3 + val4
            End With
            With Me.GridDetail
                'get the bank account number from the first item
                Me.BankAccountNumber = DirectCast(.GetData(0, 0), String)
            End With
        Catch ex As Exception
            Throw
        End Try

        Try

            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'print the total info box left-side
                        .RenderDirectText(25, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(25, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        y = 32
                        .RenderDirectText(118, y, "Cash:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 4, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 8, "Coin:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 12, "Credit Card:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, y + 18, "Total register:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(160, y, val1.ToString.Format("{0:F2}", val1), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 4, val2.ToString.Format("{0:F2}", val2), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 8, val3.ToString.Format("{0:F2}", val3), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 12, val4.ToString.Format("{0:F2}", val4), 30, 5, verdanaright8bold)
                        .RenderDirectText(160, y + 18, calcval.ToString.Format("{0:C2}", calcval), 30, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Description/Received", 40, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
                        .RenderDirectText(165, y, "Amount", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    With Me.GridDetail
                        trcptnum = DirectCast(.GetData(index, 2), String)
                        trecon = DirectCast(.GetData(index, 4), String).ToUpper
                        tpaytype = DirectCast(.GetData(index, 7), String)
                        trcptamt = CDbl(.GetData(index, 8))
                        tlineamt = CDbl(.GetData(index, 9))
                        tapplied = CDate(.GetData(index, 10))
                        tissued = CDate(.GetData(index, 11))
                        tacctnum = DirectCast(.GetData(index, 12), String)
                        tsubacctnum = DirectCast(.GetData(index, 13), String)
                        trcvdfrom = DirectCast(.GetData(index, 14), String)
                        tremarks = DirectCast(.GetData(index, 15), String)
                        tvoidapplied = CDate(.GetData(index, 17))
                        tvoidissued = CDate(.GetData(index, 18))
                        tvoidremarks = DirectCast(.GetData(index, 19), String)
                        sumamount += tlineamt
                    End With

                    If trcptnum <> prevtrcptnum Then
                        rcptcount += 1
                        If currow > 1 Then y += 5
                        .RenderDirectText(1, y, "Voided on", 20, 5, verdanaleft8bold)
                        .RenderDirectText(15, y, tvoidissued.ToShortDateString, 23, 5, verdanaright8)
                        .RenderDirectText(40, y, tvoidremarks, 150, 5, verdanaleft8)
                        y += 5
                        .RenderDirectText(1, y, trcptnum, 20, 5, verdanaleft8)
                        .RenderDirectText(18, y, tissued.ToShortDateString, 20, 5, verdanaright8)
                        .RenderDirectText(40, y, trcvdfrom, 45, 10, verdanaleft8)
                        .RenderDirectText(165, y, trcptamt.ToString.Format("{0:F2}", trcptamt), 25, 5, verdanaright8)
                    End If
                    .RenderDirectText(85, y, tacctnum & "-" & tsubacctnum, 20, 5, verdanaleft8)
                    .RenderDirectText(103, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 22, 5, verdanaright8)
                    .RenderDirectText(125, y, tpaytype, 15, 5, verdanaleft8)
                    .RenderDirectText(137, y, tremarks, 33, 10, verdanaleft8)
                    y += 7
                    'get the current rcptnumber
                    prevtrcptnum = trcptnum

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
                        .RenderDirectText(22, y, "Issued", 20, 5, verdanaleft8bold)
                        .RenderDirectText(40, y, "Description/Received", 40, 5, verdanaleft8bold)
                        .RenderDirectText(85, y, "Account", 25, 5, verdanaleft8bold)
                        .RenderDirectText(103, y, "Line", 22, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Type", 25, 5, verdanaleft8bold)
                        .RenderDirectText(137, y, "Remarks", 25, 5, verdanaleft8bold)
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
                .RenderDirectText(60, y, "Total Amount Voided:", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                .RenderDirectText(60, y + 4, "Number Receipts Voided:", 50, 5, verdanaright8bold)
                .RenderDirectText(165, y + 4, rcptcount.ToString.Format("{0:D2}", rcptcount), 25, 5, verdanaright8bold)
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

    Private Sub Print1098TReport()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '     0        1         2        3        4       5 
        '   year    number    account  applied   rcvd    amount 
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "ReceiptRegister"
        Me.ReportName = "1098-T Tuition (IRS)"
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
        DefineStyles()

        'special style for this report
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

        Dim index, currow, x, y, rcptcount As Int32
        Dim amount, tempamount As Decimal
        Dim applieddate As Date
        Dim calendaryear As Int32
        Dim accountnumber, nextreceived, prevreceived, receipt, received As String

        Try

            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        'handle first record;
                        y = 34
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(25, y, "Receipted", 25, 5, verdanaright8bold)
                        .RenderDirectText(60, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(125, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(155, y, "Total", 25, 5, verdanaright8bold)
                        y += 7
                    End If

                    With Me.GridDetail
                        calendaryear = CInt(.GetData(index, 0))
                        receipt = DirectCast(.GetData(index, 1), String).ToUpper
                        accountnumber = DirectCast(.GetData(index, 2), String).ToUpper
                        applieddate = CDate(.GetData(index, 3))
                        received = DirectCast(.GetData(index, 4), String).ToUpper
                        amount = CDec(.GetData(index, 5))
                        'get the next record for comparison;
                        If index < (.Rows.Count - 1) Then nextreceived = DirectCast(.GetData(index + 1, 4), String).ToUpper
                        If index = .Rows.Count - 1 Then nextreceived = ""
                        tempamount += amount
                    End With

                    .RenderDirectText(2, y, receipt, 20, 5, verdanaleft8)
                    .RenderDirectText(25, y, applieddate.ToString.Format("{0:MM/dd/yyyy}", applieddate), 25, 10, verdanaright8)
                    .RenderDirectText(60, y, received, 95, 10, verdanaleft8)
                    .RenderDirectText(125, y, amount.ToString.Format("{0:F2}", amount), 25, 5, verdanaright8)
                    If received.Compare(received, nextreceived, False) <> 0 Then
                        .RenderDirectText(155, y, tempamount.ToString.Format("{0:F2}", tempamount), 25, 5, verdanaright8)
                        tempamount = 0
                    End If

                    y += 7

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        y = 34
                        .RenderDirectText(2, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(25, y, "Receipted", 25, 5, verdanaright8bold)
                        .RenderDirectText(60, y, "Received", 25, 5, verdanaleft8bold)
                        .RenderDirectText(125, y, "Amount", 25, 5, verdanaright8bold)
                        .RenderDirectText(155, y, "Total", 25, 5, verdanaright8bold)
                        y += 7
                    End If
                    prevreceived = received
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                '''''y += 10
                '''''If y > 240 Then
                '''''    .NewPage()
                '''''    y = 65
                '''''End If
                ''''''draw top of total box
                '''''.RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                '''''.RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                '''''.RenderDirectText(60, y, "Total Received", 50, 5, verdanaright8bold)
                '''''''''''.RenderDirectText(165, y, sumamount.ToString.Format("{0:C2}", sumamount), 25, 5, verdanaright8bold)
                '''''.RenderDirectText(60, y + 4, "Total Receipts", 50, 5, verdanaright8bold)
                '''''.RenderDirectText(165, y + 4, rcptcount.ToString.Format("{0:D2}", rcptcount), 25, 5, verdanaright8bold)
                '''''y += 10
                ''''''draw bottom of total box
                '''''.RenderDirectLine(59, y, 190, y, Color.Black, 0.25)
                '''''.RenderDirectLine(59, y + 0.5, 190, y + 0.5, Color.Black, 0.25)
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
                .RenderDirectLine(0, 264, 191, 264, Color.Black, 0.5)
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
        Catch ex As Exception
            Throw
        End Try

        Try
            PrintFooter()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RenderReceiptTickets()
        Me.DocumentName = "ReceiptTicket"
        Me.ReportName = "Activity Fund - Receipt Ticket"
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
        'define the styles
        DefineStyles()
        'define the document
        DefineDocumentSettings(Me.DocumentName)

        Try
            Dim tcreated, tapplied As Date
            Dim tacctnum, tsubacctnum, trcptnum, trcvdfrom, tremarks, trevcode As String
            Dim tacctname, tsubacctname As String
            Dim tpaytype, trcptchk, tstatus, trecon, prevtrcptnum, prtstatus, prtcode As String
            Dim trcptamt, tlineamt As Double
            Dim y, index As Int32

            For index = 0 To Me.GridDetail.Rows.Count - 1
                With Me.GridDetail
                    trcptnum = DirectCast(.GetData(index, 2), String)
                    tstatus = DirectCast(.GetData(index, 3), String).ToUpper
                    trecon = DirectCast(.GetData(index, 4), String).ToUpper
                    tapplied = CDate(.GetData(index, 6))
                    tcreated = CDate(.GetData(index, 7))
                    trcvdfrom = DirectCast(.GetData(index, 8), String)
                    tpaytype = DirectCast(.GetData(index, 10), String)
                    trcptchk = DirectCast(.GetData(index, 11), String)
                    tlineamt = CDbl(.GetData(index, 12))
                    trcptamt = CDbl(.GetData(index, 13))
                    tacctnum = DirectCast(.GetData(index, 14), String)
                    tsubacctnum = DirectCast(.GetData(index, 15), String)
                    tremarks = DirectCast(.GetData(index, 16), String)
                    trevcode = DirectCast(.GetData(index, 17), String)
                    prtcode = FormatRevenueCode(trevcode)
                    tacctname = DirectCast(.GetData(index, 20), String)
                    tsubacctname = DirectCast(.GetData(index, 21), String)
                    Select Case tstatus
                        Case "O"
                            prtstatus = "Open"
                        Case "C"    'closed, now check if recon is on or off
                            If trecon = "Y" Then
                                prtstatus = "Cleared"
                            Else
                                prtstatus = "Outstanding"
                            End If
                        Case "F"    'cleared & closedout
                            prtstatus = "Cleared"
                        Case Else
                            prtstatus = tstatus
                    End Select
                End With

                With Me.Doc1
                    If index = 0 Then
                        'do the header stuff
                        y = 20
                        'change the color of a style
                        arialleft8.TextColor = Color.Gray
                        'print top line
                        .RenderDirectLine(8, y - 1, 182, y - 1, Color.Gray, 0.5)
                        'left side
                        .RenderDirectText(10, y, Me.SchoolName, 130, 5, verdanaleft10)
                        .RenderDirectText(10, y + 5, Me.SchoolAddress1, 100, 5, verdanaleft10)
                        .RenderDirectText(10, y + 10, Me.SchoolCityStateZip, 100, 5, verdanaleft10)
                        'right side
                        .RenderDirectText(130, y, "Activity Fund Receipt", 50, 5, verdanaright10)
                        .RenderDirectText(108, y + 5, "Receipt number:", 50, 5, verdanaright10)
                        .RenderDirectText(130, y + 5, trcptnum, 50, 5, verdanaright10bold)
                        .RenderDirectText(130, y + 10, "Issued: " & tcreated.ToShortDateString, 50, 5, verdanaright10)
                        'left side
                        .RenderDirectText(10, y + 25, "Received from:", 35, 5, arialleft10)
                        .RenderDirectText(40, y + 25, trcvdfrom, 80, 10, arialleft10)
                        'right side
                        .RenderDirectText(130, y + 25, "Receipt amount:", 50, 5, arialright10)
                        .RenderDirectText(130, y + 30, trcptamt.ToString.Format("{0:C2}", trcptamt), 50, 5, verdanaright10bold)

                        ''''''''''''''''''''' Signatures ''''''''''''''''''''''''''''''''''''
                        'only print the signatures on the first page

                        If Me.DoSignatures Then
                            Dim imgalign As New C1.C1PrintDocument.ImageAlignDef
                            imgalign.AlignHorz = ImageAlignHorzEnum.Left
                            imgalign.StretchHorz = True
                            imgalign.StretchVert = True
                            imgalign.KeepAspectRatio = True
                            'if primary signature image is available, then print the image
                            If Not Me.Signature1 Is Nothing Then Doc1.RenderDirectImage(116, y + 39, Me.Signature1, 300, 14, imgalign)
                            'print the name of the primary signer under the first line
                            .RenderDirectText(116, y + 50, Me.SignatureTextLine1, 80, 5, arialleft8)
                            'if secondary signature image is available, then print the image
                            'If Not Me.Signature2 Is Nothing Then Doc1.RenderDirectImage(101, 40, Me.Signature2, 300, 14, imgalign)
                            'print the name of the secondary signer under the second line
                            '.RenderDirectText(101, 51, Me.SignatureTextLine2, 80, 5, arialleft8)
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Else
                            .RenderDirectText(118, y + 50, "Activity Fund Receipt Signature", 80, 5, arialleft8)
                        End If

                        'print the signature underline
                        .RenderDirectLine(115, y + 50, 178, y + 50, Color.Gray, 1.0)
                        'mark if multiple
                        If Me.GridDetail.Rows.Count > 1 Then .RenderDirectText(10, y + 40, "Multiple", 20, 5, arialleft8)
                        'print the applied date
                        .RenderDirectText(10, y + 50, "Applied period:  " & tapplied.ToShortDateString, 50, 5, arialleft8)
                        'print left-side line
                        .RenderDirectLine(8, y - 1, 8, y + 56, Color.Gray, 0.5)
                        'print right-side line
                        .RenderDirectLine(182, y - 1, 182, y + 56, Color.Gray, 0.5)
                        'print the bottom header line
                        .RenderDirectLine(8, y + 56, 182, y + 56, Color.Gray, 0.5)
                        'print detail information headers
                        arialleft8.TextColor = Color.Black
                        y += 58     '78
                        .RenderDirectText(11, y, "Account", 20, 5, arialleft8)
                        .RenderDirectText(31, y, "Name", 25, 5, arialleft8)
                        .RenderDirectText(66, y, "Remarks", 20, 5, arialleft8)
                        .RenderDirectText(128, y, "Type", 20, 5, arialleft8)
                        .RenderDirectText(142, y, "Check", 20, 5, arialleft8)
                        .RenderDirectText(166, y, "Amount", 20, 5, arialleft8)
                        'print the bottom column header line
                        .RenderDirectLine(8, y + 4.5, 182, y + 4.5, Color.Gray, 0.5)
                        y += 6      '84
                    End If

                    'check if it's a page break
                    If y >= 260 Then
                        .NewPage()
                        y = 20
                        'print detail information headers
                        arialleft8.TextColor = Color.Black
                        'print top line
                        .RenderDirectLine(8, y - 1, 182, y - 1, Color.Gray, 0.5)
                        .RenderDirectText(12, y, "...continued from previous page", 70, 5, arialleft10)
                        .RenderDirectText(130, y, "Activity Fund Receipt", 50, 5, arialright10)
                        y += 5  '25
                        .RenderDirectText(108, y, "Receipt number:", 50, 5, arialright10)
                        .RenderDirectText(130, y, trcptnum, 50, 5, verdanaright10bold)
                        y += 10 '35
                        'print the top column header line
                        .RenderDirectText(11, y, "Account", 20, 5, arialleft8)
                        .RenderDirectText(31, y, "Name", 25, 5, arialleft8)
                        .RenderDirectText(66, y, "Remarks", 20, 5, arialleft8)
                        .RenderDirectText(128, y, "Type", 20, 5, arialleft8)
                        .RenderDirectText(142, y, "Check", 20, 5, arialleft8)
                        .RenderDirectText(166, y, "Amount", 20, 5, arialleft8)
                        'print the bottom column header line
                        .RenderDirectLine(8, y + 4.5, 182, y + 4.5, Color.Gray, 0.5)
                        y += 7  '42
                    End If
                    'print the detail information;
                    .RenderDirectText(10, y, tacctnum & "-" & tsubacctnum, 20, 5, arialleft8)
                    .RenderDirectText(30, y, tsubacctname, 60, 5, arialleft8)
                    .RenderDirectText(128, y, tpaytype, 45, 5, arialleft8)
                    .RenderDirectText(142, y, trcptchk, 20, 5, arialleft8)
                    .RenderDirectText(153, y, tlineamt.ToString.Format("{0:F2}", tlineamt), 25, 5, arialright8)
                    y += 4
                    If Me.UseOcas Then .RenderDirectText(10, y, prtcode, 45, 5, arialleft8)
                    .RenderDirectText(65, y, tremarks, 115, 10, arialleft8)
                    y += 8
                End With
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region "  Methods Retrieval "

    Private Sub GetSignatureDetails()
        'this method retrieves all siganture details as saved in school info
        '& returns a datatable cast as a generic object
        Dim SSQL As String
        SSQL = "SELECT sign_title, sign_fname, sign_mi, sign_lname," _
        & " sign_signature, sign_rcpt_sw" _
        & " FROM signatures" _
        & " WHERE sign_rcpt_sw = 'Y'" _
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
        '     0           1          2          3          4          5
        ' sign_title  sign_fname  sign_mi  sign_lname  signatures  rcptsw
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If tbl.Rows.Count < 1 Then Exit Sub
        Dim currow As Int32
        Dim arrayimage1 As Byte()
        Dim title, fname, mi, lname, rcptsw As String
        Dim ms As MemoryStream

        With tbl
            _dosignatures = False

            For currow = 0 To .Rows.Count - 1
                rcptsw = DirectCast(.Rows(currow).Item(5), String).Trim
                If rcptsw.ToUpper = "Y" Then _dosignatures = True
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

                '''''Try
                '''''    'get signature 2
                '''''    If currow = 1 Then
                '''''        SignatureTextLine2 = fname + " " + mi + " " + lname
                '''''        'if no image is available, then continue
                '''''        arrayimage1 = CType(.Rows(currow).Item(4), Byte())
                '''''        ms = New MemoryStream(arrayimage1)
                '''''        If ms.Length > 0 Then Me.Signature2 = Image.FromStream(ms)
                '''''    End If
                '''''Catch ex As Exception

                '''''End Try
            Next
        End With
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

    Private Property DoSignatures() As Boolean
        Get
            Return _dosignatures
        End Get
        Set(ByVal Value As Boolean)
            _dosignatures = Value
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

#End Region

End Class
