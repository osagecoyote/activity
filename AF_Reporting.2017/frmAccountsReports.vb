Imports C1.C1PrintDocument
Imports C1.Win.C1FlexGrid
Imports System.Data
Imports System.Data.SqlClient

Public Class frmAccountsReports
    Inherits System.Windows.Forms.Form

#Region "  Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Dim authobj As AF_Master.Authuser
        Try
            Me.AppliedDate = authobj.CurrentAppliedDate
            Me.ConnectionString = authobj.ConnectionString
            Me.CurrentMonthString = authobj.CurrentMonthString
            Me.FiscalYear = authobj.FiscalYear
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
    Friend WithEvents GridDetail As C1.Win.C1FlexGrid.C1FlexGrid
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
    Friend WithEvents GridTotals As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents Doc1 As C1.C1PrintDocument.C1PrintDocument
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAccountsReports))
        Me.GridDetail = New C1.Win.C1FlexGrid.C1FlexGrid
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
        Me.GridTotals = New C1.Win.C1FlexGrid.C1FlexGrid
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridDetail
        '
        Me.GridDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridDetail.BackColor = System.Drawing.SystemColors.Window
        Me.GridDetail.ColumnInfo = "10,0,0,250,0,85,Columns:"
        Me.GridDetail.ExtendLastCol = True
        Me.GridDetail.Location = New System.Drawing.Point(10, 0)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.Rows.Fixed = 0
        Me.GridDetail.Size = New System.Drawing.Size(646, 336)
        Me.GridDetail.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridDetail.TabIndex = 2
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
        Me.Prev1.TabIndex = 3
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
        'GridTotals
        '
        Me.GridTotals.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridTotals.ColumnInfo = "10,0,0,100,0,85,Columns:"
        Me.GridTotals.Location = New System.Drawing.Point(8, 8)
        Me.GridTotals.Name = "GridTotals"
        Me.GridTotals.Rows.Fixed = 0
        Me.GridTotals.Size = New System.Drawing.Size(648, 328)
        Me.GridTotals.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridTotals.TabIndex = 4
        '
        'frmAccountsReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 373)
        Me.Controls.Add(Me.GridTotals)
        Me.Controls.Add(Me.GridDetail)
        Me.Controls.Add(Me.Prev1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAccountsReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Accounts Reporting"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  C1Doc Events "

    Private Sub Doc1_NewPageStarted(ByVal sender As C1.C1PrintDocument.C1PrintDocument, ByVal e As C1.C1PrintDocument.NewPageStartedEventArgs) Handles Doc1.NewPageStarted
        Select Case Me.DocumentName
            Case "BalanceSheet", "ChartOfAccounts", "ChartOfSubAccounts", "SummaryOfAccounts"
                PrintHeader()
            Case Else
                'receipt ticket does not use the PrintHeader routine
        End Select
    End Sub

#End Region

#Region "  Class Members "

    'styles;
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
    'header values;
    Private CellMiddleBottom As String = ""
    Private CellMiddleMiddle As String = ""
    Private CellMiddleTop As String = ""
    Private CellRightBottom As String = ""
    Private CellRightMiddle As String = ""
    Private CellRightTop As String = ""
    'property vars;
    Private _bankaccountnumber As String
    Private _applieddate As Date
    Private _connectionstring As String
    Private _currentmonthstring As String
    Private _documentname As String
    Private _fiscalyear As Int32
    Private _fiscalyearselected As Int32
    Private _reportname As String
    Private _schoolname As String
    Private _schooladdress1 As String
    Private _schooladdress2 As String
    Private _schoolcitystatezip As String
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
                Case "BalanceSheet", "ChartOfAccounts", "ChartOfSubAccounts", "SummaryOfAccounts"
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

        DefineFooterStyle()

    End Sub

    Private Sub DefineFooterStyle()
        'style for the footer
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

#Region "  Methods Generation "

    Public Function GenerateStatementOfChangePriorYear(ByVal selectedfisyr As Int32, ByVal ebankaccountnumber As String, ByVal eincludesubaccounts As Boolean, ByVal ecurrentyear As Boolean) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this report was required by Miami Public Schools (p/Wendy) due to auditor
        'requirements on 2011.01.25 (Fred)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname    hdrbeg   hdrrev
        '    7         8         9         10        11         12        13
        ' hdrexp    hdradj    hdrend     subbeg    subrev     subexp   subadj
        '   14        15        16         17        18         19        20        
        ' subend   nwhdrrev  nwhdrexp   nwsubrev  nwsubexp     
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim tbl As DataTable
        Dim cmd As SqlCommand

        ecurrentyear = False
        Me.FiscalYearSelected = selectedfisyr

        Try
            'collect the accounts/subaccounts with placeholders for balances;
            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " ahst_beg_year_balance, ahst_mtd_receipts + ahst_ytd_receipts," _
            & " ahst_mtd_expend + ahst_ytd_expend," _
            & " ahst_mtd_adjust + ahst_ytd_adjust," _
            & " ahst_beg_year_balance + (ahst_mtd_receipts + ahst_ytd_receipts)" _
            & "  - (ahst_mtd_expend + ahst_ytd_expend) + (ahst_mtd_adjust + ahst_ytd_adjust) AS Accounttotal," _
            & " ahst_beg_year_balance, ahst_mtd_receipts + ahst_ytd_receipts, ahst_mtd_expend + ahst_ytd_expend," _
            & " ahst_mtd_adjust + ahst_ytd_adjust, ahst_beg_year_balance + (ahst_mtd_receipts + ahst_ytd_receipts)" _
            & " - (ahst_mtd_expend + ahst_ytd_expend) + (ahst_mtd_adjust + ahst_ytd_adjust) AS Accounttotal," _
            & " 0.0 AS NewHdrRev, 0.0 AS NewHdrExp, 0.0 AS NewSubRev, 0.0 AS NewSubExp" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1 and ahst_fisyr = @p2" _
            & " AND ahst_current_month = 6" _
            & " ORDER BY af_acct_num, as_acct_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", Me.FiscalYearSelected)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("accountbalances")
            da.Fill(tbl)
            cn.Close()
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
        End Try

        Try
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail;
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        Dim index As Int32
        Dim holdvalue As Decimal
        Dim tbankacct, tacctname, tacctnum, tsubname, tsubnum As String
        Dim hdrbeg, hdrexp, hdrrev, hdrend, subbeg, subexp, subrev, subend As Decimal
        Dim newhdrexp, newhdrrev, newsubexp, newsubrev As Decimal

        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tbankacct = CType(.GetData(index, 0), String)
                tacctnum = CType(.GetData(index, 1), String)
                tacctname = CType(.GetData(index, 2), String)
                tsubnum = CType(.GetData(index, 3), String)
                tsubname = CType(.GetData(index, 4), String)
                hdrbeg = CType(.GetData(index, 5), Decimal)
                hdrrev = CType(.GetData(index, 6), Decimal)
                hdrexp = CType(.GetData(index, 7), Decimal)
                subbeg = CType(.GetData(index, 10), Decimal)
                subrev = CType(.GetData(index, 11), Decimal)
                subexp = CType(.GetData(index, 12), Decimal)

                Try
                    'collect the revenue adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND (tran_type = 'I' OR tran_type = 'N' OR tran_type = 'R');"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND (tran_type = 'I' OR tran_type = 'N' OR tran_type = 'R')"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", Me.FiscalYearSelected)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("revadjust")
                    da.Fill(ds)
                    cn.Close()
                    'add the revenue adjustment;
                    newsubrev += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrrev += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    If cn.State <> ConnectionState.Closed Then cn.Close()
                End Try

                Try
                    'collect the bank charge adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND tran_type = 'B';"
                    '& " AND (tran_type = 'B' or tran_type = 'E');"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND tran_type = 'B'"
                    '& " AND (tran_type = 'B' or tran_type = 'E')"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", Me.FiscalYearSelected)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("expadjust")
                    da.Fill(ds)
                    cn.Close()
                    'convert bank charge adjustment to check;
                    holdvalue *= -1
                    holdvalue = CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    'holdvalue *= -1
                    newsubexp += holdvalue
                    holdvalue = CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                    'holdvalue *= -1
                    newhdrexp += holdvalue
                Catch ex As Exception
                    Throw
                Finally
                    If cn.State <> ConnectionState.Closed Then cn.Close()
                End Try

                Try
                    'collect the expense adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND tran_type = 'E';"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND tran_type = 'E'"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", Me.FiscalYearSelected)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("expadjust")
                    da.Fill(ds)
                    cn.Close()
                    'convert expense adjustment to check;
                    holdvalue = CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    holdvalue *= -1
                    newsubexp += holdvalue
                    holdvalue = CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                    holdvalue *= -1
                    newhdrexp += holdvalue
                Catch ex As Exception
                    Throw
                Finally
                    If cn.State <> ConnectionState.Closed Then cn.Close()
                End Try


                Try
                    'collect the revenue transfers;
                    SSQL = "SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_to = @p3" _
                    & " AND as_acct_num_to = @p4;"
                    SSQL += " SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_to = @p3"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", Me.FiscalYearSelected)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("revtransfer")
                    da.Fill(ds)
                    cn.Close()
                    'add the revenue transfer;
                    newsubrev += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrrev += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    If cn.State <> ConnectionState.Closed Then cn.Close()
                End Try

                Try
                    'collect the expense transfers;
                    SSQL = "SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_from = @p3" _
                    & " AND as_acct_num_from = @p4;"
                    SSQL += " SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_from = @p3"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", Me.FiscalYearSelected)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("exptransfer")
                    da.Fill(ds)
                    cn.Close()
                    'add the expense transfer;
                    newsubexp += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrexp += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    If cn.State <> ConnectionState.Closed Then cn.Close()
                End Try

                'add ytd rev to adj/trx rev;
                newsubrev += subrev
                newhdrrev += hdrrev
                'add ytd exp to adj/trx exp;
                newsubexp += subexp
                newhdrexp += hdrexp
                'calc ending balance;
                subend = subbeg + newsubrev - newsubexp
                hdrend = hdrbeg + newhdrrev - newhdrexp
                'set the new revenue & expense amounts;
                .SetData(index, 15, newhdrrev)
                .SetData(index, 16, newhdrexp)
                .SetData(index, 17, newsubrev)
                .SetData(index, 18, newsubexp)
                'reinitialise variables;
                newhdrrev = 0D
                newhdrexp = 0D
                newsubrev = 0D
                newsubexp = 0D
            Next
        End With

        '''''Me.GridDetail.Visible = True
        '''''Me.GridTotals.Visible = False
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYearSelected.ToString
            Me.CellMiddleBottom = "YTD Summary"
            Application.DoEvents()
            'render the table;
            'Call PrintStatementOfChange(Me.FiscalYear)
            Call PrintStatementOfChange(Me.FiscalYearSelected)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateStatementOfChange(ByVal efiscalyear As Int32, ByVal ebankaccountnumber As String, ByVal eincludesubaccounts As Boolean, ByVal ecurrentyear As Boolean) As Boolean
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'this report was required by Miami Public Schools (p/Wendy) due to auditor
        'requirements on 2011.01.25 (Fred)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname    hdrbeg   hdrrev
        '    7         8         9         10        11         12        13
        ' hdrexp    hdradj    hdrend     subbeg    subrev     subexp   subadj
        '   14        15        16         17        18         19        20        
        ' subend   nwhdrrev  nwhdrexp   nwsubrev  nwsubexp     
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim tbl As DataTable
        Dim cmd As SqlCommand


        If ecurrentyear = True Then
            Try
                'collect the account balances from the account table including subaccounts;
                SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
                & " af_beg_year_balance, af_mtd_receipts + af_ytd_receipts, af_mtd_expend + af_ytd_expend," _
                & " af_mtd_adjust + af_ytd_adjust, af_beg_year_balance + (af_mtd_receipts + af_ytd_receipts) - (af_mtd_expend + af_ytd_expend) + (af_mtd_adjust + af_ytd_adjust) AS Accounttotal," _
                & " as_beg_year_balance, as_mtd_receipts + as_ytd_receipts, as_mtd_expend + as_ytd_expend," _
                & " as_mtd_adjust + as_ytd_adjust, as_beg_year_balance + (as_mtd_receipts + as_ytd_receipts) - (as_mtd_expend + as_ytd_expend) + (as_mtd_adjust + as_ytd_adjust) AS Accounttotal," _
                & " 0.0 AS NewHdrRev, 0.0 AS NewHdrExp, 0.0 AS NewSubRev, 0.0 AS NewSubExp" _
                & " FROM acct_info AS h, acct_sub AS d" _
                & " WHERE h.bank_acct_num = @p1" _
                & " AND h.bank_acct_num = d.bank_acct_num" _
                & " AND h.af_acct_num = d.af_acct_num" _
                & " AND af_status = 'O' AND as_status = 'O'" _
                & " ORDER BY h.af_acct_num, as_acct_num"
                cn = New SqlConnection(Me.ConnectionString)
                cmd = New SqlCommand(SSQL, cn)
                cmd.Parameters.Add("@p1", ebankaccountnumber)
                da = New SqlDataAdapter(cmd)
                tbl = New DataTable("accountbalances")
                da.Fill(tbl)
            Catch ex As Exception
                Throw
            Finally
                cn.Close()
                da.Dispose()
                cmd.Dispose()
            End Try
        End If


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If ecurrentyear = False Then
        'Try
        ''collect the account balances from the account history table including subaccounts;
        'SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
        '& " ahst_beg_year_balance, af_mtd_receipts + af_ytd_receipts, af_mtd_expend + af_ytd_expend," _
        '& " af_mtd_adjust + af_ytd_adjust, af_beg_year_balance + (af_mtd_receipts + af_ytd_receipts) - (af_mtd_expend + af_ytd_expend) + (af_mtd_adjust + af_ytd_adjust) AS Accounttotal," _
        '& " as_beg_year_balance, as_mtd_receipts + as_ytd_receipts, as_mtd_expend + as_ytd_expend," _
        '& " as_mtd_adjust + as_ytd_adjust, as_beg_year_balance + (as_mtd_receipts + as_ytd_receipts) - (as_mtd_expend + as_ytd_expend) + (as_mtd_adjust + as_ytd_adjust) AS Accounttotal," _
        '& " 0.0 AS NewHdrRev, 0.0 AS NewHdrExp, 0.0 AS NewSubRev, 0.0 AS NewSubExp" _
        '& " FROM acct_history" _
        '& " WHERE h.bank_acct_num = @p1" _
        '& " AND h.bank_acct_num = d.bank_acct_num" _
        '& " AND h.af_acct_num = d.af_acct_num" _
        '& " AND af_status = 'O' AND as_status = 'O'" _
        '& " ORDER BY h.af_acct_num, as_acct_num"
        'cn = New SqlConnection(Me.ConnectionString)
        'cmd = New SqlCommand(SSQL, cn)
        'cmd.Parameters.Add("@p1", ebankaccountnumber)
        'da = New SqlDataAdapter(cmd)
        'tbl = New DataTable("accountbalances")
        'da.Fill(tbl)
        'Catch ex As Exception
        'Throw
        'Finally
        'cn.Close()
        'da.Dispose()
        'cmd.Dispose()
        'End Try
        'End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try
            If tbl.Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail;
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try

        Dim index As Int32
        Dim holdvalue As Decimal
        Dim tbankacct, tacctname, tacctnum, tsubname, tsubnum As String
        Dim hdrbeg, hdrexp, hdrrev, hdrend, subbeg, subexp, subrev, subend As Decimal
        Dim newhdrexp, newhdrrev, newsubexp, newsubrev As Decimal

        With Me.GridDetail
            For index = 0 To .Rows.Count - 1
                tbankacct = CType(.GetData(index, 0), String)
                tacctnum = CType(.GetData(index, 1), String)
                tacctname = CType(.GetData(index, 2), String)
                tsubnum = CType(.GetData(index, 3), String)
                tsubname = CType(.GetData(index, 4), String)
                hdrbeg = CType(.GetData(index, 5), Decimal)
                hdrrev = CType(.GetData(index, 6), Decimal)
                hdrexp = CType(.GetData(index, 7), Decimal)
                subbeg = CType(.GetData(index, 10), Decimal)
                subrev = CType(.GetData(index, 11), Decimal)
                subexp = CType(.GetData(index, 12), Decimal)

                Try
                    'collect the revenue adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND (tran_type = 'I' OR tran_type = 'N' OR tran_type = 'R');"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND (tran_type = 'I' OR tran_type = 'N' OR tran_type = 'R')"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", efiscalyear)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("revadjust")
                    da.Fill(ds)
                    'add the revenue adjustment;
                    newsubrev += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrrev += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                    ds.Dispose()
                End Try

                Try
                    'collect the bank charge adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND tran_type = 'B';"
                    '& " AND (tran_type = 'B' or tran_type = 'E');"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND tran_type = 'B'"
                    '& " AND (tran_type = 'B' or tran_type = 'E')"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", efiscalyear)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("expadjust")
                    da.Fill(ds)
                    'convert bank charge adjustment to check;
                    holdvalue *= -1
                    holdvalue = CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    'holdvalue *= -1
                    newsubexp += holdvalue
                    holdvalue = CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                    'holdvalue *= -1
                    newhdrexp += holdvalue
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                    ds.Dispose()
                End Try

                Try
                    'collect the expense adjustments;
                    SSQL = "SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND as_acct_num = @p4" _
                    & " AND tran_type = 'E';"
                    SSQL += " SELECT ISNULL(SUM(tran_amt), 0) FROM transactions" _
                    & " WHERE tran_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num = @p3" _
                    & " AND tran_type = 'E'"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", efiscalyear)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("expadjust")
                    da.Fill(ds)
                    'convert expense adjustment to check;
                    holdvalue = CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    holdvalue *= -1
                    newsubexp += holdvalue
                    holdvalue = CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                    holdvalue *= -1
                    newhdrexp += holdvalue
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                    ds.Dispose()
                End Try


                Try
                    'collect the revenue transfers;
                    SSQL = "SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_to = @p3" _
                    & " AND as_acct_num_to = @p4;"
                    SSQL += " SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_to = @p3"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", efiscalyear)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("revtransfer")
                    da.Fill(ds)
                    'add the revenue transfer;
                    newsubrev += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrrev += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                    ds.Dispose()
                End Try

                Try
                    'collect the expense transfers;
                    SSQL = "SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_from = @p3" _
                    & " AND as_acct_num_from = @p4;"
                    SSQL += " SELECT ISNULL(SUM(trx_amt), 0) FROM transfers" _
                    & " WHERE trx_fisyr = @p1" _
                    & " AND bank_acct_num = @p2" _
                    & " AND af_acct_num_from = @p3"
                    cn = New SqlConnection(Me.ConnectionString)
                    cmd = New SqlCommand(SSQL, cn)
                    cmd.Parameters.Add("@p1", efiscalyear)
                    cmd.Parameters.Add("@p2", tbankacct)
                    cmd.Parameters.Add("@p3", tacctnum)
                    cmd.Parameters.Add("@p4", tsubnum)
                    da = New SqlDataAdapter(cmd)
                    ds = New DataSet("exptransfer")
                    da.Fill(ds)
                    'add the expense transfer;
                    newsubexp += CType(ds.Tables(0).Rows(0).Item(0), Decimal)
                    newhdrexp += CType(ds.Tables(1).Rows(0).Item(0), Decimal)
                Catch ex As Exception
                    Throw
                Finally
                    cn.Close()
                    cmd.Dispose()
                    ds.Dispose()
                End Try

                'add ytd rev to adj/trx rev;
                newsubrev += subrev
                newhdrrev += hdrrev
                'add ytd exp to adj/trx exp;
                newsubexp += subexp
                newhdrexp += hdrexp
                'calc ending balance;
                subend = subbeg + newsubrev - newsubexp
                hdrend = hdrbeg + newhdrrev - newhdrexp
                'set the new revenue & expense amounts;
                .SetData(index, 15, newhdrrev)
                .SetData(index, 16, newhdrexp)
                .SetData(index, 17, newsubrev)
                .SetData(index, 18, newsubexp)
                'reinitialise variables;
                newhdrrev = 0D
                newhdrexp = 0D
                newsubrev = 0D
                newsubexp = 0D
            Next
        End With

        '''''Me.GridDetail.Visible = True
        '''''Me.GridTotals.Visible = False
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "YTD Summary"
            Application.DoEvents()
            'render the table;
            Call PrintStatementOfChange(Me.FiscalYear)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateBalanceSheet(ByVal ebankaccountnumber As String, ByVal eincludesubaccounts As Boolean) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        If eincludesubaccounts Then
            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " af_beg_month_balance, af_mtd_receipts, af_mtd_expend, af_mtd_adjust," _
            & " (af_beg_month_balance + (af_mtd_receipts - af_mtd_expend + af_mtd_adjust)) AS computedaccttotal," _
            & " as_beg_month_balance, as_mtd_receipts, as_mtd_expend, as_mtd_adjust," _
            & " (as_beg_month_balance + (as_mtd_receipts - as_mtd_expend + as_mtd_adjust)) AS computedsubtotal" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND af_status = 'O' AND as_status = 'O'" _
            & " ORDER BY h.af_acct_num, as_acct_num; "
        Else
            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name, '', ''," _
            & " af_beg_month_balance, af_mtd_receipts, af_mtd_expend, af_mtd_adjust," _
            & " (af_beg_month_balance + (af_mtd_receipts - af_mtd_expend + af_mtd_adjust)) AS computedaccttotal," _
            & " 0.0, 0.0, 0.0, 0.0, 0.0" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & " ORDER BY af_acct_num; "
        End If

        SSQL += "SELECT SUM(af_beg_month_balance) AS begbal, SUM(af_mtd_receipts) AS mtdrcpt," _
        & " SUM(af_mtd_expend) AS mtdexpend, SUM(af_mtd_adjust) AS mtdadj" _
        & " FROM acct_info" _
        & " WHERE bank_acct_num = @p1" _
        & " AND af_status = 'O'"

        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
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
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        '''''Me.GridDetail.Visible = False
        '''''Me.GridTotals.Visible = True
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = Module1.ConvertCardinalMonthToString(Me.AppliedDate.Month)
            Me.CellMiddleBottom = "FY-" & Me.FiscalYear.ToString
            Application.DoEvents()
            'render the table
            If eincludesubaccounts Then
                PrintBalanceSheetSubaccounts(Me.FiscalYear, Me.AppliedDate.Month)
            Else
                PrintBalanceSheetAccounts(Me.FiscalYear, Me.AppliedDate.Month)
            End If
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateChartOfAccounts(ByVal ebankaccountnumber As String, ByVal eincludesubaccounts As Boolean) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5        6
        '   bank      acctnum   acctname  subnum   subname   hdrkey   subkey
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        If eincludesubaccounts Then
            'headers & subs
            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, h.af_acct_name," _
            & " d.as_acct_num, d.as_acct_name, h.af_autoinc_key, d.as_autoinc_key" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND (h.af_status <> 'C')" _
            & " AND (d.as_status <> 'C')" _
            & " AND h.bank_acct_num = @p1" _
            & " ORDER BY d.af_acct_num, d.as_acct_num; "
            SSQL += "SELECT COUNT(*) FROM acct_info"
        Else
            'headers only
            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, h.af_acct_name," _
             & " '', '', h.af_autoinc_key, 0" _
             & " FROM acct_info AS h" _
             & " WHERE h.bank_acct_num = @p1" _
             & " AND h.af_status <> 'C'" _
             & " ORDER BY h.af_acct_num; "
            SSQL += "SELECT COUNT(*) FROM acct_info"
        End If
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
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
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleBottom = "FY-" & Me.FiscalYear.ToString
            Application.DoEvents()
            'render the table
            If eincludesubaccounts Then
                PrintChartOfSubAccounts()
            Else
                PrintChartOfAccounts()
            End If
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateExpenditureYTDEncumbranceBalances(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'This report is called from the Expenditure Reports screen;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''''''
        '0         1         2          3         4          5           6    
        'bank     acctnum   acctname   subnum    subname   hdrbalance  subbalance
        '7         8          9        10         11        12          13
        'hdrencum  subencum  hdrspent  subspent  hdrunpaid  subunpaid  hdravailable
        '14:
        'subavailable()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim outer, inner, holdindex As Int32
        Dim account, flagindex, subaccount, prevaccount As String
        Dim HdrVoidamount, SubVoidamount, Amount As Decimal

        Try
            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, h.af_acct_name, d.as_acct_num, d.as_acct_name," _
            & " (af_beg_month_balance + af_mtd_receipts - af_mtd_expend + af_mtd_adjust) AS CurrentBalance," _
            & " (as_beg_month_balance + as_mtd_receipts - as_mtd_expend + as_mtd_adjust) AS SCurrentBalance," _
            & " (af_ytd_encumbered + af_mtd_encumbered) as Encumbered," _
            & " (as_ytd_encumbered + as_mtd_encumbered) as SEncumbered," _
            & " (af_mtd_expend + af_ytd_expend) AS TotalSpent," _
            & " (as_mtd_expend + as_ytd_expend) AS STotalSpent," _
            & " (af_ytd_encumbered + af_mtd_encumbered) - (af_mtd_expend + af_ytd_expend) AS Unpaid," _
            & " (as_ytd_encumbered + as_mtd_encumbered) - (as_mtd_expend + as_ytd_expend) AS SUnpaid," _
            & " (af_beg_month_balance + af_mtd_receipts + af_mtd_adjust - af_mtd_expend) - ((af_ytd_encumbered + af_mtd_encumbered) - (af_mtd_expend + af_ytd_expend)) AS Available," _
            & " (as_beg_month_balance + as_mtd_receipts + as_mtd_adjust - as_mtd_expend) - ((as_ytd_encumbered + as_mtd_encumbered) - (as_mtd_expend + as_ytd_expend)) AS SAvailable," _
            & " 0.00 AS HdrVoidamount, 0.00 AS SubVoidamount" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1 AND h.bank_acct_num = d.bank_acct_num AND h.af_acct_num = d.af_acct_num" _
            & " And af_status = 'O' And as_status = 'O'" _
            & " ORDER BY h.af_acct_num, d.as_acct_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("Encumbrances")
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
            SSQL = "SELECT af_acct_num, as_acct_num, ISNULL(sum(invc_amount), 0.0) AS Voids" _
            & " FROM invoices WHERE bank_acct_num = @p1 AND invc_fisyr = @p2 AND invc_status = 'V'" _
            & " GROUP BY af_acct_num, as_acct_num" _
            & " UNION" _
            & " SELECT af_acct_num, '-1', ISNULL(sum(invc_amount), 0.0) AS Voids" _
            & " FROM invoices" _
            & " WHERE bank_acct_num = @p1 AND invc_fisyr = @p2 AND invc_status = 'V'" _
            & " GROUP BY af_acct_num" _
            & " ORDER BY af_acct_num, as_acct_num"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("Voids")
            da.Fill(tbl)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            For outer = 0 To tbl.Rows.Count - 1
                account = CType(tbl.Rows(outer).Item(0), String)
                subaccount = CType(tbl.Rows(outer).Item(1), String)
                If subaccount = "-1" Then
                    HdrVoidamount = CType(tbl.Rows(outer).Item(2), Decimal)
                Else
                    Amount = CType(tbl.Rows(outer).Item(2), Decimal)
                End If

                If subaccount = "-1" Then
                    With Me.GridDetail
                        For inner = 0 To .Rows.Count - 1
                            If CType(.GetData(inner, 1), String) = account Then
                                .SetData(inner, 15, HdrVoidamount)
                            End If
                        Next
                    End With
                End If

                If subaccount <> "-1" Then
                    With Me.GridDetail
                        For inner = 0 To .Rows.Count - 1
                            If CType(.GetData(inner, 1), String) = account Then
                                If CType(.GetData(inner, 3), String) = subaccount Then
                                    .SetData(inner, 16, Amount)
                                    Exit For
                                End If
                            End If
                        Next
                    End With
                End If

                Application.DoEvents()
            Next
        Catch ex As Exception
            Throw
        End Try

        ''Me.GridDetail.Visible = True
        ''Me.GridTotals.Visible = False
        ''Me.Prev1.Visible = False
        ''Me.ShowDialog()
        ''Exit Function

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "YTD Encumbrance Summary"
            Application.DoEvents()
            'render the table;
            Call PrintExpenditureSummaryOfEncumbrances(Me.FiscalYear)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateHistoricalMTDSummaryOfAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eselectedmonth As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '    0           1         2         3         4         5        6
        '   bank      acctnum   acctname   begbal   mtdrcpt   mtdexp   mtdadj
        '    7           8 
        '  total     closedate
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim SSQL As String
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim cmd As SqlCommand
        Dim monthval As Int32
        Dim closedate As Date

        Try
            'set the current month string;
            Me.CurrentMonthString = eselectedmonth
            'get the cardinal value for the current month;
            monthval = Module1.ConvertMonthStringToCardinal(eselectedmonth)
            '
            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
            & " SUM(ahst_beg_month_balance), SUM(ahst_mtd_receipts), SUM(ahst_mtd_expend), SUM(ahst_mtd_adjust)," _
            & " SUM(ahst_beg_month_balance + ahst_mtd_receipts - ahst_mtd_expend + ahst_mtd_adjust) AS computedtotal" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = @p2" _
            & " AND ahst_current_month = @p3" _
            & " GROUP BY bank_acct_num, af_acct_num, af_acct_name" _
            & " ORDER BY bank_acct_num, af_acct_num; "
            SSQL += "SELECT SUM(ahst_beg_month_balance) AS begbal, SUM(ahst_mtd_receipts) AS mtdrcpt," _
            & " SUM(ahst_mtd_expend) AS mtdexpend, SUM(ahst_mtd_adjust) AS mtdadj" _
            & " FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = @p2" _
            & " AND ahst_current_month = @p3; "
            SSQL += "SELECT TOP 1 ahst_datetime FROM acct_history" _
            & " WHERE bank_acct_num = @p1" _
            & " AND ahst_fisyr = @p2" _
            & " AND ahst_current_month = @p3"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", monthval)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned;
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail;
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
            'get the closeout date from the 3rd table;
            If ds.Tables(2).Rows.Count = 1 Then closedate = CDate(ds.Tables(2).Rows(0).Item(0))
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "Closed: " & closedate.ToShortDateString
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccounts(efiscalyear, False, False)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateHistoricalMTDSummaryOfSubaccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32, ByVal eselectedmonth As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14        15  
        ' astotal  closedate
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'get the cardinal value for the current month 
        If eselectedmonth Is Nothing OrElse eselectedmonth.Trim.Length = 0 Then Throw New ArgumentException("No applied month has been selected...")
        Dim monthval As Int32 = Module1.ConvertMonthStringToCardinal(eselectedmonth)
        'set the current month string
        Me.CurrentMonthString = eselectedmonth

        Dim SSQL As String
        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
        & " 0.0 AS afbegmobal, 0.0 AS afmtdrcpt, 0.0 AS afmtdexpend, 0.0 AS afmtdadj, 0.0 AS computedaccttotal," _
        & " ahst_beg_month_balance, ahst_mtd_receipts, ahst_mtd_expend, ahst_mtd_adjust," _
        & " (ahst_beg_month_balance + ahst_mtd_receipts - ahst_mtd_expend + ahst_mtd_adjust) AS computedsubtotal," _
        & " ahst_datetime" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = @p3" _
        & " ORDER BY bank_acct_num, af_acct_num, as_acct_num; "
        SSQL += "SELECT SUM(ahst_beg_month_balance) AS begbal, SUM(ahst_mtd_receipts) AS mtdrcpt," _
        & " SUM(ahst_mtd_expend) AS mtdexpend, SUM(ahst_mtd_adjust) AS mtdadj" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = @p3"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        cmd.Parameters.Add("@p3", monthval)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("summary")
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
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'NOTE:  The following does not have to be performed on a regular Summary
        '       Of Accounts since the header accounts are already summarised, but
        '       but the history record is a breakdown of the subaccounts only and
        '       does not have a summary value for the header account;
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'summarise the subaccount balances into the header balances
        Dim closedate As Date
        Dim currow, totalrecs, j, k As Int32
        Dim curacct, holdacct As String
        Dim subbegbal, subrcpt, subexp, subadj, subtotal As Double
        Dim mtdbegbal, mtdrcpt, mtdexp, mtdadj, mtdtotal As Double
        Try
            With Me.GridDetail
                '.Cols(5).Width = 50
                '.Cols(6).Width = 50
                '.Cols(7).Width = 50
                '.Cols(8).Width = 50
                '.Cols(9).Width = 70
                '.Cols(0).Visible = False
                '.Cols(2).Visible = False
                '.Cols(4).Visible = False

                'get the closeout date from the first rec
                closedate = CDate(.GetData(0, 15))

                'summarise the detail amounts into the header record
                For currow = 0 To .Rows.Count - 1
                    curacct = DirectCast(.GetData(currow, 1), String)
                    For j = currow To .Rows.Count - 1
                        holdacct = DirectCast(.GetData(j, 1), String)
                        subbegbal = CDbl(.GetData(j, 10))
                        subrcpt = CDbl(.GetData(j, 11))
                        subexp = CDbl(.GetData(j, 12))
                        subadj = CDbl(.GetData(j, 13))
                        subtotal = CDbl(.GetData(j, 14))
                        If curacct = holdacct Then
                            mtdbegbal += subbegbal
                            mtdrcpt += subrcpt
                            mtdexp += subexp
                            mtdadj += subadj
                            mtdtotal += subtotal
                        Else
                            Exit For
                        End If
                    Next
                    'now set the totals for all pertinent records
                    For k = currow To j - 1
                        .SetData(k, 5, mtdbegbal)
                        .SetData(k, 6, mtdrcpt)
                        .SetData(k, 7, mtdexp)
                        .SetData(k, 8, mtdadj)
                        .SetData(k, 9, mtdtotal)
                    Next
                    currow = j - 1
                    mtdbegbal = 0
                    mtdrcpt = 0
                    mtdexp = 0
                    mtdadj = 0
                    mtdtotal = 0
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            'define header cells for the report 
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "Closed: " & closedate.ToShortDateString
            Application.DoEvents()
            'render the table
            Call PrintSummaryOfSubaccounts(efiscalyear, False, False)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateHistoricalYTDSummaryOfAccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '    0           1         2         3         4         5        6
        '   bank      acctnum   acctname   begbal   mtdrcpt   mtdexp   mtdadj
        '    7
        '  total
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
        & " SUM(ahst_beg_year_balance) AS beginningbalance," _
        & " SUM(ahst_mtd_receipts + ahst_ytd_receipts) AS totalreceipts," _
        & " SUM(ahst_mtd_expend + ahst_ytd_expend) AS totalexpenditures," _
        & " SUM(ahst_mtd_adjust + ahst_ytd_adjust) AS totaladjustments," _
        & " SUM(ahst_beg_year_balance +" _
        & " (ahst_mtd_receipts - ahst_mtd_expend + ahst_mtd_adjust) +" _
        & " (ahst_ytd_receipts - ahst_ytd_expend + ahst_ytd_adjust)) AS grandtotal" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = 6" _
        & " GROUP BY bank_acct_num, af_acct_num, af_acct_name" _
        & " ORDER BY af_acct_num; "
        SSQL += "SELECT SUM(ahst_beg_year_balance) AS begbal, SUM(ahst_mtd_receipts + ahst_ytd_receipts) AS ytdrcpt," _
        & " SUM(ahst_mtd_expend + ahst_ytd_expend) AS ytdexpend, SUM(ahst_mtd_adjust + ahst_ytd_adjust) AS ytdadj" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = 6"
        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("summary")
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
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "Historical YTD Summary"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccounts(efiscalyear, False, False)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateHistoricalYTDSummaryOfSubaccounts(ByVal ebankaccountnumber As String, ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afytdrcpt
        '    7         8         9         10        11         12        13
        ' afytdexp  afytdadj  afending  asbegbal  asytdrcpt  asytdexp  asytdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
        & " 0.0 AS AFBeg, 0.0 AS AFYTDRcpt, 0.0 AS AFYTDExp, 0.0 AS AFYTDAdj, 0.0 AS AFYTDEnd," _
        & " ahst_beg_year_balance, (ahst_mtd_receipts + ahst_ytd_receipts), (ahst_mtd_expend + ahst_ytd_expend)," _
        & " (ahst_mtd_adjust + ahst_ytd_adjust)," _
        & " (ahst_beg_year_balance + (ahst_mtd_receipts + ahst_ytd_receipts) - (ahst_mtd_expend + ahst_ytd_expend) +" _
        & " (ahst_mtd_adjust + ahst_ytd_adjust)) AS computedsubtotal," _
        & " af_acct_num + as_acct_num" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = 6" _
        & " ORDER BY af_acct_num, as_acct_num; "

        SSQL += "SELECT SUM(ahst_beg_year_balance) AS begbal, SUM(ahst_mtd_receipts + ahst_ytd_receipts) AS ytdrcpt," _
        & " SUM(ahst_mtd_expend + ahst_ytd_expend) AS ytdexpend, SUM(ahst_mtd_adjust + ahst_ytd_adjust) AS ytdadj" _
        & " FROM acct_history" _
        & " WHERE bank_acct_num = @p1" _
        & " AND ahst_fisyr = @p2" _
        & " AND ahst_current_month = 6"

        cn = New SqlConnection(Me.ConnectionString)
        Dim cmd As New SqlCommand(SSQL, cn)
        cmd.Parameters.Add("@p1", ebankaccountnumber)
        cmd.Parameters.Add("@p2", efiscalyear)
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet("summary")
        Try
            da.Fill(ds)
            cn.Close()
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afytdrcpt
        '    7         8         9         10        11         12        13
        ' afytdexp  afytdadj  aftotal   asbegbal  asytdrcpt  asytdexp  asytdadj
        '   14        15  
        ' astotal   acct-sub
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim acct, tempacct As String
        Dim sumbal, sumrcpt, sumexpend, sumadj, sumall As Double
        Dim tempbal, temprcpt, tempexpend, tempadj As Double
        Dim j, k, index As Int32
        With Me.GridDetail
            'summarise the receipt lines into a header amt
            For index = 0 To .Rows.Count - 1
                acct = DirectCast(.GetData(index, 1), String)

                For j = index To .Rows.Count - 1
                    tempacct = DirectCast(.GetData(j, 1), String)
                    tempbal = CDbl(.GetData(j, 10))
                    temprcpt = CDbl(.GetData(j, 11))
                    tempexpend = CDbl(.GetData(j, 12))
                    tempadj = CDbl(.GetData(j, 13))
                    If acct = tempacct Then
                        sumbal += tempbal
                        sumrcpt += temprcpt
                        sumexpend += tempexpend
                        sumadj += tempadj
                    Else
                        sumall = sumbal + sumrcpt - sumexpend + sumadj
                        Exit For
                    End If
                    'handle last record;
                    If j = .Rows.Count - 1 Then
                        sumall = sumbal + sumrcpt - sumexpend + sumadj
                    End If
                Next


                For k = index To j - 1
                    .SetData(k, 5, sumbal)
                    .SetData(k, 6, sumrcpt)
                    .SetData(k, 7, sumexpend)
                    .SetData(k, 8, sumadj)
                    .SetData(k, 9, sumall)
                Next
                index = j - 1
                sumbal = 0
                sumrcpt = 0
                sumexpend = 0
                sumadj = 0
                sumall = 0
            Next
        End With

        '''''With Me.GridDetail
        '''''    .Cols(0).Visible = False
        '''''    .Cols(5).Width = 75
        '''''    .Cols(6).Width = 75
        '''''    .Cols(7).Width = 75
        '''''    .Cols(8).Width = 75
        '''''    .Cols(9).Width = 75
        '''''    .Visible = True
        '''''End With
        '''''Me.GridTotals.Visible = False
        '''''Me.Prev1.Visible = False
        '''''Me.ShowDialog()
        '''''Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & efiscalyear.ToString
            Me.CellMiddleBottom = "Historical YTD Summary"
            Application.DoEvents()
            'render the table
            PrintSummaryOfSubaccounts(Me.FiscalYear, False, False)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'MTD SUMMARY (ALL OR ACCOUNT RANGE);
    Public Function GenerateMTDSummaryOfAccounts(ByVal ebankaccountnumber As String, ByVal esuppresszero As Boolean, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '    0           1         2         3         4         5        6
        '   bank      acctnum   acctname   begbal   mtdrcpt   mtdexp   mtdadj
        '    7
        '  total
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""

        Try
            If euserange = True Then
                filter = " AND af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
            & " af_beg_month_balance, af_mtd_receipts, af_mtd_expend, af_mtd_adjust," _
            & " (af_beg_month_balance + af_mtd_receipts - af_mtd_expend + af_mtd_adjust) AS computedtotal" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & filter _
            & " ORDER BY af_acct_num; "
            SSQL += "SELECT SUM(af_beg_month_balance) AS begbal, SUM(af_mtd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend) AS mtdexpend, SUM(af_mtd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & filter
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "MTD Partial Summary" Else Me.CellMiddleBottom = "MTD Summary"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccounts(Me.FiscalYear, esuppresszero, euserange)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'MTD SUMMARY ENCUMBRANCE (ALL OR ACCOUNT RANGE);
    Public Function GenerateMTDSummaryOfAccountsWithEncumbrance(ByVal ebankaccountnumber As String, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname  begmobal  mtdrcpt   mtdencum   encbalance
        '    7         8         9         10        11         12        13
        ' mtdexp    mtdadj    current  
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""


        Try
            If euserange = True Then
                filter = " AND af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
            & " af_beg_month_balance, af_mtd_receipts, af_ytd_encumbered," _
            & " af_ytd_encumbered - (af_mtd_expend + af_ytd_expend) AS EncBalance," _
            & " af_mtd_expend, af_mtd_adjust," _
            & " (af_beg_month_balance + af_mtd_receipts - af_mtd_expend + af_mtd_adjust) AS CurrentBalance" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & filter _
            & " ORDER BY af_acct_num; "
            SSQL += "SELECT 0.00 AS begbal, 0.00 AS mtdrcpt, 0.00 AS mtdexpend, 0.00 AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
            '''''Me.GridDetail.Visible = True
            '''''Me.GridTotals.Visible = False
            '''''Me.Prev1.Visible = False
            '''''Me.ShowDialog()
            '''''Exit Function
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "MTD Summary with Encumbrance"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccountsWithEncumbrance()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'MTD SUMMARY SUBACCOUNTS (ALL OR ACCOUNT RANGE);
    Public Function GenerateMTDSummaryOfSubaccounts(ByVal ebankaccountnumber As String, ByVal esuppresszero As Boolean, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""

        Try
            If euserange = True Then
                filter = " AND h.af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " af_beg_month_balance, af_mtd_receipts, af_mtd_expend, af_mtd_adjust," _
            & " (af_beg_month_balance + af_mtd_receipts - af_mtd_expend + af_mtd_adjust) AS computedaccttotal," _
            & " as_beg_month_balance, as_mtd_receipts, as_mtd_expend, as_mtd_adjust," _
            & " (as_beg_month_balance + as_mtd_receipts - as_mtd_expend + as_mtd_adjust) AS computedsubtotal" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND af_status = 'O' AND as_status = 'O'" _
            & filter _
            & " ORDER BY h.af_acct_num, as_acct_num; "
            SSQL += "SELECT SUM(af_beg_month_balance) AS begbal, SUM(af_mtd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend) AS mtdexpend, SUM(af_mtd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "MTD Partial Summary" Else Me.CellMiddleBottom = "MTD Summary"
            Application.DoEvents()
            'render the table
            Call PrintSummaryOfSubaccounts(Me.FiscalYear, esuppresszero, euserange)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'MTD SUMMARY SUBACCOUNTS ENCUMBRANCE (ALL OR ACCOUNT RANGE);
    Public Function GenerateMTDSummaryOfSubaccountsWithEncumbrance(ByVal ebankaccountnumber As String, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   begmobal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afytdenc ytdencbal  afmtdexp  afmtdadj  afcurrent  begmobal  asmtdrcpt
        '   14        15        16         17        18
        ' asytdenc ytdencbal  asmtdexp  asmtdadj  ascurrent
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""

        Try
            If euserange = True Then
                filter = " AND h.af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " af_beg_month_balance, af_mtd_receipts, af_ytd_encumbered, af_ytd_encumbered - (af_mtd_expend + af_ytd_expend), af_mtd_expend," _
            & " af_mtd_adjust, (af_beg_month_balance + af_mtd_receipts - af_mtd_expend + af_mtd_adjust) AS computedaccttotal," _
            & " as_beg_month_balance, as_mtd_receipts, as_ytd_encumbered, as_ytd_encumbered - (as_mtd_expend + as_ytd_expend), as_mtd_expend," _
            & " as_mtd_adjust, (as_beg_month_balance + as_mtd_receipts - as_mtd_expend + as_mtd_adjust) AS computedsubtotal" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND af_status = 'O'" _
            & filter _
            & " ORDER BY h.af_acct_num, as_acct_num; "
            SSQL += "SELECT SUM(af_beg_month_balance) AS begbal, SUM(af_mtd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend) AS mtdexpend, SUM(af_mtd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
            '''''Me.GridDetail.Visible = True
            '''''Me.GridTotals.Visible = False
            '''''Me.Prev1.Visible = False
            '''''Me.ShowDialog()
            '''''Exit Function
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleMiddle = Me.CurrentMonthString & ", FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "MTD Partial Summary" Else Me.CellMiddleBottom = "MTD Summary with Encumbrance"
            Application.DoEvents()
            'render the table
            Call PrintSummaryOfSubaccountsWithEncumbrance(Me.FiscalYear)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'YTD SUMMARY (ALL OR ACCOUNT RANGE);
    Public Function GenerateYTDSummaryOfAccounts(ByVal ebankaccountnumber As String, ByVal esuppresszero As Boolean, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '    0           1         2         3         4         5        6
        '   bank      acctnum   acctname   begbal   ytdrcpt   ytdexp   ytdadj
        '    7
        '  total
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String

        Try
            If euserange = True Then
                filter = " AND af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
            & " af_beg_year_balance, af_mtd_receipts + af_ytd_receipts," _
            & " af_mtd_expend + af_ytd_expend, af_mtd_adjust + af_ytd_adjust," _
            & " (af_beg_year_balance +" _
            & " (af_mtd_receipts - af_mtd_expend + af_mtd_adjust) +" _
            & " (af_ytd_receipts - af_ytd_expend + af_ytd_adjust))" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & filter _
            & " ORDER BY af_acct_num; "
            SSQL += "SELECT SUM(af_beg_year_balance) AS begbal, SUM(af_mtd_receipts + af_ytd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend + af_ytd_expend) AS mtdexpend, SUM(af_mtd_adjust + af_ytd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
        Catch ex As Exception
            Throw
        End Try

        'Me.GridDetail.Visible = True
        'Me.GridTotals.Visible = False
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "YTD Partial Summary" Else Me.CellMiddleBottom = "YTD Summary"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccounts(Me.FiscalYear, esuppresszero, euserange)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'YTD SUMMARY SUBACCOUNTS (ALL OR ACCOUNT RANGE);
    Public Function GenerateYTDSummaryOfSubaccounts(ByVal ebankaccountnumber As String, ByVal esuppresszero As Boolean, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afytdrcpt
        '    7         8         9         10        11         12        13
        ' afytdexp  afytdadj  aftotal   asbegbal  asytdrcpt  asytdexp  asytdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String

        Try
            If euserange = True Then
                filter = " AND h.af_acct_num BETWEEN @p2 AND @p3"
            End If

            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " af_beg_year_balance, (af_mtd_receipts + af_ytd_receipts), (af_mtd_expend + af_ytd_expend)," _
            & " (af_mtd_adjust + af_ytd_adjust)," _
            & " (af_beg_year_balance + (af_mtd_receipts + af_ytd_receipts) - (af_mtd_expend + af_ytd_expend) +" _
            & " (af_mtd_adjust + af_ytd_adjust)) AS computedaccttotal," _
            & " as_beg_year_balance, (as_mtd_receipts + as_ytd_receipts), (as_mtd_expend + as_ytd_expend)," _
            & " (as_mtd_adjust + as_ytd_adjust)," _
            & " (as_beg_year_balance + (as_mtd_receipts + as_ytd_receipts) - (as_mtd_expend + as_ytd_expend) +" _
            & " (as_mtd_adjust + as_ytd_adjust)) AS computedsubtotal" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND af_status = 'O' AND as_status = 'O'" _
            & filter _
            & " ORDER BY h.af_acct_num, as_acct_num; "
            SSQL += "SELECT SUM(af_beg_year_balance) AS begbal, SUM(af_mtd_receipts + af_ytd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend + af_ytd_expend) AS mtdexpend, SUM(af_mtd_adjust + af_ytd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", eaccountbeginning)
            cmd.Parameters.Add("@p3", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)
            '''''Me.GridDetail.Visible = True
            '''''Me.GridTotals.Visible = False
            '''''Me.Prev1.Visible = False
            '''''Me.ShowDialog()
            '''''Exit Function
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "YTD Partial Summary" Else Me.CellMiddleBottom = "YTD Summary"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfSubaccounts(Me.FiscalYear, esuppresszero, False)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'YTD SUMMARY ENCUMBRANCE (ALL OR ACCOUNT RANGE);
    Public Function GenerateYTDSummaryOfAccountsWithEncumbrance(ByVal ebankaccountnumber As String, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname  begyrbal  ytdrcpt   ytdencum   encbalance
        '    7         8         9         10        11         12        13
        ' ytdexp    ytdadj    current    voided
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""
        Dim efiscalyear As Int32

        Try
            If euserange = True Then
                filter = " AND af_acct_num BETWEEN @p3 AND @p4"
            End If
            efiscalyear = Me.FiscalYear

            SSQL = "SELECT bank_acct_num, af_acct_num, af_acct_name," _
            & " af_beg_year_balance, af_mtd_receipts + af_ytd_receipts," _
            & " af_mtd_encumbered + af_ytd_encumbered, 0," _
            & " af_mtd_expend + af_ytd_expend, af_mtd_adjust + af_ytd_adjust," _
            & " (af_beg_year_balance +" _
            & " (af_mtd_receipts - af_mtd_expend + af_mtd_adjust) +" _
            & " (af_ytd_receipts - af_ytd_expend + af_ytd_adjust)), 0.0" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'" _
            & filter _
            & " ORDER BY af_acct_num; "

            SSQL += "SELECT SUM(af_beg_year_balance) AS begbal, SUM(af_mtd_receipts + af_ytd_receipts) AS ytdrcpt," _
            & " SUM(af_mtd_expend + af_ytd_expend) AS ytdexpend, SUM(af_mtd_adjust + af_ytd_adjust) AS ytdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O'; "

            SSQL += "SELECT af_acct_num, ISNULL(SUM(invc_amount), 0.0) AS invc_void_amt" _
            & " FROM invoices WHERE bank_acct_num = @p1 AND invc_fisyr = @p2" _
            & " AND invc_status = 'V'" _
            & " GROUP BY af_acct_num" _
            & " ORDER BY af_acct_num"

            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", eaccountbeginning)
            cmd.Parameters.Add("@p4", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)

            Dim account As String
            Dim voidamount As Decimal

            'iterate thru the void table and update detail grid;
            For Each row As DataRow In ds.Tables(2).Rows
                account = CType(row.Item(0), String)
                voidamount = CType(row.Item(1), Decimal)
                With Me.GridDetail
                    For index As Int32 = 0 To Me.GridDetail.Rows.Count - 1
                        If String.Compare(account, CType(.GetData(index, 1), String)) = 0 Then
                            .SetData(index, 10, voidamount)
                            Exit For
                        End If
                    Next
                End With
            Next

            '''''Me.GridDetail.Visible = True
            '''''Me.GridTotals.Visible = False
            '''''Me.Prev1.Visible = False
            '''''Me.ShowDialog()
            '''''Exit Function
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYear.ToString
            Me.CellMiddleBottom = "YTD Summary with Encumbrance"
            If euserange = True Then Me.CellMiddleBottom = "YTD Partial Summary with Encumbrance" Else Me.CellMiddleBottom = "YTD Summary with Encumbrance"
            Application.DoEvents()
            'render the table;
            Call PrintSummaryOfAccountsWithEncumbrance()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    'YTD SUMMARY SUBACCOUNTS ENCUMBRANCE (ALL OR ACCOUNT RANGE);
    Public Function GenerateYTDSummaryOfSubaccountsWithEncumbrance(ByVal ebankaccountnumber As String, ByVal euserange As Boolean, ByVal eaccountbeginning As String, ByVal eaccountending As String) As Boolean
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   begmobal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afytdenc ytdencbal  afmtdexp  afmtdadj  afending  begmobal  asmtdrcpt
        '   14        15        16         17        18         19 
        ' asytdenc ytdencbal  asmtdexp  asmtdadj  ascurrent   voided 
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet
        Dim filter As String = ""
        Dim efiscalyear As Int32

        Try
            If euserange = True Then
                filter = " AND h.af_acct_num BETWEEN @p3 AND @p4"
            End If
            efiscalyear = Me.FiscalYear

            SSQL = "SELECT h.bank_acct_num, h.af_acct_num, af_acct_name, as_acct_num, as_acct_name," _
            & " af_beg_year_balance, af_mtd_receipts + af_ytd_receipts AS YtdRevenue, af_mtd_encumbered + af_ytd_encumbered," _
            & " 0, af_mtd_expend + af_ytd_expend AS YtdExpense, af_mtd_adjust + af_ytd_adjust AS YtdAdjust," _
            & " (af_beg_month_balance + (af_mtd_receipts + af_ytd_receipts) - (af_mtd_expend + af_ytd_expend) + af_mtd_adjust) AS AccountTotal," _
            & " as_beg_year_balance, as_mtd_receipts + as_ytd_receipts AS SubRevenue, as_mtd_encumbered + as_ytd_encumbered," _
            & " 0, as_mtd_expend + as_ytd_expend AS SubExpense, as_mtd_adjust + as_ytd_adjust AS SubAdjust," _
            & " (as_beg_month_balance + (as_mtd_receipts + as_ytd_receipts) - (af_mtd_expend + as_ytd_expend) + as_mtd_adjust) AS SubaccountTotal, 0.0 AS Void" _
            & " FROM acct_info AS h, acct_sub AS d" _
            & " WHERE h.bank_acct_num = @p1" _
            & " AND h.bank_acct_num = d.bank_acct_num" _
            & " AND h.af_acct_num = d.af_acct_num" _
            & " AND af_status = 'O' And as_status = 'O'" _
            & filter _
            & " ORDER BY h.af_acct_num, as_acct_num; "

            SSQL += "SELECT SUM(af_beg_month_balance) AS begbal, SUM(af_mtd_receipts) AS mtdrcpt," _
            & " SUM(af_mtd_expend) AS mtdexpend, SUM(af_mtd_adjust) AS mtdadj" _
            & " FROM acct_info" _
            & " WHERE bank_acct_num = @p1" _
            & " AND af_status = 'O';"

            SSQL += "SELECT af_acct_num,as_acct_num, ISNULL(SUM(invc_amount), 0.0) AS invc_void_amt" _
            & " FROM invoices WHERE bank_acct_num = @p1 AND invc_fisyr = @p2" _
            & " AND invc_status = 'V'" _
            & " GROUP BY af_acct_num, as_acct_num" _
            & " ORDER BY af_acct_num, as_acct_num"

            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", ebankaccountnumber)
            cmd.Parameters.Add("@p2", efiscalyear)
            cmd.Parameters.Add("@p3", eaccountbeginning)
            cmd.Parameters.Add("@p4", eaccountending)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("Summary")
            da.Fill(ds)
        Catch ex As Exception
            Throw
        Finally
            cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''this was added in january 2015''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try
            'throw error if no accounts are returned
            If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
            'datasource the detail
            Me.GridDetail.DataSource = ds.Tables(0)
            Me.GridTotals.DataSource = ds.Tables(1)

            Dim account As String
            Dim acctsub As String
            Dim voidamount As Decimal

            'iterate thru the void table and update detail grid;
            For Each row As DataRow In ds.Tables(2).Rows
                account = CType(row.Item(0), String)
                acctsub = CType(row.Item(1), String)
                voidamount = CType(row.Item(2), Decimal)
                With Me.GridDetail
                    For index As Int32 = 0 To Me.GridDetail.Rows.Count - 1
                        If String.Compare(account, CType(.GetData(index, 1), String)) = 0 Then
                            If String.Compare(acctsub, CType(.GetData(index, 3), String)) = 0 Then
                                .SetData(index, 19, voidamount)
                                Exit For
                            End If
                        End If
                    Next
                End With
            Next

            ''Me.GridDetail.Visible = True
            ''Me.GridTotals.Visible = False
            ''Me.Prev1.Visible = True
            ''Me.ShowDialog()
            ''Exit Function
        Catch ex As Exception
            Throw
        End Try

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''Try
        ''''''    'throw error if no accounts are returned
        ''''''    If ds.Tables(0).Rows.Count < 1 Then Throw New ArgumentException("No records found for the selected criteria...")
        ''''''    'datasource the detail
        ''''''    Me.GridDetail.DataSource = ds.Tables(0)
        ''''''    Me.GridTotals.DataSource = ds.Tables(1)
        ''''''    '''''Me.GridDetail.Visible = True
        ''''''    '''''Me.GridTotals.Visible = False
        ''''''    '''''Me.Prev1.Visible = False
        ''''''    '''''Me.ShowDialog()
        ''''''    '''''Exit Function
        ''''''Catch ex As Exception
        ''''''    Throw
        ''''''End Try

        Try
            Me.CellMiddleMiddle = "FY-" & Me.FiscalYear.ToString
            If euserange = True Then Me.CellMiddleBottom = "YTD Partial Summary with Encumbrance" Else Me.CellMiddleBottom = "YTD Summary with Encumbrance"
            Application.DoEvents()
            'render the table
            Call PrintSummaryOfSubaccountsWithEncumbrance(Me.FiscalYear)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region "  Methods Rendering "

    Private Sub PrintBalanceSheetAccounts(ByVal efiscalyear As Int32, ByVal emonth As Int32)
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "BalanceSheet"
        Me.ReportName = "Activity Fund Balance Sheet"
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

        Dim index As Int32

        Try
            Dim tacctnum, tacctname, tsubnum, tsubname As String
            Dim prevbank, prevacctnum As String
            Dim x, y, totalaccounts, totalsubaccounts As Int32
            Dim hdrbegbalance, hdrrevenue, hdrexpenditure, hdradjustment, hdrnewbalance As Double
            Dim sumbegbalance, sumrcpt, sumexpend, sumadj, tempbal As Double
            Dim totbegbalance, totrevenue, totexpenditure As Double
            Dim dailyhdrbalance As Double
            Dim days As Int32
            Dim dopagebreak As Boolean

            'get the total number of days in the current month
            days = Me.AppliedDate.DaysInMonth(efiscalyear, emonth)

            With Me.GridTotals
                'get the sum of the mtd totals
                sumbegbalance = CDbl(.GetData(0, 0))
                sumrcpt = CDbl(.GetData(0, 1))
                sumexpend = CDbl(.GetData(0, 2))
                sumadj = CDbl(.GetData(0, 3))
            End With

            'start the document
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    'hdr vals
                    hdrbegbalance = CDbl(.GetData(index, 5))
                    hdrrevenue = CDbl(.GetData(index, 6))
                    hdrexpenditure = CDbl(.GetData(index, 7))
                    hdradjustment = CDbl(.GetData(index, 8))
                    'since the adjustment column does not exist in this report, we'll add
                    'the adjustment to either the revenue or expenditure amount if + or -;
                    If hdradjustment >= 0 Then
                        hdrrevenue += hdradjustment
                    Else
                        hdrexpenditure += (hdradjustment * -1)
                    End If
                    hdrnewbalance = CDbl(.GetData(index, 9))
                    hdradjustment = 0
                End With

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        .RenderDirectText(118, 32, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 36, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 40, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 44, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 50, "Ending balance:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(150, 32, sumbegbalance.ToString.Format("{0:F2}", sumbegbalance), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 36, sumrcpt.ToString.Format("{0:F2}", sumrcpt), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 40, sumexpend.ToString.Format("{0:F2}", sumexpend), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 44, sumadj.ToString.Format("{0:F2}", sumadj), 40, 5, verdanaright8bold)
                        'draw a line under the sums
                        .RenderDirectLine(162, 48.5, 189, 48.5, Color.Black, 0.75)
                        'calc the ending balance
                        tempbal = sumbegbalance + sumrcpt + sumadj - sumexpend
                        .RenderDirectText(150, 50, tempbal.ToString.Format("{0:C2}", tempbal), 40, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Revenues", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditures", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "End.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Daily Balance", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        'calc the daily balance for the header
                        dailyhdrbalance = hdrnewbalance / days
                        'print the header account information
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                        .RenderDirectText(65, y, hdrbegbalance.ToString.Format("{0:F2}", hdrbegbalance), 25, 5, verdanaright8)
                        .RenderDirectText(90, y, hdrrevenue.ToString.Format("{0:F2}", hdrrevenue), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, hdrexpenditure.ToString.Format("{0:F2}", hdrexpenditure), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, hdrnewbalance.ToString.Format("{0:F2}", hdrnewbalance), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, dailyhdrbalance.ToString.Format("{0:F2}", dailyhdrbalance), 25, 5, verdanaright8)
                        prevacctnum = tacctnum
                    End If

                    y += 8  'crlf

                    'check for page break & print new column headers if true
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break
                        .NewPage()
                        y = 33
                        'print the column headers
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Revenues", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditures", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "End.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Daily Balance", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        'currec = 1
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

    Private Sub PrintStatementOfChange(ByVal efiscalyear As Int32)
        Me.DocumentName = "BalanceSheet"
        Me.ReportName = "Statement Of Change"
        'define styles;
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
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim currec, index, x, y As Int32
        Dim sumbeginning, sumexpense, sumrevenue, sumending As Decimal
        Dim tbanknum, tacctnum, tacctname, tsubnum, tsubname As String
        Dim prevbank, prevacctnum As String
        Dim beghdr, endhdr, expensehdr, revhdr As Decimal
        Dim begsub, endsub, expensesub, revsub As Decimal
        Dim dopagebreak As Boolean

        Try
            'iterate thru the table and collect summary balances from sub balances;
            With Me.GridDetail
                For index = 0 To .Rows.Count - 1
                    sumbeginning += CType(.GetData(index, 10), Decimal)
                    sumrevenue += CType(.GetData(index, 17), Decimal)
                    sumexpense += CType(.GetData(index, 18), Decimal)
                Next
                sumending = sumbeginning + sumrevenue - sumexpense
            End With
        Catch ex As Exception
            Throw
        End Try


        Try
            'start the document;
            Me.Doc1.StartDoc()
            currec = 1

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid;
                With Me.GridDetail
                    'tbanknum = CType(.GetData(index, 0), String)
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = CType(.GetData(index, 1), String)
                    tacctname = CType(.GetData(index, 2), String)
                    tsubnum = CType(.GetData(index, 3), String)
                    tsubname = CType(.GetData(index, 4), String)
                    'get the new revenue & expense header amounts and calc ending balance;
                    beghdr = CType(.GetData(index, 5), Decimal)
                    revhdr = CType(.GetData(index, 15), Decimal)
                    expensehdr = CType(.GetData(index, 16), Decimal)
                    endhdr = beghdr + revhdr - expensehdr
                    'get the new revenue & expense sub amounts and calc ending balance;
                    begsub = CType(.GetData(index, 10), Decimal)
                    revsub = CType(.GetData(index, 17), Decimal)
                    expensesub = CType(.GetData(index, 18), Decimal)
                    endsub = begsub + revsub - expensesub
                End With

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the total info box left-side;
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        '.RenderDirectText(1, 40, tbanknum, 40, 5, verdanaright8)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side;
                        .RenderDirectText(118, 32, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 37, "Revenue:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 42, "Expense:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 49, "Ending balance:", 40, 5, verdanaright8bold)
                        'print the money fields;
                        .RenderDirectText(150, 32, sumbeginning.ToString.Format("{0:C2}", sumbeginning), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 37, sumrevenue.ToString.Format("{0:C2}", sumrevenue), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 42, sumexpense.ToString.Format("{0:C2}", sumexpense), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 49, sumending.ToString.Format("{0:C2}", sumending), 40, 5, verdanaright8bold)
                        'draw a line under the sums;
                        .RenderDirectLine(162, 47.5, 189, 47.5, Color.Black, 0.75)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Beginning", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Revenue", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Expense", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        If currec <> 1 Then y += 5
                        'print the header account information;
                        .RenderDirectRectangle(0, y, 9, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(10, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        'print account number & name;
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 80, 10, arialleft8)
                        'print account balances;
                        .RenderDirectText(90, y, beghdr.ToString.Format("{0:F2}", beghdr), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, revhdr.ToString.Format("{0:F2}", revhdr), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, expensehdr.ToString.Format("{0:F2}", expensehdr), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, endhdr.ToString.Format("{0:F2}", endhdr), 25, 5, verdanaright8)
                        prevacctnum = tacctnum
                        y += 5
                    End If
                    'print the sub account information;
                    .RenderDirectText(2, y, tsubnum, 8, 5, arialleft8)
                    .RenderDirectText(12, y, tsubname, 80, 10, arialleft8)
                    .RenderDirectText(90, y, begsub.ToString.Format("{0:F2}", begsub), 25, 5, verdanaright8)
                    .RenderDirectText(115, y, revsub.ToString.Format("{0:F2}", revsub), 25, 5, verdanaright8)
                    .RenderDirectText(140, y, expensesub.ToString.Format("{0:F2}", expensesub), 25, 5, verdanaright8)
                    .RenderDirectText(165, y, endsub.ToString.Format("{0:F2}", endsub), 25, 5, verdanaright8)
                    '
                    currec += 1
                    y += 5  'crlf;

                    'check for page break & print new column headers if true;
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 34
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(90, y, "Beginning", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Revenue", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Expense", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
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

    Private Sub PrintBalanceSheetSubaccounts(ByVal efiscalyear As Int32, ByVal emonth As Int32)
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "BalanceSheet"
        Me.ReportName = "Activity Fund Balance Sheet"
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

        Dim index, currec As Int32

        Try

            Dim tacctnum, tacctname, tsubnum, tsubname As String
            Dim prevbank, prevacctnum As String
            Dim x, y, totalaccounts, totalsubaccounts As Int32
            Dim hdrbegbalance, hdrrevenue, hdrexpenditure, hdradjustment, hdrnewbalance As Double
            Dim subbegbalance, subrevenue, subexpenditure, subadjustment, subnewbalance As Double
            Dim sumbegbalance, sumrcpt, sumexpend, sumadj, tempbal As Double
            Dim totbegbalance, totrevenue, totexpenditure As Double
            Dim dailyhdrbalance, dailydetlbalance As Double
            Dim days As Int32
            Dim dopagebreak As Boolean

            'get the total number of days in the current month
            days = Me.AppliedDate.DaysInMonth(efiscalyear, emonth)

            With Me.GridTotals
                'get the sum of the mtd totals
                sumbegbalance = CDbl(.GetData(0, 0))
                sumrcpt = CDbl(.GetData(0, 1))
                sumexpend = CDbl(.GetData(0, 2))
                sumadj = CDbl(.GetData(0, 3))
            End With

            'start the document
            Me.Doc1.StartDoc()
            currec = 1
            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    tsubnum = DirectCast(.GetData(index, 3), String)
                    tsubname = DirectCast(.GetData(index, 4), String)
                    'hdr vals
                    hdrbegbalance = CDbl(.GetData(index, 5))
                    hdrrevenue = CDbl(.GetData(index, 6))
                    hdrexpenditure = CDbl(.GetData(index, 7))
                    hdradjustment = CDbl(.GetData(index, 8))
                    'since the adjustment column does not exist in this report, we'll add
                    'the adjustment to either the revenue or expenditure amount if + or -;
                    If hdradjustment >= 0 Then
                        hdrrevenue += hdradjustment
                    Else
                        hdrexpenditure += (hdradjustment * -1)
                    End If
                    hdrnewbalance = CDbl(.GetData(index, 9))
                    'sub vals
                    subbegbalance = CDbl(.GetData(index, 10))
                    subrevenue = CDbl(.GetData(index, 11))
                    subexpenditure = CDbl(.GetData(index, 12))
                    subadjustment = CDbl(.GetData(index, 13))
                    subnewbalance = CDbl(.GetData(index, 14))
                    If subadjustment >= 0 Then
                        subrevenue += subadjustment
                    Else
                        subexpenditure += (subadjustment * -1)
                    End If
                    hdradjustment = 0
                    subadjustment = 0
                End With

                'sum the lines for ending totals on last page
                totbegbalance += subbegbalance
                totrevenue += subrevenue
                totexpenditure += subexpenditure

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the total info box left-side
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        .RenderDirectText(118, 32, "Beginning balance:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 36, "Receipts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 40, "Checks:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 44, "Adjustments:", 40, 5, verdanaright8bold)
                        .RenderDirectText(118, 50, "Ending balance:", 40, 5, verdanaright8bold)
                        'print the money fields
                        .RenderDirectText(150, 32, sumbegbalance.ToString.Format("{0:F2}", sumbegbalance), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 36, sumrcpt.ToString.Format("{0:F2}", sumrcpt), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 40, sumexpend.ToString.Format("{0:F2}", sumexpend), 40, 5, verdanaright8bold)
                        .RenderDirectText(150, 44, sumadj.ToString.Format("{0:F2}", sumadj), 40, 5, verdanaright8bold)
                        'draw a line under the sums
                        .RenderDirectLine(162, 48.5, 189, 48.5, Color.Black, 0.75)
                        'calc the ending balance
                        tempbal = sumbegbalance + sumrcpt + sumadj - sumexpend
                        .RenderDirectText(150, 50, tempbal.ToString.Format("{0:C2}", tempbal), 40, 5, verdanaright8bold)
                        'print line above the column headers
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Revenues", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditures", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "End.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Daily Balance", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        If currec <> 1 Then y += 5
                        'calc the daily balance for the header
                        dailyhdrbalance = hdrnewbalance / days
                        'print the header account information
                        .RenderDirectRectangle(0, y, 9, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(10, y, 63, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                        .RenderDirectText(65, y, hdrbegbalance.ToString.Format("{0:F2}", hdrbegbalance), 25, 5, verdanaright8)
                        .RenderDirectText(90, y, hdrrevenue.ToString.Format("{0:F2}", hdrrevenue), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, hdrexpenditure.ToString.Format("{0:F2}", hdrexpenditure), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, hdrnewbalance.ToString.Format("{0:F2}", hdrnewbalance), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, dailyhdrbalance.ToString.Format("{0:F2}", dailyhdrbalance), 25, 5, verdanaright8)
                        prevacctnum = tacctnum
                        y += 5
                    End If
                    'calc the daily balance for the detail
                    dailydetlbalance = subnewbalance / days
                    'print the sub account information
                    .RenderDirectText(2, y, tsubnum, 8, 5, arialleft8)
                    .RenderDirectText(12, y, tsubname, 58, 10, arialleft8)
                    .RenderDirectText(65, y, subbegbalance.ToString.Format("{0:F2}", subbegbalance), 25, 5, verdanaright8)
                    .RenderDirectText(90, y, subrevenue.ToString.Format("{0:F2}", subrevenue), 25, 5, verdanaright8)
                    .RenderDirectText(115, y, subexpenditure.ToString.Format("{0:F2}", subexpenditure), 25, 5, verdanaright8)
                    .RenderDirectText(140, y, subnewbalance.ToString.Format("{0:F2}", subnewbalance), 25, 5, verdanaright8)
                    .RenderDirectText(165, y, dailydetlbalance.ToString.Format("{0:F2}", dailydetlbalance), 25, 5, verdanaright8)

                    currec += 1
                    y += 5  'crlf

                    'check for page break & print new column headers if true
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break
                        .NewPage()
                        y = 33
                        'print the column headers
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Revenues", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditures", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "End.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Daily Balance", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
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

    Private Sub PrintChartOfAccounts()
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5        6
        '   bank      acctnum   acctname  subnum   subname   hdrkey   subkey
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "ChartOfAccounts"
        Me.ReportName = "Chart Of Accounts"
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

        Dim currow, index As Int32

        Try
            Dim tacctnum, tacctname, tsubacctnum, tsubacctname As String
            Dim prevbank As String
            Dim acctkey, subacctkey, prevacctkey, prevsubacctkey, x, y As Int32
            Dim dopagebreak As Boolean

            'start the document
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    tsubacctnum = DirectCast(.GetData(index, 3), String)
                    tsubacctname = DirectCast(.GetData(index, 4), String)
                    acctkey = CInt(.GetData(index, 5))
                    subacctkey = CInt(.GetData(index, 6))
                End With

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the info box left-side
                        .RenderDirectText(4, y, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(4, y + 4, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        .RenderDirectText(100, y, "Total accounts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(100, y + 4, Me.GridDetail.Rows.Count.ToString, 40, 5, verdanaright8)
                        'print the lines under the info box
                        y = 45
                        .RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        .RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        'print the column headers
                        y = 45
                        .RenderDirectText(10, y, "Account", 25, 5, verdanaleft10)
                        .RenderDirectText(55, y, "Name", 25, 5, verdanaleft10)
                        y = 52
                    End If

                    'print the line
                    .RenderDirectText(4, y, tacctnum, 20, 5, verdanaright10bold)
                    .RenderDirectText(55, y, tacctname, 100, 5, verdanaleft10)
                    y += 8
                    currow += 1

                    'check for page break & print new column headers if true
                    If .CurrentPage = 1 And currow = 26 Then dopagebreak = True
                    If .CurrentPage > 1 And currow Mod 29 = 0 Then dopagebreak = True
                    If dopagebreak Then
                        .NewPage()
                        y = 32
                        '.RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        .RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        .RenderDirectText(10, y, "Account", 25, 5, verdanaleft10)
                        .RenderDirectText(55, y, "Name", 25, 5, verdanaleft10)
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

    Private Sub PrintChartOfSubAccounts()
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '     0           1         2        3        4         5        6
        '   bank      acctnum   acctname  subnum   subname   hdrkey   subkey
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "ChartOfSubAccounts"
        Me.ReportName = "Chart Of Accounts"
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

        Dim currow, index As Int32

        Try
            Dim tacctnum, tacctname, tsubacctnum, tsubacctname As String
            Dim prevbank As String
            Dim acctkey, subacctkey, prevacctkey As Int32
            Dim x, y, totalaccounts, totalsubaccounts As Int32
            Dim dopagebreak As Boolean

            With Me.GridTotals
                'get the total accounts/subaccounts
                totalaccounts = CInt(.GetData(0, 0))
                totalsubaccounts = Me.GridDetail.Rows.Count
            End With

            'start the document
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    tsubacctnum = DirectCast(.GetData(index, 3), String)
                    tsubacctname = DirectCast(.GetData(index, 4), String)
                    acctkey = CInt(.GetData(index, 5))
                    subacctkey = CInt(.GetData(index, 6))
                End With

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the total info box left-side
                        .RenderDirectText(4, y, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(4, y + 4, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side
                        .RenderDirectText(100, y, "Total accounts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(100, y + 4, totalaccounts.ToString, 40, 5, verdanaright8)
                        .RenderDirectText(140, y, "Total subaccounts:", 40, 5, verdanaright8bold)
                        .RenderDirectText(140, y + 4, totalsubaccounts.ToString, 40, 5, verdanaright8)
                        'print the lines under the info box
                        y = 45
                        .RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        .RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        'print the column headers
                        y = 45
                        .RenderDirectText(10, y, "Account", 25, 5, verdanaleft10)
                        .RenderDirectText(35, y, "Sub", 25, 5, verdanaleft10)
                        .RenderDirectText(55, y, "Name", 25, 5, verdanaleft10)
                        .RenderDirectText(95, y, "Name", 25, 5, verdanaleft10)
                        y = 52
                    End If

                    'print the detail
                    If acctkey <> prevacctkey Then
                        If currow > 1 Then y += 5
                        'this is a new account so print the header and shading
                        .RenderDirectRectangle(10, y, 53, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(55, y, 180, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectText(4, y, tacctnum, 20, 5, verdanaright10bold)
                        .RenderDirectText(55, y, tacctname, 122, 5, verdanaleft10)
                        currow += 1
                        y += 5
                    End If
                    .RenderDirectText(35, y, tsubacctnum, 100, 5, verdanaleft10bold)
                    .RenderDirectText(95, y, tsubacctname, 90, 5, verdanaleft10)
                    prevacctkey = acctkey
                    y += 5
                    currow += 1

                    'check for page break & print new column headers if true
                    If .CurrentPage = 1 And y >= 245 Then dopagebreak = True
                    If .CurrentPage > 1 And y >= 249 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break
                        .NewPage()
                        y = 32
                        '.RenderDirectLine(1, y, 189, y, Color.Gray, 0.5)
                        .RenderDirectLine(1, y + 5, 189, y + 5, Color.Gray, 0.5)
                        .RenderDirectText(10, y, "Account", 25, 5, verdanaleft10)
                        .RenderDirectText(35, y, "Sub", 25, 5, verdanaleft10)
                        .RenderDirectText(55, y, "Name", 25, 5, verdanaleft10)
                        .RenderDirectText(95, y, "Name", 25, 5, verdanaleft10)
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

    Private Sub PrintExpenditureSummaryOfEncumbrances(ByVal efiscalyear As Int32)
        Me.DocumentName = "SummaryOfAccounts"
        Me.ReportName = "Encumbrance Account Report"
        'define styles;
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
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim currec, index As Int32
        Dim taccount, taccountname, tsubaccount, tsubaccountname, hldaccount As String
        Dim tprevaccount As String
        Dim x, y As Int32
        Dim hdravailable, hdrcurrent, hdrencumbered, hdrpaid, hdrunpaid, hldunpaid As Decimal
        Dim subavailable, subcurrent, subencumbered, subpaid, subunpaid As Decimal
        Dim totalcurbalance, totalencumbered, totalexpenditure, totalunpaid, totalavailable As Decimal
        Dim dopagebreak As Boolean
        Dim HdrVoidamount, SubVoidamount As Decimal

        Try
            'start the document
            Me.Doc1.StartDoc()
            '
            currec = 1
            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid;
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    taccount = DirectCast(.GetData(index, 1), String)
                    taccountname = DirectCast(.GetData(index, 2), String)
                    tsubaccount = DirectCast(.GetData(index, 3), String)
                    tsubaccountname = DirectCast(.GetData(index, 4), String)
                    hdrcurrent = CType(.GetData(index, 5), Decimal)
                    subcurrent = CType(.GetData(index, 6), Decimal)
                    hdrencumbered = CType(.GetData(index, 7), Decimal)
                    subencumbered = CType(.GetData(index, 8), Decimal)
                    hdrpaid = CType(.GetData(index, 9), Decimal)
                    subpaid = CType(.GetData(index, 10), Decimal)
                    hdrunpaid = CType(.GetData(index, 11), Decimal)
                    subunpaid = CType(.GetData(index, 12), Decimal)
                    hdravailable = CType(.GetData(index, 13), Decimal)
                    subavailable = CType(.GetData(index, 14), Decimal)
                    HdrVoidamount = CType(.GetData(index, 15), Decimal)
                    SubVoidamount = CType(.GetData(index, 16), Decimal)
                End With

                If hldaccount = Nothing Then
                    hldaccount = taccount
                Else
                    If taccount > hldaccount Then
                        hldaccount = taccount
                        hldunpaid = 0
                    End If
                End If
                'the voids
                If taccount = hldaccount Then
                    hldunpaid = hldunpaid + hdrunpaid + HdrVoidamount
                    subunpaid += SubVoidamount
                    hdravailable -= HdrVoidamount
                    subavailable -= SubVoidamount
                Else
                    hldunpaid = 0
                    hldunpaid = hdrunpaid + HdrVoidamount
                    subunpaid += SubVoidamount
                    hdravailable -= HdrVoidamount
                    subavailable -= SubVoidamount
                End If


                'sum the lines for ending totals on last page;
                totalcurbalance += subcurrent
                totalencumbered += subencumbered
                totalexpenditure += subpaid
                totalunpaid += subunpaid
                totalavailable += subavailable

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the bank info box left-side;
                        .RenderDirectText(1, 36, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 40, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'draw a line under the sums;
                        '.RenderDirectLine(162, 44.5, 189, 44.5, Color.Black, 0.75)
                        '.RenderDirectText(10, 50, "** Expenditures Does Include Voids **", 80, 5, verdanaleft8)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Cur.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditure", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Unpaid", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Available", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If taccount <> tprevaccount Then
                        If currec <> 1 Then y += 5
                        'print the header account information;
                        .RenderDirectRectangle(0, y, 9, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(10, y, 63, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectText(0, y, taccount, 10, 5, arialleft8)
                        .RenderDirectText(10, y, taccountname, 60, 10, arialleft8)
                        .RenderDirectText(65, y, hdrcurrent.ToString.Format("{0:F2}", hdrcurrent), 25, 5, verdanaright8)
                        .RenderDirectText(90, y, hdrencumbered.ToString.Format("{0:F2}", hdrencumbered), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, hdrpaid.ToString.Format("{0:F2}", hdrpaid), 25, 5, verdanaright8)
                        '.RenderDirectText(140, y, hdrunpaid.ToString.Format("{0:F2}", hdrunpaid), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, hldunpaid.ToString.Format("{0:F2}", hldunpaid), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, hdravailable.ToString.Format("{0:F2}", hdravailable), 25, 5, verdanaright8)
                        tprevaccount = taccount
                        y += 5
                    End If

                    'print the sub account information;
                    .RenderDirectText(2, y, tsubaccount, 8, 5, arialleft8)
                    .RenderDirectText(12, y, tsubaccountname, 58, 10, arialleft8)
                    .RenderDirectText(65, y, subcurrent.ToString.Format("{0:F2}", subcurrent), 25, 5, verdanaright8)
                    .RenderDirectText(90, y, subencumbered.ToString.Format("{0:F2}", subencumbered), 25, 5, verdanaright8)
                    .RenderDirectText(115, y, subpaid.ToString.Format("{0:F2}", subpaid), 25, 5, verdanaright8)
                    .RenderDirectText(140, y, subunpaid.ToString.Format("{0:F2}", subunpaid), 25, 5, verdanaright8)
                    .RenderDirectText(165, y, subavailable.ToString.Format("{0:F2}", subavailable), 25, 5, verdanaright8)
                    currec += 1
                    y += 5  'crlf;
                    'check for page break & print new column headers if true;
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 33
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Cur.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Encumbered", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Expenditure", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Unpaid", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Available", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
                    End If
                End With
                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            Next

            'print the totals
            With Me.Doc1
                y += 8
                If y > 249 Then .NewPage() : y = 58
                'column headers for total;
                .RenderDirectText(65, y, "Cur.Balance", 25, 5, verdanaright8bold)
                .RenderDirectText(90, y, "Encumbered", 25, 5, verdanaright8bold)
                .RenderDirectText(115, y, "Expenditure", 25, 5, verdanaright8bold)
                .RenderDirectText(140, y, "Unpaid", 25, 5, verdanaright8bold)
                .RenderDirectText(165, y, "Available", 25, 5, verdanaright8bold)
                y += 6
                'highlight the totals;
                .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                'print the text;
                .RenderDirectText(10, y, "TOTALS:", 60, 10, verdanaleft8bold)
                .RenderDirectText(65, y, totalcurbalance.ToString.Format("{0:F2}", totalcurbalance), 25, 5, verdanaright8)
                .RenderDirectText(90, y, totalencumbered.ToString.Format("{0:F2}", totalencumbered), 25, 5, verdanaright8)
                .RenderDirectText(115, y, totalexpenditure.ToString.Format("{0:F2}", totalexpenditure), 25, 5, verdanaright8)
                .RenderDirectText(140, y, totalunpaid.ToString.Format("{0:F2}", totalunpaid), 25, 5, verdanaright8)
                .RenderDirectText(165, y, totalavailable.ToString.Format("{0:F2}", totalavailable), 25, 5, verdanaright8)
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

    'MTD SUMMARY (ALL OR RANGE);
    Private Sub PrintSummaryOfAccounts(ByVal efiscalyear As Int32, ByVal esuppresszero As Boolean, ByVal euserange As Boolean)
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''
        '    0        1         2         3         4         5        6
        '   bank   acctnum   acctname   begbal   mtdrcpt   mtdexp   mtdadj
        '    7
        '  total
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "SummaryOfAccounts"
        Me.ReportName = "Summary Of Accounts"

        'define styles;
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
        'define the styles;
        DefineStyles()
        'define the document;
        DefineDocumentSettings(Me.DocumentName)

        Dim index As Int32
        Dim tacctnum, tacctname As String
        Dim x, y, totalaccounts, totalsubaccounts As Int32
        Dim begbalance, mtdrcpt, mtdexpend, mtdadj, newbalance As Double
        Dim sumbegbalance, sumrcpt, sumexpend, sumadj, tempbal As Double
        Dim totbegbalance, totmtdrcpt, totmtdexpend, totmtdadj As Double
        Dim dopagebreak As Boolean
        Dim suppresshdr As Boolean

        Try
            With Me.GridTotals
                'get the sum of the mtd totals;
                sumbegbalance = CDbl(.GetData(0, 0))
                sumrcpt = CDbl(.GetData(0, 1))
                sumexpend = CDbl(.GetData(0, 2))
                sumadj = CDbl(.GetData(0, 3))
            End With

            'start the document;
            Me.Doc1.StartDoc()

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid;
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    begbalance = CDbl(.GetData(index, 3))
                    mtdrcpt = CDbl(.GetData(index, 4))
                    mtdexpend = CDbl(.GetData(index, 5))
                    mtdadj = CDbl(.GetData(index, 6))
                    newbalance = CDbl(.GetData(index, 7))
                End With

                'is zero suppression turned on;
                If esuppresszero = True Then
                    'test header for zero suppression;
                    If (begbalance = 0D) And (mtdrcpt = 0D) And (mtdexpend = 0D) _
                        And (mtdadj = 0D) And (newbalance = 0D) Then
                        'the header record is zeroes so suppress;
                        suppresshdr = True
                    Else
                        suppresshdr = False
                    End If
                End If

                'sum the lines for ending totals on last page;
                totbegbalance += begbalance
                totmtdrcpt += mtdrcpt
                totmtdexpend += mtdexpend
                totmtdadj += mtdadj

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the signature rectangle box & contents;
                        .RenderDirectRectangle(1, 32, 110, 54, Color.Gray, 0.5, Color.WhiteSmoke)
                        .RenderDirectText(50, 32, "This Report Is True And Correct", 70, 5, verdanaleft8bold)
                        .RenderDirectText(50, 36, "To The Best Of My Knowledge.", 70, 5, verdanaleft8bold)
                        .RenderDirectText(2, 48, "Date: ____/____/_____", 50, 5, verdanaleft8bold)
                        .RenderDirectLine(50, 51.5, 108, 51.5, Color.Black, 0.75)
                        'print the total info box left-side;
                        .RenderDirectText(1, 32, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 36, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print the info box right-side if All Accounts;
                        If euserange = False Then
                            .RenderDirectText(118, 32, "Beginning balance:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 36, "Receipts:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 40, "Checks:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 44, "Adjustments:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 50, "Ending balance:", 40, 5, verdanaright8bold)
                            'print the money fields;
                            .RenderDirectText(150, 32, sumbegbalance.ToString.Format("{0:F2}", sumbegbalance), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 36, sumrcpt.ToString.Format("{0:F2}", sumrcpt), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 40, sumexpend.ToString.Format("{0:F2}", sumexpend), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 44, sumadj.ToString.Format("{0:F2}", sumadj), 40, 5, verdanaright8bold)
                            'draw a line under the sums;
                            .RenderDirectLine(162, 48.5, 189, 48.5, Color.Black, 0.75)
                            'calc the ending balance;
                            tempbal = sumbegbalance + sumrcpt + sumadj - sumexpend
                            .RenderDirectText(150, 50, tempbal.ToString.Format("{0:C2}", tempbal), 40, 5, verdanaright8bold)
                        End If

                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If suppresshdr = False Then
                        'print the header record;
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                        .RenderDirectText(65, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, verdanaright8)
                        .RenderDirectText(90, y, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, mtdadj.ToString.Format("{0:F2}", mtdadj), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, newbalance.ToString.Format("{0:F2}", newbalance), 25, 5, verdanaright8)
                        y += 8
                    End If

                    'check for page break & print new column headers if true;
                    If y > 249 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 34
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 41
                        dopagebreak = False
                    End If
                End With
                'expose the current record & count to the caller;
                'EventRecordProcessed((reccurrent), reccount)
            Next

            'print the totals;
            With Me.Doc1
                y += 8
                If y > 249 Then .NewPage() : y = 58
                'highlight the totals;
                .RenderDirectRectangle(10, y, 63, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                'print the totals;
                .RenderDirectText(10, y, "TOTALS:", 60, 10, verdanaleft8bold)
                .RenderDirectText(65, y, totbegbalance.ToString.Format("{0:F2}", totbegbalance), 25, 5, verdanaright8)
                .RenderDirectText(90, y, totmtdrcpt.ToString.Format("{0:F2}", totmtdrcpt), 25, 5, verdanaright8)
                .RenderDirectText(115, y, totmtdexpend.ToString.Format("{0:F2}", totmtdexpend), 25, 5, verdanaright8)
                .RenderDirectText(140, y, totmtdadj.ToString.Format("{0:F2}", totmtdadj), 25, 5, verdanaright8)
                newbalance = totbegbalance + totmtdrcpt + totmtdadj - totmtdexpend
                .RenderDirectText(165, y, newbalance.ToString.Format("{0:C2}", newbalance), 25, 5, verdanaright8)
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

    'MTD SUMMARY ENCUMBRANCE (ALL OR RANGE);
    Private Sub PrintSummaryOfAccountsWithEncumbrance()
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname  begyrbal  ytdrcpt   ytdencum   encbalance
        '    7         8         9         10        11         12        13
        ' ytdexp    ytdadj    current    voided
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "SummaryOfAccounts"
        Me.ReportName = "Summary Of Accounts"

        'define styles;
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
        'define the styles;
        DefineStyles()
        'define the document;
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currec As Int32

        Try
            Dim tacctnum, tacctname, tsubnum, tsubname As String
            Dim prevbank, prevacctnum, tbankaccount As String
            Dim x, y, totalaccounts, totalsubaccounts, efiscalyear As Int32
            Dim begbalance, mtdrcpt, mtdencumber, mtdexpend, mtdadj, mtdvoids, projbalance As Decimal
            Dim sumbegbalance, sumrcpt, sumencumber, sumexpend, sumadj, tempbal As Decimal
            Dim totbegbalance, totmtdrcpt, totmtdencumber, totmtdexpend, totmtdadj, totalvoids As Decimal
            Dim dopagebreak As Boolean

            With Me.GridTotals
                'get the sum of the mtd totals
                sumbegbalance = CDec(.GetData(0, 0))
                sumrcpt = CDec(.GetData(0, 1))
                sumexpend = CDec(.GetData(0, 2))
                sumadj = CDec(.GetData(0, 3))
            End With

            'start the document
            Me.Doc1.StartDoc()
            currec = 1

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid;
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    'hdr vals;
                    begbalance = CDec(.GetData(index, 3))
                    mtdrcpt = CDec(.GetData(index, 4))
                    mtdencumber = CDec(.GetData(index, 5))   'outstanding encumbrance balance for account;
                    mtdexpend = CDec(.GetData(index, 7))
                    mtdadj = CDec(.GetData(index, 8))
                    mtdvoids = CDec(.GetData(index, 10))

                    'calc the projected balance;
                    'projbalance = begbalance + mtdrcpt - (mtdencumber - mtdexpend) - mtdexpend + mtdadj
                    projbalance = begbalance + mtdrcpt - (mtdencumber - mtdexpend) - mtdexpend + mtdadj - mtdvoids

                End With

                'sum the lines for ending totals on last page;
                totbegbalance += begbalance
                totmtdrcpt += mtdrcpt
                totmtdencumber += mtdencumber - mtdexpend
                totmtdexpend += mtdexpend
                totmtdadj += mtdadj
                totalvoids += mtdvoids

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the signature rectangle box & contents;
                        .RenderDirectRectangle(1, 32, 110, 54, Color.Gray, 0.5, Color.WhiteSmoke)
                        .RenderDirectText(50, 32, "This Report Is True And Correct", 70, 5, verdanaleft8bold)
                        .RenderDirectText(50, 36, "To The Best Of My Knowledge.", 70, 5, verdanaleft8bold)
                        .RenderDirectText(2, 48, "Date: ____/____/_____", 50, 5, verdanaleft8bold)
                        .RenderDirectLine(50, 51.5, 108, 51.5, Color.Black, 0.75)
                        'print the total info box left-side;
                        .RenderDirectText(1, 32, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 36, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(85, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(105, y, "Encum.", 25, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Projected", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        If currec <> 1 Then y += 5
                        'print the header account information;
                        '.RenderDirectRectangle(0, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                        .RenderDirectText(60, y, begbalance.ToString.Format("{0:F2}", begbalance), 25, 5, verdanaright8)
                        .RenderDirectText(85, y, mtdrcpt.ToString.Format("{0:F2}", mtdrcpt), 25, 5, verdanaright8)
                        .RenderDirectText(105, y, mtdencumber.ToString.Format("{0:F2}", mtdencumber), 25, 5, verdanaright8)
                        .RenderDirectText(125, y, mtdexpend.ToString.Format("{0:F2}", mtdexpend), 25, 5, verdanaright8)
                        .RenderDirectText(145, y, mtdadj.ToString.Format("{0:F2}", mtdadj), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, projbalance.ToString.Format("{0:F2}", projbalance), 25, 5, verdanaright8)
                        prevacctnum = tacctnum
                        y += 5
                    End If

                    currec += 1

                    'check for page break & print new column headers if true;
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 33
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(85, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(105, y, "Encum.", 25, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Projected", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
                    End If
                End With
                'expose the current record & count to the caller;
                'EventRecordProcessed((reccurrent), reccount)
            Next

            'print the totals
            With Me.Doc1
                y += 8
                If y > 225 Then .NewPage() : y = 40
                'highlight the totals
                .RenderDirectRectangle(10, y, 190, y + 30, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                'print the text
                .RenderDirectText(10, y, "TOTALS:", 60, 10, verdanaleft8bold)
                .RenderDirectText(0, y, "Beginning balance:", 100, 5, verdanaright8)
                .RenderDirectText(165, y, totbegbalance.ToString.Format("{0:F2}", totbegbalance), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 5, "Add receipts:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 5, totmtdrcpt.ToString.Format("{0:F2}", totmtdrcpt), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 10, "Less outstanding encumbrance:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 10, totmtdencumber.ToString.Format("{0:F2}", totmtdencumber), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 15, "Less checks:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 15, totmtdexpend.ToString.Format("{0:F2}", totmtdexpend), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 20, "Add total voids:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 20, totalvoids.ToString.Format("{0:F2}", totalvoids), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 25, "Add adjustments:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 25, totmtdadj.ToString.Format("{0:F2}", totmtdadj), 25, 5, verdanaright8)
                'calc the total projected balance;
                projbalance = totbegbalance + totmtdrcpt + totmtdadj - totmtdexpend - totmtdencumber - totalvoids
                .RenderDirectText(0, y + 30, "Projected balance:", 100, 5, verdanaright8bold)
                .RenderDirectText(165, y + 30, projbalance.ToString.Format("{0:C2}", projbalance), 25, 5, verdanaright8bold)
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

    'MTD/YTD SUMMARY SUBACCOUNTS (ALL OR RANGE);
    Private Sub PrintSummaryOfSubaccounts(ByVal efiscalyear As Int32, ByVal esuppresszero As Boolean, ByVal euserange As Boolean)
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
        '   14
        ' astotal
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "SummaryOfAccounts"
        Me.ReportName = "Summary Of Accounts"

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
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim index, currec As Int32
        Dim suppresshdr, suppressdetl As Boolean

        Try
            Dim tacctnum, tacctname, tsubnum, tsubname As String
            Dim prevbank, prevacctnum As String
            Dim x, y, totalaccounts, totalsubaccounts As Int32
            Dim hdrbegbalance, hdrmtdrcpt, hdrmtdexpend, hdrmtdadj, hdrnewbalance As Decimal
            Dim subbegbalance, submtdrcpt, submtdexpend, submtdadj, subnewbalance As Decimal
            Dim sumbegbalance, sumrcpt, sumexpend, sumadj, tempbal As Decimal
            Dim totbegbalance, totmtdrcpt, totmtdexpend, totmtdadj As Decimal
            Dim dopagebreak As Boolean

            With Me.GridTotals
                'get the sum of the mtd totals
                sumbegbalance = CDec(.GetData(0, 0))
                sumrcpt = CDec(.GetData(0, 1))
                sumexpend = CDec(.GetData(0, 2))
                sumadj = CDec(.GetData(0, 3))
            End With

            'start the document
            Me.Doc1.StartDoc()
            ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
            '    0         1         2          3         4          5         6
            '  bank     acctnum   acctname   subnum    subname   afbegbal  afmtdrcpt
            '    7         8         9         10        11         12        13
            ' afmtdexp  afmtdadj  aftotal   asbegbal  asmtdrcpt  asmtdexp  asmtdadj
            '   14
            ' astotal
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            currec = 1
            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    tsubnum = DirectCast(.GetData(index, 3), String)
                    tsubname = DirectCast(.GetData(index, 4), String)
                    'hdr vals;
                    hdrbegbalance = CDec(.GetData(index, 5))
                    hdrmtdrcpt = CDec(.GetData(index, 6))
                    hdrmtdexpend = CDec(.GetData(index, 7))
                    hdrmtdadj = CDec(.GetData(index, 8))
                    hdrnewbalance = CDec(.GetData(index, 9))

                    'sub vals;
                    subbegbalance = CDec(.GetData(index, 10))
                    submtdrcpt = CDec(.GetData(index, 11))
                    submtdexpend = CDec(.GetData(index, 12))
                    submtdadj = CDec(.GetData(index, 13))
                    subnewbalance = CDec(.GetData(index, 14))

                    'is zero suppression turned on;
                    If esuppresszero = True Then
                        'test header for zero suppression;
                        If (hdrbegbalance = 0D) And (hdrmtdrcpt = 0D) And (hdrmtdexpend = 0D) _
                            And (hdrmtdadj = 0D) And (hdrnewbalance = 0D) Then
                            'the header record is zeroes so suppress;
                            suppresshdr = True
                        Else
                            suppresshdr = False
                        End If
                        'test detail for zero suppression;
                        If (subbegbalance = 0D) And (submtdrcpt = 0D) And (submtdexpend = 0D) _
                            And (submtdadj = 0D) And (subnewbalance = 0D) Then
                            'the header record is zeroes so suppress;
                            suppressdetl = True
                        Else
                            suppressdetl = False
                        End If
                    End If
                End With

                'sum the lines for ending totals on last page;
                totbegbalance += subbegbalance
                totmtdrcpt += submtdrcpt
                totmtdexpend += submtdexpend
                totmtdadj += submtdadj

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the signature rectangle box & contents;
                        .RenderDirectRectangle(1, 32, 110, 54, Color.Gray, 0.5, Color.WhiteSmoke)
                        .RenderDirectText(50, 32, "This Report Is True And Correct", 70, 5, verdanaleft8bold)
                        .RenderDirectText(50, 36, "To The Best Of My Knowledge.", 70, 5, verdanaleft8bold)
                        .RenderDirectText(2, 48, "Date: ____/____/_____", 50, 5, verdanaleft8bold)
                        .RenderDirectLine(50, 51.5, 108, 51.5, Color.Black, 0.75)
                        'print the total info box left-side;
                        .RenderDirectText(1, 32, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 36, Me.BankAccountNumber, 40, 5, verdanaright8)
                        If euserange = False Then
                            'print the info box right-side;
                            .RenderDirectText(118, 32, "Beginning balance:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 36, "Receipts:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 40, "Checks:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 44, "Adjustments:", 40, 5, verdanaright8bold)
                            .RenderDirectText(118, 50, "Ending balance:", 40, 5, verdanaright8bold)
                            'print the money fields;
                            .RenderDirectText(150, 32, sumbegbalance.ToString.Format("{0:F2}", sumbegbalance), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 36, sumrcpt.ToString.Format("{0:F2}", sumrcpt), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 40, sumexpend.ToString.Format("{0:F2}", sumexpend), 40, 5, verdanaright8bold)
                            .RenderDirectText(150, 44, sumadj.ToString.Format("{0:F2}", sumadj), 40, 5, verdanaright8bold)
                            'draw a line under the sums;
                            .RenderDirectLine(162, 48.5, 189, 48.5, Color.Black, 0.75)
                            'calc the ending balance;
                            tempbal = sumbegbalance + sumrcpt + sumadj - sumexpend
                            .RenderDirectText(150, 50, tempbal.ToString.Format("{0:C2}", tempbal), 40, 5, verdanaright8bold)
                        End If
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'zero supression check;
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If suppresshdr = False Then
                            If currec <> 1 Then y += 5
                            'print the header account information;
                            .RenderDirectRectangle(0, y, 9, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(10, y, 63, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                            .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                            .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                            .RenderDirectText(65, y, hdrbegbalance.ToString.Format("{0:F2}", hdrbegbalance), 25, 5, verdanaright8)
                            .RenderDirectText(90, y, hdrmtdrcpt.ToString.Format("{0:F2}", hdrmtdrcpt), 25, 5, verdanaright8)
                            .RenderDirectText(115, y, hdrmtdexpend.ToString.Format("{0:F2}", hdrmtdexpend), 25, 5, verdanaright8)
                            .RenderDirectText(140, y, hdrmtdadj.ToString.Format("{0:F2}", hdrmtdadj), 25, 5, verdanaright8)
                            .RenderDirectText(165, y, hdrnewbalance.ToString.Format("{0:F2}", hdrnewbalance), 25, 5, verdanaright8)
                            y += 5
                        End If
                        'store the previous account number;
                        prevacctnum = tacctnum
                    End If

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'zero suppression check;
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If suppresshdr = False And suppressdetl = False Then
                        'print the sub account information;
                        .RenderDirectText(2, y, tsubnum, 8, 5, arialleft8)
                        .RenderDirectText(12, y, tsubname, 58, 10, arialleft8)
                        .RenderDirectText(65, y, subbegbalance.ToString.Format("{0:F2}", subbegbalance), 25, 5, verdanaright8)
                        .RenderDirectText(90, y, submtdrcpt.ToString.Format("{0:F2}", submtdrcpt), 25, 5, verdanaright8)
                        .RenderDirectText(115, y, submtdexpend.ToString.Format("{0:F2}", submtdexpend), 25, 5, verdanaright8)
                        .RenderDirectText(140, y, submtdadj.ToString.Format("{0:F2}", submtdadj), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, subnewbalance.ToString.Format("{0:F2}", subnewbalance), 25, 5, verdanaright8)
                        currec += 1
                        y += 5
                    End If


                    'check for page break & print new column headers if true;
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 33
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(90, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(115, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(140, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Ending", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
                    End If
                End With
                'expose the current record & count to the caller
                'EventRecordProcessed((reccurrent), reccount)
            Next

            'print the totals
            With Me.Doc1
                y += 8
                If y > 249 Then .NewPage() : y = 58
                'highlight the totals
                .RenderDirectRectangle(10, y, 63, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(65, y, 90, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(92, y, 115, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(117, y, 140, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(142, y, 165, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                .RenderDirectRectangle(167, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                'print the text
                .RenderDirectText(10, y, "TOTALS:", 60, 10, verdanaleft8bold)
                .RenderDirectText(65, y, totbegbalance.ToString.Format("{0:F2}", totbegbalance), 25, 5, verdanaright8)
                .RenderDirectText(90, y, totmtdrcpt.ToString.Format("{0:F2}", totmtdrcpt), 25, 5, verdanaright8)
                .RenderDirectText(115, y, totmtdexpend.ToString.Format("{0:F2}", totmtdexpend), 25, 5, verdanaright8)
                .RenderDirectText(140, y, totmtdadj.ToString.Format("{0:F2}", totmtdadj), 25, 5, verdanaright8)
                subnewbalance = totbegbalance + totmtdrcpt + totmtdadj - totmtdexpend
                .RenderDirectText(165, y, subnewbalance.ToString.Format("{0:C2}", subnewbalance), 25, 5, verdanaright8)
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

    'MTD/YTD SUMMARY SUBACCOUNTS ENCUMBRANCE (ALL OR RANGE);
    Private Sub PrintSummaryOfSubaccountsWithEncumbrance(ByVal efiscalyear As Int32)
        ''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''''''''''''''''
        '    0         1         2          3         4          5         6
        '  bank     acctnum   acctname   subnum    subname   begmobal  afmtdrcpt
        '    7         8         9         10        11         12        13
        ' afytdenc ytdencbal  afmtdexp  afmtdadj  afcurrent  begmobal  asmtdrcpt
        '   14        15        16         17        18
        ' asytdenc ytdencbal  asmtdexp  asmtdadj  ascurrent
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "SummaryOfAccounts"
        Me.ReportName = "Summary Of Accounts"

        'define styles;
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
        'define the styles;
        DefineStyles()
        'define the document;
        DefineDocumentSettings(Me.DocumentName)

        Dim index, currec As Int32

        Try
            Dim tacctnum, tacctname, tsubnum, tsubname As String
            Dim prevbank, prevacctnum As String
            Dim x, y, totalaccounts, totalsubaccounts As Int32
            Dim hdrbegbalance, hdrrcpt, hdrmtdencumber, hdrmtdexpend, hdrmtdadj, hdrprojbalance, hdrvoids As Double
            Dim subbegbalance, subrcpt, submtdencumber, submtdexpend, submtdadj, subprojbalance As Double
            Dim sumbegbalance, sumrcpt, sumencumber, sumexpend, sumadj, tempbal, subvoids As Double
            Dim totbegbalance, totmtdrcpt, totmtdencumber, totmtdexpend, totmtdadj, totalvoids As Double
            Dim dopagebreak As Boolean

            With Me.GridTotals
                'get the sum of the mtd totals
                sumbegbalance = CDbl(.GetData(0, 0))
                sumrcpt = CDbl(.GetData(0, 1))
                sumexpend = CDbl(.GetData(0, 2))
                sumadj = CDbl(.GetData(0, 3))
            End With

            'start the document
            Me.Doc1.StartDoc()
            currec = 1

            For index = 0 To Me.GridDetail.Rows.Count - 1
                'collect the information from the grid;
                With Me.GridDetail
                    Me.BankAccountNumber = DirectCast(.GetData(index, 0), String)
                    tacctnum = DirectCast(.GetData(index, 1), String)
                    tacctname = DirectCast(.GetData(index, 2), String)
                    tsubnum = DirectCast(.GetData(index, 3), String)
                    tsubname = DirectCast(.GetData(index, 4), String)
                    'hdr vals
                    hdrbegbalance = CDbl(.GetData(index, 5))
                    hdrrcpt = CDbl(.GetData(index, 6))
                    hdrmtdencumber = CDbl(.GetData(index, 7))   'encumbrance;
                    hdrmtdexpend = CDbl(.GetData(index, 9))
                    hdrmtdadj = CDbl(.GetData(index, 10))
                    hdrvoids = CDbl(.GetData(index, 19))
                    hdrprojbalance = hdrbegbalance + hdrrcpt - (hdrmtdencumber - hdrmtdexpend) - hdrmtdexpend + hdrmtdadj - hdrvoids
                    'sub vals
                    subbegbalance = CDbl(.GetData(index, 12))
                    subrcpt = CDbl(.GetData(index, 13))
                    submtdencumber = CDbl(.GetData(index, 14))  'outstanding encumbrance balance for subaccount;
                    'submtdencumber = CDbl(.GetData(index, 15))  'outstanding encumbrance balance for subaccount;
                    submtdexpend = CDbl(.GetData(index, 16))
                    submtdadj = CDbl(.GetData(index, 17))
                    subvoids = CDbl(.GetData(index, 19))
                    subprojbalance = subbegbalance + subrcpt - (submtdencumber - submtdexpend) - submtdexpend + submtdadj - subvoids

                    'calc the projected balance;
                    'projbalance = begbalance + mtdrcpt - (mtdencumber - mtdexpend) - mtdexpend + mtdadj
                    'projbalance = hdrbegbalance + mtdrcpt - (mtdencumber - mtdexpend) - mtdexpend + mtdadj - subvoids

                End With

                'sum the lines for ending totals on last page;
                totbegbalance += subbegbalance
                totmtdrcpt += subrcpt
                totmtdencumber += submtdencumber - submtdexpend
                totmtdexpend += submtdexpend
                totmtdadj += submtdadj
                totalvoids += subvoids

                With Me.Doc1
                    If index = 0 Then
                        x = 19
                        y = 33
                        'print the signature rectangle box & contents;
                        .RenderDirectRectangle(1, 32, 110, 54, Color.Gray, 0.5, Color.WhiteSmoke)
                        .RenderDirectText(50, 32, "This Report Is True And Correct", 70, 5, verdanaleft8bold)
                        .RenderDirectText(50, 36, "To The Best Of My Knowledge.", 70, 5, verdanaleft8bold)
                        .RenderDirectText(2, 48, "Date: ____/____/_____", 50, 5, verdanaleft8bold)
                        .RenderDirectLine(50, 51.5, 108, 51.5, Color.Black, 0.75)
                        'print the total info box left-side;
                        .RenderDirectText(1, 32, "For Bank Account:", 40, 5, verdanaright8bold)
                        .RenderDirectText(1, 36, Me.BankAccountNumber, 40, 5, verdanaright8)
                        'print line above the column headers;
                        .RenderDirectLine(0, 55, 190, 55, Color.Gray, 0.5)
                        y = 58
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(85, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(105, y, "Encum.", 25, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Projected", 25, 5, verdanaright8bold)
                        y = 65
                    End If

                    If tacctnum <> prevacctnum Then
                        If currec <> 1 Then y += 5
                        'print the header account information;
                        .RenderDirectRectangle(0, y, 190, y + 4.5, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                        .RenderDirectText(0, y, tacctnum, 10, 5, arialleft8)
                        .RenderDirectText(10, y, tacctname, 60, 10, arialleft8)
                        .RenderDirectText(60, y, hdrbegbalance.ToString.Format("{0:F2}", hdrbegbalance), 25, 5, verdanaright8)
                        .RenderDirectText(85, y, hdrrcpt.ToString.Format("{0:F2}", hdrrcpt), 25, 5, verdanaright8)
                        .RenderDirectText(105, y, hdrmtdencumber.ToString.Format("{0:F2}", hdrmtdencumber), 25, 5, verdanaright8)
                        .RenderDirectText(125, y, hdrmtdexpend.ToString.Format("{0:F2}", hdrmtdexpend), 25, 5, verdanaright8)
                        .RenderDirectText(145, y, hdrmtdadj.ToString.Format("{0:F2}", hdrmtdadj), 25, 5, verdanaright8)
                        .RenderDirectText(165, y, hdrprojbalance.ToString.Format("{0:F2}", hdrprojbalance), 25, 5, verdanaright8)
                        prevacctnum = tacctnum
                        y += 5
                    End If

                    'print the sub account information;
                    .RenderDirectText(2, y, tsubnum, 8, 5, arialleft8)
                    .RenderDirectText(12, y, tsubname, 58, 10, arialleft8)
                    .RenderDirectText(60, y, subbegbalance.ToString.Format("{0:F2}", subbegbalance), 25, 5, verdanaright8)
                    .RenderDirectText(85, y, subrcpt.ToString.Format("{0:F2}", subrcpt), 25, 5, verdanaright8)
                    .RenderDirectText(105, y, submtdencumber.ToString.Format("{0:F2}", submtdencumber), 25, 5, verdanaright8)
                    .RenderDirectText(125, y, submtdexpend.ToString.Format("{0:F2}", submtdexpend), 25, 5, verdanaright8)
                    .RenderDirectText(145, y, submtdadj.ToString.Format("{0:F2}", submtdadj), 25, 5, verdanaright8)
                    .RenderDirectText(165, y, subprojbalance.ToString.Format("{0:F2}", subprojbalance), 25, 5, verdanaright8)

                    currec += 1
                    y += 5  'crlf;

                    'check for page break & print new column headers if true;
                    If y >= 250 Then dopagebreak = True

                    If dopagebreak Then
                        'don't page break if it's the last record;
                        If index = Me.GridDetail.Rows.Count - 1 Then Exit For
                        'do the page break;
                        .NewPage()
                        y = 33
                        'print the column headers;
                        .RenderDirectText(0, y, "Acct.", 10, 5, verdanaleft8bold)
                        .RenderDirectText(10, y, "Name", 50, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Beg.Balance", 25, 5, verdanaright8bold)
                        .RenderDirectText(85, y, "Receipts", 25, 5, verdanaright8bold)
                        .RenderDirectText(105, y, "Encum.", 25, 5, verdanaright8bold)
                        .RenderDirectText(125, y, "Checks", 25, 5, verdanaright8bold)
                        .RenderDirectText(145, y, "Adjust.", 25, 5, verdanaright8bold)
                        .RenderDirectText(165, y, "Projected", 25, 5, verdanaright8bold)
                        y = 40
                        dopagebreak = False
                        currec = 1
                    End If
                End With
                'expose the current record & count to the caller;
                'EventRecordProcessed((reccurrent), reccount)
            Next

            'print the totals
            With Me.Doc1
                y += 8
                If y > 225 Then .NewPage() : y = 40
                'highlight the totals
                .RenderDirectRectangle(10, y, 190, y + 30, Color.WhiteSmoke, 0.25, Color.WhiteSmoke)
                'print the text
                .RenderDirectText(10, y, "TOTALS:", 60, 10, verdanaleft8bold)
                .RenderDirectText(0, y, "Beginning balance:", 100, 5, verdanaright8)
                .RenderDirectText(165, y, totbegbalance.ToString.Format("{0:F2}", totbegbalance), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 5, "Add receipts:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 5, totmtdrcpt.ToString.Format("{0:F2}", totmtdrcpt), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 10, "Less outstanding encumbrance:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 10, totmtdencumber.ToString.Format("{0:F2}", totmtdencumber), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 15, "Less checks:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 15, totmtdexpend.ToString.Format("{0:F2}", totmtdexpend), 25, 5, verdanaright8)
                .RenderDirectText(0, y + 20, "Add adjustments:", 100, 5, verdanaright8)
                .RenderDirectText(165, y + 20, totmtdadj.ToString.Format("{0:F2}", totmtdadj), 25, 5, verdanaright8)
                subprojbalance = totbegbalance + totmtdrcpt + totmtdadj - totmtdexpend - totmtdencumber
                .RenderDirectText(0, y + 25, "Projected balance:", 100, 5, verdanaright8bold)
                .RenderDirectText(165, y + 25, subprojbalance.ToString.Format("{0:C2}", subprojbalance), 25, 5, verdanaright8bold)
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
                .RenderDirectText(65, 25, Me.CellMiddleBottom, 80, 5, verdanaleft8)
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

#End Region

#Region "  Properties "

    Private Property AppliedDate() As Date
        Get
            Return _applieddate
        End Get
        Set(ByVal Value As Date)
            _applieddate = Value
        End Set
    End Property

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

    Private Property CurrentMonthString() As String
        Get
            Return _currentmonthstring
        End Get
        Set(ByVal Value As String)
            _currentmonthstring = Value.Trim
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

    Private Property FiscalYear() As Int32
        Get
            Return _fiscalyear
        End Get
        Set(ByVal Value As Int32)
            _fiscalyear = Value
        End Set
    End Property

    Private Property FiscalYearSelected() As Int32
        Get
            Return _fiscalyearselected
        End Get
        Set(ByVal Value As Int32)
            _fiscalyearselected = Value
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

#End Region

End Class
