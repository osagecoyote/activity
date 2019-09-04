Imports C1.C1PrintDocument
Imports C1.Win.C1FlexGrid
Imports System.Data
Imports System.Data.SqlClient

Public Class frmVendorReports
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
    Friend WithEvents GridDetail As C1.Win.C1FlexGrid.C1FlexGrid
    Friend WithEvents GridTotals As C1.Win.C1FlexGrid.C1FlexGrid
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmVendorReports))
        Me.GridTotals = New C1.Win.C1FlexGrid.C1FlexGrid
        Me.GridDetail = New C1.Win.C1FlexGrid.C1FlexGrid
        Me.Prev1 = New C1.Win.C1PrintPreview.C1PrintPreview
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
        Me.Doc1 = New C1.C1PrintDocument.C1PrintDocument
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GridTotals
        '
        Me.GridTotals.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridTotals.BackColor = System.Drawing.SystemColors.Window
        Me.GridTotals.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridTotals.Location = New System.Drawing.Point(8, 8)
        Me.GridTotals.Name = "GridTotals"
        Me.GridTotals.Rows.Fixed = 0
        Me.GridTotals.Size = New System.Drawing.Size(528, 376)
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
        Me.GridTotals.Visible = False
        '
        'GridDetail
        '
        Me.GridDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridDetail.BackColor = System.Drawing.SystemColors.Window
        Me.GridDetail.ColumnInfo = "10,0,0,0,0,85,Columns:"
        Me.GridDetail.Location = New System.Drawing.Point(8, 8)
        Me.GridDetail.Name = "GridDetail"
        Me.GridDetail.Rows.Fixed = 0
        Me.GridDetail.Size = New System.Drawing.Size(528, 376)
        Me.GridDetail.Styles = New C1.Win.C1FlexGrid.CellStyleCollection("Fixed{BackColor:Control;ForeColor:ControlText;Border:Flat,1,ControlDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "Hi" & _
        "ghlight{BackColor:Highlight;ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Search{BackColor:Highlight" & _
        ";ForeColor:HighlightText;}" & Microsoft.VisualBasic.ChrW(9) & "Frozen{BackColor:Beige;}" & Microsoft.VisualBasic.ChrW(9) & "EmptyArea{BackColor:AppWorks" & _
        "pace;Border:Flat,1,ControlDarkDark,Both;}" & Microsoft.VisualBasic.ChrW(9) & "GrandTotal{BackColor:Black;ForeColor:W" & _
        "hite;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal0{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal1{BackColor" & _
        ":ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal2{BackColor:ControlDarkDark;ForeColor" & _
        ":White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal3{BackColor:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal4{BackCol" & _
        "or:ControlDarkDark;ForeColor:White;}" & Microsoft.VisualBasic.ChrW(9) & "Subtotal5{BackColor:ControlDarkDark;ForeCol" & _
        "or:White;}" & Microsoft.VisualBasic.ChrW(9))
        Me.GridDetail.TabIndex = 5
        Me.GridDetail.Visible = False
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
        Me.Prev1.Size = New System.Drawing.Size(544, 376)
        Me.Prev1.Splitter.Cursor = System.Windows.Forms.Cursors.VSplit
        Me.Prev1.Splitter.Width = 3
        Me.Prev1.StatusBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.StatusBar.TabIndex = 4
        Me.Prev1.TabIndex = 6
        Me.Prev1.ToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.c1pBtnFileOpen1, Me.c1pBtnFileSave1, Me.c1pBtnFilePrint1, Me.c1pBtnPageSetup1, Me.c1pBtnReflow1, Me.c1pBtnStop1, Me.c1pBtnDocInfo1, Me.c1pBtnSeparator1, Me.c1pBtnShowNavigationBar1, Me.c1pBtnSeparator2, Me.c1pBtnMouseHand1, Me.c1pBtnMouseZoom1, Me.c1pBtnMouseZoomOut1, Me.c1pBtnMouseSelect1, Me.c1pBtnFindText1, Me.c1pBtnSeparator3, Me.c1pBtnGoFirst1, Me.c1pBtnGoPrev1, Me.c1pBtnGoNext1, Me.c1pBtnGoLast1, Me.c1pBtnSeparator4, Me.c1pBtnHistoryPrev1, Me.c1pBtnHistoryNext1, Me.c1pBtnSeparator5, Me.c1pBtnZoomOut1, Me.c1pBtnZoomIn1, Me.c1pBtnSeparator6, Me.c1pBtnViewActualSize1, Me.c1pBtnViewFullPage1, Me.c1pBtnViewPageWidth1, Me.c1pBtnViewTwoPages1, Me.c1pBtnViewFourPages1, Me.c1pBtnSeparator7, Me.c1pBtnHelp1})
        Me.Prev1.ToolBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.ToolBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        'frmVendorReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 389)
        Me.Controls.Add(Me.Prev1)
        Me.Controls.Add(Me.GridDetail)
        Me.Controls.Add(Me.GridTotals)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmVendorReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Vendor Reporting"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.GridTotals, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  C1Doc Events "

    Private Sub Doc1_NewPageStarted(ByVal sender As C1.C1PrintDocument.C1PrintDocument, ByVal e As C1.C1PrintDocument.NewPageStartedEventArgs) Handles Doc1.NewPageStarted
        PrintHeader()
    End Sub

#End Region

#Region "  Class Members "

    'styles;
    Private docstyle As C1DocStyle
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
    'header values;
    Private CellMiddleBottom As String = ""
    Private CellMiddleMiddle As String = ""
    Private CellMiddleTop As String = ""
    Private CellRightBottom As String = ""
    Private CellRightMiddle As String = ""
    Private CellRightTop As String = ""
    'property vars;
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
    Private MSGTITLE As String = "Activity Fund Reports"
    Private _reportname As String
    Private _schoolname As String
    Private _schooladdress1 As String
    Private _schooladdress2 As String
    Private _schoolcitystatezip As String
    Private SignatureTextLine1 As String = ""
    Private SignatureTextLine2 As String = ""
    Private Signature1 As Image
    Private Signature2 As Image
    '
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

    Public Function Generate1099Vendors(ByVal ecalendaryear As Int32) As Boolean
        '1099 vendor report;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '    key      ssn        name     addr1    addr2     city       state
        '     7        8          9        10       11        12         13   
        '    zip      ext       total     count 
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim begdate, enddate As Date
        Try
            begdate = CDate("01/01/" & ecalendaryear)
            enddate = begdate.AddYears(1)
            enddate = enddate.AddSeconds(-1)
            begdate = begdate.AddSeconds(1)
        Catch ex As Exception
            Throw
        End Try

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable

        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'this union query no longer needed since the vendors are fully tied by the
            'vendor key and the vendor number isn't needed for ties; 2014.01.24;
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'SSQL = "SELECT v.vend_autoinc_key, vend_ssn, vend_name, vend_addr1, vend_addr2, vend_city, vend_state," _
            ' & " vend_zip, vend_zip_ext, SUM(chks_amount) AS Total, COUNT(*) AS Items" _
            ' & " FROM vend_info AS v, chks_info AS h" _
            ' & " WHERE v.vend_number = h.vend_number" _
            ' & " AND (vend_status = 'O')" _
            ' & " AND (chks_status <> 'V')" _
            ' & " AND v.vend_1099_sw = 'Y'" _
            ' & " AND chks_applied_date BETWEEN @p1 AND @p2" _
            ' & " GROUP BY v.vend_autoinc_key, vend_ssn, vend_name, vend_addr1, vend_addr2," _
            ' & " vend_city, vend_state, vend_zip, vend_zip_ext" _
            ' & " UNION"
            SSQL = "SELECT v.vend_autoinc_key, vend_ssn, vend_name, vend_addr1, vend_addr2, vend_city, vend_state," _
             & " vend_zip, vend_zip_ext, SUM(chks_amount) AS Total, COUNT(*) AS Items" _
             & " FROM vend_info AS v, chks_info AS h" _
             & " WHERE h.vend_autoinc_key = v.vend_autoinc_key" _
             & " AND (vend_status = 'O')" _
             & " AND (chks_status <> 'V')" _
             & " AND v.vend_1099_sw = 'Y'" _
             & " AND chks_applied_date BETWEEN @p1 AND @p2" _
             & " GROUP BY v.vend_autoinc_key, vend_ssn, vend_name, vend_addr1, vend_addr2," _
             & " vend_city, vend_state, vend_zip, vend_zip_ext" _
             & " ORDER BY vend_name, vend_ssn"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", begdate)
            cmd.Parameters.Add("@p2", enddate)
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
            'prompt for no records found;
            If tbl.Rows.Count < 1 Then
                MsgBox("No records found for the selected criteria.", MsgBoxStyle.Information, "Activity Fund Reporting")
                Exit Function
            End If
            'datasource the detail;
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        End Try


        'Me.GridDetail.Visible = True
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellRightMiddle = "For Calendar Year " & ecalendaryear.ToString
            Application.DoEvents()
            'render the table
            Call Print1099Vendors()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateVendorAudit(ByVal efiscalyear As Int32) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'generates a report detailing the vendors that have changed names;
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0         1        2         3          4        5
        '  vendkey   vname    audname   audtime    userkey  username
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim index As Int32
        Dim puserkey, userkey As Int32
        Dim begdate, enddate As Date
        Dim username As String

        Try
            'calc the beginning & ending date for the fiscal year;
            begdate = CDate("07/01/" & efiscalyear - 1)
            enddate = CDate("07/01/" & efiscalyear)
            enddate = enddate.AddSeconds(-1)
        Catch ex As Exception
            Throw
        End Try

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable

        Try
            SSQL = "SELECT a.vend_autoinc_key, v.vend_name, a.vend_name, vaud_datetime, user_autoinc_key, '' AS UserNm" _
            & " FROM vend_audit AS a, vend_info AS v" _
            & " WHERE v.vend_autoinc_key = a.vend_autoinc_key" _
            & " AND vaud_datetime BETWEEN @p1 AND @p2" _
            & " ORDER BY v.vend_name, vaud_datetime"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", begdate)
            cmd.Parameters.Add("@p2", enddate)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("auditvendor")
            da.Fill(tbl)
            'success;
            cn.Close()
            'return if no records returned;
            If tbl.Rows.Count < 1 Then
                MsgBox("No records found for the selected fiscal year.", MsgBoxStyle.Information, "Activity Fund Reporting")
                Exit Function
            End If
            'datasource the detail;
            Me.GridDetail.DataSource = tbl
        Catch ex As Exception
            Throw
        Finally
            If cn.State <> ConnectionState.Closed Then cn.Close()
            da.Dispose()
            cmd.Dispose()
        End Try

        Try
            'update the user field for each entry;
            With Me.GridDetail
                For index = 0 To .Rows.Count - 1
                    userkey = CType(.GetData(index, 4), Int32)
                    'only collect username if its different than prior user;
                    If puserkey <> userkey And (userkey > 0) Then
                        SSQL = "SELECT user_fullname FROM user_info WHERE user_autoinc_key = @p1"
                        cn = New SqlConnection(Me.ConnectionString)
                        cmd = New SqlCommand(SSQL, cn)
                        cmd.Parameters.Add("@p1", userkey)
                        If cn.State <> ConnectionState.Open Then cn.Open()
                        username = CType(cmd.ExecuteScalar, String)
                        cn.Close()
                    End If
                    If userkey < 1 Then username = "ADMINISTRATOR"
                    'store username into the grid;
                    .SetData(index, 5, username.Trim)
                    'store the prior userkey;
                    puserkey = userkey
                Next
            End With
        Catch ex As Exception
            Throw
        End Try

        Try
            Me.CellRightMiddle = "FY-" & efiscalyear.ToString
            Application.DoEvents()
            'render the table;
            PrintAuditListing()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GenerateVendorExpenditureSummary(ByVal erawyear As Int32, ByVal eprint1099only As Boolean, ByVal eincludezerobalances As Boolean, ByVal euseminimum600 As Boolean, ByVal eusefiscal As Boolean) As Boolean
        'this is the 1099 summary report;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        '    21       22         23        24 
        ' vendkey   expamt     expcnt     W9sw
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim begindate, enddate As Date
        Try
            If eusefiscal Then
                'calc the beginning & ending date for the calendar year
                begindate = CDate("07/01/" & erawyear - 1)
                enddate = CDate("07/01/" & erawyear)
            Else
                'calc the beginning & ending date for the calendar year
                begindate = CDate("01/01/" & erawyear)
                enddate = CDate("01/01/" & erawyear + 1)
            End If
        Catch ex As Exception
            Throw
        End Try

        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim ds As DataSet

        Try
            SSQL = "SELECT vend_number, vend_name, vend_1099_sw, vend_status," _
            & " vend_addr1, vend_addr2, vend_addr3, vend_city, vend_state," _
            & " vend_zip, vend_zip_ext, vend_phone1, vend_phone1_ext," _
            & " vend_phone2, vend_phone2_ext, vend_fax, vend_email_addr," _
            & " vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
            & " vend_autoinc_key, 0.0 AS expense, 0 AS items, vend_w9_sw" _
            & " FROM vend_info" _
            & " WHERE vend_status = 'O'" _
            & " ORDER BY vend_name;"
            SSQL += "SELECT vend_autoinc_key, SUM(chks_amount), COUNT(*)" _
            & " FROM chks_info" _
            & " WHERE (chks_status <> 'V')" _
            & " AND (chks_applied_date BETWEEN @p1 AND @p2)" _
            & " GROUP BY vend_autoinc_key" _
            & " ORDER BY vend_autoinc_key"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", begindate)
            cmd.Parameters.Add("@p2", enddate)
            da = New SqlDataAdapter(cmd)
            ds = New DataSet("register")
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

        Dim index, rowindex, items As Int32
        Dim expense As Double
        Dim vendorkey As Int32


        Try
            'iterate the detail grid & search for a matching vendor number & info; * vend_number - is char!!!!
            With Me.GridDetail
                For index = 0 To .Rows.Count - 1
                    vendorkey = DirectCast(.GetData(index, 21), Int32)
                    'find the matching vendor
                    rowindex = Me.GridTotals.FindRow(vendorkey.ToString, 0, 0, False, True, False)
                    If rowindex >= 0 Then
                        'the vendor has expenditures
                        expense = CDbl(Me.GridTotals.GetData(rowindex, 1))
                        items = CInt(Me.GridTotals.GetData(rowindex, 2))
                        'load the expense info into the detail grid
                        .SetData(index, 22, expense)
                        .SetData(index, 23, items)
                    End If
                Next
            End With

        Catch ex As Exception

        End Try

        Try
            If eusefiscal Then
                Me.CellMiddleBottom = "Fiscal year " & erawyear.ToString
            Else
                Me.CellMiddleBottom = "Calendar year " & erawyear.ToString
            End If
            If euseminimum600 Then Me.CellRightMiddle = "$600 Minimum"
            Application.DoEvents()
            'render the table;
            PrintVendorExpenditureSummary(eprint1099only, eincludezerobalances, euseminimum600)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateVendorExpenditureDetail(ByVal erawyear As Int32, ByVal eprint1099only As Boolean, ByVal eprintssn As Boolean, ByVal eusefiscal As Boolean) As Boolean
        'this is the 1099 detail report;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        '    21       22         23        24       25        26         27 
        ' chkfisyr  chknum     chkamt    lineamt   acctnum   subnum   chkappl   
        '    28       29         30        31       32
        ' chkissue remarks    vendkey    hdrkey    detlkey
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim begindate, enddate As Date
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable

        Try
            If eusefiscal Then
                'calc the beginning & ending date for the calendar year
                begindate = CDate("07/01/" & erawyear - 1)
                enddate = CDate("07/01/" & erawyear)
            Else
                'calc the beginning & ending date for the calendar year
                begindate = CDate("01/01/" & erawyear)
                enddate = CDate("01/01/" & erawyear + 1)
            End If
        Catch ex As Exception
            Throw
        End Try

        If eprint1099only Then
            'get all 1099 expenditures by vendor using vendor number only;
            SSQL = "SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1," _
            & " vend_addr2, vend_addr3, vend_city, vend_state, vend_zip, vend_zip_ext," _
            & " vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax," _
            & " vend_email_addr, vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
            & " chks_fisyr, chks_num, chks_amount, ckdt_amount, af_acct_num, as_acct_num," _
            & " chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key," _
            & " h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND v.vend_number = h.vend_number" _
            & " AND h.vend_autoinc_key = 0" _
            & " AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V')" _
            & " AND v.vend_1099_sw = 'Y'" _
            & " AND chks_applied_date BETWEEN @p1 AND @p2" _
            & " UNION "
            'get all 1099 expenditures by vendor using payee key only;
            SSQL &= "SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1," _
            & " vend_addr2, vend_addr3, vend_city, vend_state, vend_zip, vend_zip_ext," _
            & " vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax," _
            & " vend_email_addr, vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
            & " chks_fisyr, chks_num, chks_amount, ckdt_amount, af_acct_num, as_acct_num," _
            & " chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key," _
            & " h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND v.vend_autoinc_key = h.vend_autoinc_key" _
            & " AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V')" _
            & " AND v.vend_1099_sw = 'Y'" _
            & " AND chks_applied_date BETWEEN @p1 AND @p2" _
            & " ORDER BY vend_name, h.chks_autoinc_key, d.ckdt_autoinc_key"
        Else
            'get all expenditures by vendor using vendor number only;
            SSQL = "SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1," _
            & " vend_addr2, vend_addr3, vend_city, vend_state, vend_zip, vend_zip_ext," _
            & " vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax," _
            & " vend_email_addr, vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
            & " chks_fisyr, chks_num, chks_amount, ckdt_amount, af_acct_num, as_acct_num," _
            & " chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key," _
            & " h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND v.vend_number = h.vend_number" _
            & " AND h.vend_autoinc_key = 0" _
            & " AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V')" _
            & " AND chks_applied_date BETWEEN @p1 AND @p2" _
            & " UNION "
            'get all expenditures by vendor using payee key only;
            SSQL &= "SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1," _
            & " vend_addr2, vend_addr3, vend_city, vend_state, vend_zip, vend_zip_ext," _
            & " vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax," _
            & " vend_email_addr, vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
            & " chks_fisyr, chks_num, chks_amount, ckdt_amount, af_acct_num, as_acct_num," _
            & " chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key," _
            & " h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND v.vend_autoinc_key = h.vend_autoinc_key" _
            & " AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V')" _
            & " AND chks_applied_date BETWEEN @p1 AND @p2" _
            & " ORDER BY vend_name, h.chks_autoinc_key, d.ckdt_autoinc_key"
        End If

        Try
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", begindate)
            cmd.Parameters.Add("@p2", enddate)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("payees")
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

        Try
            If eusefiscal Then
                Me.CellMiddleBottom = "Fiscal year " & erawyear.ToString
            Else
                Me.CellMiddleBottom = "Calendar year " & erawyear.ToString
            End If
            Application.DoEvents()
            'render the table
            Call PrintVendorExpenditureDetail(eprintssn)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateVendorExpenditureDetailSingle(ByVal evendorkey As Int32, ByVal erawyear As Int32, ByVal eprintssn As Boolean, ByVal eusefiscal As Boolean) As Boolean
        'this is the 1099 detail report for a single vendor;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        '    21       22         23        24       25        26         27 
        ' chkfisyr  chknum     chkamt    lineamt   acctnum   subnum   chkappl   
        '    28       29         30        31       32
        ' chkissue remarks    vendkey    hdrkey    detlkey
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim begindate, enddate As Date
        Dim SSQL As String
        Dim cmd As SqlCommand
        Dim da As SqlDataAdapter
        Dim tbl As DataTable
        Dim VendorName, VendorNumber As String

        Try
            If eusefiscal Then
                'calc the beginning & ending date for the calendar year
                begindate = CDate("07/01/" & erawyear - 1)
                enddate = CDate("07/01/" & erawyear)
            Else
                'calc the beginning & ending date for the calendar year
                begindate = CDate("01/01/" & erawyear)
                enddate = CDate("01/01/" & erawyear + 1)
            End If
        Catch ex As Exception
            Throw
        End Try

        Try
            SSQL = "SELECT vend_name, vend_number FROM vend_info AS v WHERE vend_autoinc_key = @p1"
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", evendorkey)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("vendstuff")
            da.Fill(tbl)
            With tbl
                If .Rows.Count = 1 Then
                    VendorName = CType(.Rows(0).Item(0), String)
                    VendorNumber = CType(.Rows(0).Item(1), String)
                Else
                    Throw New ArgumentException("Invalid vendor record for vendor key " & evendorkey)
                End If
            End With
        Catch ex As Exception
            Throw
        End Try


        Try
            '-pull by vendnum;
            SSQL = "SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1, vend_addr2, vend_addr3, vend_city, vend_state," _
            & " vend_zip, vend_zip_ext, vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax, vend_email_addr," _
            & " vend_ssn, vend_ssn_format, vend_transdate, vend_datetime, chks_fisyr, chks_num, chks_amount, ckdt_amount," _
            & " af_acct_num, as_acct_num, chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key, h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND (v.vend_number = h.vend_number) AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V') AND chks_applied_date BETWEEN @p1 AND @p2" _
            & " AND (v.vend_number = @p3)" _
            & " UNION"
            'pull by vend_autoinc_key;
            SSQL += " SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1, vend_addr2, vend_addr3, vend_city, vend_state," _
            & " vend_zip, vend_zip_ext, vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax, vend_email_addr," _
            & " vend_ssn, vend_ssn_format, vend_transdate, vend_datetime, chks_fisyr, chks_num, chks_amount, ckdt_amount," _
            & " af_acct_num, as_acct_num, chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key, h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND (v.vend_autoinc_key = h.vend_autoinc_key) AND (vend_status <> 'D')" _
            & " AND (chks_status <> 'V') AND chks_applied_date BETWEEN @p1 AND @p2 AND v.vend_autoinc_key = @p4" _
            & " UNION"
            'pull by vend_name/chks_payee_name;
            SSQL += " SELECT v.vend_number, vend_name, vend_1099_sw, vend_status, vend_addr1, vend_addr2, vend_addr3, vend_city, vend_state," _
            & " vend_zip, vend_zip_ext, vend_phone1, vend_phone1_ext, vend_phone2, vend_phone2_ext, vend_fax, vend_email_addr," _
            & " vend_ssn, vend_ssn_format, vend_transdate, vend_datetime, chks_fisyr, chks_num, chks_amount, ckdt_amount," _
            & " af_acct_num, as_acct_num, chks_applied_date, chks_datetime, ckdt_descr, v.vend_autoinc_key, h.chks_autoinc_key, d.ckdt_autoinc_key" _
            & " FROM vend_info AS v, chks_info AS h, chks_detl AS d" _
            & " WHERE h.chks_autoinc_key = d.chks_autoinc_key" _
            & " AND v.vend_autoinc_key = h.vend_autoinc_key" _
            & " AND (vend_status <> 'D') AND (chks_status <> 'V')" _
            & " AND (chks_applied_date BETWEEN @p1 AND @p2)" _
            & " AND (vend_name = chks_payee_name)" _
            & " AND (chks_payee_name = @p5)" _
            & " ORDER BY vend_name, h.chks_autoinc_key, d.ckdt_autoinc_key"
            '
            cn = New SqlConnection(Me.ConnectionString)
            cmd = New SqlCommand(SSQL, cn)
            cmd.Parameters.Add("@p1", begindate)
            cmd.Parameters.Add("@p2", enddate)
            cmd.Parameters.Add("@p3", VendorNumber)
            cmd.Parameters.Add("@p4", evendorkey)
            cmd.Parameters.Add("@p5", VendorName)
            da = New SqlDataAdapter(cmd)
            tbl = New DataTable("vendors")
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

        Try
            If eusefiscal Then
                Me.CellMiddleBottom = "Fiscal year " & erawyear.ToString
            Else
                Me.CellMiddleBottom = "Calendar year " & erawyear.ToString
            End If
            Application.DoEvents()
            'render the table
            Call PrintVendorExpenditureDetail(eprintssn)
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

    Public Function GenerateVendorListing() As Boolean
        'this method retrieves all available vendors;
        ''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SSQL As String
        Dim cmd As SqlCommand
        cn = New SqlConnection(Me.ConnectionString)
        SSQL = "SELECT vend_number, vend_name, vend_1099_sw, vend_status," _
        & " vend_addr1, vend_addr2, vend_addr3, vend_city, vend_state," _
        & " vend_zip, vend_zip_ext, vend_phone1, vend_phone1_ext," _
        & " vend_phone2, vend_phone2_ext, vend_fax, vend_email_addr," _
        & " vend_ssn, vend_ssn_format, vend_transdate, vend_datetime," _
        & " vend_autoinc_key" _
        & " FROM vend_info" _
        & " WHERE (vend_status <> 'D')" _
        & " AND (vend_number <> '00000')" _
        & " ORDER BY vend_name"
        cmd = New SqlCommand(SSQL, cn)
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

        'With Me.GridDetail
        '    .Visible = True
        'End With
        'Me.Prev1.Visible = False
        'Me.ShowDialog()
        'Exit Function

        Try
            Me.CellMiddleBottom = "FY-" & Me.FiscalYear.ToString
            Application.DoEvents()
            'render the table
            PrintVendorListing()
            Application.DoEvents()
            Me.ShowDialog()
        Catch ex As Exception
            Throw
        End Try

    End Function

#End Region

#Region "  Methods Rendering "

    Private Sub Print1099Vendors()
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '    key      ssn        name     addr1    addr2     city       state
        '     7        8          9        10       11        12         13   
        '    zip      ext       total     count 
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VendorExpenditures"
        Me.ReportName = "Activity Fund 1099 Vendors"
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
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim x, y, count, index, items, key, prevkey, pagecheck As Int32
        Dim checkamount, totalamount As Decimal
        Dim addr1, addr2, city, name, ssn, state, zip, zipext As String
        Dim remarks As String
        Dim dopage As Boolean

        Try
            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    With Me.GridDetail
                        key = CInt(.GetData(index, 0))
                        ssn = DirectCast(.GetData(index, 1), String).Trim
                        name = DirectCast(.GetData(index, 2), String).Trim
                        addr1 = DirectCast(.GetData(index, 3), String).Trim
                        addr2 = DirectCast(.GetData(index, 4), String).Trim
                        city = DirectCast(.GetData(index, 5), String).Trim
                        state = DirectCast(.GetData(index, 6), String).Trim
                        zip = DirectCast(.GetData(index, 7), String).Trim
                        zipext = DirectCast(.GetData(index, 8), String).Trim
                        checkamount = CDec(.GetData(index, 9))
                        items = CInt(.GetData(index, 10))
                        'clear the remarks for this line;
                        remarks = ""
                        'test 1099 for validation;
                        If ssn.Length < 1 Then remarks = "SSN/FEI REQUIRED FOR 1099 VENDORS"
                        If name.Length < 1 Then remarks = "VENDOR NAME REQUIRED FOR 1099 VENDORS"
                    End With

                    If index = 0 Then dopage = True

                    If dopage Then
                        y = 32
                        'print the column headers;
                        .RenderDirectText(2, y, "SSN/FEI", 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Vendor", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Items", 15, 5, verdanaright8bold)
                        .RenderDirectText(80, y, "Amount", 20, 5, verdanaright8bold)
                        .RenderDirectText(110, y, "Remarks", 80, 5, verdanaleft8bold)
                        'print line below the column headers;
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 40
                        prevkey = key
                        pagecheck = 0
                        dopage = False
                    End If

                    'render the line;
                    .RenderDirectText(0, y, ssn, 20, 5, verdanaleft8)
                    .RenderDirectText(20, y, name, 50, 5, verdanaleft8)
                    .RenderDirectText(70, y, items.ToString, 10, 10, verdanaright8)
                    .RenderDirectText(80, y, checkamount.ToString.Format("{0:F2}", checkamount), 20, 10, verdanaright8)
                    .RenderDirectText(110, y, remarks, 80, 10, verdanaleft8bold)

                    If y >= 245 Then
                        .NewPage()
                        y = 32
                        'print the column headers
                        .RenderDirectText(2, y, "SSN/FEI", 20, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Vendor", 50, 5, verdanaleft8bold)
                        .RenderDirectText(65, y, "Items", 15, 5, verdanaright8bold)
                        .RenderDirectText(80, y, "Amount", 20, 5, verdanaright8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 40
                        prevkey = key
                        pagecheck = 0
                        dopage = False
                    Else
                        y += 8
                    End If
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
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

    Private Sub PrintAuditListing()
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0         1        2         3          4        5
        '  vendkey   vname    audname   audtime    userkey  username
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VendorListing"
        Me.ReportName = "Activity Fund Vendor Audit"
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
        'define the styles;
        Call DefineStyles()
        'define the document;
        Call DefineDocumentSettings(Me.DocumentName)

        Dim x, y, index, currow, pvendorkey, vendorkey As Int32
        Dim auditname, fullname, vendorname As String
        Dim auditdate As Date

        Try
            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        y = 36
                        'print the column headers;
                        .RenderDirectText(0, y, "Vendor name", 60, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Original", 20, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "Changed", 20, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Changed By", 25, 5, verdanaleft8bold)
                        'print line below the column headers;
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                    End If

                    With Me.GridDetail
                        vendorkey = CType(.GetData(index, 0), Int32)
                        vendorname = CType(.GetData(index, 1), String)
                        auditname = CType(.GetData(index, 2), String)
                        auditdate = CType(.GetData(index, 3), Date)
                        fullname = CType(.GetData(index, 5), String)
                    End With

                    If currow > 1 Then y += 5

                    'mask the vendor name if on same vendor;
                    If pvendorkey = vendorkey Then vendorname = """"
                    'render the record;
                    .RenderDirectText(0, y, vendorname, 60, 10, verdanaleft8)
                    .RenderDirectText(60, y, auditname, 60, 10, verdanaleft8)
                    .RenderDirectText(120, y + 4, auditdate.ToShortTimeString, 60, 5, verdanaleft8)
                    .RenderDirectText(120, y, auditdate.ToString.Format("{0:MM/dd/yyyy}", auditdate), 20, 5, verdanaleft8)
                    .RenderDirectText(150, y, fullname, 40, 5, verdanaleft8)
                    'store the vendorkey;
                    pvendorkey = vendorkey
                    '
                    y += 7

                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        y = 36
                        'print the column headers;
                        .RenderDirectText(0, y, "Vendor name", 60, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Original", 20, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "Changed", 20, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Changed By", 25, 5, verdanaleft8bold)
                        'print line below the column headers;
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                        currow = 0
                    End If
                Next
                Application.DoEvents()
                'end-of-report;
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

    Private Sub PrintVendorExpenditureDetail(ByVal eprintssn As Boolean)
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        '    21       22         23        24       25        26         27 
        ' chkfisyr  chknum     chkamt    lineamt   acctnum   subnum   chkappl   
        '    28       29         30        31       32
        ' chkissue remarks    vendkey    hdrkey    detlkey
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VendorExpenditures"
        Me.ReportName = "Activity Fund Vendor Expenditures"
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

        Dim x, y, fisyr, index, count, hdrkey, prevhdrkey, pagecheck As Int32
        Dim lineamount, checkamount, sumcheckamount As Double
        Dim vendorname, vendornumber, prevvendornumber, nextvendornumber As String
        Dim flag1099, checknumber, account, subaccount, remarks As String
        Dim ssn, ssnfmt As String
        Dim applied, issued As Date
        Dim docheckheader, dopage As Boolean

        Try
            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    With Me.GridDetail
                        vendornumber = DirectCast(.GetData(index, 0), String).Trim
                        vendorname = DirectCast(.GetData(index, 1), String).Trim
                        flag1099 = DirectCast(.GetData(index, 2), String).Trim.ToUpper
                        If flag1099 = "Y" Then flag1099 = "Yes" Else flag1099 = "No"
                        ssn = DirectCast(.GetData(index, 17), String).Trim
                        ssnfmt = DirectCast(.GetData(index, 18), String).Trim
                        fisyr = CInt(.GetData(index, 21))
                        checknumber = DirectCast(.GetData(index, 22), String).Trim
                        account = DirectCast(.GetData(index, 25), String).Trim
                        subaccount = DirectCast(.GetData(index, 26), String).Trim
                        applied = CDate(.GetData(index, 27))
                        issued = CDate(.GetData(index, 28))
                        remarks = DirectCast(.GetData(index, 29), String).Trim
                        checkamount = CDbl(.GetData(index, 23))
                        lineamount = CDbl(.GetData(index, 24))
                        hdrkey = CInt(.GetData(index, 31))
                        'check for ssn printing ability
                        If Not eprintssn Then ssn = ""
                        If flag1099 <> "Yes" Then ssn = ""
                    End With

                    'has the vendor changed
                    If vendornumber.Compare(vendornumber, prevvendornumber) <> 0 Then dopage = True

                    If dopage Then
                        If index > 0 Then .NewPage()
                        y = 39
                        If ssn.Length > 0 Then .RenderDirectText(0, y, ssnfmt, 30, 5, verdanaright10bold)
                        .RenderDirectText(33, y, vendornumber, 20, 5, verdanaright10bold)
                        .RenderDirectText(55, y, vendorname, 50, 5, verdanaleft10bold)
                        .RenderDirectText(120, y, "1099 Vendor:    " + flag1099, 40, 5, verdanaright10bold)
                        y = 50
                        'print the column headers
                        .RenderDirectText(2, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Year", 10, 5, verdanaright8bold)
                        .RenderDirectText(32, y, "Number", 23, 5, verdanaleft8bold)
                        .RenderDirectText(50, y, "Account", 20, 5, verdanaleft8bold)
                        .RenderDirectText(70, y, "Remarks", 30, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Line item", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Amount", 20, 5, verdanaright8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 58
                        prevvendornumber = vendornumber
                        pagecheck = 0
                        dopage = False
                        docheckheader = True
                    End If

                    'has the check header changed
                    If hdrkey <> prevhdrkey Then docheckheader = True

                    If docheckheader Then
                        If pagecheck >= 1 Then y += 3
                        .RenderDirectText(0, y, issued.ToString.Format("{0:MM/dd/yyyy}", issued), 20, 5, verdanaleft8)
                        .RenderDirectText(20, y, fisyr.ToString, 10, 5, verdanaright8)
                        .RenderDirectText(32, y, checknumber, 20, 5, verdanaleft8)
                        .RenderDirectText(50, y, account + "-" + subaccount, 20, 5, verdanaleft8)
                        .RenderDirectText(70, y, remarks, 80, 5, verdanaleft8)
                        .RenderDirectText(150, y, lineamount.ToString.Format("{0:F2}", lineamount), 20, 5, verdanaright8)
                        .RenderDirectText(170, y, checkamount.ToString.Format("{0:F2}", checkamount), 20, 5, verdanaright8)
                        prevhdrkey = hdrkey
                        pagecheck += 1
                        docheckheader = False
                    Else
                        .RenderDirectText(50, y, account + "-" + subaccount, 20, 5, verdanaleft8)
                        .RenderDirectText(70, y, remarks, 80, 5, verdanaleft8)
                        .RenderDirectText(150, y, lineamount.ToString.Format("{0:F2}", lineamount), 20, 5, verdanaright8)
                    End If

                    'tally the values
                    sumcheckamount += lineamount

                    If y >= 250 Then
                        .NewPage()
                        y = 39
                        .RenderDirectText(30, y, vendornumber, 20, 5, verdanaright10bold)
                        .RenderDirectText(55, y, vendorname, 50, 5, verdanaleft10bold)
                        .RenderDirectText(120, y, "1099 Vendor:    " + flag1099, 40, 5, verdanaright10bold)
                        y = 50
                        'print the column headers
                        .RenderDirectText(2, y, "Issued", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Year", 10, 5, verdanaright8bold)
                        .RenderDirectText(32, y, "Number", 23, 5, verdanaleft8bold)
                        .RenderDirectText(50, y, "Account", 20, 5, verdanaleft8bold)
                        .RenderDirectText(70, y, "Remarks", 30, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, "Line item", 20, 5, verdanaright8bold)
                        .RenderDirectText(170, y, "Amount", 20, 5, verdanaright8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 53
                    End If

                    'check to print vendor total
                    If index < (Me.GridDetail.Rows.Count - 1) Then
                        nextvendornumber = DirectCast(Me.GridDetail.GetData(index + 1, 0), String).Trim
                        'will the vendor change
                        If vendornumber.Compare(vendornumber, nextvendornumber) <> 0 Then
                            y += 9
                            .RenderDirectLine(100, y - 0.5, 190, y - 0.5, Color.Gray, 0.5)
                            .RenderDirectText(100, y, "Total for vendor:", 50, 5, verdanaleft8bold)
                            .RenderDirectText(150, y, sumcheckamount.ToString.Format("{0:C2}", sumcheckamount), 40, 5, verdanaright8bold)
                            .RenderDirectLine(100, y + 5, 190, y + 5, Color.Gray, 0.5)
                            sumcheckamount = 0
                        End If
                    End If
                    If index = (Me.GridDetail.Rows.Count - 1) Then
                        y += 9
                        .RenderDirectLine(100, y - 0.5, 190, y - 0.5, Color.Gray, 0.5)
                        .RenderDirectText(100, y, "Total for vendor:", 50, 5, verdanaleft8bold)
                        .RenderDirectText(150, y, sumcheckamount.ToString.Format("{0:C2}", sumcheckamount), 40, 5, verdanaright8bold)
                        .RenderDirectLine(100, y + 5, 190, y + 5, Color.Gray, 0.5)
                        sumcheckamount = 0
                    End If

                    y += 5
                    count += 1
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
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

    Private Sub PrintVendorExpenditureSummary(ByVal eprint1099only As Boolean, ByVal eprintzero As Boolean, ByVal euseminimum As Boolean)
        ''''''''''''''''''''''''''''''''' GRIDDETAIL ''''''''''''''''''''''''''''' 
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        '    21       22         23        24 
        ' vendkey   expamt     expcnt     W9sw
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VendorExpenditures"
        Me.ReportName = "Activity Fund Vendor Expenditures"
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


        Dim x, y, index, currow, count, checkitems, sumcheckitems, w9sw As Int32
        Dim checkamount, sumcheckamount As Decimal
        Dim number, name, flag1099, status, addr1, addr2, addr3, city, state, w9str, zip, zipx, zipfull As String
        Dim phone1, phone1x, phone2, phone2x, fax, email, ssn, ssnfmt As String
        Dim transdate, issuedate As Date

        Try
            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    If index = 0 Then
                        y = 36
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Name", 95, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "W-9", 15, 5, verdanaleft8bold)
                        .RenderDirectText(133, y, "1099", 15, 5, verdanaleft8bold)
                        .RenderDirectText(142, y, "Items", 20, 5, verdanaright8bold)
                        .RenderDirectText(160, y, "Total Expense", 30, 5, verdanaright8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                    End If
                    With Me.GridDetail
                        number = DirectCast(.GetData(index, 0), String).Trim
                        name = DirectCast(.GetData(index, 1), String).Trim
                        flag1099 = DirectCast(.GetData(index, 2), String).Trim.ToUpper
                        status = DirectCast(.GetData(index, 3), String).Trim.ToUpper
                        addr1 = DirectCast(.GetData(index, 4), String).Trim
                        addr2 = DirectCast(.GetData(index, 5), String).Trim
                        addr3 = DirectCast(.GetData(index, 6), String).Trim
                        city = DirectCast(.GetData(index, 7), String).Trim
                        state = DirectCast(.GetData(index, 8), String).Trim.ToUpper
                        zip = DirectCast(.GetData(index, 9), String).Trim
                        zipx = DirectCast(.GetData(index, 10), String).Trim
                        zipfull = zip
                        If zipx.Length > 0 Then zipfull = zip & "-" & zipx
                        phone1 = DirectCast(.GetData(index, 11), String).Trim
                        phone1x = DirectCast(.GetData(index, 12), String).Trim
                        phone2 = DirectCast(.GetData(index, 13), String).Trim
                        phone2x = DirectCast(.GetData(index, 14), String).Trim
                        fax = DirectCast(.GetData(index, 15), String).Trim
                        email = DirectCast(.GetData(index, 16), String).Trim
                        ssn = DirectCast(.GetData(index, 17), String).Trim
                        ssnfmt = DirectCast(.GetData(index, 18), String).Trim
                        transdate = CDate(.GetData(index, 19))
                        issuedate = CDate(.GetData(index, 20))
                        checkamount = CDec(.GetData(index, 22))
                        checkitems = CInt(.GetData(index, 23))
                        'added W-9 switch on 12.22.2014;
                        w9sw = CInt(.GetData(index, 24))
                        If w9sw = -1 Then w9str = "R"
                        If w9sw = 0 Then w9str = "N"
                        If w9sw = 1 Then w9str = "Y"

                        If (flag1099 = "N") And eprint1099only Then GoTo Bypass
                        If checkamount = 0.0 And Not eprintzero Then GoTo Bypass
                        If euseminimum Then
                            If checkamount < 600 Then GoTo Bypass
                        End If
                    End With
                    currow += 1
                    If currow > 1 Then y += 4
                    .RenderDirectText(2, y, number, 20, 5, verdanaleft8)
                    .RenderDirectText(20, y, name, 95, 5, verdanaleft8)
                    .RenderDirectText(122, y, w9str, 15, 5, verdanaleft8)
                    .RenderDirectText(137, y, flag1099, 15, 5, verdanaleft8)
                    .RenderDirectText(140, y, checkitems.ToString, 20, 5, verdanaright8)
                    .RenderDirectText(160, y, checkamount.ToString.Format("{0:F2}", checkamount), 30, 5, verdanaright8)
                    'tally the values
                    sumcheckitems += checkitems
                    sumcheckamount += checkamount
                    y += 4
                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        y = 36
                        'print the column headers
                        .RenderDirectText(0, y, "Number", 15, 5, verdanaleft8bold)
                        .RenderDirectText(20, y, "Name", 95, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "W-9", 15, 5, verdanaleft8bold)
                        .RenderDirectText(133, y, "1099", 15, 5, verdanaleft8bold)
                        .RenderDirectText(142, y, "Items", 20, 5, verdanaright8bold)
                        .RenderDirectText(160, y, "Total Expense", 30, 5, verdanaright8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                        currow = 0
                    End If
                    count += 1
Bypass:
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 15
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Amount:", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 5, "Total Expenses:", 50, 5, verdanaright8bold)
                .RenderDirectText(60, y + 10, "Total Vendors:", 50, 5, verdanaright8bold)
                .RenderDirectText(160, y, sumcheckamount.ToString.Format("{0:C2}", sumcheckamount), 30, 5, verdanaright8bold)
                .RenderDirectText(160, y + 5, sumcheckitems.ToString.Format("{0:D1}", sumcheckitems), 30, 5, verdanaright8bold)
                .RenderDirectText(160, y + 10, count.ToString.Format("{0:D1}", count), 30, 5, verdanaright8bold)
                y += 16
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

    Private Sub PrintVendorListing()
        ''''''''''''''''''''''''''''''''' GRIDDETAIL '''''''''''''''''''''''''''''
        '     0        1          2         3        4         5          6   
        '  number     name      1099sw    status    addr1    addr2     addr3
        '     7        8          9        10       11        12         13   
        '   city     state       zip      zipext   phone1   phone1x    phone2
        '    14       15         16        17       18        19         20   
        ' phone2x     fax       email      ssn     ssnfmt    trans    created
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Me.DocumentName = "VendorListing"
        Me.ReportName = "Activity Fund Vendor Listing"
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

        Dim x, y, index, currow, count As Int32
        Dim number, name, flag1099, status, addr1, addr2, addr3, city, state, zip, zipx, zipfull As String
        Dim phone1, phone1x, phone2, phone2x, fax, email, ssn, ssnfmt As String
        Dim transdate, issuedate As Date

        Try
            With Me.Doc1
                .StartDoc()
                For index = 0 To Me.GridDetail.Rows.Count - 1
                    currow += 1
                    If index = 0 Then
                        y = 36
                        'print the column headers
                        .RenderDirectText(0, y, "Vendor name", 60, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Address", 20, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "City", 20, 5, verdanaleft8bold)
                        .RenderDirectText(158, y, "State", 20, 5, verdanaleft8bold)
                        .RenderDirectText(168, y, "Postal Code", 25, 5, verdanaleft8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                    End If
                    With Me.GridDetail
                        number = DirectCast(.GetData(index, 0), String).Trim
                        name = DirectCast(.GetData(index, 1), String).Trim
                        flag1099 = DirectCast(.GetData(index, 2), String).Trim.ToUpper
                        status = DirectCast(.GetData(index, 3), String).Trim.ToUpper
                        addr1 = DirectCast(.GetData(index, 4), String).Trim
                        addr2 = DirectCast(.GetData(index, 5), String).Trim
                        addr3 = DirectCast(.GetData(index, 6), String).Trim
                        city = DirectCast(.GetData(index, 7), String).Trim
                        state = DirectCast(.GetData(index, 8), String).Trim.ToUpper
                        zip = DirectCast(.GetData(index, 9), String).Trim
                        zipx = DirectCast(.GetData(index, 10), String).Trim
                        zipfull = zip
                        If zipx.Length > 0 Then zipfull = zip & "-" & zipx
                        phone1 = DirectCast(.GetData(index, 11), String).Trim
                        phone1x = DirectCast(.GetData(index, 12), String).Trim
                        phone2 = DirectCast(.GetData(index, 13), String).Trim
                        phone2x = DirectCast(.GetData(index, 14), String).Trim
                        fax = DirectCast(.GetData(index, 15), String).Trim
                        email = DirectCast(.GetData(index, 16), String).Trim
                        ssn = DirectCast(.GetData(index, 17), String).Trim
                        ssnfmt = DirectCast(.GetData(index, 18), String).Trim
                        transdate = CDate(.GetData(index, 19))
                        issuedate = CDate(.GetData(index, 20))
                    End With

                    If currow > 1 Then y += 5
                    .RenderDirectText(0, y, name, 60, 5, verdanaleft8)
                    .RenderDirectText(60, y, addr1, 60, 5, verdanaleft8)
                    .RenderDirectText(60, y + 4, addr2, 60, 5, verdanaleft8)
                    .RenderDirectText(120, y, city, 40, 5, verdanaleft8)
                    .RenderDirectText(160, y, state, 10, 5, verdanaleft8)
                    .RenderDirectText(170, y, zipfull, 20, 5, verdanaleft8)
                    y += 5
                    If y >= 250 Then
                        If index >= (Me.GridDetail.Rows.Count - 1) Then Exit For
                        .NewPage()
                        y = 36
                        'print the column headers
                        .RenderDirectText(0, y, "Vendor name", 60, 5, verdanaleft8bold)
                        .RenderDirectText(60, y, "Address", 20, 5, verdanaleft8bold)
                        .RenderDirectText(120, y, "City", 20, 5, verdanaleft8bold)
                        .RenderDirectText(158, y, "State", 20, 5, verdanaleft8bold)
                        .RenderDirectText(168, y, "Postal Code", 25, 5, verdanaleft8bold)
                        'print line below the column headers
                        .RenderDirectLine(0, y + 5, 190, y + 5, Color.Gray, 0.5)
                        y = 45
                        currow = 0
                    End If
                    count += 1
                    'expose the current record & count to the caller
                    'EventRecordProcessed((reccurrent), reccount)
                Next

                'print totals
                y += 15
                If y > 240 Then
                    .NewPage()
                    y = 65
                End If
                'draw top of total box
                .RenderDirectLine(59, y - 2, 190, y - 2, Color.Black, 0.25)
                .RenderDirectLine(59, y - 1.5, 190, y - 1.5, Color.Black, 0.25)
                .RenderDirectText(60, y, "Total Vendors:", 50, 5, verdanaright8bold)
                .RenderDirectText(150, y, count.ToString.Format("{0:D2}", count), 20, 5, verdanaright8bold)
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

#End Region

End Class
