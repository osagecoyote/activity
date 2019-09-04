Imports C1.C1PrintDocument
Imports System.IO

Public Class frmVendors
    Inherits System.Windows.Forms.Form

#Region "  Windows Form Designer generated code "

    Public Sub New()
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
        Me.Bankaccountnum = authobj.BankAccountNumber
        Me.currentmonth = authobj.CurrentMonthString()
        Me.fiscalyear = authobj.FiscalYear
        Me.monthbegindate = authobj.CurrentMonthBeginning
        Me.monthenddate = authobj.CurrentMonthEnding


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
    Friend WithEvents PreviewToolBarButton1 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton2 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton3 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton4 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton5 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton6 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton7 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton8 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton9 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton10 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton11 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton12 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton13 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton14 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton15 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton16 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton17 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton18 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton19 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton20 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton21 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton22 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton23 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton24 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton25 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton26 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton27 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton28 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton29 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton30 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton31 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton32 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents PreviewToolBarButton33 As C1.Win.C1PrintPreview.PreviewToolBarButton
    Friend WithEvents mnuZoomIn As System.Windows.Forms.MenuItem
    Friend WithEvents mnuZoomOut As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFullPage As System.Windows.Forms.MenuItem
    Friend WithEvents mnuActualPage As System.Windows.Forms.MenuItem
    Protected WithEvents Prev1 As C1.Win.C1PrintPreview.C1PrintPreview
    Protected WithEvents Doc1 As C1.C1PrintDocument.C1PrintDocument
    Protected WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents doc2 As C1.C1PrintDocument.C1PrintDocument
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmVendors))
        Me.Prev1 = New C1.Win.C1PrintPreview.C1PrintPreview
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.mnuZoomIn = New System.Windows.Forms.MenuItem
        Me.mnuZoomOut = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuActualPage = New System.Windows.Forms.MenuItem
        Me.mnuFullPage = New System.Windows.Forms.MenuItem
        Me.Doc1 = New C1.C1PrintDocument.C1PrintDocument
        Me.PreviewToolBarButton1 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton2 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton3 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton4 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton5 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton6 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton7 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton8 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton9 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton10 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton11 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton12 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton13 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton14 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton15 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton16 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton17 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton18 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton19 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton20 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton21 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton22 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton23 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton24 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton25 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton26 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton27 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton28 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton29 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton30 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton31 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton32 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.PreviewToolBarButton33 = New C1.Win.C1PrintPreview.PreviewToolBarButton
        Me.doc2 = New C1.C1PrintDocument.C1PrintDocument
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Prev1
        '
        Me.Prev1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Prev1.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIA"
        Me.Prev1.ContextMenu = Me.ContextMenu1
        Me.Prev1.DockPadding.All = 5
        Me.Prev1.Document = Me.Doc1
        Me.Prev1.Location = New System.Drawing.Point(0, 0)
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
        Me.Prev1.Size = New System.Drawing.Size(808, 424)
        Me.Prev1.Splitter.Cursor = System.Windows.Forms.Cursors.VSplit
        Me.Prev1.Splitter.Width = 3
        Me.Prev1.StatusBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.StatusBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Prev1.StatusBar.TabIndex = 4
        Me.Prev1.TabIndex = 0
        Me.Prev1.ToolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.PreviewToolBarButton1, Me.PreviewToolBarButton2, Me.PreviewToolBarButton3, Me.PreviewToolBarButton4, Me.PreviewToolBarButton5, Me.PreviewToolBarButton6, Me.PreviewToolBarButton7, Me.PreviewToolBarButton8, Me.PreviewToolBarButton9, Me.PreviewToolBarButton10, Me.PreviewToolBarButton11, Me.PreviewToolBarButton12, Me.PreviewToolBarButton13, Me.PreviewToolBarButton14, Me.PreviewToolBarButton15, Me.PreviewToolBarButton16, Me.PreviewToolBarButton17, Me.PreviewToolBarButton18, Me.PreviewToolBarButton19, Me.PreviewToolBarButton20, Me.PreviewToolBarButton21, Me.PreviewToolBarButton22, Me.PreviewToolBarButton23, Me.PreviewToolBarButton24, Me.PreviewToolBarButton25, Me.PreviewToolBarButton26, Me.PreviewToolBarButton27, Me.PreviewToolBarButton28, Me.PreviewToolBarButton29, Me.PreviewToolBarButton30, Me.PreviewToolBarButton31, Me.PreviewToolBarButton32, Me.PreviewToolBarButton33})
        Me.Prev1.ToolBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.Prev1.ToolBar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuZoomIn, Me.mnuZoomOut, Me.MenuItem1, Me.mnuActualPage, Me.mnuFullPage})
        '
        'mnuZoomIn
        '
        Me.mnuZoomIn.Index = 0
        Me.mnuZoomIn.Text = "Zoom In"
        '
        'mnuZoomOut
        '
        Me.mnuZoomOut.Index = 1
        Me.mnuZoomOut.Text = "Zoom Out"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "-"
        '
        'mnuActualPage
        '
        Me.mnuActualPage.Index = 3
        Me.mnuActualPage.Text = "Actual Page"
        '
        'mnuFullPage
        '
        Me.mnuFullPage.Index = 4
        Me.mnuFullPage.Text = "Full Page"
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
        'PreviewToolBarButton1
        '
        Me.PreviewToolBarButton1.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileOpen
        Me.PreviewToolBarButton1.ImageIndex = 0
        Me.PreviewToolBarButton1.ToolTipText = "File Open"
        '
        'PreviewToolBarButton2
        '
        Me.PreviewToolBarButton2.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FileSave
        Me.PreviewToolBarButton2.ImageIndex = 1
        Me.PreviewToolBarButton2.ToolTipText = "File Save"
        '
        'PreviewToolBarButton3
        '
        Me.PreviewToolBarButton3.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FilePrint
        Me.PreviewToolBarButton3.ImageIndex = 2
        Me.PreviewToolBarButton3.ToolTipText = "Print"
        '
        'PreviewToolBarButton4
        '
        Me.PreviewToolBarButton4.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.PageSetup
        Me.PreviewToolBarButton4.ImageIndex = 3
        Me.PreviewToolBarButton4.ToolTipText = "Page Setup"
        '
        'PreviewToolBarButton5
        '
        Me.PreviewToolBarButton5.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Reflow
        Me.PreviewToolBarButton5.ImageIndex = 4
        Me.PreviewToolBarButton5.ToolTipText = "Reflow"
        '
        'PreviewToolBarButton6
        '
        Me.PreviewToolBarButton6.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Stop
        Me.PreviewToolBarButton6.ImageIndex = 5
        Me.PreviewToolBarButton6.ToolTipText = "Stop"
        Me.PreviewToolBarButton6.Visible = False
        '
        'PreviewToolBarButton7
        '
        Me.PreviewToolBarButton7.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton7.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton8
        '
        Me.PreviewToolBarButton8.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ShowNavigationBar
        Me.PreviewToolBarButton8.ImageIndex = 6
        Me.PreviewToolBarButton8.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton8.ToolTipText = "Show Navigation Bar"
        '
        'PreviewToolBarButton9
        '
        Me.PreviewToolBarButton9.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton9.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton10
        '
        Me.PreviewToolBarButton10.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseHand
        Me.PreviewToolBarButton10.ImageIndex = 7
        Me.PreviewToolBarButton10.Pushed = True
        Me.PreviewToolBarButton10.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton10.ToolTipText = "Hand Tool"
        '
        'PreviewToolBarButton11
        '
        Me.PreviewToolBarButton11.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseZoom
        Me.PreviewToolBarButton11.ImageIndex = 8
        Me.PreviewToolBarButton11.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.PreviewToolBarButton11.ToolTipText = "Zoom In Tool"
        '
        'PreviewToolBarButton12
        '
        Me.PreviewToolBarButton12.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseZoomOut
        Me.PreviewToolBarButton12.ImageIndex = 25
        Me.PreviewToolBarButton12.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.PreviewToolBarButton12.ToolTipText = "Zoom Out Tool"
        Me.PreviewToolBarButton12.Visible = False
        '
        'PreviewToolBarButton13
        '
        Me.PreviewToolBarButton13.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.MouseSelect
        Me.PreviewToolBarButton13.ImageIndex = 9
        Me.PreviewToolBarButton13.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton13.ToolTipText = "Select Text"
        '
        'PreviewToolBarButton14
        '
        Me.PreviewToolBarButton14.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.FindText
        Me.PreviewToolBarButton14.ImageIndex = 10
        Me.PreviewToolBarButton14.ToolTipText = "Find Text"
        '
        'PreviewToolBarButton15
        '
        Me.PreviewToolBarButton15.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton15.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton16
        '
        Me.PreviewToolBarButton16.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoFirst
        Me.PreviewToolBarButton16.Enabled = False
        Me.PreviewToolBarButton16.ImageIndex = 11
        Me.PreviewToolBarButton16.ToolTipText = "First Page"
        '
        'PreviewToolBarButton17
        '
        Me.PreviewToolBarButton17.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoPrev
        Me.PreviewToolBarButton17.Enabled = False
        Me.PreviewToolBarButton17.ImageIndex = 12
        Me.PreviewToolBarButton17.ToolTipText = "Previous Page"
        '
        'PreviewToolBarButton18
        '
        Me.PreviewToolBarButton18.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoNext
        Me.PreviewToolBarButton18.ImageIndex = 13
        Me.PreviewToolBarButton18.ToolTipText = "Next Page"
        '
        'PreviewToolBarButton19
        '
        Me.PreviewToolBarButton19.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.GoLast
        Me.PreviewToolBarButton19.ImageIndex = 14
        Me.PreviewToolBarButton19.ToolTipText = "Last Page"
        '
        'PreviewToolBarButton20
        '
        Me.PreviewToolBarButton20.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton20.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'PreviewToolBarButton21
        '
        Me.PreviewToolBarButton21.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryPrev
        Me.PreviewToolBarButton21.Enabled = False
        Me.PreviewToolBarButton21.ImageIndex = 15
        Me.PreviewToolBarButton21.ToolTipText = "Previous View"
        Me.PreviewToolBarButton21.Visible = False
        '
        'PreviewToolBarButton22
        '
        Me.PreviewToolBarButton22.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.HistoryNext
        Me.PreviewToolBarButton22.Enabled = False
        Me.PreviewToolBarButton22.ImageIndex = 16
        Me.PreviewToolBarButton22.ToolTipText = "Next View"
        Me.PreviewToolBarButton22.Visible = False
        '
        'PreviewToolBarButton23
        '
        Me.PreviewToolBarButton23.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton23.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton23.Visible = False
        '
        'PreviewToolBarButton24
        '
        Me.PreviewToolBarButton24.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomOut
        Me.PreviewToolBarButton24.ImageIndex = 17
        Me.PreviewToolBarButton24.ToolTipText = "Zoom Out"
        Me.PreviewToolBarButton24.Visible = False
        '
        'PreviewToolBarButton25
        '
        Me.PreviewToolBarButton25.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ZoomIn
        Me.PreviewToolBarButton25.ImageIndex = 18
        Me.PreviewToolBarButton25.ToolTipText = "Zoom In"
        Me.PreviewToolBarButton25.Visible = False
        '
        'PreviewToolBarButton26
        '
        Me.PreviewToolBarButton26.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton26.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton26.Visible = False
        '
        'PreviewToolBarButton27
        '
        Me.PreviewToolBarButton27.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewActualSize
        Me.PreviewToolBarButton27.ImageIndex = 19
        Me.PreviewToolBarButton27.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton27.ToolTipText = "Actual Size"
        '
        'PreviewToolBarButton28
        '
        Me.PreviewToolBarButton28.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFullPage
        Me.PreviewToolBarButton28.ImageIndex = 20
        Me.PreviewToolBarButton28.Pushed = True
        Me.PreviewToolBarButton28.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton28.ToolTipText = "Full Page"
        '
        'PreviewToolBarButton29
        '
        Me.PreviewToolBarButton29.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewPageWidth
        Me.PreviewToolBarButton29.ImageIndex = 21
        Me.PreviewToolBarButton29.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton29.ToolTipText = "Page Width"
        '
        'PreviewToolBarButton30
        '
        Me.PreviewToolBarButton30.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewTwoPages
        Me.PreviewToolBarButton30.ImageIndex = 22
        Me.PreviewToolBarButton30.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.PreviewToolBarButton30.ToolTipText = "Two Pages"
        '
        'PreviewToolBarButton31
        '
        Me.PreviewToolBarButton31.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.ViewFourPages
        Me.PreviewToolBarButton31.ImageIndex = 23
        Me.PreviewToolBarButton31.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton
        Me.PreviewToolBarButton31.ToolTipText = "Four Pages"
        '
        'PreviewToolBarButton32
        '
        Me.PreviewToolBarButton32.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.None
        Me.PreviewToolBarButton32.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.PreviewToolBarButton32.Visible = False
        '
        'PreviewToolBarButton33
        '
        Me.PreviewToolBarButton33.Action = C1.Win.C1PrintPreview.ToolBarButtonActionEnum.Help
        Me.PreviewToolBarButton33.ImageIndex = 24
        Me.PreviewToolBarButton33.ToolTipText = "Help"
        Me.PreviewToolBarButton33.Visible = False
        '
        'doc2
        '
        Me.doc2.C1DPageSettings = "color:True;landscape:False;margins:100,100,100,100;papersize:850,1100,TABlAHQAdAB" & _
        "lAHIA"
        Me.doc2.ColumnSpacingStr = "0.5in"
        Me.doc2.ColumnSpacingUnit.DefaultType = True
        Me.doc2.ColumnSpacingUnit.UnitValue = "0.5in"
        Me.doc2.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Inch
        Me.doc2.DocumentName = ""
        '
        'frmVendors
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(804, 421)
        Me.Controls.Add(Me.Prev1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmVendors"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "  Activity Fund.Net Vendor Reporting"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Prev1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "  Class Members "

    Private reportyear As Int32

    'property values
    Private _msgtitle As String = "Activity Fund.Net Reports - Vendors"


    'header
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
    'print 
    Private headerstyle As C1DocStyle
    Private docstyle As C1DocStyle
    Private footerstyle As C1DocStyle
    Private linestyle1 As C1DocStyle
    Private linestyle2 As C1DocStyle
    Private linestyle3 As C1DocStyle
    Private linestyle4 As C1DocStyle
    Private linestyle5 As C1DocStyle
    Private linestyle6 As C1DocStyle
    Private linestyle7 As C1DocStyle
    Private linestyle8 As C1DocStyle
    Private linestyle9 As C1DocStyle

    Private monthbegindate As Date
    Private monthenddate As Date
    Private calenderyear As String

    Private bankaccountnumber As String

    'Vendorlist report
    Private vendcreate As Date
    Private vendnum As String
    Private vendname As String
    Private a1099sw As String
    Private vendstatus As String
    Private a1099flag As Boolean
    'Vendor Address details
    Private vendaddr1 As String
    Private vendaddr2 As String
    Private vendcity As String
    Private vendstate As String
    Private vendzip As String
    Private FlagAddress As Boolean
    Private sortbynumber As Boolean
    Private sortbyname As Boolean
    Private sortbycity As Boolean



    'Vendor detail report
    Private vendchkdate As Date
    Private vendchkfisyr As Int32
    Private vendchknum As String
    Private afacct As String
    Private asacct As String
    Private vendchkamt As Double
    Private vendrmrks As String
    Private Totvendchks As Double
    Private endgenerate As Boolean
    Private currvendnum, prevvendnum As String
    Private rcrds As Int32

    'Vendor report flags for titles
    Private vendlist As Boolean
    Private vendtransactionledger As Boolean
    Private vendtransdetail As Boolean
    Private vendspecific As Boolean
    Private flagover600 As Boolean


#End Region

#Region "  Button Events"

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            Me.Dispose()
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "  Document Settings & Styles "

    Private Sub DefineDocumentSettings()
        'settings for all documents
        With Me.Doc1
            .DefaultUnit = UnitTypeEnum.Mm
            .DefaultUnitOfFrames = UnitTypeEnum.Mm
            .Style = docstyle
            .PageHeader.Height = 24
            .PageHeader.Style = headerstyle
            .PageFooter.Style.TextAlignHorz = AlignHorzEnum.Left
            .PageFooter.Height = 15
            .PageFooter.RenderText.Style = footerstyle
            .PageSettings.Landscape = False
            .PageSettings.Margins.Bottom = 10
            .PageSettings.Margins.Left = 35
            .PageSettings.Margins.Right = 20
            .PageSettings.Margins.Top = 60
        End With
        With Me.doc2
            .DefaultUnit = UnitTypeEnum.Mm
            .DefaultUnitOfFrames = UnitTypeEnum.Mm
            .Style = docstyle
            .PageHeader.Height = 27
            .PageHeader.Style = headerstyle
            .PageFooter.Style.TextAlignHorz = AlignHorzEnum.Left
            .PageFooter.Height = 15
            .PageFooter.RenderText.Style = footerstyle
            .PageSettings.Landscape = False
            .PageSettings.Margins.Bottom = 10
            .PageSettings.Margins.Left = 35
            .PageSettings.Margins.Right = 20
            .PageSettings.Margins.Top = 60
        End With
    End Sub

    Private Sub DefineDocumentStyle(ByRef docstyle As C1DocStyle)
        'style for the document
        With docstyle
            .Font = New Font("Arial", 48, FontStyle.Regular)
        End With
    End Sub

    Private Sub DefineFooterStyle(ByRef footerstyle As C1DocStyle)
        'style for the footer
        With footerstyle
            .Borders.Top = New LineDef(Color.DarkGray, 0.25)
            .Font = New Font("Arial", 8, FontStyle.Regular)
        End With
    End Sub

    Private Sub DefineHeaderStyle(ByRef headerstyle As C1DocStyle)
        'style for the header
        With headerstyle
            .Borders.AllEmpty = True
            .Font = New Font("Arial", 10, FontStyle.Regular)
        End With
    End Sub

    Private Sub DefineLineStyle1(ByRef style As C1DocStyle)
        'style for the header
        With style
            '.Borders.AllEmpty = True
            ' .BackColor = Color.GhostWhite
            .Font = New Font("Arial", 9, FontStyle.Italic)
        End With
    End Sub

    Private Sub DefineLineStyle2(ByRef style As C1DocStyle)
        'style for the header
        With style
            '.Borders.AllEmpty = True
            .Font = New Font("Arial", 9, FontStyle.Regular)
        End With
    End Sub

    Private Sub DefineLineStyle3(ByRef style As C1DocStyle)
        'style for the header
        With style
            '.Borders.AllEmpty = True
            .TextAlignVert = AlignVertEnum.Center
            .TextAlignHorz = AlignHorzEnum.Right
            .ShapeLine.Color = Color.DarkGray
            .TextColor = Color.Gray
            .Font = New Font("Verdana", 6, FontStyle.Regular)
        End With
    End Sub

    Private Sub DefineLineStyle4(ByRef style As C1DocStyle)
        'style for the header
        With style
            .Font = New Font("Arial", 8, FontStyle.Bold)
            '.TextAlignHorz = AlignHorzEnum.Right
        End With
    End Sub

    Private Sub DefineLineStyle5(ByRef style As C1DocStyle)
        'style for the headerbox
        With style
            .BackColor = Color.GhostWhite
            .Font = New Font("Arial", 8, FontStyle.Bold)

        End With
    End Sub
    Private Sub DefineLineStyle6(ByRef style As C1DocStyle)
        With style
            .Font = New Font("Arial", 8, FontStyle.Bold)
            .TextAlignHorz = AlignHorzEnum.Right

        End With
    End Sub

    Private Sub DefineLineStyle7(ByRef style As C1DocStyle)

        With style
            .Borders.Top = New LineDef(Color.DarkGray, 0.25)
            .Borders.Bottom = New LineDef(Color.DarkGray, 0.25)
            .Font = New Font("Arial", 8, FontStyle.Bold)

        End With
    End Sub
    Private Sub DefineLineStyle8(ByRef style As C1DocStyle)

        With style
            .Font = New Font("Arial", 8, FontStyle.Underline)
        End With
    End Sub
    Private Sub DefineLineStyle9(ByRef style As C1DocStyle)

        With style
            .ShapeLine.Color = Color.DarkGray
            .Font = New Font("Arial", 8, FontStyle.Regular)

        End With
    End Sub



#End Region

#Region "  Context Menu Events "

    Private Sub mnuActualPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuActualPage.Click
        With Prev1.PreviewPane
            .ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        End With
    End Sub

    Private Sub mnuFullPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFullPage.Click
        With Prev1.PreviewPane
            .ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.FullPage
        End With
    End Sub

    Private Sub mnuZoomIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZoomIn.Click
        With Prev1.PreviewPane
            Dim factor As Single = Prev1.PreviewPane.ZoomFactor
            factor *= 2
            .ZoomFactor = factor
        End With
    End Sub

    Private Sub mnuZoomOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuZoomOut.Click
        With Prev1.PreviewPane
            Dim factor As Single = Prev1.PreviewPane.ZoomFactor
            factor /= 2
            .ZoomFactor = factor
        End With
    End Sub

#End Region

#Region "  Methods Generate "

    Private Sub GenerateVendorlist(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   datatable Vendors '''''''''
        ''    0        1           2          3         4       '
        '' transdate number,     name,    1099_sw    status     '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try

            'define styles
            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            HeaderVendorList()

            'If Me.a1099flag = True Then
            '    MsgBox("true")
            'End If
            'If Me.a1099flag = False Then
            '    MsgBox("false")
            'End If


            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(4)
            With rendertbl
                .Columns(0).Width = 18  'create date
                .Columns(1).Width = 22  'vend number
                .Columns(2).Width = 92  'vend name
                .Columns(3).Width = 25  '1099sw
                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2

            'Style for Col Headers
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 1).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)


            'Text for col headers
            band.Cell(currow, 0).RenderText.Text = "Created"
            band.Cell(currow, 1).RenderText.Text = "Number"
            band.Cell(currow, 2).RenderText.Text = "Vendor Name"
            band.Cell(currow, 3).RenderText.Text = "1099 Vendor"
            rendertbl.Body.Rows.AddSome(2)
            currow += 2

            'iterate thru the table & render the detail lines
            'Iterate Through Vendors table for matching expenditure codes
            Dim i As Int32

            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendcreate = DirectCast(.Rows(i)(0), Date)
                    vendnum = DirectCast(.Rows(i)(1), String)
                    vendname = DirectCast(.Rows(i)(2), String)
                    a1099sw = DirectCast(.Rows(i)(3), String)
                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With
                If Me.a1099flag = True And a1099sw = CStr("Yes") Then
                    band.Cell(currow, 0).RenderText.Text = vendcreate.ToString("MM/dd/yyyy")
                    band.Cell(currow, 1).RenderText.Text = vendnum
                    band.Cell(currow, 2).RenderText.Text = vendname
                    band.Cell(currow, 3).RenderText.Text = a1099sw
                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                End If
            Next
            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendcreate = DirectCast(.Rows(i)(0), Date)
                    vendnum = DirectCast(.Rows(i)(1), String)
                    vendname = DirectCast(.Rows(i)(2), String)
                    a1099sw = DirectCast(.Rows(i)(3), String)
                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With
                If Me.a1099flag = False Then
                    band.Cell(currow, 0).RenderText.Text = vendcreate.ToString("MM/dd/yyyy")
                    band.Cell(currow, 1).RenderText.Text = vendnum
                    band.Cell(currow, 2).RenderText.Text = vendname
                    band.Cell(currow, 3).RenderText.Text = a1099sw
                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                End If
            Next

            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateVendorAddresslist(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   datatable Vendors '''''''''
        ''    0        1           2          3         4       '
        '' transdate number,     name,    1099_sw    status     '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try

            'define styles
            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            HeaderVendorList()

          

            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(4)
            With rendertbl
                .Columns(0).Width = 18  'vend number
                .Columns(1).Width = 60    'vend name
                .Columns(2).Width = 60  'addr1,addr2
                .Columns(3).Width = 60  'city,state,zip



                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2

            'Style for Col Headers
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 1).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)


            'Text for col headers
            band.Cell(currow, 0).RenderText.Text = "Number"
            band.Cell(currow, 1).RenderText.Text = "Vendor Name"
            band.Cell(currow, 2).RenderText.Text = "Address"
            band.Cell(currow, 3).RenderText.Text = "City, State, Zip"
            rendertbl.Body.Rows.AddSome(2)
            currow += 2


            'Iterate Through Vendors table & render the detail lines
            Dim i As Int32
            '   0        1       2      3     4     5     6    7
            'number    name    adr1   adr2  city  state  zip zip

            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendnum = DirectCast(.Rows(i)(0), String)
                    vendname = DirectCast(.Rows(i)(1), String)
                    vendaddr1 = DirectCast(.Rows(i)(2), String)
                    vendaddr2 = DirectCast(.Rows(i)(3), String)
                    vendcity = DirectCast(.Rows(i)(4), String)
                    vendstate = DirectCast(.Rows(i)(5), String)
                    vendzip = DirectCast(.Rows(i)(6), String)

                End With

                band.Cell(currow, 0).RenderText.Text = vendnum
                band.Cell(currow, 1).RenderText.Text = vendname
                band.Cell(currow, 2).RenderText.Text = vendaddr1 & " " & vendaddr2
                band.Cell(currow, 3).RenderText.Text = vendcity & " " & vendstate & " " & vendzip
                rendertbl.Body.Rows.AddSome(1)
                currow += 1

            Next
          

            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateVendordetailed(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   Datatable Vendors  Detailed'''''''''''''''''
        ''    0        1           2          3         4        5         6     '
        '' date      fisyr,    chksnum,   afacctnum, asacctnum amount,   remarks '
        ''    7          8          9                                            '
        '' vendnum   vendname     1099sw                                         '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try

            'define styles

            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            HeaderVendorDetailed()

            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(8)
            With rendertbl
                .Columns(0).Width = 18  'date
                .Columns(1).Width = 15  'Year
                .Columns(2).Width = 18  'Checknum
                .Columns(3).Width = 18  'account
                .Columns(4).Width = 22  'Amount
                .Columns(5).Width = 42  'Remarks
                .Columns(6).Width = 18  'Vendnum
                .Columns(7).Width = 45  'Vendname
                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2

            'Style for Col Headers
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 1).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            band.Cell(currow, 5).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 6).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 7).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)



            'Text for col headers
            band.Cell(currow, 0).RenderText.Text = "Date"
            band.Cell(currow, 1).RenderText.Text = "Year"
            band.Cell(currow, 2).RenderText.Text = "Check #"
            band.Cell(currow, 3).RenderText.Text = "Account"
            band.Cell(currow, 4).RenderText.Text = "Amount"
            band.Cell(currow, 5).RenderText.Text = "Remarks"
            band.Cell(currow, 6).RenderText.Text = "Vend #"
            band.Cell(currow, 7).RenderText.Text = "Vendor"


            rendertbl.Body.Rows.AddSome(2)
            currow += 2

            'iterate thru the table & render the detail lines
            'Iterate Through Vendors table for matching expenditure codes
            Dim i As Int32
            Dim totvend1099 As Double
            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendchkdate = DirectCast(.Rows(i)(0), Date)
                    vendchkfisyr = DirectCast(.Rows(i)(1), Int32)
                    vendchknum = DirectCast(.Rows(i)(2), String)
                    afacct = DirectCast(.Rows(i)(3), String)
                    asacct = DirectCast(.Rows(i)(4), String)
                    vendchkamt = CDbl(.Rows(i)(5))
                    vendrmrks = DirectCast(.Rows(i)(6), String)
                    vendnum = DirectCast(.Rows(i)(7), String)
                    vendname = DirectCast(.Rows(i)(8), String)
                    a1099sw = DirectCast(.Rows(i)(9), String)

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With

                If Me.a1099flag = True And a1099sw = CStr("Yes") Then

                    Dim temptot As Double
                    band.Cell(currow, 0).RenderText.Text = vendchkdate.ToString("MM/dd/yyyy")
                    band.Cell(currow, 1).RenderText.Text = vendchkfisyr.ToString
                    band.Cell(currow, 2).RenderText.Text = vendchknum
                    band.Cell(currow, 3).RenderText.Text = afacct & "-" & asacct
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt)
                    band.Cell(currow, 5).RenderText.Text = vendrmrks
                    band.Cell(currow, 6).RenderText.Text = vendnum
                    band.Cell(currow, 7).RenderText.Text = vendname

                    temptot = CDbl(tbl.Rows(i)(5))
                    totvend1099 += temptot
                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                End If
            Next

            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendchkdate = DirectCast(.Rows(i)(0), Date)
                    vendchkfisyr = DirectCast(.Rows(i)(1), Int32)
                    vendchknum = DirectCast(.Rows(i)(2), String)
                    afacct = DirectCast(.Rows(i)(3), String)
                    asacct = DirectCast(.Rows(i)(4), String)
                    vendchkamt = CDbl(.Rows(i)(5))
                    vendrmrks = DirectCast(.Rows(i)(6), String)
                    vendnum = DirectCast(.Rows(i)(7), String)
                    vendname = DirectCast(.Rows(i)(8), String)
                    a1099sw = DirectCast(.Rows(i)(9), String)

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With

                If Me.a1099flag = False Then

                    band.Cell(currow, 0).RenderText.Text = vendchkdate.ToString("MM/dd/yyyy")
                    band.Cell(currow, 1).RenderText.Text = vendchkfisyr.ToString
                    band.Cell(currow, 2).RenderText.Text = vendchknum
                    band.Cell(currow, 3).RenderText.Text = afacct & "-" & asacct
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt)
                    band.Cell(currow, 5).RenderText.Text = vendrmrks
                    band.Cell(currow, 6).RenderText.Text = vendnum
                    band.Cell(currow, 7).RenderText.Text = vendname

                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                End If
            Next

            'summary - totals @ bottom of report
            rendertbl.Body.Rows.AddSome(3)
            currow += 1

            If Me.a1099flag = True Then
                rendertbl.Body.Rows.AddSome(3)
                currow += 1
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & totvend1099.ToString.Format("{0:F2}", totvend1099)

            End If
            If Me.a1099flag = False Then
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & Totvendchks.ToString.Format("{0:F2}", Totvendchks)
            End If




            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateVendorTransdetailed(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   Datatable Vendors  Detailed'''''''''''''''''
        ''    0        1           2          3         4        5         6     '
        '' date      fisyr,    chksnum,   afacctnum, asacctnum amount,   remarks '
        ''    7          8          9                                            '
        '' vendnum   vendname     1099sw                                         '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim a1099client As Boolean

        Try

            'define styles

            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            HeaderVendorTransDetailed()

            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(8)
            With rendertbl
                .Columns(0).Width = 20 '18  'date
                .Columns(1).Width = 17 '15  'Year
                .Columns(2).Width = 20 '18  'Checknum
                .Columns(3).Width = 20 '18  'account
                .Columns(4).Width = 24 '22  'Amount
                .Columns(5).Width = 95  'Remarks
                '.Columns(6).Width = 18  'Vendnum
                ' .Columns(7).Width = 45  'Vendname
                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2

            'Style for Col Headers
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 1).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            band.Cell(currow, 5).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            'Text for col headers
            band.Cell(currow, 0).RenderText.Text = "Date"
            band.Cell(currow, 1).RenderText.Text = "Year"
            band.Cell(currow, 2).RenderText.Text = "Check #"
            band.Cell(currow, 3).RenderText.Text = "Account"
            band.Cell(currow, 4).RenderText.Text = "Amount"
            band.Cell(currow, 5).RenderText.Text = "Remarks"

            rendertbl.Body.Rows.AddSome(2)
            currow += 2

            'iterate thru the table & render the detail lines
            'Iterate Through Vendors table for matching expenditure codes
            Dim i As Int32
            Dim j As Int32

            Dim totvend1099 As Double
            ' ________________________________________________________________
            'ALL VENDORS
          
            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    'For Title Lines
                    Dim vendornumber As String = DirectCast(.Rows(i)(7), String)
                    Dim Vendorname As String = DirectCast(.Rows(i)(8), String)
                    'band detail lines
                    vendchkdate = DirectCast(.Rows(i)(0), Date)
                    vendchkfisyr = DirectCast(.Rows(i)(1), Int32)
                    vendchknum = DirectCast(.Rows(i)(2), String)
                    afacct = DirectCast(.Rows(i)(3), String)
                    asacct = DirectCast(.Rows(i)(4), String)
                    vendchkamt = CDbl(.Rows(i)(5))
                    vendrmrks = DirectCast(.Rows(i)(6), String)
                    vendnum = DirectCast(.Rows(i)(7), String)
                    vendname = DirectCast(.Rows(i)(8), String)
                    a1099sw = DirectCast(.Rows(i)(9), String)
                    If a1099sw = "Y" Then
                        a1099client = True
                    Else
                        a1099client = False
                    End If

                    Dim tempamt As Double
                    tempamt += vendchkamt

                    currvendnum = vendnum
                    If vendnum = prevvendnum Then
                        currvendnum = ""
                    End If
                    If Me.a1099flag = False Then
                        If currvendnum = "" Then
                            rendertbl.Body.Rows.AddSome(2)
                            rcrds = 1

                        Else
                            currow += 1
                            rendertbl.Body.Rows.AddSome(1)

                            band.Cell(currow, 2).SpanColumns = 3
                            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

                            If a1099client = False Then
                                band.Cell(currow, 2).RenderText.Text = Vendorname & "    " & vendornumber
                            Else
                                band.Cell(currow, 2).RenderText.Text = Vendorname & "    " & vendornumber & "  (1099 Client)"

                            End If


                            currow += 1
                            rendertbl.Body.Rows.AddSome(1)
                        End If
                    End If

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With

                If Me.a1099flag = False Then
                    band.Cell(currow, 0).RenderText.Text = vendchkdate.ToString("MM/dd/yyyy")
                    band.Cell(currow, 1).RenderText.Text = vendchkfisyr.ToString
                    band.Cell(currow, 2).RenderText.Text = vendchknum
                    band.Cell(currow, 3).RenderText.Text = afacct & "-" & asacct
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt)
                    band.Cell(currow, 5).RenderText.Text = vendrmrks

                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                    prevvendnum = vendnum
                End If
            Next

            '_________________________________________________
            Dim x As Int32
            Dim rcrd As Int32
            Dim teststr As String
            Dim amt As Double
            Dim amtsub As Double

            'Total each vendors totals
            'Test for Blank lines or Zeros
            If Me.a1099flag = False Then
                x = rendertbl.Body.Rows.Count
                For x = 4 To rendertbl.Body.Rows.Count - 1
                    teststr = (band.Cell(x - 1, 4).RenderText.Text)
                    If teststr = "" Then
                        amt = 0
                        amtsub = 0
                    Else
                        amt = CDbl(band.Cell(x - 1, 4).RenderText.Text)
                        If amt > 0 Then
                            amtsub += amt
                        End If
                    End If
                    If band.Cell(x, 5).RenderText.Text = "" Then
                        rcrd += 1
                        If rcrd = 2 Then

                            'band.Cell(x, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
                            'band.Cell(x, 3).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                            'band.Cell(x, 3).RenderText.Text = "Sub-Tot"
                            band.Cell(x, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
                            band.Cell(x, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                            band.Cell(x, 4).RenderText.Text = "$" & amtsub.ToString.Format("{0:F2}", amtsub)


                        End If

                    Else

                        rcrd = 1
                    End If


                Next
            End If


            ''_________________________________________________
            'summary - totals @ bottom of report
            rendertbl.Body.Rows.AddSome(3)
            currow += 1

            If Me.a1099flag = True Then
                rendertbl.Body.Rows.AddSome(3)
                currow += 1
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & totvend1099.ToString.Format("{0:F2}", totvend1099)

            End If
            If Me.a1099flag = False Then
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & Totvendchks.ToString.Format("{0:F2}", Totvendchks)
            End If




            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateVendorTransdetailedOver600(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   Datatable Vendors  Detailed'''''''''''''''''
        ''    0        1           2          3         4        5         6     '
        '' date      fisyr,    chksnum,   afacctnum, asacctnum amount,   remarks '
        ''    7          8          9                                            '
        '' vendnum   vendname     1099sw                                         '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim a1099client As Boolean

        Try

            'define styles

            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            Me.HeaderVendorTransDetailedOver600()

            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(6)
            With rendertbl
                .Columns(0).Width = 15 '18  'blank
                .Columns(1).Width = 20 '15  'vend num
                .Columns(2).Width = 50 '18  '  "  name
                .Columns(3).Width = 20 '18  '1099 sw
                .Columns(4).Width = 30 '22  'Amount

                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2


            'Header line
            band.Cell(currow, 0).SpanColumns = 4
            band.Cell(currow, 0).RenderText.Style.Borders.Bottom = New LineDef(Color.DarkGray, 0.25)
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
            band.Cell(currow, 0).RenderText.Text = "Vendors Exceeding $600.00 in Transactions"

            rendertbl.Body.Rows.AddSome(2)
            currow += 2

            'Style for Col Headers

            band.Rows(currow).StyleTableCell.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right

            'Text for col headers
            ' band.Cell(currow, 0).RenderText.Text = "Date"
            band.Cell(currow, 1).RenderText.Text = "Vend #"
            band.Cell(currow, 2).RenderText.Text = "Vendor Name"
            band.Cell(currow, 3).RenderText.Text = "1099 Client"
            band.Cell(currow, 4).RenderText.Text = "Total"

            rendertbl.Body.Rows.AddSome(1)
            currow += 1

            'iterate thru the table & render the detail lines
            'Iterate Through Vendors table for matching expenditure codes
            Dim i As Int32
            Dim j As Int32

            Dim totvend1099 As Double
            ' ________________________________________________________________
            'ALL VENDORS


            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    'For Title Lines
                    vendnum = DirectCast(.Rows(i)(0), String)
                    vendname = DirectCast(.Rows(i)(1), String)
                    a1099sw = DirectCast(.Rows(i)(2), String)
                    vendchkamt = CDbl(.Rows(i)(3))

                 

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With

                If vendchkamt > 600 Then

                    '  band.Cell(currow, 0).RenderText.Text = 
                    band.Cell(currow, 1).RenderText.Text = vendnum
                    band.Cell(currow, 2).RenderText.Text = vendname
                    band.Cell(currow, 3).RenderText.Text = a1099sw
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt)

                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                    prevvendnum = vendnum
                End If
            Next

            '_________________________________________________
            'Dim x As Int32
            'Dim rcrd As Int32
            'Dim teststr As String
            'Dim amt As Double
            'Dim amtsub As Double

            ''Total each vendors totals
            ''Test for Blank lines or Zeros
            'If Me.a1099flag = False Then
            '    x = rendertbl.Body.Rows.Count
            '    For x = 4 To rendertbl.Body.Rows.Count - 1
            '        teststr = (band.Cell(x - 1, 4).RenderText.Text)
            '        If teststr = "" Then
            '            amt = 0
            '            amtsub = 0
            '        Else
            '            amt = CDbl(band.Cell(x - 1, 4).RenderText.Text)
            '            If amt > 0 Then
            '                amtsub += amt
            '            End If
            '        End If
            '        If band.Cell(x, 5).RenderText.Text = "" Then
            '            rcrd += 1
            '            If rcrd = 2 Then

            '                band.Cell(x, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            '                band.Cell(x, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            '                If amtsub > 600 Then
            '                    '  band.Cell(x, 4).SpanColumns = 2
            '                    band.Cell(x, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            '                    band.Cell(x, 4).RenderText.Style.TextColor = band.Cell(x, 4).RenderText.Style.TextColor.Blue
            '                    band.Cell(x, 4).RenderText.Text = amtsub.ToString.Format("{0:C2}", amtsub)
            '                    Dim excessamt As String
            '                    excessamt = CStr(amtsub - 600)

            '                    band.Cell(x, 0).SpanColumns = 4
            '                    band.Cell(x, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            '                    band.Cell(x, 0).RenderText.Style.TextColor = band.Cell(x, 4).RenderText.Style.TextColor.Blue
            '                    band.Cell(x, 0).RenderText.Text = "NOTE:    " & "Vendor Exceeds the $600 requirement for 1099 Client by  $" & excessamt & " In Transactions"



            '                Else

            '                    band.Cell(x, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            '                    band.Cell(x, 4).RenderText.Text = "$" & amtsub.ToString.Format("{0:F2}", amtsub)
            '                End If


            '            End If
            '        Else
            '            rcrd = 1
            '        End If
            ' Next
            '  End If


            ''_________________________________________________
            'summary - totals @ bottom of report
            'rendertbl.Body.Rows.AddSome(3)
            'currow += 1

            'If Me.a1099flag = True Then
            '    rendertbl.Body.Rows.AddSome(3)
            '    currow += 1
            '    band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
            '    band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
            '    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            '    band.Cell(currow, 3).RenderText.Text = "Total"
            '    band.Cell(currow, 4).RenderText.Text = "$" & totvend1099.ToString.Format("{0:F2}", totvend1099)

            'End If
            'If Me.a1099flag = False Then
            '    band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
            '    band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
            '    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            '    band.Cell(currow, 3).RenderText.Text = "Total"
            '    band.Cell(currow, 4).RenderText.Text = "$" & Totvendchks.ToString.Format("{0:F2}", Totvendchks)
            'End If




            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub GenerateVendorSpecific(ByVal tbl As DataTable)
        '''''''''''''''''''''''''''   Datatable Vendors  Detailed'''''''''''''''''
        ''    0        1           2          3         4        5         6     '
        '' date      fisyr,    chksnum,   afacctnum, asacctnum amount,   remarks '
        ''    7          8          9                                            '
        '' vendnum   vendname     1099sw                                         '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Try

            'define styles

            docstyle = New C1DocStyle(Doc1)
            footerstyle = New C1DocStyle(Doc1)
            headerstyle = New C1DocStyle(Doc1)
            linestyle1 = New C1DocStyle(Doc1)
            linestyle2 = New C1DocStyle(Doc1)
            linestyle3 = New C1DocStyle(Doc1)
            linestyle4 = New C1DocStyle(Doc1)
            linestyle5 = New C1DocStyle(Doc1)
            linestyle6 = New C1DocStyle(Doc1)
            linestyle7 = New C1DocStyle(Doc1)

            'define the document attributes using the styles
            DefineDocumentSettings()
            'open document for rendering
            If Doc1.IsGenerating = False Then
                Doc1.StartDoc()
            End If

            'define the document style
            DefineDocumentStyle(docstyle)
            'define the footer style
            DefineFooterStyle(footerstyle)
            'define the header style
            DefineHeaderStyle(headerstyle)
            'define a linestyle
            DefineLineStyle1(linestyle1)
            DefineLineStyle3(linestyle3)
            DefineLineStyle4(linestyle4)
            DefineLineStyle5(linestyle5)
            DefineLineStyle6(linestyle6)
            DefineLineStyle7(linestyle7)

            'was header

            '    ' define rendertable & rendertable band
            Dim rendertbl As New RenderTable(Doc1)

            rendertbl.BeginUpdate()

            Dim band As TableBand = rendertbl.Body


            'define the table format
            Dim row As DataRow
            rendertbl.Columns.AddSome(6)
            With rendertbl
                .Columns(0).Width = 18  'date
                .Columns(1).Width = 15  'Year
                .Columns(2).Width = 18  'Checknum
                .Columns(3).Width = 18  'account
                .Columns(4).Width = 22  'Amount
                .Columns(5).Width = 90  'Remarks

                .Style.Borders.AllEmpty = True
            End With
            ''set the style of the body
            With rendertbl.Body
                .StyleTableCell.BorderTableVert.Empty = True
                .StyleTableCell.BorderTableHorz.Empty = True
                .StyleTableCell.Font = New Font("Arial", 8, FontStyle.Regular)
            End With

            'set currow & allow for header space.
            Dim currow As Int32
            rendertbl.Body.Rows.AddSome(3)
            currow = 2

            'Style for Col Headers
            band.Cell(currow, 0).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 1).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 2).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)
            band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
            band.Cell(currow, 5).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Underline)

            'Text for col headers
            band.Cell(currow, 0).RenderText.Text = "Date"
            band.Cell(currow, 1).RenderText.Text = "Year"
            band.Cell(currow, 2).RenderText.Text = "Check #"
            band.Cell(currow, 3).RenderText.Text = "Account"
            band.Cell(currow, 4).RenderText.Text = "Amount"
            band.Cell(currow, 5).RenderText.Text = "Remarks"


            rendertbl.Body.Rows.AddSome(2)
            currow += 2

            'iterate thru the table & render the detail lines
            'Iterate Through Vendors table for matching expenditure codes
            Dim i As Int32
            Dim totvend1099 As Double
            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendchkdate = DirectCast(.Rows(i)(0), Date)
                    vendchkfisyr = DirectCast(.Rows(i)(1), Int32)
                    vendchknum = DirectCast(.Rows(i)(2), String)
                    afacct = DirectCast(.Rows(i)(3), String)
                    asacct = DirectCast(.Rows(i)(4), String)
                    vendchkamt = CDbl(.Rows(i)(5))
                    vendrmrks = DirectCast(.Rows(i)(6), String)
                    vendnum = DirectCast(.Rows(i)(7), String)
                    vendname = DirectCast(.Rows(i)(8), String)
                    a1099sw = DirectCast(.Rows(i)(9), String)

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With


                If Me.a1099flag = True And a1099sw = CStr("Yes") Then

                    Dim temptot As Double
                    band.Cell(currow, 0).RenderText.Text = vendchkdate.ToShortDateString
                    band.Cell(currow, 1).RenderText.Text = vendchkfisyr.ToString
                    band.Cell(currow, 2).RenderText.Text = vendchknum
                    band.Cell(currow, 3).RenderText.Text = afacct & "-" & asacct
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt.ToString)
                    band.Cell(currow, 5).RenderText.Text = vendrmrks


                    temptot = CDbl(tbl.Rows(i)(5))
                    totvend1099 += temptot
                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1


                End If
            Next

            For i = 0 To tbl.Rows.Count - 1
                With tbl
                    vendchkdate = DirectCast(.Rows(i)(0), Date)
                    vendchkfisyr = DirectCast(.Rows(i)(1), Int32)
                    vendchknum = DirectCast(.Rows(i)(2), String)
                    afacct = DirectCast(.Rows(i)(3), String)
                    asacct = DirectCast(.Rows(i)(4), String)
                    vendchkamt = CDbl(.Rows(i)(5))
                    vendrmrks = DirectCast(.Rows(i)(6), String)
                    vendnum = DirectCast(.Rows(i)(7), String)
                    vendname = DirectCast(.Rows(i)(8), String)
                    a1099sw = DirectCast(.Rows(i)(9), String)

                    Select Case a1099sw
                        Case "N"
                            a1099sw = CStr("No")
                        Case "Y"
                            a1099sw = CStr("Yes")
                    End Select
                End With

                If Me.a1099flag = False Then

                    band.Cell(currow, 0).RenderText.Text = vendchkdate.ToShortDateString
                    band.Cell(currow, 1).RenderText.Text = vendchkfisyr.ToString
                    band.Cell(currow, 2).RenderText.Text = vendchknum
                    band.Cell(currow, 3).RenderText.Text = afacct & "-" & asacct
                    band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                    band.Cell(currow, 4).RenderText.Text = vendchkamt.ToString.Format("{0:F2}", vendchkamt)
                    band.Cell(currow, 5).RenderText.Text = vendrmrks


                    rendertbl.Body.Rows.AddSome(1)
                    currow += 1
                End If
            Next
            HeaderVendorSpecific()
            'summary - totals @ bottom of report
            rendertbl.Body.Rows.AddSome(3)
            currow += 1

            If Me.a1099flag = True Then
                rendertbl.Body.Rows.AddSome(3)
                currow += 1
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & totvend1099.ToString.Format("{0:F2}", totvend1099)
                If totvend1099 = 0 Then
                    MsgBox("This Vendor Is Not Flagged As a 1099 Vendor...", MsgBoxStyle.Information, _msgtitle)
                    endgenerate = True
                    Exit Sub

                End If

            End If
            If Me.a1099flag = False Then
                band.Cell(currow, 3).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.Font = New Font("Arial", 8, FontStyle.Bold)
                band.Cell(currow, 4).RenderText.Style.TextAlignHorz = AlignHorzEnum.Right
                band.Cell(currow, 3).RenderText.Text = "Total"
                band.Cell(currow, 4).RenderText.Text = "$" & Totvendchks.ToString.Format("{0:F2}", Totvendchks)
            End If

            rendertbl.EndUpdate()
            Doc1.RenderBlock(rendertbl)
            Doc1.EndDoc()
            Prev1.PreviewPane.ZoomMode = C1.Win.C1PrintPreview.ZoomModeEnum.ActualSize
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

#End Region

#Region "  Methods Retrieval "

    Public Sub PrintVendorList(ByVal a1099flag As Boolean)

        'Print All Vendors regardless of transactions effected or not
        'User Has Option to Select only 1099 flagged vendors
        'collect the data and pass to the generate method
        Dim obj As New ClassVendors
        Dim tbl As DataTable

        If a1099flag = True Then
            Me.a1099flag = True
        Else
            Me.a1099flag = False
        End If

        Try
            tbl = DirectCast(obj.GetVendorlist, DataTable)
            If tbl.Rows.Count < 1 Then
                MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                Exit Sub
            End If
            vendlist = True
            ProgressBarUpdate(tbl)
            GenerateVendorlist(tbl)
            Doc1.EndDoc()
            Show()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            tbl.Dispose()
            vendlist = False

        End Try
    End Sub

    Public Sub PrintVendorAddressList(ByVal sortbynumber As Boolean, ByVal sortbyname As Boolean, ByVal sortbycity As Boolean)

        'Print All Vendors regardless of transactions effected or not
        'User Has Option to Select only 1099 flagged vendors
        'collect the data and pass to the generate method
        Dim obj As New ClassVendors
        Dim tbl As DataTable


        Me.a1099flag = False

        Try
            If sortbynumber = True Then
                tbl = DirectCast(obj.GetVendorAddresslist, DataTable)
                If tbl.Rows.Count < 1 Then
                    MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                    Exit Sub
                End If
            End If

            If sortbyname = True Then
                tbl = DirectCast(obj.GetVendorAddresslistbyname, DataTable)
                If tbl.Rows.Count < 1 Then
                    MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                    Exit Sub
                End If
            End If

            If sortbycity = True Then
                tbl = DirectCast(obj.GetVendorAddresslistbycity, DataTable)
                If tbl.Rows.Count < 1 Then
                    MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                    Exit Sub
                End If
            End If

            FlagAddress = True
            ProgressBarUpdate(tbl)
            GenerateVendorAddresslist(tbl)
            Doc1.EndDoc()
            Show()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            tbl.Dispose()

            FlagAddress = False


        End Try
    End Sub

    Public Sub PrintVendorsDetail(ByVal a1099flag As Boolean)
        'Print All Vendors with transactions effected (checks issued)
        'User Has Option to Select only 1099 flagged vendors
        'collect the data and pass to the generate method
        '
        Dim obj As New ClassVendors
        Dim tbl As DataTable

        If a1099flag = True Then
            Me.a1099flag = True
        Else
            Me.a1099flag = False
        End If

        Try
            tbl = DirectCast(obj.GetVendorsDetail(a1099flag), DataTable)
            If tbl.Rows.Count < 1 Then
                MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                Exit Sub
            End If
            Dim i As Int32

            Dim temptot As Double
            For i = 0 To tbl.Rows.Count - 1
                temptot = CDbl(tbl.Rows(i)(5))
                Totvendchks += temptot
            Next
            vendtransactionledger = True
            ProgressBarUpdate(tbl)
            GenerateVendordetailed(tbl)

            Doc1.EndDoc()
            Show()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            tbl.Dispose()
            vendtransactionledger = False

        End Try
    End Sub

    Public Sub PrintVendorsTransDetail(ByVal a1099flag As Boolean)
        'Print All Vendors with transactions effected (checks issued)
        'User Has Option to Select only 1099 flagged vendors
        'collect the data and pass to the generate method
        '
        Dim obj As New ClassVendors
        Dim tbl As DataTable

        Try
            tbl = DirectCast(obj.GetVendorsDetail(a1099flag), DataTable)
            If tbl.Rows.Count < 1 Then
                MsgBox("Database Contains No Vendors...", MsgBoxStyle.Information, _msgtitle)
                Exit Sub
            End If
            Dim i As Int32

            Dim temptot As Double
            For i = 0 To tbl.Rows.Count - 1
                temptot = CDbl(tbl.Rows(i)(5))
                Totvendchks += temptot
            Next
            vendtransdetail = True
            ProgressBarUpdate(tbl)
            GenerateVendorTransdetailed(tbl)

            Doc1.EndDoc()
            Show()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            tbl.Dispose()
            vendtransdetail = False


        End Try
    End Sub

    Public Sub PrintVendorsTransDetailOver600(ByVal begyear As String, ByVal endyear As String)

        'Print All Vendors with transactions effected (checks issued)
        'collect the data and pass to the generate method
        '
        Dim obj As New ClassVendors
        Dim tbl As DataTable

        Try
            calenderyear = begyear.Remove(4, 6)



            tbl = DirectCast(obj.GetVendorsDetail1099Over600(begyear, endyear), DataTable)
            If tbl.Rows.Count < 1 Then
                MsgBox("Database Contains No Vendors over $600 in transactions for this calender year...", MsgBoxStyle.Information, _msgtitle)
                Exit Sub
            End If
            Dim i As Int32

            Dim temptot As Double
            For i = 0 To tbl.Rows.Count - 1
                temptot = CDbl(tbl.Rows(i)(3))
                Totvendchks += temptot
            Next
            Me.flagover600 = True

            ProgressBarUpdate(tbl)
            GenerateVendorTransdetailedOver600(tbl)

            Doc1.EndDoc()
            Show()

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            flagover600 = False
            tbl.Dispose()



        End Try
    End Sub

    Public Sub PrintVendorSpecificDetail(ByVal vendnum As String, ByVal a1099flag As Boolean)
        'Print Specific Vendor Information with transactions effected (checks issued)
        'User Has Option to Select only 1099 flagged vendors
        'collect the data and pass to the generate method
        Dim obj As New ClassVendors
        Dim tbl As DataTable

        If a1099flag = True Then
            Me.a1099flag = True
        Else
            Me.a1099flag = False
        End If

        Try
            tbl = DirectCast(obj.GetVendorsSpecificDetail(vendnum), DataTable)
            If tbl.Rows.Count < 1 Then
                MsgBox("Database Contains No Record OR Transactions for Vendor Number: " & vendnum, MsgBoxStyle.Information, _msgtitle)
                Exit Sub
            End If
            Dim i As Int32

            Dim temptot As Double
            For i = 0 To tbl.Rows.Count - 1
                temptot = CDbl(tbl.Rows(i)(5))
                Totvendchks += temptot
            Next
            vendspecific = True
            ProgressBarUpdate(tbl)
            GenerateVendorSpecific(tbl)

            Doc1.EndDoc()
            If endgenerate = True Then
            Else
                Show()
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)

            Throw
        Finally
            obj.Dispose()
            tbl.Dispose()
            vendspecific = True

        End Try
    End Sub

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

#Region "  Report Headers "


    Private Sub HeaderVendorList()

        'For 1st Page of Report
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle3 = New C1DocStyle(Doc1)
        DefineLineStyle3(linestyle3)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)

        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)

        With Doc1
            If Doc1.CurrentPage = 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                'Report Title
                If FlagAddress = True Then
                    .RenderDirectText(75, 1, "Activity Fund - Vendor Address Details", 95, 5, headerstyle)
                Else

                    .RenderDirectText(80, 1, "Activity Fund - Vendor Listing", 95, 5, headerstyle)
                End If

                If Me.a1099flag = True Then
                    .RenderDirectText(22, 6, "Lists Only 1099 Vendors", 95, 5, linestyle3)
                End If

                'Todays date on Header Info
                .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
                'Double line below header
                Doc1.RenderDirectLine(200, 20, 0, 20)
                Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            End If
            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With
    End Sub

    Private Sub HeaderVendorDetailed()

        'For 1st Page of Report
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle3 = New C1DocStyle(Doc1)
        DefineLineStyle3(linestyle3)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)

        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)

        With Doc1
            If Doc1.CurrentPage = 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                'Report Title

                .RenderDirectText(80, 1, "Activity Fund -  Vendor Transaction Ledger", 95, 5, headerstyle)
                If Me.a1099flag = True Then
                    .RenderDirectText(22, 6, "Lists Only 1099 Vendors", 95, 5, linestyle3)
                End If

                'Todays date on Header Info
                .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
                'Double line below header
                Doc1.RenderDirectLine(200, 20, 0, 20)
                Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            End If
            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With
    End Sub

    Private Sub HeaderVendorTransDetailed()

        'For 1st Page of Report
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle3 = New C1DocStyle(Doc1)
        DefineLineStyle3(linestyle3)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)

        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)

        With Doc1
            If Doc1.CurrentPage = 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                'Report Title

                .RenderDirectText(70, 1, "Activity Fund -  Vendor Detailed Transaction Ledger (All Vendors)", 150, 5, headerstyle)
                If Me.a1099flag = True Then
                    .RenderDirectText(22, 6, "Lists Only 1099 Vendors", 95, 5, linestyle3)
                End If

                'Todays date on Header Info
                .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
                'Double line below header
                Doc1.RenderDirectLine(200, 20, 0, 20)
                Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            End If
            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With
    End Sub

    Private Sub HeaderVendorTransDetailedOver600()

        'For 1st Page of Report
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle3 = New C1DocStyle(Doc1)
        DefineLineStyle3(linestyle3)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)

        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)

        With Doc1
            If Doc1.CurrentPage = 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                'Report Title

                .RenderDirectText(70, 1, "Activity Fund -  Vendors With Transactions Over $600.00", 150, 5, headerstyle)

                'Todays date on Header Info
                .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
                'Double line below header
                Doc1.RenderDirectLine(200, 20, 0, 20)
                Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            End If
            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With
    End Sub
    Private Sub HeaderVendorSpecific()

        'For 1st Page of Report
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle3 = New C1DocStyle(Doc1)
        DefineLineStyle3(linestyle3)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)

        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)

        With Doc1
            If Doc1.CurrentPage = 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                'Report Title

                .RenderDirectText(80, 1, "Activity Fund -  Selected Vendor Ledger", 95, 5, headerstyle)

                .RenderDirectText(90, 15, "For: " & vendname & " " & "     " & vendnum, 110, 5, linestyle4)


                If Me.a1099flag = True Then
                    .RenderDirectText(22, 6, "Lists Only 1099 Vendors", 95, 5, linestyle3)
                End If

                'Todays date on Header Info
                .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
                'Double line below header
                Doc1.RenderDirectLine(200, 20, 0, 20)
                Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            End If
            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With
    End Sub
#End Region

#Region "  New Page Events"

    Private Sub Doc1_NewPageStarted(ByVal sender As C1.C1PrintDocument.C1PrintDocument, ByVal e As C1.C1PrintDocument.NewPageStartedEventArgs) Handles Doc1.NewPageStarted

        Dim Page As Int32
        'These Details are for pages Greater than 1
        headerstyle = New C1DocStyle(Doc1)
        DefineHeaderStyle(headerstyle)
        footerstyle = New C1DocStyle(Doc1)
        DefineFooterStyle(footerstyle)
        linestyle4 = New C1DocStyle(Doc1)
        DefineLineStyle4(linestyle4)
        linestyle9 = New C1DocStyle(Doc1)
        DefineLineStyle9(linestyle9)
        linestyle8 = New C1DocStyle(Doc1)
        DefineLineStyle8(linestyle8)


        With Doc1


            If .CurrentPage > 1 Then
                Dim socialaddress As String
                socialaddress = Me._schoolcity & ", " & Me._schoolstate & " " & Me._schoolzip & " " & Me._schoolzipext

                .RenderDirectText(2, 1, Me._schoolname, 70, 5, headerstyle)
                .RenderDirectText(2, 6, Me._schooladdress1, 70, 5, headerstyle)
                .RenderDirectText(2, 11, socialaddress, 70, 5, headerstyle)
                .RenderDirectText(2, 16, Me._schooltelephone, 80, 5, headerstyle)

                If Me.flagover600 = True Then
                    .RenderDirectText(92, 7, "Lists All Possible 1099 Vendors", 95, 5, linestyle8)
                End If

                If Me.a1099flag = True Then
                    .RenderDirectText(92, 7, "Lists Only 1099 Vendors", 95, 5, linestyle8)
                End If

                'New page titles for respective reports
                If vendlist = True Then

                    .RenderDirectText(80, 1, "Vendor Listing", 95, 5, headerstyle)
                    .RenderDirectText(2, 21, "Created", 70, 5, linestyle8)
                    .RenderDirectText(20, 21, "Number", 70, 5, linestyle8)
                    .RenderDirectText(41, 21, "Vendor Name", 70, 5, linestyle8)
                    .RenderDirectText(133, 21, "1099 Vendor", 70, 5, linestyle8)
                End If

                If FlagAddress = True Then
                    .RenderDirectText(80, 1, "Vendor Address Details", 95, 5, headerstyle)
                    .RenderDirectText(2, 21, "Number", 70, 5, linestyle8)
                    .RenderDirectText(20, 21, "Vendor Name", 70, 5, linestyle8)
                    .RenderDirectText(82, 21, "Address", 70, 5, linestyle8)
                    .RenderDirectText(145, 21, "City, State, Zip", 70, 5, linestyle8)

                End If

                If vendtransactionledger = True Then
                    .RenderDirectText(80, 1, "Activity Fund -  Vendor Transaction Ledger", 95, 5, headerstyle)
                    'Text for col headers
                    .RenderDirectText(2, 21, "Date", 70, 5, linestyle8)
                    .RenderDirectText(19, 21, "Year", 70, 5, linestyle8)
                    .RenderDirectText(35, 21, "Check #", 70, 5, linestyle8)
                    .RenderDirectText(52, 21, "Account", 70, 5, linestyle8)
                    .RenderDirectText(79, 21, "Amount", 70, 5, linestyle8)
                    .RenderDirectText(92, 21, "Remarks", 70, 5, linestyle8)
                    .RenderDirectText(134, 21, "Vend #", 70, 5, linestyle8)
                    .RenderDirectText(152, 21, "Vendor", 70, 5, linestyle8)


                End If
                If vendtransdetail = True Then
                    .RenderDirectText(70, 1, "Activity Fund -  Vendor Transaction Ledger (All Vendors)", 120, 5, headerstyle)
                    'Text for col headers
                    .RenderDirectText(2, 21, "Date", 70, 5, linestyle8)
                    .RenderDirectText(21, 21, "Year", 70, 5, linestyle8)
                    .RenderDirectText(38.5, 21, "Check #", 70, 5, linestyle8)
                    .RenderDirectText(58.5, 21, "Account", 70, 5, linestyle8)
                    .RenderDirectText(89.5, 21, "Amount", 70, 5, linestyle8)
                    .RenderDirectText(102.5, 21, "Remarks", 70, 5, linestyle8)
                End If

                If vendspecific = True Then
                    .RenderDirectText(80, 1, "Activity Fund -  Selected Vendor Ledger", 95, 5, headerstyle)
                    'Text for col headers
                    .RenderDirectText(2, 21, "Date", 70, 5, linestyle8)
                    .RenderDirectText(19, 21, "Year", 70, 5, linestyle8)
                    .RenderDirectText(35, 21, "Check #", 70, 5, linestyle8)
                    .RenderDirectText(52, 21, "Account", 70, 5, linestyle8)
                    .RenderDirectText(79, 21, "Amount", 70, 5, linestyle8)
                    .RenderDirectText(92, 21, "Remarks", 70, 5, linestyle8)
                End If



            End If

            'Dim year As Int32

            'Report Year
            If flagover600 = True Then
                .RenderDirectText(92, 11, "Calender Year: " & calenderyear, 40, 7, linestyle4)
            Else

                .RenderDirectText(96, 11, "Year: " & Me.fiscalyear, 15, 7, linestyle4)
            End If
            'Report Year



            'Todays date on Header Info
            .RenderDirectText(164, 7, Date.Now.ToString("MMMM, d  yyyy"), 60, 7, linestyle4)
            'Double line below header
            Doc1.RenderDirectLine(200, 20, 0, 20)
            Doc1.RenderDirectLine(200, 20.5, 0, 20.5)

            'FOOTER
            'Footer Note - Page #
            .RenderDirectText(0, 255, "Page [@@PageNo@@] of [@@PageCount@@]", 95, 5, footerstyle)

            'Footer Note - Comments 
            .RenderDirectText(80, 255, "Private and Confidential", 75, 5, footerstyle)
            .RenderDirectText(121, 255, "Activity Fund.Net  Product of ADPC Inc 1-800-747-2372", 83, 5, footerstyle)

        End With

    End Sub


#End Region

#Region "  Menu Events"

    Private Sub MenuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose()
    End Sub

#End Region

#Region "  Progress Bar Events"

    Private Sub ProgressBarUpdate(ByVal tbl As DataTable)
        'Statusbar update on mainmenu...

        Dim x As Int32 = tbl.Rows.Count()
        Dim i As Int32

        recordstotal = x
        For i = 0 To recordstotal - 1
            If i Mod 1 = 0 Then
                RaiseEvent RecordStatus(i, recordstotal)
                System.Threading.Thread.Sleep(1)
            End If
        Next i
    End Sub

#Region "  Events & Delegates"

    Public Delegate Sub RecordCountEvents()
    Public Event RecordStatus(ByVal Records As Int32, ByVal recordstotal As Int32)
    Public recordstotal As Int32
#End Region

#End Region




End Class


