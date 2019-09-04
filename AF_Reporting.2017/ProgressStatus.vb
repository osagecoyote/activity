Public Class ProgressStatus
    Inherits StatusBar

#Region "  Class Members "

    Public ProgressBar As New ProgressBar
    Private _panel As Int32 = -1

#End Region

#Region "  Initialisation & Disposal  "

    Sub New()
        Me.ProgressBar.Hide()
        Me.Controls.Add(Me.ProgressBar)
    End Sub

#End Region

#Region "  Methods "

    Private Sub Reposition(ByVal sender As Object, ByVal sbdevent As System.Windows.Forms.StatusBarDrawItemEventArgs) Handles MyBase.DrawItem
        ProgressBar.Location = New Point(sbdevent.Bounds.X, sbdevent.Bounds.Y)
        ProgressBar.Size = New Size(sbdevent.Bounds.Width, sbdevent.Bounds.Height)
        ProgressBar.Show()
    End Sub

#End Region

#Region "  Properties "

    Public Property Panel() As Int32
        'set the position of the progress bar
        Get
            Return _panel
        End Get
        Set(ByVal Value As Int32)
            _panel = Value
            Me.Panels(Value).Style = StatusBarPanelStyle.OwnerDraw
        End Set
    End Property

#End Region

End Class
