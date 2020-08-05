Public Class DoubleBuferGrid
    Inherits DataGridView

    Public Sub New()

        'Me.DoubleBuffered = True


        Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
        Me.UpdateStyles()

    End Sub

    'Protected Overridable Overloads ReadOnly Property DoubleBuffered As Boolean
    '    Get
    '        Return True
    '    End Get
    'End Property



End Class
