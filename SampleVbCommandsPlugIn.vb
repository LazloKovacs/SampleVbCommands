Namespace SampleVbCommands
 
  Public Class SampleVbCommandsPlugIn
    Inherits Rhino.PlugIns.PlugIn

    Shared _instance As SampleVbCommandsPlugIn

    Public Sub New()
      _instance = Me
    End Sub

    Public Shared ReadOnly Property Instance() As SampleVbCommandsPlugIn
      Get
        Return _instance
      End Get
    End Property

  End Class
End Namespace