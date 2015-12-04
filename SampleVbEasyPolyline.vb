Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("dfe1df8e-3851-4091-b733-d8fa1573fbac")> _
  Public Class SampleVbEasyPolyline
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbEasyPolyline"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim pline As New Polyline
      Dim rc As Result = RhinoGet.GetPolyline(pline)
      If rc = Result.Success Then
        doc.Objects.AddPolyline(pline)
        doc.Views.Redraw()
      End If

      Return Result.Success

    End Function
  End Class

End Namespace