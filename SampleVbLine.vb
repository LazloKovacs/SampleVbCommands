Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("c6701745-21d5-4a87-9058-b527f5bf1c13")> _
  Public Class SampleVbLine
    Inherits Command

    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbLine"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim gp As New GetPoint()
      gp.SetCommandPrompt("Start of line")
      gp.Get()
      If (gp.CommandResult() <> Result.Success) Then
        Return gp.CommandResult()
      End If

      Dim startPt As Point3d = gp.Point()

      gp.SetCommandPrompt("End of line")
      gp.SetBasePoint(startPt, True)
      gp.DrawLineFromPoint(startPt, True)
      gp.Get()
      If (gp.CommandResult() <> Result.Success) Then
        Return gp.CommandResult()
      End If

      Dim endPt As Point3d = gp.Point()

      If (startPt.DistanceTo(endPt) < doc.ModelAbsoluteTolerance) Then
        Return Result.Nothing
      End If

      doc.Objects.AddLine(startPt, endPt)
      doc.Views.Redraw()

      Return Result.Success

    End Function

  End Class

End Namespace