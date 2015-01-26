Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("c6701745-21d5-4a87-9058-b527f5bf1c13")> _
  Public Class SampleVbLine
    Inherits Command

    Shared _instance As SampleVbLine

    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a refence in a static field.
      _instance = Me
    End Sub

    '''<summary>The only instance of this command.</summary>
    Public Shared ReadOnly Property Instance() As SampleVbLine
      Get
        Return _instance
      End Get
    End Property

    '''<returns>The command name as it appears on the Rhino command line.</returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbLine"
      End Get
    End Property

    ''' <summary>Runs the command.</summary>
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