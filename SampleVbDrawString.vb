Imports System
Imports System.Drawing
Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("f183bd79-07de-4248-b4c6-1984d93d8064")> _
  Public Class SampleVbDrawString
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbDrawString"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim basePt As New Point3d
      Dim rc As Result = RhinoGet.GetPoint("Start of line", False, basePt)
      If (rc <> Result.Success) Then
        Return rc
      End If

      Dim gp As New GetPoint()
      gp.SetCommandPrompt("End of line")
      gp.SetBasePoint(basePt, True)
      gp.DrawLineFromPoint(basePt, True)
      AddHandler gp.DynamicDraw, AddressOf gp_DynamicDraw
      gp.Get()
      If (gp.CommandResult() <> Result.Success) Then
        Return gp.CommandResult()
      End If

      Dim endPt As Point3d = gp.Point()
      Dim vec As Vector3d = endPt - basePt
      If Not vec.IsTiny() Then
        Dim line As New Line(basePt, endPt)
        doc.Objects.AddLine(line)
        doc.Views.Redraw()
      End If

      Return Result.Success

    End Function

    Private Sub gp_DynamicDraw(sender As Object, e As Rhino.Input.Custom.GetPointDrawEventArgs)

      Dim basePt As New Point3d
      If (e.Source.TryGetBasePoint(basePt)) Then

        ' Format distance as string
        Dim distance As Double = basePt.DistanceTo(e.CurrentPoint)
        Dim text As String = String.Format("{0:0.000}", distance)

        ' Get world-to-screen coordinate transformation
        Dim xform As Transform = e.Viewport.GetTransform(Rhino.DocObjects.CoordinateSystem.World, Rhino.DocObjects.CoordinateSystem.Screen)

        ' Transform point from world to screen coordinates
        Dim screenPt As Point3d = xform * e.CurrentPoint

        ' Offset point so text does not overlap cursor
        screenPt.X += 5.0
        screenPt.Y -= 5.0

        ' Draw the string 
        e.Display.Draw2dText(text, Color.Black, New Point2d(screenPt.X, screenPt.Y), False)

      End If

    End Sub


  End Class

End Namespace