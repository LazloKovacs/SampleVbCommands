Imports System.Globalization
Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("a0559241-fdf1-43e0-8fe7-1caa93074d41")> _
  Public Class SampleVbGetPoint
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbGetPoint"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim gp = New GetPoint()
      gp.SetCommandPrompt("Pick a point")
      gp.Get()
      If (gp.CommandResult() <> Result.Success) Then
        Return gp.CommandResult()
      End If

      Dim point = gp.Point()

      doc.Objects.AddPoint(point)
      doc.Views.Redraw()

      Dim format = String.Format("F{0}", doc.DistanceDisplayPrecision)
      Dim provider = CultureInfo.InvariantCulture

      Dim x = point.X.ToString(format, provider)
      Dim y = point.Y.ToString(format, provider)
      Dim z = point.Z.ToString(format, provider)
      RhinoApp.WriteLine("World coordinates: {0},{1},{2}", x, y, z)

      Dim view = gp.View()
      If view IsNot Nothing Then
        Dim plane = view.ActiveViewport.ConstructionPlane()
        Dim xform = Transform.ChangeBasis(plane.WorldXY, plane)

        point.Transform(xform)

        x = point.X.ToString(format, provider)
        y = point.Y.ToString(format, provider)
        z = point.Z.ToString(format, provider)
        RhinoApp.WriteLine("CPlane coordinates: {0},{1},{2}", x, y, z)
      End If

      Return Result.Success

    End Function
  End Class
End Namespace