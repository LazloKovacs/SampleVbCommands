Imports System
Imports System.Runtime.InteropServices
Imports Rhino.DocObjects
Imports Rhino
Imports Rhino.Commands
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("f9686b82-c4eb-4f9b-9f66-85ceb6c24600")> _
  Public Class SampleVbScreenPoint
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbScreenPoint"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim gp As New GetPoint()
      gp.Get()
      If (gp.CommandResult() <> Result.Success) Then
        Return gp.Result()
      End If

      Dim view = gp.View()
      If view Is Nothing Then
        Return Result.Failure
      End If

      Dim xform = view.ActiveViewport.GetTransform(CoordinateSystem.World, CoordinateSystem.Screen)

      Dim worldPt = gp.Point()

      Dim screenPt = worldPt
      screenPt.Transform(xform)

      Dim clientPoint As New System.Drawing.Point()
      clientPoint.X = Int(screenPt.X)
      clientPoint.Y = Int(screenPt.Y)

      Dim screenPoint = clientPoint
      ClientToScreen(view.Handle, screenPoint)

      RhinoApp.WriteLine(String.Format("Point in world coordinates: {0}", worldPt.ToString()))
      RhinoApp.WriteLine(String.Format("Point in screen coordinates: {0}", screenPoint.ToString()))
      RhinoApp.WriteLine(String.Format("Point in client coordinates: {0}", clientPoint.ToString()))

      Return Result.Success

    End Function

    <DllImport("user32.dll")> _
    Private Shared Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As System.Drawing.Point) As Boolean
    End Function

  End Class

End Namespace