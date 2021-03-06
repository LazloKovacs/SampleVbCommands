﻿Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("a1ca97f5-f39d-49cc-aff4-808a701d2e33")> _
  Public Class SampleVbConstructionPlane
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbConstructionPlane"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      ' Get the active view
      Dim view As Rhino.Display.RhinoView = doc.Views.ActiveView
      If view Is Nothing Then
        Return Rhino.Commands.Result.Failure
      End If

      ' Get the view's construction plane
      Dim cplane As New Rhino.DocObjects.ConstructionPlane()
      cplane = view.ActiveViewport.GetConstructionPlane()

      ' Make a copy of the the construction plane's plane
      Dim plane As New Rhino.Geometry.Plane(cplane.Plane)

      ' Nudge it in the z-axis direction a bit
      plane.Origin += New Rhino.Geometry.Vector3d(0, 0, 5)

      ' Update the construction plane's plane with our modified version
      cplane.Plane = plane

      ' Update the active view's construction plane
      view.ActiveViewport.SetConstructionPlane(cplane)
      view.Redraw()

      Return Result.Success

    End Function
  End Class
End Namespace