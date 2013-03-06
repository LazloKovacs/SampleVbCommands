Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("d283a54a-728b-40c5-a079-71609852e015")> _
  Public Class SampleVbExtrusion
    Inherits Command

    Shared _instance As SampleVbExtrusion

    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a refence in a static field.
      _instance = Me
    End Sub

    '''<summary>The only instance of this command.</summary>
    Public Shared ReadOnly Property Instance() As SampleVbExtrusion
      Get
        Return _instance
      End Get
    End Property

    '''<returns>The command name as it appears on the Rhino command line.</returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbExtrusion"
      End Get
    End Property

    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim dir0 As New Rhino.Geometry.Vector3d(0, 0, 5)

      Dim pts0 As New List(Of Rhino.Geometry.Point3d)
      pts0.Add(New Rhino.Geometry.Point3d(0, 0, 0))
      pts0.Add(New Rhino.Geometry.Point3d(25, 0, 0))
      pts0.Add(New Rhino.Geometry.Point3d(25, 25, 0))
      pts0.Add(New Rhino.Geometry.Point3d(0, 25, 0))
      pts0.Add(New Rhino.Geometry.Point3d(0, 0, 0))

      Dim pline0 As New Rhino.Geometry.PolylineCurve(pts0)
      Dim srf0 As Rhino.Geometry.Surface = Rhino.Geometry.Surface.CreateExtrusion(pline0, dir0)

      Dim temp0 As Rhino.Geometry.Brep = Rhino.Geometry.Brep.CreateFromSurface(srf0)
      Dim brep0 As Rhino.Geometry.Brep = temp0.CapPlanarHoles(doc.ModelAbsoluteTolerance)
      Dim id0 As System.Guid = doc.Objects.AddBrep(brep0)
      Dim obj0 As Rhino.DocObjects.RhinoObject = doc.Objects.Find(id0)
      brep0 = obj0.Geometry


      Dim dir1 As New Rhino.Geometry.Vector3d(0, 0, 7)

      Dim pts1 As New List(Of Rhino.Geometry.Point3d)
      pts1.Add(New Rhino.Geometry.Point3d(3, 2, -1))
      pts1.Add(New Rhino.Geometry.Point3d(22, 2, -1))
      pts1.Add(New Rhino.Geometry.Point3d(22, 23, -1))
      pts1.Add(New Rhino.Geometry.Point3d(3, 23, -1))
      pts1.Add(New Rhino.Geometry.Point3d(3, 2, -1))

      Dim pline1 As New Rhino.Geometry.PolylineCurve(pts1)
      Dim srf1 As Rhino.Geometry.Surface = Rhino.Geometry.Surface.CreateExtrusion(pline1, dir1)

      Dim temp1 As Rhino.Geometry.Brep = Rhino.Geometry.Brep.CreateFromSurface(srf1)
      temp1.CapPlanarHoles(doc.ModelAbsoluteTolerance)
      Dim brep1 As Rhino.Geometry.Brep = temp1.CapPlanarHoles(doc.ModelAbsoluteTolerance)
      Dim id1 As System.Guid = doc.Objects.AddBrep(brep1)
      Dim obj1 As Rhino.DocObjects.RhinoObject = doc.Objects.Find(id1)
      brep1 = obj1.Geometry

      Dim results As Rhino.Geometry.Brep() = Rhino.Geometry.Brep.CreateBooleanDifference(brep0, brep1, doc.ModelAbsoluteTolerance)
      If (results IsNot Nothing) Then
        For i As Integer = 0 To results.Length - 1
          doc.Objects.AddBrep(results(i))
        Next
        doc.Views.Redraw()
      End If

      Return Result.Success

    End Function
  End Class
End Namespace