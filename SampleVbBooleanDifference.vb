Imports System
Imports System.Collections.Generic
Imports Rhino
Imports Rhino.Commands
Imports Rhino.DocObjects
Imports Rhino.Geometry

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("49c533f1-58bc-495f-b11a-f879750a8b0b")> _
  Public Class SampleVbBooleanDifference
    Inherits Command

    Shared _instance As SampleVbBooleanDifference

    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a refence in a static field.
      _instance = Me
    End Sub

    '''<summary>The only instance of this command.</summary>
    Public Shared ReadOnly Property Instance() As SampleVbBooleanDifference
      Get
        Return _instance
      End Get
    End Property

    '''<returns>The command name as it appears on the Rhino command line.</returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbBooleanDifference"
      End Get
    End Property

    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim rc As Rhino.Commands.Result
      Dim objrefs As Rhino.DocObjects.ObjRef() = Nothing
      rc = Rhino.Input.RhinoGet.GetMultipleObjects("Select first set of polysurfaces", False, Rhino.DocObjects.ObjectType.PolysrfFilter, objrefs)
      If rc <> Rhino.Commands.Result.Success Then
        Return rc
      End If
      If objrefs Is Nothing OrElse objrefs.Length < 1 Then
        Return Rhino.Commands.Result.Failure
      End If

      Dim in_breps0 As New List(Of Rhino.Geometry.Brep)()
      For i As Integer = 0 To objrefs.Length - 1
        Dim brep As Rhino.Geometry.Brep = objrefs(i).Brep()
        If brep IsNot Nothing Then
          in_breps0.Add(brep)
        End If
      Next

      doc.Objects.UnselectAll()
      rc = Rhino.Input.RhinoGet.GetMultipleObjects("Select second set of polysurfaces", False, Rhino.DocObjects.ObjectType.PolysrfFilter, objrefs)
      If rc <> Rhino.Commands.Result.Success Then
        Return rc
      End If
      If objrefs Is Nothing OrElse objrefs.Length < 1 Then
        Return Rhino.Commands.Result.Failure
      End If

      Dim in_breps1 As New List(Of Rhino.Geometry.Brep)()
      For i As Integer = 0 To objrefs.Length - 1
        Dim brep As Rhino.Geometry.Brep = objrefs(i).Brep()
        If brep IsNot Nothing Then
          in_breps1.Add(brep)
        End If
      Next

      Dim tolerance As Double = doc.ModelAbsoluteTolerance
      Dim breps As Rhino.Geometry.Brep() = Rhino.Geometry.Brep.CreateBooleanDifference(in_breps0, in_breps1, tolerance)
      If breps.Length < 1 Then
        Return Rhino.Commands.Result.[Nothing]
      End If
      For i As Integer = 0 To breps.Length - 1
        doc.Objects.AddBrep(breps(i))
      Next
      doc.Views.Redraw()

      Return Result.Success

    End Function
  End Class
End Namespace