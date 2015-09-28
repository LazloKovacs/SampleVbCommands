Imports Rhino
Imports Rhino.Commands
Imports Rhino.DocObjects
Imports Rhino.Geometry
Imports Rhino.Input.Custom

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("1a848bc2-b286-4bc4-8d64-3c7dc9732e50")> _
  Public Class SampleVbSweep1
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbSweep1"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim gr = New GetObject()
      gr.SetCommandPrompt("Select rail")
      gr.GeometryFilter = ObjectType.Curve
      gr.SubObjectSelect = False
      gr.Get()
      If gr.CommandResult() <> Result.Success Then
        Return gr.CommandResult()
      End If

      Dim railCurve As Curve = gr.Object(0).Curve()
      If railCurve Is Nothing Then
        Return Result.Failure
      End If

      Dim gs = New GetObject()
      gs.SetCommandPrompt("Select cross section curve")
      gs.GeometryFilter = ObjectType.Curve
      gs.SubObjectSelect = False
      gs.EnablePreSelect(False, True)
      gs.DeselectAllBeforePostSelect = False
      gs.Get()
      If gs.CommandResult() <> Result.Success Then
        Return gs.CommandResult()
      End If

      Dim shapeCurve As Curve = gs.Object(0).Curve()
      If shapeCurve Is Nothing Then
        Return Result.Failure
      End If

      Dim breps = Brep.CreateFromSweep(railCurve, shapeCurve, False, doc.ModelAbsoluteTolerance)
      For i As Integer = 0 To breps.Length - 1
        doc.Objects.AddBrep(breps(i))
      Next
      doc.Views.Redraw()

      Return Result.Success

    End Function

  End Class

End Namespace