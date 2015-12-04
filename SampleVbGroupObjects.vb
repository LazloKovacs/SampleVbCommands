Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("94796de9-daeb-431a-bc1f-febbdf9bb9d5")> _
  Public Class SampleVbGroupObjects
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbGroupObjects"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim go As New Rhino.Input.Custom.GetObject()
      go.SetCommandPrompt("Select objects to group")
      go.GroupSelect = True
      go.GetMultiple(1, 0)
      If (go.CommandResult() <> Rhino.Commands.Result.Success) Then
        Return go.CommandResult()
      End If

      Dim objids As New List(Of Guid)()
      For i As Integer = 0 To go.ObjectCount - 1
        Dim objref As Rhino.DocObjects.ObjRef = go.Object(i)
        objids.Add(objref.ObjectId)
      Next

      Dim group_index As Integer = doc.Groups.Add(objids)

      Return Result.Success

    End Function
  End Class
End Namespace