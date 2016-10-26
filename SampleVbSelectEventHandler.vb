Imports Rhino
Imports Rhino.Commands

''' <summary>
''' SampleVbSelectEventHandler command
''' </summary>
<System.Runtime.InteropServices.Guid("be04a17a-d693-4b25-9337-53b6ba104888")> _
Public Class SampleVbSelectEventHandler
  Inherits Command

  ''' <summary>
  ''' The English command name
  ''' </summary>
  Public Overrides ReadOnly Property EnglishName() As String
    Get
      Return "SampleVbSelectEventHandler"
    End Get
  End Property

  ''' <summary>
  ''' Called by Rhino to run the command
  ''' </summary>
  Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

    ' Running the command toggles the enabled state of our event handlers

    If (IsEnabled) Then
      ' Remove event handlers (disable)
      RemoveHandler RhinoDoc.SelectObjects, AddressOf OnSelectObjects
      RemoveHandler RhinoDoc.DeselectObjects, AddressOf OnDeselectObjects
    Else
      ' Add event handlers (enable)
      AddHandler RhinoDoc.SelectObjects, AddressOf OnSelectObjects
      AddHandler RhinoDoc.DeselectObjects, AddressOf OnDeselectObjects
    End If
    IsEnabled = Not IsEnabled

    RhinoApp.WriteLine("SampleVbSelectEventHandler enabled = {0}", IsEnabled.ToString())

    Return Result.Success

  End Function

  ''' <summary>
  ''' RhinoDoc.SelectObjects event handler
  ''' </summary>
  Public Shared Sub OnSelectObjects(sender As Object, e As DocObjects.RhinoObjectSelectionEventArgs)
    RhinoApp.WriteLine("** EVENT: Select Objects **")
  End Sub

  ''' <summary>
  ''' RhinoDoc.DeselectObjects event handler
  ''' </summary>
  Public Shared Sub OnDeselectObjects(sender As Object, e As DocObjects.RhinoObjectSelectionEventArgs)
    RhinoApp.WriteLine("** EVENT: Deselect Objects **")
  End Sub

  ''' <summary>
  ''' Event handling enabled property
  ''' </summary>
  Private Property IsEnabled As Boolean

End Class