Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("2164aa26-fa4d-49e8-9b9e-f84a973c4faf")> _
  Public Class SampleVbSetDisplayMode
    Inherits Command

    ''' <returns>
    ''' The command name as it appears on the Rhino command line.
    ''' </returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbSetDisplayMode"
      End Get
    End Property

    ''' <summary>
    ''' Remove characters that are invalid for command options display
    ''' </summary>
    Private Shared Function FormatValidDisplayName(displayName As String) As String
      displayName = displayName.Replace("_", "")
      displayName = displayName.Replace(" ", "")
      displayName = displayName.Replace("-", "")
      displayName = displayName.Replace(",", "")
      displayName = displayName.Replace(".", "")
      displayName = displayName.Replace("/", "")
      displayName = displayName.Replace("(", "")
      displayName = displayName.Replace(")", "")
      Return displayName
    End Function

    ''' <summary>
    ''' Called by Rhino to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(doc As RhinoDoc, mode As RunMode) As Result
      Dim view As Rhino.Display.RhinoView = doc.Views.ActiveView
      If view Is Nothing Then
        Return Result.Failure
      End If

      Dim viewport As Rhino.Display.RhinoViewport = view.ActiveViewport
      Dim currentDisplayMode As Rhino.Display.DisplayModeDescription = viewport.DisplayMode
      Rhino.RhinoApp.WriteLine("Viewport in {0} display.", currentDisplayMode.EnglishName)

      Dim displayModes As Rhino.Display.DisplayModeDescription() = Rhino.Display.DisplayModeDescription.GetDisplayModes()

      Dim go As New Rhino.Input.Custom.GetOption()
      go.SetCommandPrompt("Select new display mode")
      go.AcceptNothing(True)

      For Each displayMode As Rhino.Display.DisplayModeDescription In displayModes
        Dim displayName As String = FormatValidDisplayName(displayMode.EnglishName)
        go.AddOption(displayName)
      Next

      Dim rc As Rhino.Input.GetResult = go.[Get]()
      Select Case rc
        Case Rhino.Input.GetResult.[Option]
          If True Then
            Dim optionIndex As Integer = go.[Option]().Index
            If optionIndex > 0 AndAlso optionIndex <= displayModes.Length Then
              Dim newDisplayMode As Rhino.Display.DisplayModeDescription = displayModes(optionIndex - 1)
              If newDisplayMode.Id <> currentDisplayMode.Id Then
                viewport.DisplayMode = newDisplayMode
                view.Redraw()
                Rhino.RhinoApp.WriteLine("Viewport set to {0} display.", viewport.DisplayMode.EnglishName)
              End If
            End If
          End If
          Exit Select

        Case Rhino.Input.GetResult.Cancel
          Return Result.Cancel
        Case Else

          Exit Select
      End Select

      Return Result.Success
    End Function

  End Class
End Namespace