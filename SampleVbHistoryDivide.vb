Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCommands

  <System.Runtime.InteropServices.Guid("e4ee2abd-aee8-4406-b85d-0c980bd7c5a5")> _
  Public Class SampleVbHistoryDivide
    Inherits Command

    Shared _instance As SampleVbHistoryDivide
    Shared _historyVersion As Integer = 20121101

    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a refence in a static field.
      _instance = Me
    End Sub

    '''<summary>The only instance of this command.</summary>
    Public Shared ReadOnly Property Instance() As SampleVbHistoryDivide
      Get
        Return _instance
      End Get
    End Property

    '''<returns>The command name as it appears on the Rhino command line.</returns>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbHistoryDivide"
      End Get
    End Property

    ''' <summary>
    ''' Rhino calls this function to run the command.
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result
      Const filter As Rhino.DocObjects.ObjectType = Rhino.DocObjects.ObjectType.Curve
      Dim objref As Rhino.DocObjects.ObjRef = Nothing
      Dim rc As Rhino.Commands.Result = Rhino.Input.RhinoGet.GetOneObject("Select curve to divide", False, filter, objref)
      If rc <> Rhino.Commands.Result.Success OrElse objref Is Nothing Then
        Return rc
      End If

      Dim curve As Rhino.Geometry.Curve = objref.Curve()
      If curve Is Nothing OrElse curve.IsShort(Rhino.RhinoMath.ZeroTolerance) Then
        Return Rhino.Commands.Result.Failure
      End If

      Dim segmentCount As Integer = 2
      rc = Rhino.Input.RhinoGet.GetInteger("Number of segments", False, segmentCount, 2, 100)
      If rc <> Rhino.Commands.Result.Success Then
        Return rc
      End If

      Dim points As Rhino.Geometry.Point3d() = Nothing
      curve.DivideByCount(segmentCount, True, points)
      If points Is Nothing Then
        Return Rhino.Commands.Result.Failure
      End If

      ' Create a history record
      Dim history As New Rhino.DocObjects.HistoryRecord(Me, _historyVersion)
      WriteHistory(history, objref, segmentCount, points.Length)

      For i As Integer = 0 To points.Length - 1
        doc.Objects.AddPoint(points(i), Nothing, history, False)
      Next

      doc.Views.Redraw()

      Return Rhino.Commands.Result.Success

    End Function

    ''' <summary>
    ''' Rhino calls this function to remake objects when inputs have changed. 
    ''' </summary>
    Protected Overrides Function ReplayHistory(replay As Rhino.DocObjects.ReplayHistoryData) As Boolean
      Dim objref As Rhino.DocObjects.ObjRef = Nothing
      Dim segmentCount As Integer = 0
      Dim pointCount As Integer = 0

      If Not ReadHistory(replay, objref, segmentCount, pointCount) Then
        Return False
      End If

      Dim curve As Rhino.Geometry.Curve = objref.Curve()
      If curve Is Nothing Then
        Return False
      End If

      If pointCount <> replay.Results.Length Then
        Return False
      End If

      Dim points As Rhino.Geometry.Point3d() = Nothing
      curve.DivideByCount(segmentCount, True, points)
      If points Is Nothing Then
        Return False
      End If

      For i As Integer = 0 To points.Length - 1
        replay.Results(i).UpdateToPoint(points(i), Nothing)
      Next

      Return True
    End Function

    Private Function ReadHistory(replay As Rhino.DocObjects.ReplayHistoryData, ByRef objref As Rhino.DocObjects.ObjRef, ByRef segmentCount As Integer, ByRef pointCount As Integer) As Boolean
      If _historyVersion <> replay.HistoryVersion Then
        Return False
      End If

      objref = replay.GetRhinoObjRef(0)
      If objref Is Nothing Then
        Return False
      End If

      If Not replay.TryGetInt(1, segmentCount) Then
        Return False
      End If

      If Not replay.TryGetInt(2, pointCount) Then
        Return False
      End If

      Return True
    End Function

    Private Function WriteHistory(history As Rhino.DocObjects.HistoryRecord, objref As Rhino.DocObjects.ObjRef, segmentCount As Integer, pointCount As Integer) As Boolean
      If Not history.SetObjRef(0, objref) Then
        Return False
      End If

      If Not history.SetInt(1, segmentCount) Then
        Return False
      End If

      If Not history.SetInt(2, pointCount) Then
        Return False
      End If

      Return True
    End Function

  End Class
End Namespace