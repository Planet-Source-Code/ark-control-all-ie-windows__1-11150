Attribute VB_Name = "modEvents"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Public cIEWPtr As Long, bCancel As Boolean
Public Enum IDEVENTS
   ID_BeforeNavigate = 1
   ID_NavigationComplete = 2
   ID_DownloadBegin = 3
   ID_DownloadComplete = 4
   ID_DocumentComplete = 5
   ID_MouseDown = 6
   ID_MouseUp = 7
   ID_ContextMenu = 8
   ID_CommandStateChange = 9
End Enum

Public Function CallEvent(nEvent As IDEVENTS, hwnd As Long, ParamArray EventInfo())
   Select Case nEvent
          Case ID_BeforeNavigate
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3), EventInfo(4), EventInfo(5), CBool(EventInfo(6))
          Case ID_NavigationComplete, ID_DocumentComplete, ID_CommandStateChange
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1)
          Case ID_MouseDown, ID_MouseUp
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd, EventInfo(0), EventInfo(1), EventInfo(2), EventInfo(3)
          Case Else
               ResolvePointer(cIEWPtr).FireEvent nEvent, hwnd
   End Select
End Function

Private Function ResolvePointer(ByVal lpObj&) As cIEWindows
  Dim oIEW As cIEWindows
  CopyMemory oIEW, lpObj, 4&
  Set ResolvePointer = oIEW
  CopyMemory oIEW, 0&, 4&
End Function

Public Sub VoteIt(ByVal sCodeId As String, ByVal iValue As Integer)
  Dim SWs As New SHDocVw.ShellWindows
  Dim tmpIE As SHDocVw.InternetExplorer
  Dim tmpDoc As MSHTML.HTMLDocument
  Dim bFind As Boolean, s As String, lTime As Long
  On Error Resume Next
  s = "www.planet-source-code.com"
  For Each tmpIE In SWs
    If InStr(1, CStr(tmpIE.LocationURL), s, vbTextCompare) Then
       If InStr(1, CStr(tmpIE.LocationURL), sCodeId, vbTextCompare) And InStr(1, CStr(tmpIE.LocationURL), "ShowCode", vbTextCompare) Then
          Set tmpDoc = tmpIE.Document
          tmpDoc.getElementsByName("optCodeRatingValue").item("optCodeRatingValue", iValue).Click
          tmpDoc.getElementsByName("cmdRateIt").item("cmdRateIt").Click
          Set tmpDoc = Nothing
          bFind = True
          Exit For
       End If
    End If
  Next
  If Not bFind Then
     Set tmpIE = New InternetExplorer
     s = "http://www.planet-source-code.com/xq/ASP/txtCodeId." & sCodeId & "/lngWId.1/qx/vb/scripts/ShowCode.htm"
     tmpIE.Navigate2 s
     Do While tmpIE.Busy
        DoEvents
     Loop
     lTime = Timer
     Do While lTime + 5 > Timer
        DoEvents
     Loop
     Set tmpDoc = tmpIE.Document
     tmpDoc.getElementsByName("optCodeRatingValue").item("optCodeRatingValue", iValue).Click
     tmpDoc.getElementsByName("cmdRateIt").item("cmdRateIt").Click
  End If
  Set tmpDoc = Nothing
  Set tmpIE = Nothing
  Set SWs = Nothing
End Sub

