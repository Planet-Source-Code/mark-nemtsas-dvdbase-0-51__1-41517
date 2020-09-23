Attribute VB_Name = "globalCode"
Option Explicit

Public DEDVDBase As New ADODB.Connection
Public Const strConnectionStart = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const strConnectionEnd = ";Mode=Read|Write;Persist Security Info=False"
Public Const txtIniFile = "DVDBase.ini"
Public strConnection As String
Public Const strCaption = "DVDBase - The Easy DVD Database"


'
'API Declares
'
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public intFormAction As Integer
Public Const ADD_NEW = 1
Public Const EDIT = 2
Public Const DELETE = 3

Public cmdSelectDVDs As New ADODB.Command
Public cmdSelectDVDByID As New ADODB.Command
Public cmdSelectLatestDVD As New ADODB.Command
Public cmdGenres As New ADODB.Command
Public cmdRegions As New ADODB.Command
Public cmdRatings As New ADODB.Command
Public cmdCaseTypes As New ADODB.Command
Public cmdCurrentLocations As New ADODB.Command

Public dvdCurrent As clsDVD

Public lngDVDIDArray() As Long
Public intDVDIDArraySize As Integer

Public lngGenreArray() As Long
Public lngRegionArray() As Long
Public lngRatingArray() As Long
Public lngCaseTypeArray() As Long
Public lngCurrentLocationArray() As Long
Public intGenreArraySize As Integer
Public intRegionArraySize As Integer
Public intRatingArraySize As Integer
Public intCaseTypeArraySize As Integer
Public intCurrentLocationArraySize As Integer


Sub initialise()
  makeConnection
  defineCommandObjects
  
  Set dvdCurrent = New clsDVD
End Sub

Private Sub makeConnection()
  Dim strReadIni As String
  
  'INI File Processing
  '
  '
  'Read INI file and get database location
  '
  'strReadIni = readINIFile("DatabaseLoc")
  '
  'Construct database connection string on the basis of this, note that if no string is returned then
  'assume a default location
  '
  If Len(strReadIni) = 0 Then
    strConnection = strConnectionStart & App.Path & "\" & "DVDBase.mdb" & strConnectionEnd
  Else
    strConnection = strConnectionStart & strReadIni & strConnectionEnd
  End If
  DEDVDBase.ConnectionString = strConnection
  DEDVDBase.CursorLocation = adUseClient
  DEDVDBase.Open strConnection, "Admin"
End Sub

Public Sub exitSystem()
  '
  'Close off the database connections
  '
  DEDVDBase.Close
  Set DEDVDBase = Nothing
  '
  'Release Form Memory
  '
  Set frmMain = Nothing
  End

End Sub

Public Sub defineCommandObjects()
  Dim strSQL As String
  
  With cmdSelectDVDs
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT Discs.lngID,Discs.strTitle from Discs ORDER BY Discs.strTitle"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectLatestDVD
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT Discs.lngID,Discs.strTitle from Discs ORDER BY Discs.lngID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdSelectDVDByID
    .ActiveConnection = DEDVDBase
    .CommandText = "PARAMETERS ID Text; SELECT * from Discs where Discs.lngID =ID"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdGenres
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT * from Genres ORDER BY Genres.lngGenre"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
   
  With cmdRegions
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT * from Regions ORDER BY Regions.lngRegion"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdRatings
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT * from Ratings ORDER BY Ratings.lngRating"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdCaseTypes
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT * from `Case Types` ORDER BY `Case Types`.lngCaseType"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
  With cmdCurrentLocations
    .ActiveConnection = DEDVDBase
    .CommandText = "SELECT * from `Current Locations` ORDER BY `Current Locations`.lngCurrentLocation"
    .CommandType = adCmdText
    .Parameters.Refresh
  End With
  
End Sub

Function returnRS(cmdCommand As ADODB.Command) As ADODB.Recordset
  Dim rstReturnRS As New ADODB.Recordset
  
  With rstReturnRS
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmdCommand
  End With
  Set returnRS = rstReturnRS
End Function

Public Function fillDVDListView(lvwBox As ListView, Optional lngSelectedID As Long)
  Dim rstDVDs As New ADODB.Recordset
  Dim intArraySize As Integer
  Dim lstItem As ListItem
  Dim intSelectedIndex As Integer
  
  Set rstDVDs = returnRS(cmdSelectDVDs)
  ReDim lngDVDIDArray(1)
  
  lvwBox.ListItems.Clear
  If rstDVDs.EOF = False Then
    rstDVDs.MoveFirst
    intArraySize = 1
    While rstDVDs.EOF <> True
      ReDim Preserve lngDVDIDArray(intArraySize)
      lngDVDIDArray(intArraySize) = rstDVDs![lngID]
      Set lstItem = lvwBox.ListItems.Add
      lstItem = rstDVDs![strTitle]
      If lngSelectedID = rstDVDs![lngID] Then
        intSelectedIndex = intArraySize
      End If
      rstDVDs.MoveNext
      intArraySize = intArraySize + 1
    Wend
  Else
    intArraySize = 0
  End If
  intDVDIDArraySize = intArraySize - 1
  rstDVDs.Close
  Set rstDVDs = Nothing
  If intSelectedIndex > 0 Then
    Set lvwBox.SelectedItem = lvwBox.ListItems(intSelectedIndex)
    discCode.enableDiscDisplay
    intFormAction = EDIT
    dvdCurrent.fillDVD lngDVDIDArray(intSelectedIndex)
    dvdCurrent.displayDVD
  End If
End Function

Public Function fillSelectCombo(strWhichCombo As String)
  Dim rstTmp As New ADODB.Recordset
  Dim cboTmp As ComboBox
  Dim cmdTmp As ADODB.Command
  Dim lngIDArray() As Long
  Dim intArraySize As Integer
  
  Select Case strWhichCombo
    Case "cboGenre"
      Set cboTmp = frmMain.cboGenre
      Set cmdTmp = cmdGenres
    Case "cboRegion"
      Set cboTmp = frmMain.cboRegion
      Set cmdTmp = cmdRegions
    Case "cboRating"
      Set cboTmp = frmMain.cboRating
      Set cmdTmp = cmdRatings
    Case "cboCurrentLocation"
      Set cboTmp = frmMain.cboCurrentLocation
      Set cmdTmp = cmdCurrentLocations
    Case "cboCaseType"
      Set cboTmp = frmMain.cboCaseType
      Set cmdTmp = cmdCaseTypes
  End Select
  
  Set rstTmp = returnRS(cmdTmp)
  ReDim lngIDArray(1)
  
  cboTmp.Clear
  If rstTmp.EOF = False Then
    rstTmp.MoveFirst
    intArraySize = 1
    While rstTmp.EOF <> True
      Select Case strWhichCombo
        Case "cboGenre"
          ReDim Preserve lngGenreArray(intArraySize)
          lngGenreArray(intArraySize) = rstTmp![lngGenre]
          cboTmp.AddItem rstTmp![strGenre]
        Case "cboRegion"
          ReDim Preserve lngRegionArray(intArraySize)
          lngRegionArray(intArraySize) = rstTmp![lngRegion]
          cboTmp.AddItem rstTmp![strRegion]
        Case "cboRating"
          ReDim Preserve lngRatingArray(intArraySize)
          lngRatingArray(intArraySize) = rstTmp![lngRating]
          cboTmp.AddItem rstTmp![strRating]
        Case "cboCurrentLocation"
          ReDim Preserve lngCurrentLocationArray(intArraySize)
          lngCurrentLocationArray(intArraySize) = rstTmp![lngCurrentLocation]
          cboTmp.AddItem rstTmp![strCurrentLocation]
        Case "cboCaseType"
          ReDim Preserve lngCaseTypeArray(intArraySize)
          lngCaseTypeArray(intArraySize) = rstTmp![lngCaseType]
          cboTmp.AddItem rstTmp![strCaseType]
     End Select
      rstTmp.MoveNext
      intArraySize = intArraySize + 1
    Wend
    Select Case strWhichCombo
      Case "cboGenre"
        intGenreArraySize = intArraySize - 1
      Case "cboRegion"
        intRegionArraySize = intArraySize - 1
      Case "cboRating"
        intRatingArraySize = intArraySize - 1
      Case "cboCurrentLocation"
        intCurrentLocationArraySize = intArraySize - 1
      Case "cboCaseType"
        intCaseTypeArraySize = intArraySize - 1
    End Select
  Else
    intArraySize = 0
  End If
  
  rstTmp.Close
  Set rstTmp = Nothing
End Function

Public Function checkSelected(lvwTmp As Variant) As Integer
  Dim intLoop As Integer
  checkSelected = -1
  If lvwTmp.ListItems.Count = 0 Then
    checkSelected = -1
    Exit Function
  End If
  For intLoop = 1 To lvwTmp.ListItems.Count
    If lvwTmp.ListItems(intLoop).Selected = True Then
      checkSelected = intLoop
      Exit For
    End If
  Next intLoop
End Function

Public Function parseDirectory(strFilename) As String
  Dim strTmp As String
  Dim lngPosition As Long
  
  If IsNull(strFilename) = False Then
    strTmp = StrReverse(strFilename)
    lngPosition = InStr(strTmp, "\")
    strTmp = Left(strFilename, Len(strFilename) - lngPosition)
    parseDirectory = strTmp
    If Len(parseDirectory) = 0 Then
      parseDirectory = App.Path
    End If
  Else
    parseDirectory = App.Path
  End If
End Function


