Attribute VB_Name = "mInvioReti"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
   ByVal lpDefault As String, ByVal lpReturnedString As String, _
   ByVal nSize As Long, ByVal lpFileName As String) As Long

Public db As Database, wrkCurrent As Workspace
'   Public IdContainer As Long, rsSearch As Recordset
'   Public PesiCheck(0 To 25) As Integer, frm As Form, asCriteri() As String
'Public rsTabsGrid As Recordset, sDbPath As String
Public sDbPath As String, DbName As String  '   potrebbero non servire
'
'   Seguono dei parametri che possono essere utili
'   public parIdSpuntaStrada As String
'   Public parIdClienteVuoti As String, parIdTerminalVuoti As String
'   Public parKmMinimi As Integer
Public fSelectField As Boolean, fDBOk As Boolean

Public Enum cColors
   cGray = &H8000000F
   cWhite = &H80000005
   cLightRed = &H8080FF
End Enum

Public Enum FormStatus
   fDisplay = 1
   fInsert = 2
   fModify = 3
   fSearch = 4
End Enum

Public Enum Menu
   cmnuApri = 0
'   cmnuNuovo = 1
   cmnuSetPrinter = 2
   cmnuEsci = 4
    
   cmnuCliente = 0
   cmnuCondizPagam = 1
   cmnuTreno = 2
   cmnuFornitoreTreno = 3
   cmnuViaggioNave = 4
   cmnuNave = 5
   cmnuPorto = 6
   cmnuTransitPorto = 7
   cmnuTraspStrada = 8
   cmnuSpedizioniere = 9
   cmnuTerminalInterno = 11
   cmnuTerminalVuoti = 12
   cmnuDogana = 13
   cmnuTipoContainer = 14
   cmnuServizio = 15
   cmnuCompagnia = 16
   cmnuCodiciPorto = 17
   cmnuCVuoti = 18
   cmnuDistanze = 19
    
   cmnuSpStrada = 0
   cmnuSpIntl = 1
   cmnuSpMsc = 2
    
   cmnuGPieniN = 0
   cmnuGPieniC = 1
   
   cmnuGVuoti = 1
   cmnuGUVuoti = 3
   
   cmnuSTCompleto = 0
   cmnuSTCliente = 1
   
   cmnuOpzClear = 0
   cmnuOpzTreno = 2
   cmnuOpzUpdate = 4
   cmnuOpzSelect = 6
   
   cmnuSetTreno = 0
   cmnuuRecalcTreno = 1
   cmnuExport = 3
   cmnuCompact = 5
   cmnuRepair = 6
   cmnuPVuoti = 8
   cmnuStatistiche = 10
    
End Enum

Public Function DataUsa(DataI As String) As Date
   Dim gg As Integer, mm As Integer, aa  As Integer
   gg = Val(Mid$(DataI, 1, 2))
   mm = Val(Mid$(DataI, 4, 2))
   aa = Val(Mid$(DataI, 7, 4))
   DataUsa = DateSerial(aa, mm, gg)
'   DataUsa = "#" & Format$(DataTemp, "mm/dd/Yyyy") & "#"
End Function

Sub Main()
   Dim sCmdLine As String
   Screen.MousePointer = vbArrowHourglass
   sDbPath = App.Path & "\..\Reclami\"
   sCmdLine = Command()
   If Len(sCmdLine) > 0 Then
       sDbPath = sDbPath & sCmdLine
   Else
       sDbPath = sDbPath & "Reclami.mdb"
   End If
   
   Set wrkCurrent = DBEngine.Workspaces(0)
   fDBOk = DbOpen(sDbPath)
   
   If fDBOk Then
       SetAllParameters
   End If
   
   Screen.MousePointer = vbNormal

   Load frmInvioReti
   
End Sub
Sub SetAllParameters()
'   parPreview = SetParameters("Preview")
'   parAliquotaIva = SetParameters("Aliquota Iva")
'   parIdClienteVuoti = SetParameters("Id Cliente Vuoti")
'   parIdTerminalVuoti = SetParameters("Id Terminal Vuoti")
'   parKmMinimi = Val(SetParameters("KmMinimi"))
'   parDirTreniExp = SetParameters("Dir TreniExp")
End Sub
Public Sub CenterForm(f As Form)
   f.Top = (Screen.Height - f.Height) / 2
   f.Left = (Screen.Width - f.Width) / 2
End Sub
Public Function DbOpen(DbName As String)
   Dim sCheck As String
   On Error GoTo DbOpenError
   Set db = wrkCurrent.OpenDatabase(DbName, False, False)
   On Error Resume Next
   sCheck = db.TableDefs("$$CheckReclami").Name
   If Err.Number <> 0 Then
       DbOpen = False
       MsgBox "Database NON corretto !!", vbCritical + vbOKOnly, "Apertura DataBase"
   Else
       DbOpen = True
   End If
   Exit Function
    
DbOpenError:
   Dim sMsg As String
   sMsg = "Errore nell'apertura del Database" & vbCrLf & _
       DbName & vbCrLf & vbCrLf & _
       "Numero Errore: " & Err.Number & vbCrLf & Err.Description
   MsgBox sMsg
   DbOpen = False
End Function
