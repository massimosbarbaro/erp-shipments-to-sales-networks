VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInvioReti 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Invio Reti"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   Icon            =   "frmInvioReti.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   855
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmInvioReti.frx":0CCA
      Top             =   5400
      Width           =   5655
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   6720
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "C:\Programmi\reclami\txt\resi282.txt"
      Top             =   1200
      Width           =   5415
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsci 
      Cancel          =   -1  'True
      Caption         =   "Esci"
      Default         =   -1  'True
      Height          =   615
      Left            =   6600
      Picture         =   "frmInvioReti.frx":0CD5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Messaggi Postali"
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   5160
      Width           =   1185
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   0
      Picture         =   "frmInvioReti.frx":0E1F
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Invio dati da MFG alle reti"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   5730
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1200
      TabIndex        =   3
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Label lblLog 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2295
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   5655
   End
End
Attribute VB_Name = "frmInvioReti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" _
   (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
   ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const INFINITE = &HFFFF
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = 259
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

'  Private ImpMfg(0 To 19, 1 To 2) As Variant
Private Const dbTempName As String = "$$Temp.mdb"
Private Const MaxParms As Byte = 10

Private dbTemp As Database
Private ImpMfg(0 To 25) As String
Private dbTempV As Database
Private Const dbTempNameV As String = "ReclamiVenditore.mdb"

Private Const sFieldSep As String = "\"

Private sKeys(0 To MaxParms - 1) As String
Private sKeysDefaults(0 To MaxParms - 1) As String

Private sSmtpServer As String, sSmtpTo As String
Private sSmtpSenderName As String, sSmtpSenderEMail As String
Private sSmtpRecipientName As String, sSmtpRecipientEMail As String
Private iSmtpServerPort As Integer, iSmtpTimeout As Integer
Private sSmtpStartStr As String, sSmtpEndStr As String, sSmtpSubject As String

Private sBuff As String, fOkToSend As Boolean

Private Const iMaxLenBodyRow As Integer = 60


Private Function IsActive(hProg) As Long
   Dim hProc, RetVal As Long
   Const PROCESS_QUERY_INFORMATION = &H400
   Const STILL_ACTIVE = 259
   hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
   If hProc <> 0 Then
      GetExitCodeProcess hProc, RetVal
   End If
   IsActive = (RetVal = STILL_ACTIVE)
   CloseHandle hProc
End Function

Private Sub cmdEsci_Click()
   Unload Me
End Sub
Private Sub AggMFG()
   Dim nFile As Integer, sFileName As String
   Dim sqlc As String, rsImport As Recordset, rsMFG As Recordset
   Dim sLine As String, i As Integer, sTmp As Variant
   Dim nFields As Integer, nRecs As Integer
   Dim sFields(0 To 25) As String
'   *************************************************************************
   Dim hProg, hProc, RetVal As Long
   Dim sCmd As String, sPrg As String, sFile As String, sLabel As String
   
'   Const PROCESS_ALL_ACCESS = 0
   Const PROCESS_ALL_ACCESS = &H1F0FFF
   cmdEsci.Enabled = False
   Me.Enabled = False
    
' Parte lo script di reflection per creare il file resi282.txt
'   sPrg = "c:\programmi\r1win\r1win.exe"
   sPrg = "C:\program files\r1win\r1win.exe"

   sFile = App.Path & "\" & "reclami.rbs"
   sCmd = sPrg & " " & "/rbs" & " " & sFile
'   sCmd = "calc.exe"       'prova di funzionamento sincrono
   
   hProg = Shell(sCmd, vbNormalFocus)
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
   End If
'  MsgBox "Fine scarico da MFG"
 
 '   sCmd = App.Path & "\" & "teln.bat"
   
'  trasporto il file txt dal server sul pc
'   sCmd = "ftp -s:" & App.Path & "\" & "ftp.txt" '
   sCmd = "ftp -s:" & App.Path & "\" & "ftp.scr" ' originale
'   sCmd = "ftp -s:" & App.Path & "\" & "prova.txt"
   hProg = Shell(sCmd, vbNormalFocus)
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
   End If
   
   Do While IsActive(hProg)
      DoEvents
   Loop
   hProc = OpenProcess(PROCESS_ALL_ACCESS, False, hProg)
   If hProc <> 0 Then
      RetVal = WaitForSingleObject(hProc, INFINITE)
      CloseHandle hProc
  End If
 ' MsgBox "Fine trasporto FTP"
     
  Me.Enabled = True
  Me.SetFocus
'   ******************************************************************************
   Screen.MousePointer = vbArrowHourglass
   
   cmdImport.Enabled = False
   Me.Enabled = False
'
'  Apertura Db Temporaneo
'
   Set dbTemp = DAO.OpenDatabase(App.Path & "\..\reclami\" & dbTempName)
'solo test
'      Set dbTemp = DAO.OpenDatabase(App.Path & "\..\reclami2000\" & dbTempName)
   lblLog.Caption = ""
   lblLog.Visible = True
   With lblProgress
      .Caption = ""
      .Top = lblLog.Top + lblLog.Height - (.Height * 1.5)
      .Visible = True
   End With
   Me.Refresh
   
   nRecs = 0
   
   lblLog.Caption = "E' in corso " & vbCrLf & _
                  " la cancellazione della tabella temporanea"
   DoEvents
   
   sqlc = "Delete from [$$ImportMFG]"
   dbTemp.Execute (sqlc)
   
   sqlc = "Select * from [$$ImportMFG]"
   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenDynaset)
   nFields = rsImport.Fields.Count
   
   lblLog.Caption = lblLog.Caption & vbCrLf & _
      "Stiamo caricando una nuova tabella "
   DoEvents
   
   sFileName = txtFileName.Text
   nFile = FreeFile
   Open sFileName For Input As nFile
   Do While Not EOF(nFile)
      Line Input #nFile, sLine
      nRecs = nRecs + 1
      If (nRecs Mod 100) = 0 Then
         DoEvents
         lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
      End If
      For i = 0 To nFields - 1
         sFields(i) = Trim$(Mid$(sLine, (i * 50 + 1), 50))
      Next i
      rsImport.AddNew
      For i = 0 To nFields - 1
         sTmp = Trim$(sFields(i))
         Select Case rsImport.Fields(ImpMfg(i)).Type
            Case dbText
               If sTmp <> "" Then
                  rsImport(ImpMfg(i)) = Left$(sTmp, rsImport(ImpMfg(i)).Size)
               Else
                  rsImport(ImpMfg(i)) = Null
               End If
    
            Case dbDouble
               If sTmp = "" Then
                  rsImport(ImpMfg(i)) = Null
               Else
                  rsImport(ImpMfg(i)) = Format(sTmp, "0.00#")
               End If
'            Case dbLong
'            Case dbBigInt
'            Case dbBinary
'            Case dbBoolean
'            Case dbByte
'            Case dbChar
'            Case dbCurrency
'            Case dbDate
'            Case dbDecimal
 '           Case dbDouble
 ''           Case dbFloat
 '           Case dbGUID
 '           Case dbInteger
 '           Case dbLong
 '           Case dbLongBinary
 '           Case dbMemo
 '           Case dbNumeric
 '           Case dbSingle
 '           Case dbText
 '           Case dbTime
 '           Case dbTimeStamp
  '          Case dbVarBinary

            Case dbDate
               If sTmp = "" Then
                  rsImport(ImpMfg(i)) = Null
               Else
                  rsImport(ImpMfg(i)) = sTmp
               End If
            Case Else
               If sTmp = "" Then
                  rsImport(ImpMfg(i)) = Null
               Else
                  rsImport(ImpMfg(i)) = Val(sTmp)
               End If
         End Select
      Next i
      rsImport.Update
   Loop
   Close nFile
   rsImport.Close
   Set rsImport = Nothing
   lblProgress.Visible = False
   
   lblLog.Caption = vbCrLf & _
         "Fine caricamento tabella temporanea: " & _
         Str$(nRecs) & " Records"
   DoEvents
   sLabel = "Caricati " & Str$(nRecs) & " Records"
'  Caricamento tabella MFG con i soli ODL Distinti e non nulli
   lblLog.Caption = ""
   With lblLog
'      .Caption = lblLog.Caption & vbCrLf & _
         "Cancellazione  vecchi dati "
      .ForeColor = &H80000007
      .FontSize = 10
      .Alignment = 0
   End With
   DoEvents
'Comincio il confronto dei dati
' rsImport è per noi il rsNew

'   sqlc = "SELECT Transazione, Causale, Bolla, DataBolla, ODL, Rete, CodCli, Cliente, CodProd, Prodotto, LineaProdotto, " & _
'            "Colore, Finitura, Spessore, Altezza, Larghezza, " & _
'            "Sum(Qta) AS SommaDiQta, ODV, Esterni, Ubicazione, Bobinatrice, " & _
'            "DataBobinatrice, Calandra, DataCalandra " & _
'            "FROM [$$ImportMFG] " & _
'            "Where ODL Is Not Null "
'            "GROUP BY ODL, Rete, CodCli, Cliente, CodProd, Prodotto, " & _
'            "LineaProdotto, Colore, Finitura, Spessore, Altezza, Larghezza, " & _
'            "ODV, Esterni, Ubicazione, Bobinatrice, DataBobinatrice, Calandra, " & _
'            "DataCalandra"
   
'   sqlc = "SELECT Max([$$ImportMFG].Transazione) AS MaxDiTransazione, " & _
            "[$$ImportMFG].Causale, Max([$$ImportMFG].Bolla) AS MaxDiBolla, " & _
            "Max([$$ImportMFG].DataBolla) AS MaxDiDataBolla, [$$ImportMFG].ODL, " & _
            "[$$ImportMFG].Rete, [$$ImportMFG].CodCli, [$$ImportMFG].Cliente, " & _
            "[$$ImportMFG].CodProd, [$$ImportMFG].Prodotto, " & _
            "[$$ImportMFG].LineaProdotto, [$$ImportMFG].Colore, " & _
            "[$$ImportMFG].Finitura, [$$ImportMFG].Spessore, " & _
            "[$$ImportMFG].Altezza, [$$ImportMFG].Larghezza, Sum([$$ImportMFG].Qta) " & _
            "AS SommaDiQta, [$$ImportMFG].ODV, [$$ImportMFG].Esterni, " & _
            "[$$ImportMFG].Ubicazione, [$$ImportMFG].Bobinatrice, " & _
            "Max([$$ImportMFG].DataBobinatrice) AS MaxDiDataBobinatrice, " & _
            "[$$ImportMFG].Calandra, Max([$$ImportMFG].DataCalandra) " & _
            "AS MaxDiDataCalandra,  [$$ImportMFG].Costo, [$$ImportMFG].Fattore " & _
            "FROM [$$ImportMFG] INNER JOIN Causale ON " & _
            "[$$ImportMFG].Causale = Causale.Causale " & _
            "GROUP BY [$$ImportMFG].Causale, [$$ImportMFG].ODL, " & _
            "[$$ImportMFG].Rete, [$$ImportMFG].CodCli, [$$ImportMFG].Cliente, " & _
            "[$$ImportMFG].CodProd, [$$ImportMFG].Prodotto, " & _
            "[$$ImportMFG].LineaProdotto, [$$ImportMFG].Colore, " & _
            "[$$ImportMFG].Finitura, [$$ImportMFG].Spessore, " & _
            "[$$ImportMFG].Altezza, [$$ImportMFG].Larghezza, " & _
            "[$$ImportMFG].ODV, [$$ImportMFG].Esterni, [$$ImportMFG].Ubicazione, " & _
            "[$$ImportMFG].Bobinatrice, [$$ImportMFG].Calandra, [$$ImportMFG].Costo, [$$ImportMFG].Fattore " & _
            "Having ((([$$ImportMFG].ODL) Is Not Null))" & _
            "ORDER BY [$$ImportMFG].ODL"
            
sqlc = "SELECT Max([$$ImportMFG].Transazione) AS MaxDiTransazione, " & _
            "[$$ImportMFG].Causale, Max([$$ImportMFG].Bolla) AS MaxDiBolla, " & _
            "Max([$$ImportMFG].DataBolla) AS MaxDiDataBolla, [$$ImportMFG].ODL, " & _
            "[$$ImportMFG].Rete, [$$ImportMFG].CodCli, [$$ImportMFG].Cliente, " & _
            "[$$ImportMFG].CodProd, [$$ImportMFG].Prodotto, [$$ImportMFG].LineaProdotto, " & _
            "[$$ImportMFG].Colore, [$$ImportMFG].Finitura, [$$ImportMFG].Spessore, " & _
            "[$$ImportMFG].Altezza, [$$ImportMFG].Larghezza, Sum([$$ImportMFG].Qta) " & _
            "AS SommaDiQta, [$$ImportMFG].ODV, [$$ImportMFG].Esterni, " & _
            "[$$ImportMFG].Ubicazione, [$$ImportMFG].Bobinatrice, " & _
            "Max([$$ImportMFG].DataBobinatrice) AS MaxDiDataBobinatrice, " & _
            "[$$ImportMFG].Calandra, Max([$$ImportMFG].DataCalandra) AS " & _
            "MaxDiDataCalandra, [$$ImportMFG].Costo, [$$ImportMFG].Fattore " & _
        "FROM [$$ImportMFG] INNER JOIN Causale ON [$$ImportMFG].Causale = Causale.Causale " & _
        "GROUP BY [$$ImportMFG].Causale, [$$ImportMFG].ODL, [$$ImportMFG].Rete, " & _
            "[$$ImportMFG].CodCli, [$$ImportMFG].Cliente, [$$ImportMFG].CodProd, " & _
            "[$$ImportMFG].Prodotto, [$$ImportMFG].LineaProdotto, [$$ImportMFG].Colore, " & _
            "[$$ImportMFG].Finitura, [$$ImportMFG].Spessore, [$$ImportMFG].Altezza, " & _
            "[$$ImportMFG].Larghezza, [$$ImportMFG].ODV, [$$ImportMFG].Esterni, " & _
            "[$$ImportMFG].Ubicazione, [$$ImportMFG].Bobinatrice, " & _
            "[$$ImportMFG].Calandra, [$$ImportMFG].Costo, [$$ImportMFG].Fattore " & _
        "Having ((([$$ImportMFG].ODL) Is Not Null))" & _
        "ORDER BY [$$ImportMFG].ODL "
            
   
   Set rsImport = dbTemp.OpenRecordset(sqlc, dbOpenSnapshot)
        nRecs = 0
        With lblProgress
            .Caption = ""
            .Top = lblLog.Top + (.Height * 0.5)
            .Visible = True
        End With
        nFields = rsImport.Fields.Count
    
    Do While Not rsImport.EOF
        sqlc = "SELECT * FROM MFG " & _
                "where ODL='" & rsImport("ODL") & "'"
        Set rsMFG = db.OpenRecordset(sqlc, dbOpenDynaset)
        
        On Error Resume Next
        
      If Not rsMFG.EOF Then
         rsMFG.Edit
      Else
         rsMFG.AddNew
         rsMFG("ODL") = rsImport("ODL")
      End If
' Questa parte serve solo a caricare la barra ---------------------
        nRecs = nRecs + 1
        If (nRecs Mod 100) = 0 Then
           lblProgress.Caption = "Caricamento rec. #" & Format$(nRecs, "####0")
           DoEvents
        End If
' Qui si fa il confronto tra i dati per la verifica degli ODL

      For i = 0 To nFields - 1
         If rsMFG.Fields(i).Name <> "ODL" Then
            rsMFG(i) = rsImport(i)
         End If
      Next i
        rsMFG.Update
' Controllo la possibilità di errori di duplicatura e li segnalo
        If Err.Number <> 0 Then
           Select Case Err.Number
              Case 3022
                 MsgBox "Duplicato ODL: " & rsMFG("ODL")
              Case Else
 '                MsgBox "Errore: " & Str$(Err.Number) & vbCrLf & _
                       Err.Description
           End Select
           Err.Clear
        End If
        rsMFG.Close
        rsImport.MoveNext
   Loop
   lblProgress.Visible = False
   With lblLog
'      .Caption = lblLog.Caption & vbCrLf &
'      .Caption = vbCrLf & _
       "Caricati " & Str$(rsMFG.RecordCount) & _
               " ODL distinti in tabella MFG"
      .ForeColor = &HFF0000
      .FontSize = 12
      .Alignment = 2
   End With
   DoEvents
   sLabel = sLabel & vbCrLf & _
       " di cui  " & Str$(rsImport.RecordCount) & _
               " ODL  distinti"

   rsImport.Close
  
   Set rsImport = Nothing
   Set rsMFG = Nothing
'  Chiusura db temporaneo
   dbTemp.Close
   lblLog.Caption = vbCrLf & vbCrLf & sLabel & vbCrLf & vbCrLf & _
            "Fine esecuzione"
                                                                                                                                                                                                                                                               
   Screen.MousePointer = vbNormal
   Me.Enabled = True
   cmdEsci.Enabled = True

End Sub
Private Sub cmdImport_Click()
   Dim rsConfronto As Recordset, rsInput As Recordset
'   dim rsDatiV As Recordset
   Dim sqlc As String, nRecs As Long, sRetePrec As String, sSend As String
   Dim fRes As Boolean, rsTemp As Recordset, rsInV As Recordset
   Dim dOggi As Date
   
   nRecs = 0
   dOggi = Date
   
  Call AggMFG
   
   sqlc = " SELECT m.ODL AS ODLM, m.Qta AS QtaM, i.ODL AS ODLI, " & _
            "ucase(left(m.[Rete], 1)) as ReteRagg " & _
            "FROM MFG as m LEFT JOIN daInviareVenditore as i ON m.ODL = i.ODL " & _
            "Where i.Odl is null  and  m.Rete is not null"
   Set rsConfronto = db.OpenRecordset(sqlc, dbOpenDynaset)
   
   Set rsInput = db.OpenRecordset("Select ODL, ReteRagg, DataElaborazione from daInviareVenditore")

   Do While Not rsConfronto.EOF
      If IsNull(rsConfronto("ODLI")) Then
         'Inserisci nuovo
         rsInput.AddNew
         rsInput("ODL") = rsConfronto("ODLM")
         rsInput("ReteRagg") = rsConfronto("ReteRagg")
         rsInput("DataElaborazione") = dOggi
         rsInput.Update
         nRecs = nRecs + 1
      Else
         If rsConfronto("QtaM") <> rsConfronto("QtaI") Then
            'non dovrebbe mai succedere perchè nel sqlc ho inserito odli is null
            'Inserisci nuovo
               rsInput.AddNew
               rsInput("ODL") = rsConfronto("ODLM")
               rsInput("ReteRagg") = rsConfronto("ReteRagg")
               rsInput("DataElaborazione") = dOggi
               rsInput.Update
         nRecs = nRecs + 1
         End If
      End If
      rsConfronto.MoveNext
   Loop
'MsgBox nRecs


'   sqlc = "SELECT I.ODL, MFG.Cliente, MFG.Prodotto, MFG.Qta, I.ReteRagg " & _
'            "FROM MFG INNER JOIN daInviareVenditore AS I ON MFG.ODL = I.ODL " & _
'            "ORDER BY  I.ReteRagg, I.ODL"
'
'   Set rsDatiV = db.OpenRecordset(sqlc, dbOpenDynaset)
'
'Procedura temporanea
'********************************************************************
'
'         Set dbTempV = DAO.OpenDatabase(App.Path & "\..\Venditori\" & dbTempNameV)
'         Me.Refresh
'         DoEvents
'
'         sqlc = "SELECT IV.ODL, IV.Cliente, IV.Prodotto, IV.QtaSpedita " & _
'                  "FROM InvioVenditore AS IV"
'
'         Set rsTemp = dbTempV.OpenRecordset(sqlc, dbOpenDynaset)
'
'         sqlc = "SELECT * " & _
'                  "FROM InviatiVenditore AS InV"
'
'         Set rsInV = db.OpenRecordset(sqlc, dbOpenDynaset)
'
'      On Error Resume Next
'        Do While Not rsDatiV.EOF
'            rsTemp.AddNew
'            rsTemp("ODL") = rsDatiV("ODL")
'            rsTemp("Cliente") = rsDatiV("Cliente")
'            rsTemp("Prodotto") = rsDatiV("Prodotto")
'            rsTemp("QtaSpedita") = rsDatiV("Qta")
'            rsTemp.Update
'            If Err.Number <> 0 Then
'               Debug.Print "Err ", Err.Number, Err.Description
'               Err.Clear
'            End If
'            if err.Number <>3022 t
'
'            rsInV.AddNew
'            rsInV("ODL") = rsDatiV("ODL")
'            rsInV("Qta") = rsDatiV("Qta")
'            rsInV("ReteRagg") = rsDatiV("ReteRagg")
'            rsInV.Update
'            If Err.Number <> 0 Then
'               Debug.Print "Err ", Err.Number, Err.Description
'               Err.Clear
'            End If
'            if err.Number <>3022 t
'
'            rsDatiV.MoveNext
'
'        Loop
'   ' Chiudo il rsTEmp
'         rsTemp.Close
'         Set rsTemp = Nothing
'         rsInV.Close
'         Set rsInV = Nothing
''Fine Modifiche Temporanee
'______________________________________________________________________________

   

   
   
   
   
   '________________________________________________________________________________________
'Da utilizzarsi con la posta
'

'   If Not rsDatiV.EOF Then
'      sRetePrec = rsDatiV("ReteRagg")
'   Else
'      sRetePrec = "zzz"
'   End If
   
'   Do While Not rsDatiV.EOF
'      If rsDatiV("ReteRagg") <> sRetePrec Then
'  Invia le righe per sRetePrec
'         Call Spedisci(sSend, sRetePrec)
'         sSend = ""
'         sRetePrec = rsDatiV("ReteRagg")
'      End If
'      sSend = sSend & "X!" & sFieldSep & Trim$(rsDatiV("ODL")) & sFieldSep & _
'                 Trim$(rsDatiV("Cliente")) & sFieldSep & _
'                 Trim$(rsDatiV("Prodotto")) & sFieldSep & _
'                 Trim$(Str$(rsDatiV("Qta"))) & vbCrLf
'
'      Spedisci (rsDatiV(0))
'      rsDatiV.MoveNext
'   Loop
'
'  Spedisco gli ultimi messaggi
'
'   If sRetePrec <> "zzz" Then
'      Call Spedisci(sSend, sRetePrec)
'   End If
'
'
'_________________________________________________________________________________________________


' aggiorno i dati contenuti nella tab inviati con la data dell'invio
'   sqlc = "INSERT INTO InviatiVenditore ( ODL, Qta, ReteRagg, DataInvio ) " & _
'            "SELECT I.ODL, m.Qta, I.ReteRagg, Now() " & _
'            "FROM daInviareVenditore AS I inner join MFG as m  on i.odl = m.odl "
'   db.Execute (sqlc)
'   sqlc = "Delete * from daInviareVenditore"
'   db.Execute (sqlc)
   
'Chiudere tutti i RS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

   rsInput.Close
   rsConfronto.Close
'   rsDatiV.Close
   Set rsInput = Nothing
   Set rsConfronto = Nothing
'   Set rsDatiV = Nothing

'**********************************************************************************+

End Sub
Private Sub Spedisci(sBody As String, sRete As String)
   Dim sqlc As String, fRes As Boolean, rsEMail As Recordset
   Dim sRN As String, sREmail As String
   
   sqlc = "Select SMTPRecipientName, SMTPRecipientEMail " & _
            "From EMailxRete " & _
            "Where ReteRagg = '" & sRete & "'"
   Set rsEMail = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsEMail.EOF Then
      sRN = rsEMail(0)
      sREmail = rsEMail(1)
   Else
      sRN = sSmtpRecipientName
      sREmail = sSmtpRecipientEMail
   End If

   fRes = MailTo(sBody, sRN, sREmail)
   If Not fRes Then
      MsgBox "fallito invio per rete " & sRete
   End If

End Sub
Private Sub Form_Load()
   Dim sqlc As String, rsTemp As Recordset
   Dim sAppName As String, sFileNameIni As String
   Dim sTmp As String, nSize As Long, sKey As String, i As Integer
   Dim sNullString As String
   
   sAppName = App.EXEName & Chr$(0)
   
   sFileNameIni = App.Path & "\" & App.EXEName & ".ini" & Chr$(0)
   
   txtFileName.Text = App.Path & "\..\reclami\" & "resi282.txt"
'   txtFileName.Text = App.Path & "\" & "20000328.txt"
'CLP
'      txtFileName.Text = App.Path & "\..\" & "resi282.txt"

   lblProgress.Visible = False
   lblLog.Visible = False
   
   sNullString = ""
   
   Me.Show
'
   sqlc = "Select CampoMFG from CorrImportMFG order by CampoImport"
   Set rsTemp = db.OpenRecordset(sqlc, dbOpenSnapshot)
   i = 0
   Do While Not rsTemp.EOF
      ImpMfg(i) = rsTemp(0)
      rsTemp.MoveNext
      i = i + 1
   Loop
   rsTemp.Close
   Set rsTemp = Nothing
   
   sTmp = Space$(255)
   sKey = "SMTPServer" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               sNullString, sTmp, Len(sTmp), sFileNameIni)
   sSmtpServer = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPSenderName" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "", sTmp, Len(sTmp), sFileNameIni)
   sSmtpSenderName = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPSenderEMail" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "", sTmp, Len(sTmp), sFileNameIni)
   sSmtpSenderEMail = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPRecipientName" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "", sTmp, Len(sTmp), sFileNameIni)
   sSmtpRecipientName = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPRecipientEMail" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "", sTmp, Len(sTmp), sFileNameIni)
   sSmtpRecipientEMail = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPServerPort" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "25", sTmp, Len(sTmp), sFileNameIni)
   iSmtpServerPort = Val(Left$(sTmp, nSize))

   sTmp = Space$(255)
   sKey = "SMTPTimeout" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "10", sTmp, Len(sTmp), sFileNameIni)
   iSmtpTimeout = Val(Left$(sTmp, nSize))

   sTmp = Space$(255)
   sKey = "SMTPStartStr" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "$$**", sTmp, Len(sTmp), sFileNameIni)
   sSmtpStartStr = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPEndStr" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "**$$", sTmp, Len(sTmp), sFileNameIni)
   sSmtpEndStr = Left$(sTmp, nSize)

   sTmp = Space$(255)
   sKey = "SMTPSubject" & Chr$(0)
   nSize = GetPrivateProfileString(sAppName, sKey, _
               "$$RECLAMI$$", sTmp, Len(sTmp), sFileNameIni)
   sSmtpSubject = Left$(sTmp, nSize)

    cmdImport_Click


End Sub


Function MailTo(sBody As String, sRN, sREmail) As Boolean
   
   Dim dTimer As Double, fOk As Boolean, sReply As String
   Dim iLenBody As Integer, i As Integer, sBodyRow As String, sBodyTmp As String
   Dim iPos As Integer, iLenBodyTmp As Integer, iTmp As Integer
   
   fOk = True
   
   With wsk
      .Protocol = sckTCPProtocol 'inizializza WINSOCK con protocollo TCP
      .Connect sSmtpServer, iSmtpServerPort
   
      dTimer = Timer + iSmtpTimeout

      Do While Timer < dTimer And .State <> sckConnected
         DoEvents
      Loop
      If .State <> sckConnected Then
         fOk = False
         AddLog "Connessione NON effettuata "
         GoTo ExitMailTo
      End If
 
      dTimer = Timer + iSmtpTimeout
      Do While Timer < dTimer And InStr(sBuff, vbCrLf) = 0
         DoEvents
      Loop
      If Not GetReply(sReply) Then
         fOk = False
         AddLog "Connessione NON effettuata "
         GoTo ExitMailTo
      End If
      AddLog "Connessione effettuata "
'  il programma si 'presenta' al server inviando il proprio IP address
      .SendData "HELO " & .LocalIP & vbCrLf
      If Not GetReply(sReply) Then
         fOk = False
         AddLog "Helo Rifiutato "
         GoTo ExitMailTo
      End If
      AddLog "Helo OK "
'  specifica l'indirizzo da cui proviene la mail
      .SendData "MAIL FROM:<" & sSmtpSenderEMail & ">" & vbCrLf
      If Not GetReply(sReply) Then
         fOk = False
         AddLog "MAIL FROM Rifiutato "
         GoTo ExitMailTo
      End If
      AddLog "MAIL FROM Ok "
'  specifica l'indirizzo a cui inviare la mail
      .SendData "RCPT TO:<" & sREmail & ">" & vbCrLf
      If Not GetReply(sReply) Then
         fOk = False
         GoTo ExitMailTo
      End If
'  inizio del messaggio (comando DATA)
      fOkToSend = False
      .SendData "DATA" & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
      If Not GetReply(sReply) Then
         fOk = False
         GoTo ExitMailTo
      End If
'
'  Header
'
'  Date:
      fOkToSend = False
      .SendData "Date: " & Format$(Now, "dd mmm yy hh:mm:ss") & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
'  From:
      fOkToSend = False
      .SendData "From: " & sSmtpSenderName & " <" & sSmtpSenderEMail & ">" & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
'  Subject:
      fOkToSend = False
      .SendData "Subject: " & sSmtpSubject & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
'  To:
      fOkToSend = False
      .SendData "To: " & sRN & "<" & sREmail & "> " & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
'  Termina la header con vbcrlf & vbcrlf (riga nulla)
      fOkToSend = False
      .SendData vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
'  Invio del Corpo del messaggio
      fOkToSend = False
      iPos = InStr(1, sBody, vbCrLf)
      iTmp = 1
      Do While iPos > 0
         sBodyTmp = Mid$(sBody, iTmp, iPos - iTmp)
         i = 1
         iLenBody = Len(sBodyTmp)
         Do While i < iLenBody
            sBodyRow = sSmtpStartStr & Mid$(sBodyTmp, i, iMaxLenBodyRow) & sSmtpEndStr
            .SendData sBodyRow & vbCrLf
            Do Until fOkToSend
               DoEvents
            Loop
            i = i + iMaxLenBodyRow
         Loop
         iTmp = iPos + 2
         iPos = InStr(iTmp, sBody, vbCrLf)
      Loop
'*****************
'  termina il messaggio e lo invia con una riga  composta solo da un "."
      fOkToSend = False
      .SendData "." & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
      If Not GetReply(sReply) Then
         fOk = False
         GoTo ExitMailTo
      End If
'  chiude la connessione con il server
      fOkToSend = False
      .SendData "QUIT" & vbCrLf
      Do Until fOkToSend
         DoEvents
      Loop
      If Not GetReply(sReply) Then
         fOk = False
         GoTo ExitMailTo
      End If
   End With
ExitMailTo:
   wsk.Close
   MailTo = fOk
End Function

Private Sub wsk_SendComplete()
'  pone a true il flag di avvenuta consegna del messaggio
   fOkToSend = True
End Sub
Function GetReply(Reply As String) As Boolean
'  attende una risposta dal server, terminata con CR e/o LF
   Dim dTimer As Double, fOk As Boolean
   dTimer = Timer + iSmtpTimeout
   Do While Timer < dTimer _
         And InStr(sBuff, vbCr) = 0 And InStr(sBuff, vbLf) = 0
      DoEvents
   Loop
   fOk = False
   If InStr(sBuff, Chr$(13)) Then
      sBuff = Left$(sBuff, InStr(sBuff, vbCr) - 1)
      fOk = True
   End If
   If InStr(sBuff, vbLf) Then
      sBuff = Left$(sBuff, InStr(sBuff, vbLf) - 1)
      fOk = True
   End If
   If fOk Then
'  se tutto OK mette la stringa di risposta nella variabile Reply
      Reply = sBuff
   Else
      Reply = ""
   End If
   sBuff = ""
   GetReply = fOk
End Function

Private Sub AddLog(sText As String)
   txtLog.Text = txtLog.Text & sText & vbCrLf
   txtLog.SelStart = Len(txtLog.Text)
End Sub


Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
   Dim sInp As String
   sInp = ""
   wsk.GetData sInp, vbString
   sBuff = sBuff + sInp
End Sub

Private Sub wsk_Error(ByVal Number As Integer, Description As String, _
                      ByVal Scode As Long, ByVal Source As String, _
                      ByVal HelpFile As String, ByVal HelpContext As Long, _
                      CancelDisplay As Boolean)
   MsgBox "Errore nel controllo WinSock: " & vbCrLf & vbCrLf & _
            Str$(Number) & " - " & Description
End Sub




