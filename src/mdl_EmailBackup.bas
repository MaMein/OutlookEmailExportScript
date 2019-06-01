Attribute VB_Name = "mdl_EmailBackup"
Option Explicit


'Email subject prefixes (such us "RE:", "FW:" etc.) to be removed. Please note that this is a
'RegEx expression, google for "regex" for further information. For instance "\s" means blank " ".
Private Const EXM_OPT_CLEANSUBJECT_REGEX As String = "RE:\s|Re:\s|AW:\s|FW:\s|WG:\s|SV:\s|Antwort:\s"

Private Const BOOL_EXPORT_OVERRIDEEXISTINGFILES As Boolean = False
Private Const STRING_FILENAME_INHALTSVERZEICHNIS As String = "ExportInhaltsverzeichnis"

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' ****************************************************************************************************************************************************************
'   EMail Backup
' ****************************************************************************************************************************************************************
'   Markus Meinhard         13.05.2015              -   Erstellung des Moduls.
'   Markus Meinhard         23.07.2015              -   Modul zum laufen gebracht
'
' ****************************************************************************************************************************************************************
Sub Backup_Eigene_Ordner()

    ' Messzeit starten
    ' --------------------------------------------
    Dim beginTime As Variant
    beginTime = mdl_TimeMeas.TimerEx



    ' Exportieren
    ' -------------------------------------------
    Dim olFolder As Outlook.MAPIFolder
    'Set olFolder = Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders("Eigene Ordner").Folders("Troubleshooting")
    Set olFolder = Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders("Eigene Ordner")
    
    
    
    Dim strTargetFileForlder As String
    strTargetFileForlder = "c:\OutlookExport"
    
    
    ' Erzeuge Inhaltsvezeichnisdatei
    Dim strInhaltsverzeichnisDatei As String
    strInhaltsverzeichnisDatei = strTargetFileForlder & "\" & STRING_FILENAME_INHALTSVERZEICHNIS & ".txt"
    Open strInhaltsverzeichnisDatei For Output As #1
    Print #1, "Exportdatei vom: " & vbTab & vbTab & Now
    Print #1, "Exportierter Outlook Ordner: " & vbTab & vbTab & olFolder.Name
    Print #1, "--------------------------------------------------------------"
    Print #1, "Filename;Besitzer;Sender;Empfangen am;Empfänger;Betreff"
    
    
    ' Rekursiver durchlauf der outlook Verzeichnisse
    Call exportFolder(olFolder, strTargetFileForlder)


     ' Inhaltsverzeichnis Datei schliessen
    Close #1


    ' Messzeit Ermitteln und Ausgeben
    ' -------------------------------------------
    Dim endtime As Variant
    endtime = mdl_TimeMeas.TimerEx



    MsgBox "Export wurde angelegt in " & strTargetFileForlder & vbCrLf & _
            "Vergangene Zeit: " & CStr(endtime - beginTime) & " s"



End Sub



' Rekursiver aufruf
Private Sub exportFolder(olFolder As Outlook.Folder, strTargetPath As String)

    ' Abbrechen wenn kein Folder Objekt vorhanden
    If olFolder Is Nothing Then Exit Sub
    
    ' Zielpfadnamen für den Ordner erstellen
    Dim strFolderPath As String
    strFolderPath = strTargetPath & "\" & CleanString(olFolder.Name)
    Call createDirectory(strFolderPath & "\")

    
    
    
    'Exportiere alle EMail Items des aktuellen Folders
    Dim it As Variant
    Dim msg As Outlook.MailItem
    Dim boolSave As Boolean
    Dim strFilename As String
    Dim strEmailname As String
    For Each it In olFolder.Items
        
        'Debug.Print it.Class
        
        If it.Class = OlObjectClass.olMail Then
            
            Set msg = it
            
            ' Erzeuge Dateinamen für die zu exportierende nachricht
            strEmailname = getEmailFilename(msg)
            
            ' Beschneide Namen wenn länger wie 255 zeichen
            If Len(strFolderPath & "\" & strEmailname & ".msg") >= 255 Then
                strFilename = Left(strFolderPath & "\" & strEmailname, 245) & "..." & ".msg"
            Else
                strFilename = strFolderPath & "\" & strEmailname & ".msg"
            End If
            
                 
            ' prüfe, ob nachricht gespeichert werden soll
            boolSave = False
            If isFileExisting(strFilename) = False Then
                boolSave = True
            ElseIf isFileExisting(strFilename) = True And BOOL_EXPORT_OVERRIDEEXISTINGFILES = True Then
                boolSave = True
            Else
                boolSave = False
            End If
            
            ' Speichere Datei wenn gewünscht
            If boolSave = True Then
                'Save file
                msg.SaveAs strFilename, olMSG
                DoEvents
            End If
                    
                    
            ' Inhaltsverzeichnis informationen Schreiben
            Print #1, strFilename _
            & ";" & msg.ReceivedByName _
            & ";" & msg.SenderName _
            & ";" & msg.ReceivedTime _
            & ";" & getRecipientsAsString(msg) _
            & ";" & msg.Subject
                    
                    
            If msg Is Nothing Then
                Exit For
            Else
                Set msg = Nothing
            End If
              
              
        End If
          
    Next it



    ' Falls Unterordner vorhanden, zuerst diese Bearbeiten
    Dim subFolder As Outlook.MAPIFolder
    For Each subFolder In olFolder.Folders
        Debug.Print "[" & Now & "] " & strFolderPath & "/" & subFolder
        Call exportFolder(subFolder, strFolderPath)
    Next



   On Error Resume Next
   Set msg = Nothing


End Sub

Private Function getRecipientsAsString(msg As Outlook.MailItem) As String

    Dim strReturn As String

    Dim rec As Recipient
    For Each rec In msg.Recipients
        strReturn = strReturn & "<" & rec.Name & ">" & " "
    Next

    getRecipientsAsString = strReturn

End Function



Private Function getEmailFilename(ByVal msg As Outlook.MailItem) As String

    ' Erzeuge Dateinamen für die zu exportierende nachricht
    Dim strEmailname As String
    
    Dim datDatum As Date
    datDatum = CDate(msg.ReceivedTime)
    
    Dim strSubject As String
    strSubject = CleanString(msg.Subject)
    
    strEmailname = Format(datDatum, "yyyy_mm_dd___hh_mm_ss") & "___" & strSubject

    getEmailFilename = strEmailname

End Function


Private Function isFileExisting(strFilename As String) As Boolean

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    isFileExisting = objFSO.FileExists(strFilename)

    On Error Resume Next
    Set objFSO = Nothing


End Function


Private Function createDirectory(strFoldername As String)

    Dim astrFolderNames As Variant
    ' Nur Backslashes erlaubt
    strFoldername = Replace(strFoldername, "/", "\")
    
    astrFolderNames = Split(strFoldername, "\")
        
    ' Verzeichnispfad
    astrFolderNames(1) = astrFolderNames(0) & "\" & astrFolderNames(1)
    
    Dim strTargetName As String
    strTargetName = astrFolderNames(1)
    Dim Folder As Variant
    For Each Folder In astrFolderNames
        
        If InStr(1, Folder, ":") = 0 Then
            On Error Resume Next
            Call MkDir(strTargetName)
            
            ' Nächsten Ebene Erzeugen
            strTargetName = strTargetName & "\" & Folder
        End If
    
    
    Next
    

End Function


Private Function CleanString(strData As String) As String

    Const PROCNAME As String = "CleanString"

    On Error GoTo ErrorHandler

    'Instantiate RegEx
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True

    'Cut out strings we don't like
    objRegExp.Pattern = EXM_OPT_CLEANSUBJECT_REGEX
    strData = objRegExp.Replace(strData, "")

    'Replace and cut out invalid strings.
    strData = Replace(strData, Chr(9), "_")
    strData = Replace(strData, Chr(10), "_")
    strData = Replace(strData, Chr(13), "_")
    strData = Replace(strData, " ", "_")
    objRegExp.Pattern = "[/\\*]"
    strData = objRegExp.Replace(strData, "-")
    objRegExp.Pattern = "[""]"
    strData = objRegExp.Replace(strData, "'")
    objRegExp.Pattern = "[:?<>\|]"
    strData = objRegExp.Replace(strData, "")
    
    'Replace multiple chars by 1 char
    objRegExp.Pattern = "\s+"
    strData = objRegExp.Replace(strData, " ")
    objRegExp.Pattern = "_+"
    strData = objRegExp.Replace(strData, "_")
    objRegExp.Pattern = "-+"
    strData = objRegExp.Replace(strData, "-")
    objRegExp.Pattern = "'+"
    strData = objRegExp.Replace(strData, "'")
            
    'Trim
    strData = Trim(strData)
    
    'Return result
    CleanString = strData
  
ExitScript:
    On Error Resume Next
    Set objRegExp = Nothing
    Exit Function
ErrorHandler:
    CleanString = "ERROR_OCCURRED:" & "Error #" & Err & ": " & Error$ & " (Procedure: " & PROCNAME & ")"
    Resume ExitScript
End Function











'
'' ============================================================
'
'
'Sub BeispielCode_olSaveAttachments()
'
'    Dim olFolder As Outlook.MAPIFolder
'    Dim msg As Outlook.MailItem
'    Dim msg2 As Outlook.MailItem
'    Dim att As Outlook.Attachment
'    Dim strFilePath As String
'    Dim strTmpMsg As String
'    Dim fsSaveFolder As String
'
'    fsSaveFolder = "C:\test\"
'
'    'path for creating attachment msg file for stripping
'    strFilePath = "C:\temp\"
'    strTmpMsg = "KillMe.msg"
'
'   'My testing done in Outlok using a "temp" folder underneath Inbox
'    Set olFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
'    Set olFolder = olFolder.Folders("Temp")
'    If Not olFolder Is Nothing Then
'
'        For Each msg In olFolder.Items
'            If msg.Attachments.Count > 0 Then
'            While msg.Attachments.Count > 0
'            bflag = False
'                If Right$(msg.Attachments(1).FileName, 3) = "msg" Then
'                    bflag = True
'                    msg.Attachments(1).SaveAsFile strFilePath & strTmpMsg
'                    Set msg2 = Application.CreateItemFromTemplate(strFilePath & strTmpMsg)
'                End If
'                If bflag Then
'                    sSavePathFS = fsSaveFolder & msg2.Attachments(1).FileName
'                    msg2.Attachments(1).SaveAsFile sSavePathFS
'                    msg2.Delete
'                Else
'                    sSavePathFS = fsSaveFolder & msg.Attachments(1).FileName
'                    msg.Attachments(1).SaveAsFile sSavePathFS
'                End If
'                msg.Attachments(1).Delete
'                Wend
'                 msg.Delete
'            End If
'        Next
'
'    End If
'
'    On Error Resume Next
'    Set olFolder = Nothing
'    Set msg = Nothing
'    Set msg2 = Nothing
'    Set att = Nothing
'
'
'End Sub
'
'Sub old_EMailBackup()
'
'    Dim olFolder As Outlook.MAPIFolder
'    Dim msg As Outlook.MailItem
'    Dim msg2 As Outlook.MailItem
'    Dim att As Outlook.Attachment
'    Dim strFilePath As String
'    Dim strTmpMsg As String
'    Dim fsSaveFolder As String
'
'    fsSaveFolder = "C:\test\"
'
'    'path for creating attachment msg file for stripping
'    strFilePath = "C:\temp\"
'    strTmpMsg = "KillMe.msg"
'
'   'My testing done in Outlok using a "temp" folder underneath Inbox
'    Set olFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
'    Set olFolder = olFolder.Folders("Eigene Ordner")
'
'
'
'
'
'
'    Dim strFoldername As String
'    strFoldername = CleanString(olFolder.Name)
'
'    For Each msg In olFolder.Items
'
'        Dim strTargetPath As String
'        strTargetPath = strFilePath & msg.fo
'
'
'
'
'    Next
'
'
'    For Each msg In olFolder.Items
'        If msg.Attachments.Count > 0 Then
'        While msg.Attachments.Count > 0
'        bflag = False
'            If Right$(msg.Attachments(1).FileName, 3) = "msg" Then
'                bflag = True
'                msg.Attachments(1).SaveAsFile strFilePath & strTmpMsg
'                Set msg2 = Application.CreateItemFromTemplate(strFilePath & strTmpMsg)
'            End If
'            If bflag Then
'                sSavePathFS = fsSaveFolder & msg2.Attachments(1).FileName
'                msg2.Attachments(1).SaveAsFile sSavePathFS
'                msg2.Delete
'            Else
'                sSavePathFS = fsSaveFolder & msg.Attachments(1).FileName
'                msg.Attachments(1).SaveAsFile sSavePathFS
'            End If
'            msg.Attachments(1).Delete
'            Wend
'             msg.Delete
'        End If
'    Next
'
'
'
'End Sub
