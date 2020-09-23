VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "RedLof Cleaner By Syed Rizwan Muhammad Rizvi, Praise to Allah who has given Muslim Enough Brain . . . Optimized for Speed"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "RedlofCrack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstFiles 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   11535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Cleaning RedLof"
      Height          =   615
      Left            =   9840
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Infected Files : "
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label lblFolder 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO 'As New FileSystemObject
Dim WsShell 'As New WshShell
Dim winpath As String
Dim defaultid As String
Dim outlookversion As String
Dim temppath As String
Dim StartUpFile As String
Dim CflCnt As Long
Dim tFlcnt As Long
Private Sub Command1_Click()
On Error Resume Next
Command1.Enabled = False
lblMessage.Caption = "Removing Registry Keys"
winpath = FSO.GetSpecialFolder(0) & "\"
defaultid = WsShell.RegRead("HKEY_CURRENT_USER\Identities\Default User ID")
outlookversion = WsShell.RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Outlook Express\MediaVer")
temppath = ""
If Not (FSO.FileExists(winpath & "WScript.exe")) Then
temppath = "system32\"
End If
If temppath = "system32\" Then
StartUpFile = winpath & "SYSTEM\Kernel32.dll"
Else
StartUpFile = winpath & "SYSTEM\Kernel.dll"
End If
msg "WinPath : " & winpath
msg "DefaultID : " & defaultid
msg "Outlook Version : " & outlookversion
msg "TempPath : " & temppath
msg "Kernel : " & StartUpFile

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Compose Use Stationery"
WsShell.RegDelete "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Compose Use Stationery"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Stationery Name"
WsShell.RegDelete "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Stationery Name"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Wide Stationery Name"
WsShell.RegDelete "HKEY_CURRENT_USER\Identities\" & defaultid & "\Software\Microsoft\Outlook Express\" & Left(outlookversion, 1) & ".0\Mail\Wide Stationery Name"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Outlook\Options\Mail\EditorPreference"
WsShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Outlook\Options\Mail\EditorPreference"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Software\Microsoft\Windows Messaging Subsystem\Profiles\Microsoft Outlook Internet Settings\0a0d020000000000c000000000000046\001e0360"
WsShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows Messaging Subsystem\Profiles\Microsoft Outlook Internet Settings\0a0d020000000000c000000000000046\001e0360"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Microsoft Outlook Internet Settings\0a0d020000000000c000000000000046\001e0360"
WsShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Microsoft Outlook Internet Settings\0a0d020000000000c000000000000046\001e0360"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Outlook\Options\Mail\EditorPreference"
WsShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Outlook\Options\Mail\EditorPreference"

msg "Deleteing Key : " & "HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Common\MailSettings\NewStationery"
WsShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Common\MailSettings\NewStationery"

msg "Deleteing Key : " & "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32"
WsShell.RegDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\Kernel32"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\.dll\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\.dll\"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\.dll\Content Type"
WsShell.RegDelete "HKEY_CLASSES_ROOT\.dll\Content Type"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\dllfile\DefaultIcon\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\dllfile\DefaultIcon\"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\dllfile\ScriptEngine\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\dllfile\ScriptEngine\"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\dllFile\Shell\Open\Command\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\dllFile\Shell\Open\Command\"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\dllFile\ShellEx\PropertySheetHandlers\WSHProps\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\dllFile\ShellEx\PropertySheetHandlers\WSHProps\"

msg "Deleteing Key : " & "HKEY_CLASSES_ROOT\dllFile\ScriptHostEncode\"
WsShell.RegDelete "HKEY_CLASSES_ROOT\dllFile\ScriptHostEncode\"

lblMessage.Caption = "Restoring Old Files . . ."

If (FSO.FileExists(winpath & "web\kjwall.gif")) Then
msg "Restoring Old Folder.htt"
FSO.CopyFile winpath & "web\kjwall.gif", winpath & "web\folder.htt", True
End If
If (FSO.FileExists(winpath & "system32\kjwall.gif")) Then
msg "Restoring Old Desktop.ini"
FSO.CopyFile winpath & "system32\kjwall.gif", winpath & "system32\desktop.ini", True
End If


If FSO.FileExists(StartUpFile) Then
    lblMessage.Caption = "Analyzing Kernel . . ."
    msg "Checking Kernel For Virus . . ."
    If InStr(1, FSO.OpenTextFile(StartUpFile, 1).ReadAll, "KJ_start()") > 0 Then
        MsgBox "Your Kernel appears to be suspicious and it will be deleted", , "Kernel.dll or Kernel32.dll Under normall circumstances deleting wont affect ur system but still!"
        FSO.DeleteFile StartUpFile, True
        msg "Kernel Deleted . . ."
    Else
        msg "Kernel Not Deleted . . ."
    End If
End If

lblMessage.Caption = "Estimating No. of  Files . . ."
lstFiles.Clear
Dim dv 'As Scripting.Drive
Set dv = CreateObject("Scripting.Drive")
For Each dv In FSO.Drives
msg "Going to Evaluate Drive : " & dv.Name
    If dv.IsReady Then '(dv.DriveType = 2 Or dv.DriveType = 3 Or dv.DriveType = 1) And
        msg "Drive is Ready . . ."
        tFlcnt = 1
        CflCnt = 1
        est dv.DriveLetter & ":\"
        doClean dv.DriveLetter & ":\"
    End If
Next
MsgBox "Done Cleaning Files . . ."
Shell "Notepad.exe C:\RedLofCleanLog.log"
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Set FSO = CreateObject("Scripting.Filesystemobject")
Set WsShell = CreateObject("WScript.Shell")
If Not IsObject(FSO) Then
    MsgBox "RedLof Cleaner Was Unable to Find MS File System Object Which means redlof virus cannot run on your system either", vbCritical, "Fatal Error!"
    End
ElseIf Not IsObject(WsShell) Then
    MsgBox "RedLof Cleaner Was Unable to Find MS File System Object Which means redlof virus cannot run on your system either", vbCritical, "Fatal Error!"
    End
End If
Open "C:\RedLofCleanLog.log" For Output As #1
msg "RedLof Cleaner Started At : " & Now
tFlcnt = 1
CflCnt = 1
End Sub

Sub doClean(fldr As String)
On Error Resume Next
Dim x 'As File
Dim ts
Dim st As String
Dim ext As String
Set x = CreateObject("Scripting.File")
Set ts = CreateObject("Scripting.TextStream")

lblFolder.Caption = "Current Folder : " & fldr
msg "Scanning Folder : " & fldr
For Each x In FSO.GetFolder(fldr).Files
    ext = UCase(FSO.GetExtensionName(x.Path))
    CflCnt = CflCnt + 1
    lblDrive.Caption = "Drive - " & x.Drive.DriveLetter & ": " & CflCnt & " Files done out of : " & tFlcnt  'Round((CflCnt / tFlcnt) * 100) & "% "
    If x.Size >= 10000 Then
        If ext = "HTM" Or ext = "HTML" Or ext = "ASP" Or ext = "PHP" Or ext = "JSP" Or ext = "VBS" Or ext = "HTT" Then
    
    lblFile.Caption = "Current File : " & x.Name
    lblMessage.Caption = "Analyzing file : " & x.Path & "\" & x.Name
    If x.Size > 50000 Then lblFile.Caption = lblFile.Caption & " - This is a large file and may take several minutes to scan and clean"
    DoEvents
    Set ts = x.OpenAsTextStream(1)
    st = ts.ReadAll
    ts.Close
    If InStr(1, st, "KJ_start()") Then
    msg x.Path & "\" & x.Name & " Found Infected . . ."
        lstFiles.AddItem "Found and Cleaned : " & x.Path & "\" & x.Name
        st = Replace(st, "<" & "script language=vbscript>" & vbCrLf & "document.write " & """" & "<" & "div style='position:absolute; left:0px; top:0px; width:0px; height:0px; z-index:28; visibility: hidden'>" & "<""&""" & "APPLET NAME=KJ""&""_guest HEIGHT=0 WIDTH=0 code=com.ms.""&""activeX.Active""&""XComponent>" & "<" & "/APPLET>" & "<" & "/div>""" & vbCrLf & "<" & "/script>", "")
        st = Replace(st, "<" & "HTML>" & vbCrLf & "<" & "BODY onload=""" & "vbscript:" & "KJ_start()""" & ">", "")
        st = Replace(st, "<" & "BODY onload=""" & "vbscript:" & "KJ_start()""" & ">", "")
        st = Replace(st, vbCrLf & "<" & "HTML>" & vbCrLf & "<" & "BODY onload=""" & "vbscript:" & "KJ_start()""" & ">", "")
        st = Replace(st, "<" & "HTML>" & vbCrLf & "<" & "BODY onload=""" & "vbscript:" & "KJ_start()""" & ">" & vbCrLf, "")
        Dim OtherArr(3) As Integer
        Dim i1, i2, i3, i4, k1, k2
        Me.Tag = "N"
        For i1 = 0 To 9
        OtherArr(0) = i1
            For i2 = 0 To 9
            OtherArr(1) = i2
                For i3 = 0 To 9
                OtherArr(2) = i3
                    For i4 = 0 To 9
                    OtherArr(3) = i4
                        If Not st = Replace(st, "Execute(""Dim KeyArr(3),ThisText""&vbCrLf&""KeyArr(0) = " & OtherArr(0) & """&vbCrLf&""KeyArr(1) = " & OtherArr(1) & """&vbCrLf&""KeyArr(2) = " & OtherArr(2) & """&vbCrLf&""KeyArr(3) = " & OtherArr(3) & """&vbCrLf&""For i=1 To Len(ExeString)""&vbCrLf&""TempNum = Asc(Mid(ExeString,i,1))""&vbCrLf&""If TempNum = 18 Then""&vbCrLf&""TempNum = 34""&vbCrLf&""End If""&vbCrLf&""TempChar = Chr(TempNum + KeyArr(i Mod 4))""&vbCrLf&""If TempChar = Chr(28) Then""&vbCrLf&""TempChar = vbCr""&vbCrLf&""ElseIf TempChar = Chr(29) Then""&vbCrLf&""TempChar = vbLf""&vbCrLf&""End If""&vbCrLf&""ThisText = ThisText & TempChar""&vbCrLf&""Next"")" & vbCrLf & "Execute(ThisText)", "") Then
                            st = Replace(st, "Execute(""Dim KeyArr(3),ThisText""&vbCrLf&""KeyArr(0) = " & OtherArr(0) & """&vbCrLf&""KeyArr(1) = " & OtherArr(1) & """&vbCrLf&""KeyArr(2) = " & OtherArr(2) & """&vbCrLf&""KeyArr(3) = " & OtherArr(3) & """&vbCrLf&""For i=1 To Len(ExeString)""&vbCrLf&""TempNum = Asc(Mid(ExeString,i,1))""&vbCrLf&""If TempNum = 18 Then""&vbCrLf&""TempNum = 34""&vbCrLf&""End If""&vbCrLf&""TempChar = Chr(TempNum + KeyArr(i Mod 4))""&vbCrLf&""If TempChar = Chr(28) Then""&vbCrLf&""TempChar = vbCr""&vbCrLf&""ElseIf TempChar = Chr(29) Then""&vbCrLf&""TempChar = vbLf""&vbCrLf&""End If""&vbCrLf&""ThisText = ThisText & TempChar""&vbCrLf&""Next"")" & vbCrLf & "Execute(ThisText)", "")
                            Me.Tag = "F"
                            Exit For
                        End If
                    Next
                    If Me.Tag = "F" Then Exit For
                Next
                If Me.Tag = "F" Then Exit For
            Next
            If Me.Tag = "F" Then Exit For
        Next
        k1 = InStr(1, st, "<script language=vbscript>" & vbCrLf & "ExeString =")
        If k1 <> 0 Then k2 = InStr(k1, st, "</script>")
        If Not (k1 = 0 Or k2 = 0) Then st = Replace(st, Mid(st, k1, (k2 - k1) + 28), "")
        Set ts = x.OpenAsTextStream(2)
        ts.Write st
        ts.Close
        msg "File Cleaned Successfully"
    End If
    lblFile.Caption = "Current File : Scanning . . ."
    lblMessage.Caption = "Analyzing File : Scanning . . ."
    End If
End If
Next
Dim fld
Set fld = CreateObject("Scripting.Folder")
For Each fld In FSO.GetFolder(fldr).SubFolders
    If Right(fld.Path, 1) = "\" Then
        doClean fld.Path & fld.Name
    Else
        doClean fld.ParentFolder & "\" & fld.Name
    End If
Next
End Sub

Sub est(fldr As String)
On Error Resume Next
Dim fd 'As Scripting.Folder
Set fd = CreateObject("Scripting.Folder")
tFlcnt = tFlcnt + FSO.GetFolder(fldr).Files.Count
For Each fd In FSO.GetFolder(fldr).SubFolders
    If Right(fd.Path, 1) = "\" Then
        est fd.Path & fd.Name
    Else
        est fd.ParentFolder & "\" & fd.Name
    End If
Next
End Sub

Private Sub Form_Terminate()
Set FSO = Nothing
Set WsShell = Nothing
msg "RedLof Cleaner Was Closed at : " & Now
Close #1
Close 0
End Sub

Sub msg(ms As String)
Print #1, ms
End Sub
