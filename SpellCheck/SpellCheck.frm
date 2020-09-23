VERSION 5.00
Begin VB.Form frmSpellCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Spell Checker Using Soundex and Levinshtein Distance"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SpellCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1065
   End
   Begin VB.CommandButton cmdCreateList 
      Caption         =   "Generate Word List"
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstLdMatches 
      Height          =   2205
      Left            =   2940
      TabIndex        =   8
      Top             =   180
      Width           =   2775
   End
   Begin VB.HScrollBar LdVal 
      Height          =   235
      Left            =   105
      TabIndex        =   4
      Top             =   1065
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtInputWord 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   60
      X2              =   2840
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   60
      X2              =   2840
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label lblLDMatchCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1470
      TabIndex        =   10
      Top             =   825
      Width           =   1185
   End
   Begin VB.Label lblLDVal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2340
      TabIndex        =   9
      Top             =   3705
      Width           =   375
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   2115
      Width           =   2550
   End
   Begin VB.Label lblAccuracy 
      Caption         =   "Adjust accuracy"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblEnterWord 
      Caption         =   "Enter a word"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmSpellCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form requires the following type library:

' Reference=*\G{C878CB53-7E75-4115-BD13-EECBC9430749}#1.1#0#MemAPIs.tlb#Memory APIs

' Or uncomment the following declares:

'Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal lByteLen As Long)
'Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
'Private Declare Function AllocStrSpPtr Lib "oleaut32" Alias "SysAllocStringLen" (ByVal lStrPtr As Long, ByVal lLen As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpMC As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal P3 As Long, ByVal P4 As Long) As Long
Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As VbFileAttribute
Private Declare Function GetInputState Lib "user32" () As Long

Private Const INVALID_FILE_ATTRIBUTES As Long = &HFFFFFFFF

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private StackLB() As Long, StackUB() As Long, StackSize As Long  ' Pending boundary stacks
Private TwisterBuf() As Long, TwisterBufSize As Long             ' Twister copymemory buffer
Private mMethod As VbCompareMethod

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Public Enum eOptimal
    ReversePretty = 1&     '0.001<<reverse-pretty-reverse>>0.002
    ReverseOptimal = 2&    '0.002<<reverse-sorting-unsorted>>0.003
    RefreshUnsorted = 3&   '0.003<<unsorted-refresh-sorting>>0.004
    #If False Then
        Dim ReversePretty, ReverseOptimal, RefreshUnsorted
    #End If
End Enum

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Const SUBooRANGE As Long = &H9
Private Const MAX_PATH As Long = 260

Private aWords() As String
Private lA_sdx() As Long
Private lA_rev() As Long
Private lA_len() As Long

Private arrLDMatches() As Long
Private arrLDItemCnt() As Long

'Private strClosestMatch As String
Private sAppPathSlash As String

Private mTag As Long
Private mCnt As Long

Private Sub SortWordsFile()
    lblStatus.Caption = " Sorting..."
    lblStatus.Refresh
    mMethod = vbTextCompare
    TwisterStringSort aWords, 1&, mCnt
End Sub

Public Property Get CorrectedWord() As String
    CorrectedWord = txtInputWord
End Property

Public Property Let CorrectedWord(sSpellWord As String)
    txtInputWord = sSpellWord
'    If Clipboard.GetFormat(vbCFText) Then
'        Text1.SelText = Clipboard.GetText(vbCFText)
'    End If
End Property

Private Sub Form_Load()
    sAppPathSlash = App.Path
    If Right$(sAppPathSlash, 1) <> "\" Then sAppPathSlash = sAppPathSlash & "\"
    LdVal.Enabled = False
    Me.Show
    Me.Refresh
    If IsValidData Then
        txtInputWord.Visible = True
        txtInputWord.SetFocus
        lblEnterWord.Visible = True
        lblAccuracy.Visible = True
        LdVal.Visible = True
        lblStatus.Caption = " Preparing..."
        Me.Refresh
        GetDataFiles
        GetWordsFile
        lblStatus.Caption = " Ready..."
        lblStatus.Refresh
    Else
        cmdCreateList.Visible = True
        cmdCreateList.SetFocus
    End If
End Sub

Private Function IsValidData() As Boolean
    Dim lFileLen As Long ' Enable updating of new words file
    If FileExists(sAppPathSlash & "words.dat") Then
        If FileExists(sAppPathSlash & "sdx_r.dat") Then
            lFileLen = CLng(GetSetting(App.EXEName, "Words Data", "File Size", "-1"))
            IsValidData = (lFileLen = FileLen(sAppPathSlash & "words.dat"))
        End If
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Must Unload this form at program termination
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "Close" Then
        ' Must Unload this form at program termination
        Me.Hide

'MsgBox CorrectedWord
'Unload Me

    Else
        cmdClose.Caption = "Close"
        mTag = (mTag Mod MAX_PATH) + 1&

        lblStatus.Caption = " Processing stopped..."
        lblStatus.Refresh

    End If
End Sub

Private Sub cmdSelect_Click()
    ' Must Unload this form at program termination
    Me.Hide
    If lstLdMatches.SelCount Then txtInputWord = lstLdMatches.Text

'    Clipboard.Clear
'    Clipboard.SetText txtInputWord, vbCFText

'MsgBox CorrectedWord
'Unload Me
End Sub

Private Sub lstLdMatches_DblClick()
    ' Must Unload this form at program termination
    Me.Hide
    txtInputWord = lstLdMatches.Text 'lstLdMatches.List(lstLdMatches.ListIndex)

'    Clipboard.Clear
'    Clipboard.SetText txtInputWord, vbCFText

'MsgBox CorrectedWord
'Unload Me
End Sub

Private Sub cmdCreateList_Click()
    Dim lFileLen As Long
    On Error GoTo ErrHandler
    cmdCreateList.Enabled = False
    lblStatus.Caption = " Preparing..."
    lblStatus.Refresh
    If Not FileExists(sAppPathSlash & "words.dat") Then
        MsgBox "Place your flat file database of words in the app path and name it words.dat"
        Exit Sub
    End If
    mCnt = 0&
    GetWordsFile
    GenerateData
    SaveDataFiles
    lFileLen = FileLen(sAppPathSlash & "words.dat")
    SaveSetting App.EXEName, "Words Data", "File Size", CStr(lFileLen)
    cmdCreateList.Visible = False
    txtInputWord.Visible = True
    txtInputWord.SetFocus
    lblEnterWord.Visible = True
    lblAccuracy.Visible = True
    LdVal.Visible = True
ErrHandler:
End Sub

Private Sub txtInputWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0&
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0&
        Beep
    End If
End Sub

Private Sub txtInputWord_Change()
    Dim strInput As String
    strInput = LCase$(Trim$(txtInputWord.Text))
    If LenB(strInput) = 0& Then Exit Sub
    mTag = (mTag Mod MAX_PATH) + 1&
    cmdClose.Caption = "Stop"
    Call GetMatches(strInput, mTag)
End Sub

Private Sub GenerateData()
    Screen.MousePointer = vbHourglass

    lblStatus.Caption = " Generating..."
    lblStatus.Refresh

    ReDim lA_sdx(1 To mCnt) As Long
    ReDim lA_rev(1 To mCnt) As Long
    ReDim lA_len(1 To mCnt) As Long

'    Dim tmpStr As String, i As Long
'    Dim lpStr As Long, lp As Long
'
'    lpStr = VarPtr(tmpStr)       ' Cache pointer to the string variable
'    lp = VarPtr(aWords(1&)) - 4& ' Cache pointer to the array (one based)
'
'    For i = 1& To mCnt
'        '// Read words from the array and add to the database
'        'tmpStr = aWords(i) ' Grab current value into variable
'        CopyMemByV lpStr, lp + (i * 4&), 4&
'
'        lA_sdx(i) = GetSoundexWord(tmpStr)
'        lA_rev(i) = GetSoundexWordR(tmpStr)
'        lA_len(i) = Len(tmpStr)
'
''        '// Prevent the UI from freezing up
''        If i Mod 1000 = 0& Then
''            lblStatus.Caption = " Adding..." & i & " words"
''            lblStatus.Refresh
''        End If
'    Next i
'
'    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to variable

    Dim lpWords As Long, lpSndex As Long
    Dim lpRevSx As Long, lpArLen As Long

    lpWords = VarPtr(aWords(1))
    lpSndex = VarPtr(lA_sdx(1))
    lpRevSx = VarPtr(lA_rev(1))
    lpArLen = VarPtr(lA_len(1))
    lA_len(1) = mCnt

' Please do not edit the following machine code
Const GENSNDX As String = "žÕ‹ìƒìˆæœÖ×ÒÑÓ³›À‰Åø‹ý”•‹Ÿ‰Ýü»ÅÙüô¹Áà‚‹Øýˆƒø‹·…éöô¢‹Îü‹ÙÑë‹ý”ƒ‹ø‰Ÿ‹ýƒÃøè®‚€€‹˜ýŒƒøè–€Ô€€‹ÅøÀ‰ŽÅøë" & _
                          "ÂÛÙÚ¸ßÞæ‹åÝ©Â€‹Æü³¥É³Ûæ³Ò»È„î€€€™Š´±€þàã‡ã€€€€þŒÁ„Ñ€€€³€þÂ„„Ì€€€þÃ„±ˆ€€€þÄÆ„Œ€€€™þÅ„­€€" & _
                          "¦€€þÆ„à˜€€€þÇá„ä€€€þŒÈ„’€€€³€þÉ„‰€Ì€€€þÊ„°É€€€þËÂ„À€€€‰þÌ„Ñ€¦€€þÍ„Õ˜€€€þÎá„Ì€€€þ˜Ïô×" & _
                          "€þÐá„Š€€€þ„Ñ„Ž€€“€þÒ„¹Ì€€€þÓ„±ü€€€€þÔÂ„€€€€±þÕô¥€þÖÆ„Ø€€€€±þ×ô—€þØÆ„×€€€€±þÙô‰€þÚÆ„É€€€ƒ" & _
                          "õÁ‚éŠÿÿÿçÃ€þáôò€™þâ„¥€€æ€€þã„©˜€€€€þäã„­€€€€þœåôÒ€þæã„…€€€€þŒç„‰€€€ó€þèô»€þœéô¶€þêôÆú€þëôõ€™þì" & _
                          "„†€€æ€€þí„Š˜€€€€þîã„€€€€þœïôŒ€þðôÆÃ€þñôË€±þòôú€þóŒôÁ€þôôÉÓ€þõ„êÿÿÿ€þöô¡Ó€þ÷„Üÿÿÿ€þøô Ó€þù„Îÿ" & _
                          "ÿÿ€þúô’ýéÄÿÿÿ€òô„»ÿÿÿµ²ë¿€ò‚ú„®ÿÿÿ²š‚ë²€òƒ½„¡ÿÿÿ²ƒÍë¥€ò„„Þ”ÿÿÿ²„ë¦˜€ò…„‡¯ÿÿÿ²…ë‹ó€ò†„úþ—ÿÿ²" & _
                          "†ˆ”»æÃƒû„‚ë¯þÿÿÃ³ÛæÖ³ÒëÃçã—üƒé‚Š´±³€þà‡ß€Ì€€€þÁôé³€þÂ„˜Ì€€€þÃ„±œ€€€þÄÆ„ €€€¹þÅôÉ€þÆÂ„ø€€€" & _
                          "‰þÇ„ü€æ€€þÈô²€¹þÉô­€þÊÂ„é€€€‰þË„à€¦€€þÌ„ñ˜€€€þÍá„õ€€€þ„Î„ì€€Ó€þÏ„÷ÿÏÿÿ€þÐ„°¦€€€þÑÂ„ª" & _
                          "€€€‰þÒ„Õ€¦€€þÓ„˜˜€€€þÔá„œ€€€þôÕ„Áÿÿÿ³€þÖ„ð€Ì€€€þ×„¾¯ÿÿÿ€þØÆ„ë€€€€éþÙ„ÿÿçÿ€þÚ„Ùè€€€éÿÿ" & _
                          "§ÿ€þá„†Ÿÿÿÿ€þâã„µ€€€€þŒã„¹€€€³€þä„½€Ì€€€þå„¿âþÿÿ€þæÆ„‘€€€€™þç„•€€æ€€þè„ÇŸþÿÿ€þéÿ„¾þÿÿ€þ˜êôþ" & _
                          "€þëôæù€þì„Š˜€€€€þíã„Ž€€€€þŒî„…€€€ó€þï„þÿÿ€þðôÃã€þñôË€þ˜òôú€þóôÆÁ€þôôÉ€éþõ„îþÿÇÿ€þöô¡€éþ÷„àþÿ" & _
                          "Çÿ€þøô €éþù„ÒþÿÇÿ€þúô’é¾Èþÿÿ€òú„«ÿÿÿ²šë¿€ò‚½„žÿÿÿ²‚Íë²€òƒ„Þ‘ÿÿÿ²ƒë¦¥€ò„„„¯ÿÿÿ²„ë˜ó€ò…„÷þ×ÿÿ²" & _
                          "…ë‹€ùò†„êþÿ‹ÿ²†ˆ”»Ãóƒû„‚Ûþ‡ÿÿÃ"
    Dim rc As Long
    Dim lpMCsndx As Long
    Dim abMCode() As Byte

    AsmEscStr2Bin abMCode, GENSNDX
    lpMCsndx = VarPtr(abMCode(0))

    rc = CallWindowProc(lpMCsndx, lpWords, lpSndex, lpRevSx, lpArLen)

    lblStatus.Caption = " Database created - " & rc & " words added"
    Screen.MousePointer = vbDefault
    Me.Refresh
End Sub

Sub AsmEscStr2Bin(b() As Byte, sEscStr As String) ' Original code jeremyxtz
    Dim cmeta As Long, remainder As Long
    Dim c7Bs As Long, i As Long, j As Long
    Dim power As Byte, meta As Byte

    j = Len(sEscStr)    ' mbbbbbbbmbbbbbbbmbbbbbbbmbb
    i = (j \ 8)         ' mbbbbbbbmbbbbbbbmbbbbbbb
    remainder = j Mod 8 ' mbb
    c7Bs = i * 7        ' bbbbbbbbbbbbbbbbbbbbb

    c7Bs = c7Bs - 1&         ' Account for zero based array
    If remainder = 0& Then
        ReDim b(c7Bs) As Byte     ' No remainder bytes (7 is a divisor of cnt)
    Else                          ' Remove remainder meta byte from count
        remainder = remainder - 1&
        ReDim b(c7Bs + remainder) ' bb
    End If

    meta = Asc(Mid$(sEscStr, 1&, 1&)) ' Assign bit flags for first set of 7 machine code bytes
    cmeta = 2&                        ' Set read pointer, dodge first meta byte

    For i = 0 To c7Bs + remainder     ' Unexcape all machine code bytes into array
        If meta And (2 ^ power) Then  ' I already extended then assign it as is
            b(i) = Asc(Mid$(sEscStr, i + cmeta, 1&))
        Else                          ' Else assign it with MSBit masked out
            b(i) = Asc(Mid$(sEscStr, i + cmeta, 1&)) And (Not 128)
        End If
        power = (i + 1) Mod 7
        If power = 0 Then             ' If set complete grab next set of meta bits
            cmeta = cmeta + 1&        ' Skip the meta byte
            If i + cmeta > j Then Exit For ' No remainder condition (so no meta byte)
            meta = Asc(Mid$(sEscStr, i + cmeta, 1&))
        End If
    Next
End Sub

Private Function SaveTextFile(sFileSpec As String, sText As String) As Long
    On Error GoTo SaveFileError
    Dim iFile As Integer
    iFile = FreeFile
    Open sFileSpec For Output Access Write Lock Write As #iFile
      Print #iFile, sText;
SaveFileError:
    Close #iFile
    SaveTextFile = Err
End Function

Private Sub GetWordsFile()
    Dim iFile As Integer, numcount As Long
    Dim sFile As String, iLen As Long
    Dim iSubLen As Long, lCnt As Long
    Dim idx1 As Long, idx2 As Long
    Dim lA() As Long
    Const PBrk As String = vbCrLf & vbCrLf

    On Error GoTo ErrorHandler
    iFile = FreeFile
    Screen.MousePointer = vbHourglass

    ' Open in binary mode
    Open sAppPathSlash & "words.dat" For Binary Access Read Lock Write As #iFile

        ' Get the data length
        iLen = LOF(iFile)
        
        ' Allocate the length first: sFile = Space$(iLen)
       'CopyMemByV VarPtr(sFile), VarPtr(AllocStrSpPtr(0&, iLen)), 4&
        sFile = AllocStr(vbNullString, iLen)

        ' Get the file in one chunk
        Get #iFile, 1&, sFile
        ' sFile = File text converted to Unicode

    Close #iFile ' Close the file

    If Not mCnt = 0& Then
        ReDim aWords(1& To mCnt) As String
        idx1 = 1&
        ' Read words from the file and add to the array
        For numcount = 1& To mCnt
            iSubLen = lA_len(numcount)
           'CopyMemByV VarPtr(aWords(numcount)), VarPtr(AllocStrSpPtr(0&, iSubLen)), 4&
            aWords(numcount) = AllocStr(vbNullString, iSubLen)
            Mid$(aWords(numcount), 1&) = Mid$(sFile, idx1, iSubLen)
            idx1 = idx1 + iSubLen + 2&
        Next
    Else
        If (InStr(1&, sFile, vbCrLf) = 0&) Then ' Fix UNIX line ends
            sFile = Replace$(sFile, vbLf, vbCrLf)
        End If
        If (InStr(1&, sFile, ",") <> 0&) Then    ' Convert comma delimited
            sFile = Replace$(sFile, ",", vbCrLf)
        End If
        Do While Not (InStr(1&, sFile, PBrk) = 0&) ' Remove blank lines
            sFile = Replace$(sFile, PBrk, vbCrLf)
        Loop
        If Not (iLen = Len(sFile)) Then ' If file has changed
            SaveTextFile sAppPathSlash & "words.dat", sFile
            iLen = Len(sFile)
        End If
        lCnt = 1000000
        ReDim lA(1& To lCnt) As Long
        idx1 = 1&
        ' Index words from the file
        Do While (idx1 < iLen)
            idx2 = InStr(idx1, sFile, vbCrLf)
            If (idx2 = 0&) Then idx2 = iLen + 1&
            numcount = numcount + 1&
            lA(numcount) = idx2
            idx1 = idx2 + 2&
        Loop
        lCnt = numcount
        ReDim aWords(1& To lCnt) As String
        idx1 = 1&
        ' Read words from the file and add to the array
        For numcount = 1& To lCnt
            idx2 = lA(numcount)
            iSubLen = idx2 - idx1
           'CopyMemByV VarPtr(aWords(numcount)), VarPtr(AllocStrSpPtr(0&, iSubLen)), 4&
            aWords(numcount) = AllocStr(vbNullString, iSubLen)
            Mid$(aWords(numcount), 1&) = LCase$(Mid$(sFile, idx1, iSubLen))
            idx1 = idx2 + 2&
        Next
        mCnt = lCnt
        SortWordsFile
        SaveTextFile sAppPathSlash & "words.dat", Join(aWords, vbCrLf)
    End If

ErrorHandler:
    Screen.MousePointer = vbDefault
    If Err = SUBooRANGE Then
        lCnt = lCnt + 100000
        ReDim Preserve lA(1& To lCnt) As Long
        Resume
    ElseIf Err Then
        Close #iFile ' Close the file
        MsgBox "Error - " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub SaveDataFiles()
    On Error GoTo Abort
    Dim iFile As Integer

    Screen.MousePointer = vbHourglass

    ' Clear the file
    iFile = FreeFile
    Open sAppPathSlash & "sdx_f.dat" For Output As #iFile
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "sdx_f.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_sdx()
    Close #iFile

    ' Clear the file
    iFile = FreeFile
    Open sAppPathSlash & "sdx_r.dat" For Output As #iFile
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "sdx_r.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_rev()
    Close #iFile

    ' Clear the file
    iFile = FreeFile
    Open sAppPathSlash & "len_w.dat" For Output As #iFile
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "len_w.dat" For Binary Access Write Lock Write As #iFile
        Put #iFile, 1&, lA_len()
Abort:
    Close #iFile
    Screen.MousePointer = vbDefault
End Sub

Private Sub GetDataFiles()
    On Error GoTo Abort
    Dim iFile As Integer
    Screen.MousePointer = vbHourglass

    mCnt = FileLen(sAppPathSlash & "len_w.dat") \ 4&

    ReDim lA_sdx(1& To mCnt) As Long
    ReDim lA_rev(1& To mCnt) As Long
    ReDim lA_len(1& To mCnt) As Long

    iFile = FreeFile
    Open sAppPathSlash & "sdx_f.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_sdx()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "sdx_r.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_rev()
    Close #iFile

    iFile = FreeFile
    Open sAppPathSlash & "len_w.dat" For Binary Access Read Lock Write As #iFile
        Get #iFile, 1&, lA_len()
Abort:
    Close #iFile ' Close the file
    Screen.MousePointer = vbDefault
End Sub

Private Function GetMatches(strInput As String, ByVal lTag As Long) As Long
    Dim lMatches() As Long, strMatch As String
    Dim lSndex As Long, lRevSndex As Long
    Dim lenTmp As Long, LD As Long, LdMax As Long
    Dim Index As Long, Total As Long
    Dim lpStr As Long, lp As Long

    On Error GoTo ExitSub
    GetMatches = -1&

    lpStr = VarPtr(strMatch)     ' Cache pointer to the string variable
    lp = VarPtr(aWords(1&)) - 4& ' Cache pointer to the array (one based)

    lblStatus.Caption = " Processing items..."
    lblStatus.Refresh

    '// Get the soundex of the input word
    lSndex = GetSoundexWord(strInput)

    '// Get the soundex of the input word reversed
    lRevSndex = GetSoundexWordR(strInput)

    If GetInputState <> 0& Then DoEvents
    If Not lTag = mTag Then GoTo ExitSub

    ReDim lMatches(1& To mCnt) As Long
    
    '// Find all entries in the database which match the soundex of the input word
    For Index = 1& To mCnt
        If GetInputState <> 0& Then DoEvents
        If Not lTag = mTag Then GoTo ExitSub

        If (lA_sdx(Index) = lSndex) Or (lA_rev(Index) = lRevSndex) Then
            Total = Total + 1&
            lMatches(Total) = Index

            '// The max length of returned matches is the
            '// maximum Leveshtein distance we can have
            lenTmp = lA_len(Index)
            If lenTmp > LdMax Then LdMax = lenTmp
        End If
    Next

    lstLdMatches.Clear
    lblLDMatchCount.Caption = "0"

    ReDim arrLDMatches(LdMax, Total) As Long
    ReDim arrLDItemCnt(LdMax) As Long
    LdMax = 0&

    '// Walk through all soundex matches
    For Index = 1& To Total
        If GetInputState <> 0& Then DoEvents
        If Not lTag = mTag Then GoTo ExitSub

        '// The one and only access to the strings
        'strMatch = aWords(lMatches(Index)) ' Grab current word
        CopyMemByV lpStr, lp + (lMatches(Index) * 4&), 4&

        '// Get all Levenshtein distances
        LD = GetLevenshteinDistance(strInput, strMatch)

        '// Add better matches higher up
        arrLDMatches(LD, arrLDItemCnt(LD)) = lMatches(Index)
        arrLDItemCnt(LD) = arrLDItemCnt(LD) + 1&

        'If LD = 0& Then Exit For ' Perfect match found

        '// Record maximum Leveshtein distance
        If LD > LdMax Then LdMax = LD
    Next

    If GetInputState <> 0& Then DoEvents
    If Not lTag = mTag Then GoTo ExitSub

    lblStatus.Caption = " " & Total & " items found"
    lblStatus.Refresh

    '// Determine lowest Levenshtein distance that produced matches
    For Index = 0& To LdMax
        If arrLDItemCnt(Index) <> 0& Then Exit For
    Next
    GetMatches = Index ' Return index of best match (if zero then a perfect match/valid word)

'    If arrLDItemCnt(Index) = 1& Then
'        strClosestMatch = aWords(arrLDMatches(Index, 0&))
'    Else
'        strClosestMatch = vbNullString
'    End If

    With LdVal
        .Min = Index
        .Max = LdMax
         ' Call Change if Value will not change
        If .Value = Index Then LdVal_Change Else .Value = Index
        .Enabled = True
    End With

    cmdClose.Caption = "Close"
ExitSub:

    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to variable
End Function

Private Sub LdVal_Change()
    Dim i As Long, j As Long

    On Error GoTo ExitSub
    lstLdMatches.Clear

    '// Add all Levenshtein distances up to the scroll(threshold) value
    For i = 0& To LdVal.Value
        For j = 0& To arrLDItemCnt(i) - 1&
            lstLdMatches.AddItem aWords(arrLDMatches(i, j))
        Next
    Next

    lblLDMatchCount.Caption = lstLdMatches.ListCount & " matches"

    LdVal_Scroll
ExitSub:
End Sub

Private Sub LdVal_Scroll()
    lblLDVal.Caption = LdVal.Value
End Sub

Private Function FileExists(sFileSpec As String) As Boolean
    Dim Attribs As VbFileAttribute
    Attribs = GetAttributes(sFileSpec)
    If (Attribs <> INVALID_FILE_ATTRIBUTES) Then
        FileExists = ((Attribs And vbDirectory) <> vbDirectory)
    End If
End Function

'// Russell Soundex

'// From Wikipedia, the free encyclopedia

'// Soundex is a phonetic algorithm for indexing names by their
'// sound when pronounced in English.

'// Soundex is the most widely known of all phonetic algorithms and
'// is often used (incorrectly) as a synonym for "phonetic algorithm".

'// The basic aim is for names with the same pronunciation to be
'// encoded to the same signature so that matching can occur despite
'// minor differences in spelling.

'// The Soundex code for a name consists of a letter followed by three
'// numbers: the letter is the first letter of the name, and the numbers
'// encode the remaining consonants.

'// Similar sounding consonants share the same number so, for example,
'// the labial B, F, P and V are all encoded as 1.

'// If two or more letters with the same number were adjacent in the
'// original name, or adjacent except for any intervening vowels, then
'// all are omitted except the first.

'// Vowels can affect the coding, but are never coded directly unless
'// they appear at the start of the name.

'// Russell Soundex for Spell Checking

'// This particular version of the Soundex algorithm has been adapted
'// from the original design in an attempt to more reliably facilitate
'// word matching for a generic English language spell checker.

'// Normally, each Soundex begins with the first letter of the given
'// name and only subsequent letters are used to produce the phonetic
'// signature, so only names beginning with the same first letter are
'// compared for similar pronunciation using the standard algorithm.

'// For example, one may seek the correct spelling for "upholstery" and
'// may inadvertently type "apolstry", "apolstery", or even "apholstery"
'// but would still not retrieve the correct spelling for this word.

'// Therefore, this version of the Soundex algorithm has been modified
'// to allow the matching of words that start with differing first
'// letters so as not to assume that the first letter is always known.

'// Consequently, encoding begins with the first letter of the word.

'// Because of this change, many more similarly spelled words are
'// returned as a match, so the Soundex's length has also been
'// extended from three numbers to four to produce a more unique
'// phonetic signature.

'// Returns the 4 character Soundex code for an English word.
Private Function GetSoundexWord(sWord As String) As Long

    Dim bSoundex(1 To 4) As Byte
    Dim i As Long, j As Long
    Dim prev As Byte
    Dim code As Byte

    If LenB(sWord) = 0& Then Exit Function

    '// Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6

    For i = 1& To Len(sWord)
        Select Case MidLcI(sWord, i) 'LCase$(Mid$(sWord, i, 1))
              ' "a", "e", "h", "i", "o", "u", "w", "y"
            Case 97, 101, 104, 105, 111, 117, 119, 121:  GoTo nexti '// do nothing

              ' "b", "f", "p", "v"
            Case 98, 102, 112, 118:                      code = 1 '// key labials

              ' "c", "g", "j", "k", "q", "s", "x", "z"
            Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2

            Case 100, 116: code = 3   ' "d", "t"
            Case 108:      code = 4   ' "l"
            Case 109, 110: code = 5   ' "m", "n"
            Case 114:      code = 6   ' "r"
        End Select

        If prev <> code Then '// do nothing if most recent
            j = j + 1&
            bSoundex(j) = code '// add new code
            If j = 4& Then Exit For
            prev = code
        End If
nexti:
    Next i

    '// Return the first four values (padded with 0's)
    CopyMemory GetSoundexWord, bSoundex(1), 4&
End Function

'// Returns the 4 character Soundex code for an English word
'// but from right to left.
Private Function GetSoundexWordR(sWord As String) As Long

    Dim bSoundex(1 To 4) As Byte
    Dim i As Long, j As Long
    Dim prev As Byte
    Dim code As Byte

    If LenB(sWord) = 0& Then Exit Function

    '// Replacement
    '   [a, e, h, i, o, u, w, y] = 0
    '   [b, f, p, v] = 1
    '   [c, g, j, k, q, s, x, z] = 2
    '   [d, t] = 3
    '   [l] = 4
    '   [m, n] = 5
    '   [r] = 6

    For i = Len(sWord) To 1& Step -1
        Select Case MidLcI(sWord, i) 'LCase$(Mid$(sWord, i, 1))
              ' "a", "e", "h", "i", "o", "u", "w", "y"
            Case 97, 101, 104, 105, 111, 117, 119, 121:  GoTo nexti '// do nothing

              ' "b", "f", "p", "v"
            Case 98, 102, 112, 118:                      code = 1 '// key labials

              ' "c", "g", "j", "k", "q", "s", "x", "z"
            Case 99, 103, 106, 107, 113, 115, 120, 122:  code = 2

            Case 100, 116: code = 3   ' "d", "t"
            Case 108:      code = 4   ' "l"
            Case 109, 110: code = 5   ' "m", "n"
            Case 114:      code = 6   ' "r"
        End Select

        If prev <> code Then '// do nothing if most recent
            j = j + 1&
            bSoundex(j) = code '// add new code
            If j = 4& Then Exit For
            prev = code
        End If
nexti:
    Next i

    '// Return the first four soundex values
    CopyMemory GetSoundexWordR, bSoundex(1), 4&
End Function

'// Returns the Minimum of 3 numbers
Private Function min3(ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long) As Long
    If n1 < n2 Then min3 = n1 Else min3 = n2
    If n3 < min3 Then min3 = n3
End Function

'// Levenshtein Distance

'// From Wikipedia, the free encyclopedia

'// In information theory and computer science, the Levenshtein distance
'// or edit distance between two strings is given by the minimum number
'// of operations needed to transform one string into the other, where
'// an operation is an insertion, deletion, or substitution of a single
'// character.

'// It is named after Vladimir Levenshtein, who considered this distance
'// in 1965.

'// It is useful in applications that need to determine how similar two
'// strings are, such as spell checkers.

'// It can be considered a generalisation of the Hamming distance, which
'// is used for strings of the same length and only considers substitution
'// edits.

'// There are also further generalisations of the Levenshtein distance
'// that consider, for example, exchanging two characters as an operation,
'// like in the Damerau-Levenshtein distance algorithm.

'// int LevenshteinDistance(char str1[1..lenStr1], char str2[1..lenStr2])
'//    // d is a table with lenStr1+1 rows and lenStr2+1 columns
'//    declare int d[0..lenStr1, 0..lenStr2]
'//    // i and j are used to iterate over str1 and str2
'//    declare int i, j, cost
'//
'//    for i from 0 to lenStr1
'//        d i, 0:=i
'//    for j from 0 to lenStr2
'//        d 0, j:=j
'//
'//    for i from 1 to lenStr1
'//        for j from 1 to lenStr2
'//            if str1[i] = str2[j] then cost := 0
'//                                 else cost := 1
'//            d[i, j] := minimum(
'//                                 d[i-1, j  ] + 1,     // deletion
'//                                 d[i  , j-1] + 1,     // insertion
'//                                 d[i-1, j-1] + cost   // substitution
'//                             )
'//
'//    return d[lenStr1, lenStr2]

'// Returns the Levenshtein Distance between 2 strings.
Private Function GetLevenshteinDistance(argStr1 As String, argStr2 As String) As Long
    Dim LenStr1 As Long, LenStr2 As Long
    Dim editMatrix() As Long, cost As Long
    Dim str1_i As Integer, str2_j As Integer
    Dim i As Long, j As Long

    LenStr1 = Len(argStr1)
    LenStr2 = Len(argStr2)

    If LenStr1 = 0& Then
        '// The length of Str2 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr2

    ElseIf LenStr2 = 0& Then
        '// The length of Str1 is the minimum number of operations
        '// needed to transform one string into the other
        GetLevenshteinDistance = LenStr1

    Else
        '// editMatrix is a table with lenStr1+1 rows and lenStr2+1 columns
        ReDim editMatrix(LenStr1, LenStr2) As Long

        '// i and j are used to iterate over str1 and str2
        For i = 0& To LenStr1
            editMatrix(i, 0&) = i
        Next
    
        For j = 0& To LenStr2
            editMatrix(0&, j) = j
        Next
    
        For i = 1& To LenStr1
            str1_i = MidLcI(argStr1, i) 'LCase$(Mid$(argStr1, i, 1))
            For j = 1& To LenStr2
                str2_j = MidLcI(argStr2, j) 'LCase$(Mid$(argStr2, j, 1))

                If str1_i = str2_j Then cost = 0& Else cost = 1&

                '//                     deletion,insertion,substitution
                editMatrix(i, j) = min3(editMatrix(i - 1&, j) + 1&, _
                                        editMatrix(i, j - 1&) + 1&, _
                                        editMatrix(i - 1&, j - 1&) + cost)
            Next j
        Next i
    
        GetLevenshteinDistance = editMatrix(LenStr1, LenStr2)
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Visual Basic is one of the few languages where you
' can't extract a character from or insert a character
' into a string at a given position without creating
' another string.
'
' The following Property fixes that limitation.
'
' Twice as fast as AscW and Mid$ when compiled.
'        iChr = AscW(Mid$(sStr, lPos, 1))
'        iChr = MidI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get MidI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemory MidI, ByVal StrPtr(sStr) + lPos + lPos - 2&, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Mid$(sStr, lPos, 1) = Chr$(iChr)
'        MidI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let MidI(sStr As String, ByVal lPos As Long, ByVal iChr As Integer)
    CopyMemory ByVal StrPtr(sStr) + lPos + lPos - 2&, iChr, 2&
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        iChr = AscW(LCase$(Mid$(sStr, lPos, 1)))
'        iChr = MidLcI(sStr, lPos)
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get MidLcI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemory MidLcI, ByVal StrPtr(sStr) + lPos + lPos - 2&, 2&
    If MidLcI > 64 And MidLcI < 91 Then MidLcI = MidLcI + 32
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Mid$(sStr, lPos, 1) = LCase$(Chr$(iChr))
'        MidLcI(sStr, lPos) = iChr
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let MidLcI(sStr As String, ByVal lPos As Long, ByVal iChr As Integer)
    If iChr > 64 And iChr < 91 Then iChr = iChr + 32
    CopyMemory ByVal StrPtr(sStr) + lPos + lPos - 2&, iChr, 2&
End Property

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure:   InitializeBuffers [Private Function]
'  Purpose:     Calculates the range of the Twister's runners
'               Initializes pending runner stacks
'               Initializes the Twisters runner buffer
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Function InitializeBuffers(ByVal lCount As Long, ByVal eOpt As eOptimal) As Long
    Dim curve As Long, optimal As Long                       ' CraZy performance curve
    Const n10K As Long = 10000&                              ' .
    Const n20K As Long = 20000&                              '    .
    If lCount > n20K Then curve = 12& * (lCount \ n10K - 2&) '      .
    optimal = lCount * (eOpt * 0.0012!) - curve + 4&         '       .
    If optimal > StackSize Then
        StackSize = optimal
        ReDim StackLB(0 To optimal) As Long    ' Stack to hold pending lower boundries
        ReDim StackUB(0 To optimal) As Long    ' Stack to hold pending upper boundries
    End If
    If optimal > TwisterBufSize Then
        TwisterBufSize = optimal
        ReDim TwisterBuf(0 To optimal) As Long ' This is a cache used when moving ranges
    End If
    InitializeBuffers = optimal
End Function

' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
'
'  Procedure: TwisterStringSort [Private Sub]
'
'  Description:
'   Stable Insert/Binary hybrid using CopyMemory
'
' ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤

Private Sub TwisterStringSort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, Optional ByVal eOptimalA As eOptimal = RefreshUnsorted)
    Dim walk As Long, find As Long, midd As Long
    Dim base As Long, ceil As Long, item As String
    Dim cast As Long, mezz As Long, run As Long
    Dim lpStr As Long, lp As Long, lpB As Long
    Dim idx As Long, optimal As Long
    Const eComp As Long = 1&                     ' Ascending
    walk = (ubA - lbA) + 1&                      ' Grab array item count
    If walk < 2& Then Exit Sub                   ' If nothing to do then exit
    optimal = InitializeBuffers(walk, eOptimalA) ' Initialize working buffers
    lpStr = VarPtr(item)                                          ' Cache pointer to the string variable
    lp = VarPtr(sA(lbA)) - (lbA * 4&)                             ' Cache pointer to the array
    lpB = VarPtr(TwisterBuf(0&))                                  ' Cache pointer to the buffer
    walk = lbA: mezz = ubA                                        ' Initialize our walker variables
    Do Until walk = mezz ' ----==============================---- ' Do the twist while there's more items
        walk = walk + 1&                                          ' Walk up the array and use binary search to insert each item down into the sorted lower array
        CopyMemByV lpStr, lp + (walk * 4&), 4&                    ' Grab current value into item
        find = walk                                               ' Default to current position
        ceil = walk - 1&                                          ' Set ceiling to current position - 1
        base = lbA                                                ' Set base to lower bound
        Do While StrComp(sA(ceil), item, mMethod) = eComp   '  .  ' While current item must move down
            midd = (base + ceil) \ 2&                             ' Find mid point
            Do Until StrComp(sA(midd), item, mMethod) = eComp     ' Step back up if equal or below
                base = midd + 1&                                  ' Bring up the base
                midd = (base + ceil) \ 2&                         ' Find mid point
                If midd = ceil Then Exit Do                       ' If we're up to ceiling
            Loop                                                  ' Out of loop > target pos
            find = midd                                           ' Set provisional to new ceiling
            If find = base Then Exit Do                           ' If we're down to base
            ceil = midd - 1&                                      ' Bring down the ceiling
        Loop '-Twister v4 ©Rd-     .      . ...  .             .  ' Out of stable binary search loops
        If (find < walk) Then                                     ' If current item needs to move down
            CopyMemByV lpStr, lp + (find * 4&), 4&
            run = walk + 1&
            Do Until run > mezz Or run - walk > optimal           ' Runner do loop
                If Not (StrComp(item, sA(run), mMethod) = eComp) Then Exit Do
                run = run + 1&
            Loop: cast = (run - walk)
            CopyMemByV lpB, lp + (walk * 4&), cast * 4&        ' Grab current value(s)
            CopyMemByV lp + ((find + cast) * 4&), lp + (find * 4&), (walk - find) * 4& ' Move up items
            CopyMemByV lp + (find * 4&), lpB, cast * 4&        ' Re-assign current value(s) into found pos
            If cast > 1& Then
                If Not run > mezz Then
                    idx = idx + 1&
                    StackLB(idx) = run - 1&  ' Will increment back
                    StackUB(idx) = mezz
                End If
                walk = find
                mezz = find + cast - 1&
        End If: End If
        If walk = mezz Then
            If idx Then
                walk = StackLB(idx)
                mezz = StackUB(idx)
                idx = idx - 1&
    End If: End If: Loop           ' Out of walker do loop
    CopyMemByR ByVal lpStr, 0&, 4& ' De-reference pointer to item variable
    ' ----=================----
End Sub
