Attribute VB_Name = "Library"
Option Explicit

Public Const glbBlankLine           As String = vbNewLine & vbNewLine                   ' Blank Line

Public Const glbPropTag_InternetReplyID         As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"                                  ' PR_IN_REPLY_TO_ID As String

'   Error.Numbers
'
Public Const glbError_None                          As Long = 0                     ' No Error
Public Const glbError_TypeMismatch                  As Long = 13                    ' First seen in Misc_OLSetProperty using a Form property variable.
Public Const glbError_PropertyNotFound              As Long = 438                   ' Property does not exist. Object doesn't support this property or method.


'   Reserved Unicode Characters
'
Public Const glbUnicode_RepSepMark      As Long = &H25B2                            ' U+25B2 - Black Up-Pointing Triangle
Public Const glbUnicode_RepSepMarkName  As String = "Responce Seperator"
Public Const glbUnicode_LineAnchor      As Long = &H25B3                            ' U+25B3 - White Up-Pointing Triangle
Public Const glbUnicode_LineAnchorName  As String = "Line Anchor"
Public Const glbUnicode_BQStartMark     As Long = &H25B6                            ' U+25B6 - Black Right-Pointing Triangle
Public Const glbUnicode_BQStartMarkName As String = "BlockQuote Start"
Public Const glbUnicode_BQEndMark       As Long = &H25C0                            ' U+25C0 - Black Left-Pointing Triangle
Public Const glbUnicode_BQEndMarkName   As String = "BlockQuote End"

'   Special Unicode Characters
'
Public Const glbUnicode_ZWNBSP          As Long = &HFEFF                            ' U+FEFF - Zero Width No-Break Space


Public Function Msg_Box( _
    Optional ByVal oErr As VBA.ErrObject = Nothing, _
    Optional ByVal Proc As String = "", _
    Optional ByVal Step As String = "", _
    Optional ByVal Text As String = "", _
    Optional ByVal Subject As String = "", _
    Optional ByVal Buttons As Long = vbOKOnly, _
    Optional ByVal Default As Long = vbDefaultButton1, _
    Optional ByVal Icon As Long = vbExclamation _
    ) As Integer
    
    '   Make the Title long so the dialog will be wide.
    '   Looks like 108 is max (depends on character widths).
    '   After that the title ends in "..."
    '
    Dim Title As String
    Title = "Outlook VBA Code" & Space(92)
    
    Dim ProcLine As String: ProcLine = ""
    If Proc <> "" Then ProcLine = vbNewLine & "Proc: '" & Proc & "'."
    
    Dim StepLine As String: StepLine = ""
    If Step <> "" Then StepLine = vbNewLine & "Step: '" & Step & "'."
    
    Dim SubjectLine As String: SubjectLine = ""
    If Subject <> "" Then SubjectLine = glbBlankLine & "Subject: '" & Subject & "'."
    
    Dim ErrText As String: ErrText = ""
    If Not oErr Is Nothing Then
        ErrText = _
        glbBlankLine & "Err.Number: " & Err.Number & " (0x" & Hex(Err.Number) & ")." & _
        glbBlankLine & "Error.Description: '" & Err.Description & "'"
    End If
    
    Dim TextLine As String: TextLine = ""
    If Text <> "" Then TextLine = glbBlankLine & Text
    
    '   Build final MsgBlock and Remove any Leading/Trailing vbNewLine
    '
    Dim MsgBlock As String
    MsgBlock = ProcLine & StepLine & SubjectLine & ErrText & TextLine
    
    While Left(MsgBlock, 2) = vbNewLine
        MsgBlock = Mid(MsgBlock, 3)
    Wend
    While Right(MsgBlock, 2) = vbNewLine
        MsgBlock = Mid(MsgBlock, 1, Len(MsgBlock) - 2)
    Wend
    
    Msg_Box = MsgBox( _
        MsgBlock, _
        Buttons + Default + Icon, _
        Title _
    )

End Function

'   Is a Meeting Responce (Cancel, Accept, Tenative, Decline)?
'
Public Function Mail_IsMeetingResponse(ByVal Item As Object) As Boolean
Mail_IsMeetingResponse = True

    Select Case Item.Class
        Case Outlook.olMeetingCancellation
            Exit Function
        Case Outlook.olMeetingForwardNotification
            Exit Function
        Case Outlook.olMeetingResponseNegative
            Exit Function
        Case Outlook.olMeetingResponsePositive
            Exit Function
        Case Outlook.olMeetingResponseTentative
            Exit Function
    End Select

Mail_IsMeetingResponse = False
End Function

'   Is a Reply/Forward?
'
'   Can not use PR_LAST_VERB_EXECUTED during a Reply/Fwd Event
'   because that won't be set on the Original until AFTER the response is sent.
'
'   All Response will have InternetReplyID because I force it in
'   Response_ForceReply. (Stupid doesn't always add this Property or set it to a value).
'
Public Function Mail_IsResponse(ByVal Item As Object) As Boolean
Mail_IsResponse = False

    '   Must have a InternetReplyID
    '
    Dim InternetReplyID As String
    If Not Misc_OLGetProperty(Item, glbPropTag_InternetReplyID, InternetReplyID) Then InternetReplyID = ""
    If InternetReplyID = "" Then Exit Function
    
    '   And must be unsent
    '
    If Mail_IsSent(Item) Then Exit Function

Mail_IsResponse = True
End Function

'   Is the Item RTF?
'
Public Function Mail_IsRTF(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_IsRTF"
Mail_IsRTF = False

    If Item.ItemProperties.Item("BodyFormat") Is Nothing Then Exit Function
    If Not (Item.BodyFormat = Outlook.olFormatRichText) Then Exit Function
    
Mail_IsRTF = True
End Function

'   Does Item have an HTMLBody?
'
Public Function Mail_HasHTMLBody(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_HasHTMLBody"
Mail_HasHTMLBody = False

    If Item.ItemProperties.Item("BodyFormat") Is Nothing Then Exit Function
    If Not (Item.BodyFormat = Outlook.olFormatHTML) Then Exit Function
    
Mail_HasHTMLBody = True
End Function

'   Has Item Been Sent?
'
Public Function Mail_IsSent(ByVal Item As Object) As Boolean
Const ThisProc = "Mail_IsSent"

    Mail_IsSent = False

    If Item.ItemProperties.Item("Sent") Is Nothing Then Exit Function
    
    '   SPOS - Even with the above check, Stupid will still try and access "Sent" when it doesn't exist.
    '   So we Error trap a glbError_PropertyNotFound
    '
    On Error Resume Next
    
        If Not Item.Sent Then Exit Function
        If Err.Number = glbError_PropertyNotFound Then Exit Function
        If Err.Number <> glbError_None Then Stop: Exit Function
    
    On Error GoTo 0
    
    Mail_IsSent = True

End Function

' ---------------------------------------------------------------------
'   Outlook MAPI Property Accessors
'
'       SPOS - Hoarked when the Get Value is a field on a form.
'
'           e.g. Misc_OLGetProperty(Item, glbPropTag_FlagRequest, FollowUpForm.Title)
'           Will leave .Title = "" even when it has a value.
'
'           Set seems to be OK with it. Below works fine.
'           e.g. Misc_OLSetProperty(Item, glbPropTag_FlagRequest, FollowUpForm.Title)
'
'       Return FALSE if the operation fails (Property does not exist/can not be set)
'
' ---------------------------------------------------------------------
'
Public Function Misc_OLGetProperty(ByVal Item As Object, ByVal PropTag As String, ByRef Value As Variant) As Boolean
Misc_OLGetProperty = False

    Dim PA As Outlook.PropertyAccessor
    
    On Error GoTo ErrExit
    
        Set PA = Item.PropertyAccessor
        Value = PA.GetProperty(PropTag)

    On Error GoTo 0
    
Misc_OLGetProperty = True
ErrExit: End Function

Public Function Misc_OLSetProperty(ByVal Item As Object, ByVal PropTag As String, ByVal Value As Variant) As Boolean
Misc_OLSetProperty = False

    Dim PA As Outlook.PropertyAccessor
    Set PA = Item.PropertyAccessor
    
    On Error Resume Next
    
        PA.SetProperty PropTag, Value
        Select Case Err.Number
            Case glbError_None
            Case glbError_TypeMismatch
                Stop: Exit Function
            Case Else
                Exit Function
        End Select

    On Error GoTo 0
    
Misc_OLSetProperty = True
End Function

'   Reset a Word .Find object to defaults
'
'   From https://gregmaxey.com/word_tip_pages/words_fickle_vba_find_property.html
'
Public Function Word_FindDefault(ByVal wRange As Word.Range) As Word.Range

    Set Word_FindDefault = wRange
    With Word_FindDefault.Find
    
        .ClearFormatting
        .Format = False
        .Forward = True
        .Highlight = wdUndefined
        .IgnorePunct = False
        .IgnoreSpace = False
        .MatchAllWordForms = False
        .MatchCase = False
        .MatchPhrase = False
        .MatchPrefix = False
        .MatchSoundsLike = False
        .MatchSuffix = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Replacement.ClearFormatting
        .Replacement.Text = ""
        .Text = ""
        .Wrap = wdFindStop

    End With
    
End Function

'   Do a Word Text Find on a Range. Returns the Found Range or Nothing.
'
'   Word_FindRange( Range, Find )
'
Public Function Word_FindRange(ByVal wRange As Word.Range, ByVal sFind As String, Optional ByVal Backwards As Boolean = False) As Word.Range

    Set Word_FindRange = Word_FindDefault(wRange.Duplicate)
    
    With Word_FindRange.Find
    
        .Forward = Not Backwards
        .Text = sFind
        
        If Not .Execute Then Set Word_FindRange = Nothing
    
    End With
    
End Function

'   Do a Word Text Find/Replace on a Range. Returns TRUE/FALSE.
'
'   Word_Replace( Range, Find, Replace, {wdReplaceAll, wdReplaceOne} )
'
'   SPOS -  The ^u notation will not work as the Replacement Text. Use CharW(nnnn).
'
Public Function Word_Replace(ByVal wRange As Word.Range, ByVal sFind As String, ByVal sReplace As String, ByVal wReplace As Word.WdReplace) As Boolean

    Word_Replace = False
    
    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(wRange.Duplicate)
    wSearch.Find.Text = sFind
    
    wSearch.Find.Replacement.Text = sReplace
    Word_Replace = wSearch.Find.Execute(Replace:=wReplace)

End Function

'   Do a Word Text Find on a Range and keep the right N chars of the Found string. Returns TRUE/FALSE
'
'   Word_DeleteLeft( Range, Find, Count, {wdReplaceAll, wdReplaceOne} )
'
'       SPOS - Some ^p can not be replaced by Word Find/Replace.
'
'       Find.Execute is TRUE but target is still there. Same in Word GUI. It finds it but
'       won't replace it. Is NOT a "<w:cr/>" in the XML that I can see.
'
'       Typically it's a naked ^p just above a table or the "End Of Doc ^p"
'
'       This proc let's me do things to what is in front of a stuck ^p.
'
Public Function Word_DeleteLeft(ByVal wRange As Word.Range, ByVal sFind As String, ByVal RCount As Long, ByVal wReplace As Word.WdReplace) As Boolean

    Word_DeleteLeft = False
    
    Dim wSearch As Word.Range
    Set wSearch = Word_FindDefault(wRange.Duplicate)
    wSearch.Find.Text = sFind
    
    Dim wDel As Word.Range
    Set wDel = wRange.Duplicate
    
    Do
    
        '   Reset the Search Range
        '
        wSearch.Start = wRange.Start
        wSearch.End = wRange.End
    
    '   While Find is True
    '
    If Not wSearch.Find.Execute Then Exit Do
    
        '   If we had at least one hit - set return TRUE
        '
        Word_DeleteLeft = True
    
        '   Delete anything left of RCount from the end
        '
        wDel.Start = wSearch.Start
        wDel.End = wSearch.End - RCount
        wDel.Delete
        
        '   If a one off - done
        '
        If wReplace = wdReplaceOne Then Exit Do
        
    Loop

End Function

