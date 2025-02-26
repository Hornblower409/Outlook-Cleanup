Attribute VB_Name = "Cleanup_"
Option Explicit

    ' =====================================================================
    '                       DEBUG Switches
    ' =====================================================================
    
    '   2024-08-19 - Normally FALSE
    '
    #Const DEBUG_Init_ShowResponse = False                   ' Show the Response Doc as soon as we clear INIT
    #Const DEBUG_Init_SaveHTML = False                       ' Init - Write HTML
    #Const DEBUG_BQMark_SaveHTML = False                     ' BQMark - Write HTML
    #Const DEBUG_BQRestore_SaveHTML = False                  ' BQRestore - Write HTML
    #Const DEBUG_RepSep_SaveHTML = False                     ' RepSep - Write HTML
    #Const DEBUG_Cleanup_Compress = False                    ' Compress - Write HTML
    
    '   Why all the DoEvents? Because I can (rarely) get caught in an infinate loop
    '   because of some malformed HTML or something I've never seen before. Without
    '   the DoEvents Ctrl-Break is almost useless.
    
    ' =====================================================================
    '   BQ Marker String
    '
    '   MLLLM =  MarkChr & Format( BQLevel, "000" ) & MarkChr
    '
    Private Const BQMarkerLen As Long = 5
    Private Const BQMarkerLevelLen As Long = 3
    
    '   BQ HTML Tags
    '
    Private Const BQStartTag As String = "<blockquote"
    Private Const BQStartTagRestore As String = _
        "<blockquote " & _
        "style='border:none;border-left:solid #CCCCCC 1.0pt;mso-border-left-alt:solid #CCCCCC .75pt;" & _
        "padding:0in 0in 0in 6.0pt;" & _
        "margin-left:4.8pt;margin-right:0in'>"
        
    Private Const BQEndTag As String = "</blockquote>"
    Private Const BQEndTagRestore As String = "</blockquote>"
        
    '   Generic HTML Tags
    '
    Private Const ParaStartTag As String = "<p"
    Private Const ParaEndTag As String = "</p>"
    
    '   My Word RepSep Para
    '
    '   M ... M^p = MarkChr & {para body} & MarkChar & ^p
        
    '   RepSep Types
    '
    Private Const RepSepType_None As Integer = 0                            '   Not a RepSep
    Private Const RepSepType_HTML As Integer = 1                            '   HTMLDiv (Border Level N)
    Private Const RepSepType_Mine As Integer = 2                            '   My Word RepSep Para
    
    '   Misc
    '
    Private Const LineBreak As String = vbVerticalTab                       '   Word Manual Line Break. Chr(11) 0x0A.
    
    '   Cleanup Module Level Globals
    '
    '       I know this is bad practice, but I really don't want to have to put
    '       a long list of args on each Call and Function Def or mess with User Types.
    '
    '       ! Be careful with "Item" when calling Cleanup_Main            !
    '       ! Your function may have an "Item" but that ain't this "Item" !
    '
    Private Item As Object              '   The Item being cleaned
    Private wDoc As Word.Document       '   WordEditor of the Item
    Private InfoText As String          '   Text added to the Cleanup message
    Private CleanupMsg As String        '   Message shown at the end of Cleanup
    Private IsResponse As Boolean       '   Is Item a Reply/Fwd?
    Private CleanupRange As Word.Range  '   Range to do Cleanup on
    Private BQStartMarkChr As String    '   ChrW(glbUnicode_BQStartMark)
    Private BQEndMarkChr As String      '   ChrW(glbUnicode_BQEndMark)
    Private RepSepMarkChr As String     '   ChrW(glbUnicode_RepSepMark)
    
    Private CleanupType As Long         '   Who started this thing?
    Private Enum CleanupTypes
                    Manual              '       Ribbion or QAT
                    Response            '       Response Inspector Activation
                    SendNew             '       Send New email
                    SendResponse        '       Send Response
    End Enum
    
'   Cleanup Manual
'
Public Sub Cleanup_Manual_Lnk()

    Cleanup_Manual

End Sub

Public Function Cleanup_Manual() As Boolean
Const ThisProc = "Cleanup_Manual"
Cleanup_Manual = False

    '   Check for an Active Inspector
    '
    If Not (TypeOf ActiveWindow Is Outlook.Inspector) Then
        Msg_Box Proc:=ThisProc, Step:="Check Inspector", Text:="Active window is not an Inspector."
        Exit Function
    End If

    '   Do Cleanup and set return
    '
    CleanupType = CleanupTypes.Manual
    Set Item = ActiveInspector.CurrentItem
    Cleanup_Manual = Cleanup_Main()
    Msg_Box Proc:=ThisProc, Step:="After Cleanup Manual", Text:=CleanupMsg
    
End Function

'   Cleanup Response
'
Public Sub Cleanup_Response_Lnk()
Const ThisProc = "Cleanup_Response_Lnk"

    '   Check for an Active Inspector
    '
    If Not (TypeOf ActiveWindow Is Outlook.Inspector) Then
        Msg_Box Proc:=ThisProc, Step:="Check Inspector", Text:="Active window is not an Inspector."
        Exit Sub
    End If

    Dim InspectorItem As Object
    Set InspectorItem = ActiveInspector.CurrentItem
    Cleanup_Response InspectorItem

End Sub

Public Function Cleanup_Response(ByVal InspectorItem As Object) As Boolean
Const ThisProc = "Cleanup_Response"
Cleanup_Response = False

    '   Do Cleanup
    '
    CleanupType = CleanupTypes.Response
    Set Item = InspectorItem
    Cleanup_Response = Cleanup_Main()
    If Not Cleanup_Response Then Msg_Box Proc:=ThisProc, Step:="After Cleanup Response", Text:=CleanupMsg

End Function

'   Cleanup Main
'
Public Function Cleanup_Main() As Boolean

    '   2024-09-15 - SPOS - Meeting Accept, Tenative, Decline, Cancel can NOT be converted
    '   from RTF to HTML even if they have body text and the Menu Format Text looks like
    '   it alllows convert to HTML.
    '
    If Not Mail_IsMeetingResponse(Item) Then

        '   Do Cleanup_Steps and a screen refresh
        '
        Cleanup_Main = Cleanup_Steps()
        If Not wDoc Is Nothing Then wDoc.Application.ScreenRefresh
        
    Else
    
        Cleanup_Main = True
        InfoText = "SPOS - Meeting Response in RTF that can't be converted to HTML. Can't do Cleanup." & vbNewLine
    
    End If

    '   Build the Cleanup final message
    '
    CleanupMsg = "Cleanup DONE."
    If Not Cleanup_Main Then CleanupMsg = "Cleanup NOT completed."
    
    If InfoText <> "" Then
    
        If Right(InfoText, 2) = vbNewLine Then InfoText = Left(InfoText, Len(InfoText) - 2)
        CleanupMsg = Trim(InfoText) & glbBlankLine & CleanupMsg
    
    End If
    
End Function

'   Cleanup Steps
'
Private Function Cleanup_Steps() As Boolean
Cleanup_Steps = False

    If Not Cleanup_Init Then Exit Function
    If Not Cleanup_CleanupRangeSet() Then Exit Function
    If Not Cleanup_MarkerCheck() Then Exit Function
    If Not Cleanup_RepSepAdd() Then Exit Function
    If Not Cleanup_Char() Then Exit Function
    If Not Cleanup_WhiteSpace() Then Exit Function
    If Not Cleanup_BQReplace() Then Exit Function
    If Not Cleanup_BQMerge() Then Exit Function
    If Not Cleanup_ParaFormat() Then Exit Function
    If Not Cleanup_Combine() Then Exit Function
    If Not Cleanup_Compress() Then Exit Function
    If Not Cleanup_LastPass() Then Exit Function
    If Not Cleanup_BQRestore() Then Exit Function
    If Not Cleanup_SpellingIgnore() Then Exit Function
    
Cleanup_Steps = True
End Function

'   Cleanup Init
'
Private Function Cleanup_Init() As Boolean
Const ThisProc = "Cleanup_Init"
Cleanup_Init = False

    Set wDoc = Nothing
    InfoText = ""
    CleanupMsg = ""
    IsResponse = Mail_IsResponse(Item)
    
    BQStartMarkChr = ChrW(glbUnicode_BQStartMark)
    BQEndMarkChr = ChrW(glbUnicode_BQEndMark)
    RepSepMarkChr = ChrW(glbUnicode_RepSepMark)
    
    '   RTF - exit
    '
    If Mail_IsRTF(Item) Then
        InfoText = InfoText & "Item is RTF." & vbNewLine
        Exit Function
    End If
    
    '   No HTMLBody - exit
    '
    If Not Mail_HasHTMLBody(Item) Then
        InfoText = InfoText & "Item has no HTMLBody." & vbNewLine
        Exit Function
    End If
    If Item.HTMLBody = "" Then
        InfoText = InfoText & "Item HTMLBody is empty." & vbNewLine
        Exit Function
    End If
    
    '   No Word Editor - exit
    '
    Set wDoc = Item.GetInspector.WordEditor
    If wDoc Is Nothing Then
        InfoText = InfoText & "Item Inspector has no WordEditor." & vbNewLine
        Exit Function
    End If
            
    '   Not editable - exit
    '
    If wDoc.ProtectionType <> wdNoProtection Then
        InfoText = InfoText & "The Item Inspector is Locked For Editing (Read Only)." & vbNewLine
        Exit Function
    End If
    
    '   Already Sent - exit
    '
    If Mail_IsSent(Item) Then
        InfoText = InfoText & "The Active Inspector has already been Sent." & vbNewLine
        Exit Function
    End If
    
    ' - - - - - - - - - - - - - - - -
    #If DEBUG_Init_ShowResponse Then
        Item.GetInspector.Activate  ' 2025-02-19 Was "Item.Display"
    #End If
    ' - - - - - - - - - - - - - - - -
    
    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_Init_SaveHTML Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_Start") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -

Cleanup_Init = True
End Function

'   CleanupRange - Set CleanupRange
'
Private Function Cleanup_CleanupRangeSet() As Boolean
Const ThisProc = "Cleanup_CleanupRangeSet"
Cleanup_CleanupRangeSet = False

    '   Define the CleanupRange based on the CleanupType
    '
    Select Case CleanupType
    
        Case CleanupTypes.Manual
            Set CleanupRange = wDoc.Content.Duplicate
    
        Case CleanupTypes.Response
            Set CleanupRange = wDoc.Content.Duplicate
            
            '   Pull CleanupRange down so it doesn't include the two para above the RepSep (3rd para).
            '   Just visual. So I don't see the first para as a Line Break after Cleanup.
            '
            If wDoc.Content.Paragraphs.Count < 3 Then Stop: Exit Function
            CleanupRange.Start = wDoc.Content.Paragraphs.Item(3).Range.Start
        
        Case CleanupTypes.SendNew
            Set CleanupRange = wDoc.Content.Duplicate
            
        Case CleanupTypes.SendResponse
        
            Dim wPara As Word.Paragraph
            Set wPara = Cleanup_RepSepFind()
            
            '   RepSep Found - CleanupRange is Start to para before the RepSep
            '
            If Not (wPara Is Nothing) Then
            
                Set CleanupRange = Cleanup_RepSepFind().Range
                CleanupRange.Start = wDoc.Content.Start
                CleanupRange.MoveEnd WdUnits.wdParagraph, -1
                
            '   No RepSep Found - Ask for bypass (change to CleanupTypes.SendNew)
            '
            Else
            
                Select Case Msg_Box( _
                        Proc:=ThisProc, Step:="Send Response RepSep Check", _
                        Icon:=vbQuestion, Buttons:=vbYesNo, Default:=vbDefaultButton2, _
                        Text:="Item is a Response but no RepSep Marker found." & glbBlankLine & _
                        "Cleanup entire Item?")
                    Case vbYes
                        CleanupType = CleanupTypes.SendNew
                        IsResponse = False
                        Set CleanupRange = wDoc.Content.Duplicate
                        InfoText = InfoText & "Bypassed Response RepSep Check. Changed to Send New." & vbNewLine
                    Case vbNo
                        Exit Function
                End Select
            
            End If
    
    End Select

Cleanup_CleanupRangeSet = True
End Function

'   Init - Check for Marker Chars in the Doc
'
Private Function Cleanup_MarkerCheck() As Boolean
Const ThisProc = "Cleanup_MarkerCheck"
Cleanup_MarkerCheck = False
    
    If Cleanup_TypeSend() Then Cleanup_MarkerCheck = True:  Exit Function
    
    '   Look for: Manual - Only BQMarks, Else - All.
    '
    Dim Marks() As Variant
    If CleanupType = CleanupTypes.Manual Then
        Marks() = Array(glbUnicode_BQStartMark, glbUnicode_BQEndMark)
    Else
        Marks() = Array(glbUnicode_BQStartMark, glbUnicode_BQEndMark, glbUnicode_RepSepMark)
    End If
    
    '   Find and notify
    '
    Dim SearchAsc As Variant
    For Each SearchAsc In Marks()

        Dim wRange As Word.Range
        Set wRange = Word_FindRange(CleanupRange, ChrW(SearchAsc))
        If Not wRange Is Nothing Then
            Msg_Box Proc:=ThisProc, Step:="Search for Markers", _
                    Text:="Found Marker: CharW(" & SearchAsc & "), U+" & Hex(SearchAsc) & ", '" & Cleanup_MarkerAscToName(SearchAsc) & "' at positon " & CStr(wRange.End) & "."
                    InfoText = InfoText & "Found Marker Char in Doc." & vbNewLine
            Exit Function
        End If
        
    Next SearchAsc

Cleanup_MarkerCheck = True
End Function

' ---------------------------------------------------------------------
'   Character
' ---------------------------------------------------------------------
'
Private Function Cleanup_Char() As Boolean
Cleanup_Char = False
    
    '   ZWNBSP - Zero Width No-Break Space.
    '
    '       Word Find White Space won't find it. So make it a
    '       plain NoBreakSpace (^s) that Word can find.
    '
    Word_Replace CleanupRange, ChrW(glbUnicode_ZWNBSP), "^s", wdReplaceAll

Cleanup_Char = True
End Function

' ---------------------------------------------------------------------
'   Paragraph Format - Whole Range
'---------------------------------------------------------------------
'
Private Function Cleanup_ParaFormat() As Boolean
Cleanup_ParaFormat = False
    
    With CleanupRange.ParagraphFormat
    
        '   SpaceAfter/Before all Zero
        '
        .SpaceAfterAuto = False
        .SpaceAfter = 0
        .SpaceBeforeAuto = False
        .SpaceBefore = 0
    
        '   Right Indent = Zero
        '   (Yes. I got an email with paras that had Right Indent)
        '
        .RightIndent = 0
        
        '   LineSpacingRule (Just "Line Spacing" in the GUI)
        '
        .LineSpacingRule = wdLineSpaceSingle
        
        '   Shading = None
        '
        '   (If you are so pretentious that you use Shading in a email, I'm gonna ignore it)
        '
        .Shading.BackgroundPatternColor = wdColorAutomatic
        .Shading.BackgroundPatternColorIndex = wdNoHighlight
        .Shading.ForegroundPatternColor = wdColorAutomatic
        .Shading.ForegroundPatternColorIndex = wdNoHighlight
        .Shading.Texture = wdTextureNone

    End With
    
Cleanup_ParaFormat = True
End Function

' ---------------------------------------------------------------------
'   Combine
'---------------------------------------------------------------------
'
'   Merge paragraphs in a Range that have compatible formatting/borders
'   by replacing the Paragraph Mark at the end of the paragraph with a Line Break.
'
'   SPOS - If Stupid sees "{empty}^l xxxxx^l {empty}^p"  on incomming mail he changes it
'   to "{empty}^l xxxxx^p" and puts an After 12 on the ^p.
'
Private Function Cleanup_Combine() As Boolean
Const ThisProc = "Cleanup_Combine"
Cleanup_Combine = False
    
    '   Setup
    '
    Dim LoopCount As Long: LoopCount = 0
    Dim OriginalParaCount As Long: OriginalParaCount = CleanupRange.Paragraphs.Count
    Dim Paras As Word.Paragraphs: Set Paras = CleanupRange.Paragraphs
    
    Dim ThisPara As Word.Paragraph
    Dim NextPara As Word.Paragraph
    
    '   Loop until PIx = Current number of paras in the CleanupRange
    '
    Dim PIx As Long: PIx = 1
    Do While PIx < Paras.Count
    
        DoEvents
    
        Set ThisPara = Paras.Item(PIx)
        Set NextPara = Paras.Item(PIx + 1)

        '   Check for Word Table Cell/Row marker paragraphs
        '
        '       SPOS - Word Table Cell/Row marker paragraphs end with Chr(13) & Chr(7)
        '       If either paragraph does not end with CR - skip it
        '
        If ThisPara.Range.Characters.Last.Text <> vbCr Then GoTo NextLoop
        If NextPara.Range.Characters.Last.Text <> vbCr Then GoTo NextLoop
        
        '   Clear any Drop Caps
        '
        '       SPOS - Stupid throws a 4605 "method or property is not available because the
        '       current paragraph has no text." if you even look at .DropCap for a para
        '       that is nothing but an ^l. But I can't test beforehand because Len(.Range.Text)
        '       also blows up and .Range.Characters.Count is greater than one. So On Error Resume.
        '
        On Error Resume Next
            ThisPara.DropCap.Position = wdDropNone
        On Error GoTo 0

        '   If P1 is a RepSep (Either Type) - SpaceBefore from Zero to 5
        '
        If Cleanup_RepSepType(ThisPara) <> RepSepType_None Then ThisPara.SpaceBefore = 5
        
        '   If Next is empty - align it with this one
        '
        If NextPara.Range.Characters.Count = 1 Then NextPara.LeftIndent = ThisPara.LeftIndent
        
        '   2024-11-26 - Add .SpaceAfter before an indent change.
        '   See Cleanup_LastPass for details.
        '
        If ThisPara.LeftIndent <> NextPara.LeftIndent Then ThisPara.SpaceAfter = 8

        '   If the the two paras are similar enough to combine
        '
        If Cleanup_CombineTest(ThisPara, NextPara) Then
        
            '   Replace the first para ^p with an ^l
            '   Backup the index (there is one less para)
            '
            ThisPara.Range.Characters.Last.Text = LineBreak
            PIx = PIx - 1
            
        End If
        
NextLoop:

        PIx = PIx + 1
        LoopCount = LoopCount + 1
        If (LoopCount Mod 200) = 0 Then GoSub Combine_LoopCheck
        
    Loop
    
Cleanup_Combine = True
Exit Function
    
'   -- Local Subs ---------------------------------------------------------------------
'
'   Why a GoSub? So it can Get all the Proc variables without having to pass them.
'   And because I still have fond memories of GoSub from my Pick Basic days and wanted
'   to write one last one.
'
'   Give the user a chance to cancel a long running Combine
'
Combine_LoopCheck:

    Select Case Msg_Box( _
            Proc:=ThisProc, Step:="Paragraph Walk - If Cleanup_CombineTest", _
            Icon:=vbQuestion, Buttons:=vbYesNoCancel, Default:=vbDefaultButton3, _
            Text:=CStr((OriginalParaCount - Paras.Count + PIx - 1)) & " paragraphs out of " & CStr(OriginalParaCount) & " processed." & glbBlankLine & _
            "Continue?")
        Case vbYes
            '  Continue
        Case vbNo
            InfoText = InfoText & "Skipped Combine because it was running too long." & vbNewLine
            Cleanup_Combine = True
            Exit Function
        Case vbCancel
            InfoText = InfoText & "Canceled Cleanup because Combine was running too long." & vbNewLine
            Exit Function
    End Select

    Return
    
End Function

'   Combine - Can two paragraphs be combined?
'
Private Function Cleanup_CombineTest(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineTest = False
    
    If Not Cleanup_CombineFormat(ThisPara, NextPara) Then Exit Function
    If Not Cleanup_CombineSpecial(ThisPara, NextPara) Then Exit Function
    If Not Cleanup_CombineTable(ThisPara, NextPara) Then Exit Function
    If Not Cleanup_CombineBorder(ThisPara, NextPara) Then Exit Function
    If Not Cleanup_CombineList(ThisPara, NextPara) Then Exit Function
    
Cleanup_CombineTest = True
End Function

'   Combine - If either para part of a List - no combine
'
Private Function Cleanup_CombineList(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineList = False

    If Not ThisPara.Range.ListFormat.List Is Nothing Then Exit Function
    If Not NextPara.Range.ListFormat.List Is Nothing Then Exit Function
    
Cleanup_CombineList = True
End Function

'   Combine - If both para do not have the same user para borders - no combine
'
'       This is only about plain para borders applied by the user.
'       RepSep Top borders (para or HTML) have already been checked in Cleanup_CombineSpecial.
'
Private Function Cleanup_CombineBorder(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineBorder = False

    Dim xBorder As Variant
    For Each xBorder In Array(Word.WdBorderType.wdBorderTop, Word.WdBorderType.wdBorderLeft, Word.WdBorderType.wdBorderRight, Word.WdBorderType.wdBorderBottom)
    
        If ThisPara.Borders.Item(xBorder).Visible <> NextPara.Borders.Item(xBorder).Visible Then Exit Function
        
    Next xBorder

Cleanup_CombineBorder = True
End Function

'   Combine - If both para not inside/outside a table - no combine
'
Private Function Cleanup_CombineTable(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineTable = False

    If Not (ThisPara.Range.Information(wdWithInTable) = NextPara.Range.Information(wdWithInTable)) Then Exit Function

Cleanup_CombineTable = True
End Function

'   Combine - If either para is special - no combine
'
Private Function Cleanup_CombineSpecial(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineSpecial = False

    If Cleanup_ParaAnySpecial(ThisPara) Then Exit Function
    If Cleanup_ParaAnySpecial(NextPara) Then Exit Function

Cleanup_CombineSpecial = True
End Function

'   Combine - Do two paragraphs have compatible formatting?
'
Private Function Cleanup_CombineFormat(ByVal ThisPara As Word.Paragraph, ByVal NextPara As Word.Paragraph) As Boolean
Cleanup_CombineFormat = False
    
    '   I have no idea how many of these really matter.
    '   I just copied everything from the Word docs.
    '
    '   Star = I set it in Cleanup_Para or Cleanup_ParaFormat. No need to test.
    '   Minus = It doesn't matter?
    
    With ThisPara
    
        If .Alignment <> NextPara.Alignment Then Exit Function
        '- If .AutoAdjustRightIndent <> NextPara.AutoAdjustRightIndent Then Exit Function
        '- If .BaseLineAlignment <> NextPara.BaseLineAlignment Then Exit Function
        '- If .Borders.Enable <> NextPara.Borders.Enable Then Exit Function
        '- If .CharacterUnitFirstLineIndent <> NextPara.CharacterUnitFirstLineIndent Then Exit Function
        '- If .CharacterUnitLeftIndent <> NextPara.CharacterUnitLeftIndent Then Exit Function
        '- If .CharacterUnitRightIndent <> NextPara.CharacterUnitRightIndent Then Exit Function
        '- If .DisableLineHeightGrid <> NextPara.DisableLineHeightGrid Then Exit Function
        If .FirstLineIndent <> NextPara.FirstLineIndent Then Exit Function
        '- If .HalfWidthPunctuationOnTopOfLine <> NextPara.HalfWidthPunctuationOnTopOfLine Then Exit Function
        '- If .HangingPunctuation <> NextPara.HangingPunctuation Then Exit Function
        '- If .Hyphenation <> NextPara.Hyphenation Then Exit Function
        '- If .KeepTogether <> NextPara.KeepTogether Then Exit Function
        '- If .KeepWithNext <> NextPara.KeepWithNext Then Exit Function
        If .LeftIndent <> NextPara.LeftIndent Then Exit Function
        '* If .LineSpacing <> NextPara.LineSpacing Then Exit Function
        '* If .LineSpacingRule <> NextPara.LineSpacingRule Then Exit Function
        '- If .LineUnitAfter <> NextPara.LineUnitAfter Then Exit Function
        '- If .LineUnitBefore <> NextPara.LineUnitBefore Then Exit Function
        '- If .MirrorIndents <> NextPara.MirrorIndents Then Exit Function
        '- If .NoLineNumber <> NextPara.NoLineNumber Then Exit Function
        If .OutlineLevel <> NextPara.OutlineLevel Then Exit Function
        '- If .PageBreakBefore <> NextPara.PageBreakBefore Then Exit Function
        '- If .ReadingOrder <> NextPara.ReadingOrder Then Exit Function
        '* If .RightIndent <> NextPara.RightIndent Then Exit Function
        '* If .Shading.BackgroundPatternColor <> NextPara.Shading.BackgroundPatternColor Then Exit Function
        '* If .Shading.BackgroundPatternColorIndex <> NextPara.Shading.BackgroundPatternColorIndex Then Exit Function
        '* If .Shading.ForegroundPatternColor <> NextPara.Shading.ForegroundPatternColor Then Exit Function
        '* If .Shading.ForegroundPatternColorIndex <> NextPara.Shading.ForegroundPatternColorIndex Then Exit Function
        '* If .Shading.Texture <> NextPara.Shading.Texture Then Exit Function
        '* If .SpaceAfter <> NextPara.SpaceAfter Then Exit Function
        '* If .SpaceAfterAuto <> NextPara.SpaceAfterAuto Then Exit Function
        '* If .SpaceBefore <> NextPara.SpaceBefore Then Exit Function
        '* If .SpaceBeforeAuto <> NextPara.SpaceBeforeAuto Then Exit Function
        '- If .TextboxTightWrap <> NextPara.TextboxTightWrap Then Exit Function
        '- If .WidowControl <> NextPara.WidowControl Then Exit Function
        '- If .WordWrap <> NextPara.WordWrap Then Exit Function
                
    End With
    
Cleanup_CombineFormat = True
End Function

' ---------------------------------------------------------------------
'   White Space
' ---------------------------------------------------------------------

'   Remove {White Space} in front of a line end.
'
Private Function Cleanup_WhiteSpace() As Boolean
Cleanup_WhiteSpace = False
    
    '   ^w^p -> ^p
    '
    ' SPOS - 2024-05-01
    '
    '   If the line ends with "<image> space ^p^p" then find "^w^p" selects the SECOND ^p.
    '   And in my Word_DeleteLeft it can get stuck in an loop.
    '   So we do the replace for ^w chars one at a time.
    '
    '   Original = Word_DeleteLeft CleanupRange, "^w^p", 1, wdReplaceAll
    '
    Word_DeleteLeft CleanupRange, " ^p", 1, wdReplaceAll
    Word_DeleteLeft CleanupRange, "^t^p", 1, wdReplaceAll
    
    '   ^w^l -> ^l
    '
    Word_Replace CleanupRange, "^w^l", "^l", wdReplaceAll
    
Cleanup_WhiteSpace = True
End Function

' ---------------------------------------------------------------------
'   Compress
' ---------------------------------------------------------------------

'   Compress - Multiple blank lines
'
Private Function Cleanup_Compress() As Boolean
Const ThisProc = "Cleanup_Compress"
Cleanup_Compress = False
    
    Dim wSearch As Word.Range
    
    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_Cleanup_Compress Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_Start") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -
    
    '   Remove any double blank lines.
    '
    '   ^p^l^l -> ^p^l
    '   ^l^l^p -> ^l^p
    '   ^l^l^l -> ^l^l
    '
    '       SPOS - Word Find Wildcards (RegEx) does not have a "start of line anchor"
    '       so this is the best I could come up with that doesn't break all the
    '       formatting Word packs into a para mark.
    '
    While Word_Replace(CleanupRange, "^p^l^l", "^p^l", wdReplaceAll):  DoEvents: Wend
    While Word_DeleteLeft(CleanupRange, "^l^l^p", 2, wdReplaceAll):    DoEvents: Wend
    While Word_Replace(CleanupRange, "^l^l^l", "^l^l", wdReplaceAll):  DoEvents: Wend
    
    '   BQStart
    '
    '   BQStart ^p^l -> BQStart ^p
    '
    While Word_Replace(CleanupRange, BQStartMarkChr & "^p^l", BQStartMarkChr & "^p", wdReplaceAll): DoEvents: Wend
    
    '   BQStart ^p^p -> BQStart ^p
    '
    '   2024-08-19 - SPOS.
    '
    '       "BQStart ^p^p {Table}" will find the "BQStart ^p^p" but won't replace the "^p^p" with "^p".
    '       So we have to do a Find on the "BQStart ^p^p" and manually delete the first ^p.
    '
    '-    While Word_Replace(CleanupRange, BQStartMarkChr & "^p^p", BQStartMarkChr & "^p", wdReplaceAll)
    '-    Wend
    Do
    
        DoEvents
    
        Set wSearch = Word_FindRange(CleanupRange, BQStartMarkChr & "^p^p")
        If wSearch Is Nothing Then Exit Do
        
        wSearch.Start = wSearch.Start + 1
        wSearch.End = wSearch.End - 1
        wSearch.Delete
        
    Loop
    
    '   BQEnd
    '
    '   ^l^p BQEnd -> ^p BQEnd
    '   ^p^p BQEnd -> ^p BQEnd
    '
    While Word_Replace(CleanupRange, "^l^p" & BQEndMarkChr, "^p" & BQEndMarkChr, wdReplaceAll): DoEvents: Wend
    While Word_Replace(CleanupRange, "^p^p" & BQEndMarkChr, "^p" & BQEndMarkChr, wdReplaceAll): DoEvents: Wend
        
    '   If the "untouchable" doc end ^p is not in the CleanupRange - done
    '
    If Not wDoc.Content.Paragraphs.Last.Range.InRange(CleanupRange) Then
        Cleanup_Compress = True
        Exit Function
    End If
    
    '   Setup for cases involving the doc end ^p
    '
    Set wSearch = Word_FindDefault(wDoc.Content.Duplicate)
    
    '   If the doc end para is empty
    '
    If wDoc.Content.Paragraphs.Last.Range.Characters.Count = 1 Then
    
        '   ^l^p^p| or ^p^l^p| -> ^p^p|
        '
        Do
            wSearch.Start = wDoc.Content.End - 3
            wSearch.Collapse wdCollapseStart
        Loop While Word_Replace(wSearch, "^l", "", wdReplaceOne)
        
    End If
    
    '   If the doc end para is preceded by two empty para
    '
    '   ^p1^p2^p| -> ^p1^p|
    '
    Do
        wSearch.Start = wDoc.Content.End - 3
        wSearch.Collapse wdCollapseStart
        Set wSearch = Word_FindRange(wSearch, "^p^p^p")
    If wSearch Is Nothing Then Exit Do
        wSearch.Start = wSearch.Start + 1
        wSearch.End = wSearch.End - 1
        wSearch.Delete
    Loop
    
Cleanup_Compress = True
End Function

' ---------------------------------------------------------------------
'   Last Pass - Runs after Combine abd Compress.
' ---------------------------------------------------------------------
'
Private Function Cleanup_LastPass() As Boolean
Const ThisProc = "Cleanup_LastPass"
Cleanup_LastPass = False

    '   At this point (after Combine and Compress) the only ^l^p left
    '   are ones between indent levels and on the Doc End ^p.
    '
    '   Compress any remaining ^l^p into the ^p and give it a
    '   SpaceAfter = 8.
    '
    '   Combine has already put a SpaceAfter = 8 on the ^p before an
    '   indent change. I'm just doing it again and additionally for
    '   any ^l^p| (Doc End ^p).
    '
    '   The rollup is so what you see mirrors the HTML that Stupid will
    '   generate. Behind you back (at Send) he does almost exactly the
    '   same thing.
    '
    Dim wSearch As Word.Range
    
    '   Using an Array for the sFinds so if I catch another one, I'm ready.
    '
    Dim sFinds() As Variant
    sFinds = Array("^l^p")
    Dim sFind As Variant
    Dim Occurance As Long
    
    Dim SearchRange As Word.Range
    Dim LastOccurance As Long
        
    For Each sFind In sFinds: Do
    
        Set SearchRange = Word_FindDefault(CleanupRange.Duplicate)
        LastOccurance = SearchRange.Start

        For Occurance = 1 To 999999: Do: DoEvents
        
            '   Pull down the Search Range to LastOccurance
            '
            SearchRange.Start = LastOccurance
            
            '   Find the first occurance of ^x^p in the Search Range
            '
            Set wSearch = Word_FindRange(SearchRange, sFind)
            If wSearch Is Nothing Then Exit For ' = Next sFind
            
            '   Update LastOccurance to the right of the ^p
            '
            LastOccurance = wSearch.End + 1
            
            '   Get the wPara for the ^p
            '
            Dim wPara As Word.Paragraph
            Set wPara = wSearch.Paragraphs.Item(1)
            
            '   Check for all the wPara, ^p or Chr(13) that I can't touch
            '
            '       If wPara is a Word Table Cell/Row marker paragraph (Ends with Chr(13) & Chr(7))
            '       If wPara is a RepSep (Either Type)
            '
            If wPara.Range.Characters.Last.Text <> vbCr Then Exit Do  ' = Next Occurance
            If Cleanup_RepSepType(wPara) <> RepSepType_None Then Exit Do ' = Next Occurance
            
            '   Add a .SpaceAfter = 8 to the wPara
            '
            wPara.SpaceAfter = 8
            
            '   Point wSearch to the ^x and Delete it,
            '   Move LastOccurance one to the left so it still points to the ^p.
            '   Pull CleanupRange up one to account for the Delete
            '
            wSearch.End = wSearch.End - 1
            wSearch.Delete
            LastOccurance = LastOccurance - 1
            CleanupRange.End = CleanupRange.End - 1
            
        Loop While False: Next Occurance
    Loop While False: Next sFind
    
Cleanup_LastPass = True
End Function

' ---------------------------------------------------------------------
'   BlockQuotes
' ---------------------------------------------------------------------

'   BlockQuotes - Mark and Remove
'
'       SPOS - If you do Item.HTMLBody <-> sHTMLBody too fast it can hoark the HTML. It won't show when you are stepping
'       through the code. Only when it is running full speed. Hence the colsolidation of BQMark and BQReplace to avoid
'       flipping HTMLBody <-> sHTMLBody twice.
'
Private Function Cleanup_BQReplace() As Boolean
Const ThisProc = "Cleanup_BQReplace"
Cleanup_BQReplace = False

    If Cleanup_TypeSend() Then Cleanup_BQReplace = True:  Exit Function
    
    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_BQMark_SaveHTML Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_Start_ItemHTMLBody") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -
    
    Dim sHTMLBody As String
    sHTMLBody = Item.HTMLBody
    If Not Cleanup_BQMark(sHTMLBody) Then Exit Function
    If Not Cleanup_BQRemove(sHTMLBody) Then Exit Function
    Item.HTMLBody = sHTMLBody
    
    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_BQMark_SaveHTML Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_AfterItemSet_ItemHTMLBody") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -
    
    If Not Cleanup_CleanupRangeSet() Then Exit Function
    
Cleanup_BQReplace = True
End Function

'   BQReplace - BQMark - Insert my BQPara into the HTML
'
Private Function Cleanup_BQMark(ByRef sHTMLBody As String) As Boolean
Cleanup_BQMark = False
    
    '   ByRefs for the recursion
    '
    Dim BQIx As Long: BQIx = 1
        
    '   Desend into Word HTML Hell
    '
    If Not Cleanup_BQDesend(sHTMLBody, BQIx, 0) Then Exit Function
        
Cleanup_BQMark = True
End Function

'   BQReplace - BQMark - Recursive Desend
'                                                                                \/  Don't touch this ByVal
Private Function Cleanup_BQDesend(ByRef sHTMLBody As String, ByRef BQIx As Long, ByVal BQLevel As Long) As Boolean
Cleanup_BQDesend = False

    Const NotFound As Long = 999999
    
    ' BQLevel                       '   My Level. Comes from my parent. Never changes.
    
    Dim BQStartTagIx As Long        '   Next BlockQuote Start tag
    Dim BQEndTagIx As Long          '   Next BlockQuote End tag
    
    Dim BQMarkerPara As String
    
    Do
        
        DoEvents
    
        '   Find the next Start/End Tags from the current position
        '
        BQStartTagIx = InStr(BQIx, sHTMLBody, BQStartTag, vbTextCompare)
        If BQStartTagIx = 0 Then BQStartTagIx = NotFound
    
        BQEndTagIx = InStr(BQIx, sHTMLBody, BQEndTag, vbTextCompare)
        If BQEndTagIx = 0 Then BQEndTagIx = NotFound
        
        '   BQEndTag
        '
        If (BQEndTagIx < BQStartTagIx) Then
        
            '   Insert a BQEnd Para before the the BQEnd Tag
            '
            BQMarkerPara = "<p>" & BQEndMarkChr & Format(BQLevel, "000") & BQEndMarkChr & "</p>"
            sHTMLBody = Left(sHTMLBody, BQEndTagIx - 1) & BQMarkerPara & Mid(sHTMLBody, BQEndTagIx)
            BQIx = BQEndTagIx + Len(BQMarkerPara) + Len(BQEndTag)
        
            Exit Do
        
        '   BQStartTag
        '
        ElseIf (BQStartTagIx < BQEndTagIx) Then
        
            '   Insert a BQStart Para after the BQStart Tag
            '
            BQIx = InStr(BQStartTagIx, sHTMLBody, ">", vbTextCompare)
            If BQIx = 0 Then Stop: Exit Function
            BQMarkerPara = "<p>" & BQStartMarkChr & Format(BQLevel + 1, "000") & BQStartMarkChr & "</p>"
            sHTMLBody = Left(sHTMLBody, BQIx) & BQMarkerPara & Mid(sHTMLBody, BQIx + 1)
            BQIx = BQIx + Len(BQMarkerPara)
        
            '   Call myself
            '
            If Not Cleanup_BQDesend(sHTMLBody, BQIx, BQLevel + 1) Then Exit Function
            
        Else
            Exit Do
        End If
            
    Loop

Cleanup_BQDesend = True
End Function

'   BQReplace - BQRemove - Remove all the original BQTags from the HTML
'
Private Function Cleanup_BQRemove(ByRef sHTMLBody As String) As Boolean
Cleanup_BQRemove = False

    Dim TagStartIx As Long
    Dim TagEndIx As Long
    Dim iX As Long

    Dim SearchTag As Variant
    For Each SearchTag In Array(BQStartTag, BQEndTag)
    
        iX = 1
        Do
        
            DoEvents
        
            '   Get the start of the next BQTag - <blockquote ...> or </blockquote>
            '
            TagStartIx = InStr(iX, sHTMLBody, SearchTag, vbTextCompare)
            If TagStartIx = 0 Then Exit Do
            
            '   Find the end of the BQTag - the trailing >
            '
            TagEndIx = InStr(TagStartIx, sHTMLBody, ">", vbBinaryCompare)
            If TagEndIx = 0 Then Stop: Exit Function
            
            '   Cut it out
            '
            sHTMLBody = Left(sHTMLBody, TagStartIx - 1) & Mid(sHTMLBody, TagEndIx + 1)
            iX = TagStartIx

        Loop
        
    Next SearchTag
        
Cleanup_BQRemove = True
End Function

'   BlockQuotes - BQMerge - Merge contigious BQs
'
'       SPOS - I can't find any way to stop him from inserting an (unneeded) "stutter"
'       when changing BlockQuote levels, or after a ^p in the same level, when I just
'       "touch" the HTML (e.g. sHTML = Item.HTMLBody, Item.HTMLBody = sHTML).
'
'       <bq1>
'       ...
'       </bq1>      |   Stutter
'       <bq1>       |
'           <bq2>
'           ...
'           </bq2>  |   Stutter
'           <bq2>   |
'           ...
'           </bq2>
'       </bq1>      |   Stutter
'       <bq1>       |
'       ...
'       </bq1>
'
'       If I take out the stutters between BQ level changes in Word (after BQRemove) then when
'       I put the BQs back in (with BQRestore), the generated HTML is a mess. The only solution
'       I can think of would be to fix them directly in the HTML. Something I'm not going to
'       try. But it seems Stupid will let me get away with removing them when the BQs are contigious
'       (at the deepest level with no nested BQs between them). So that's what I'm doing here.
'
Private Function Cleanup_BQMerge() As Boolean
Cleanup_BQMerge = False

    If Cleanup_TypeSend() Then Cleanup_BQMerge = True:  Exit Function

    '   Set the initial Search Range as CleanupRange
    '
    Dim wSearch As Word.Range
    Set wSearch = CleanupRange.Duplicate
    
    Do
    
        DoEvents
    
        '   Find next: {BQEndMark} ^p {BQStartMark} in the current Search Range
        '
        Dim wRange As Word.Range
        Set wRange = Word_FindRange(wSearch, BQEndMarkChr & "^p" & BQStartMarkChr)
    
    If wRange Is Nothing Then Exit Do
    
        '   Get the Start/End paras as ranges
        '
        Dim BQEnd As Word.Range
        Set BQEnd = wRange.Paragraphs.Item(1).Range
        
        Dim BQStart As Word.Range
        Set BQStart = wRange.Paragraphs.Item(2).Range
        
        '   Pull the top of the Searh Range down to the end of the Start Para
        '
        wSearch.Start = BQStart.End
        
        '   If the BQEnd and BQStart are NOT the same Level - next
        '
        If Mid(BQEnd.Text, 2, BQMarkerLevelLen) <> Mid(BQStart.Text, 2, BQMarkerLevelLen) Then GoTo NextLoop
        
        '   wPriorStart Range = This End Para up to the start of CleanupRange
        '
        Dim wPriorStart As Word.Range
        Set wPriorStart = BQEnd.Duplicate
        wPriorStart.Start = CleanupRange.Start
        
        '   Search backwards in wPriorStart for a Start para
        '
        Set wRange = Word_FindRange(wPriorStart, BQStartMarkChr, Backwards:=True)
        If wRange Is Nothing Then Stop: Exit Function
        Dim BQPriorStart As Word.Range
        Set BQPriorStart = wRange.Paragraphs.Item(1).Range
        
        '   If the Prior Start and this End are NOT the same Level number - next
        '   (i.e. there is a nested BQ above me)
        '
        If Mid(BQPriorStart.Text, 2, BQMarkerLevelLen) <> Mid(BQEnd.Text, 2, BQMarkerLevelLen) Then GoTo NextLoop
        
        '   wNextEnd = This Start Para to the end of the CleanupRange
        '
        Dim wNextEnd As Word.Range
        Set wNextEnd = BQStart.Duplicate
        wNextEnd.End = CleanupRange.End
        
        '   Search forwards in wNextEnd for a End Para
        '
        Set wRange = Word_FindRange(wNextEnd, BQEndMarkChr)
        If wRange Is Nothing Then Stop: Exit Function
        Dim BQNextEnd As Word.Range
        Set BQNextEnd = wRange.Paragraphs.Item(1).Range
        
        '   If the Next End and this Start are NOT the same Level number - next
        '   (i.e. there is a nested BQ below me)
        '
        If Mid(BQNextEnd.Text, 2, BQMarkerLevelLen) <> Mid(BQStart.Text, 2, BQMarkerLevelLen) Then GoTo NextLoop
        
        '   ELSE - We've got a "stutter" with no nested BQs above or below.
        '   Delete this BQEnd/BQStart pair (deepest first!)
        '
        BQStart.Delete
        BQEnd.Delete
        
NextLoop: Loop
    
Cleanup_BQMerge = True
End Function

'   BlockQuotes - BQRestore - Replace my BQPara with BQTags
'
Private Function Cleanup_BQRestore() As Boolean
Cleanup_BQRestore = False

    If Cleanup_TypeSend() Then Cleanup_BQRestore = True:  Exit Function

    If Not Cleanup_BQRestoreHTML() Then Exit Function
    If Not Cleanup_BQRestoreDoc() Then Exit Function
    If Not Cleanup_CleanupRangeSet() Then Exit Function

Cleanup_BQRestore = True
End Function

'   BQRestore - BQRestoreHTML - Replace my BQ Paras in the HTML with my standard BQTags
'
Private Function Cleanup_BQRestoreHTML() As Boolean
Const ThisProc = "Cleanup_BQRestoreHTML"
Cleanup_BQRestoreHTML = False

    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_BQRestore_SaveHTML Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_Start") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -

    Dim sHTMLBody As String
    sHTMLBody = Item.HTMLBody
    
    Dim SearchChr As Variant
    For Each SearchChr In Array(BQStartMarkChr, BQEndMarkChr)
    
        Dim BQRestoreTag As String
        If SearchChr = BQStartMarkChr Then BQRestoreTag = BQStartTagRestore
        If SearchChr = BQEndMarkChr Then BQRestoreTag = BQEndTagRestore
    
        Dim iX As Long:  iX = 1
        Do
    
            DoEvents

            '   Find the next BQPara by looking for the BQMarkerChr
            '
            iX = InStr(iX, sHTMLBody, SearchChr, vbBinaryCompare)
            If iX = 0 Then Exit Do
    
            '   Start -> "<" in "<p ..."
            '
            Dim TagStartIx As Long:  TagStartIx = InStrRev(sHTMLBody, ParaStartTag, iX, vbTextCompare)
            If TagStartIx = 0 Then Stop: Exit Function
            
            '   End -> "<" in "</p>..."
            '
            Dim TagEndIx As Long:  TagEndIx = InStr(TagStartIx, sHTMLBody, ParaEndTag, vbTextCompare)
            If TagEndIx = 0 Then Stop: Exit Function
            
            '   End -> "X" in "</p>X"
            '
            TagEndIx = TagEndIx + Len(ParaEndTag)
    
            '   Cut out the BQPara and replace it with the BQRestoreTag
            '
            sHTMLBody = Left(sHTMLBody, TagStartIx - 1) & BQRestoreTag & Mid(sHTMLBody, TagEndIx)
            
            '   Ix -> "X" in "<BQRestoreTag ...>X"
            '
            iX = TagStartIx + Len(BQRestoreTag)
            
        Loop
        
    Next SearchChr
    
    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_BQRestore_SaveHTML Then
        If Not File_SaveHTML(sHTMLBody, ThisProc & "_After") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -
    
    Item.HTMLBody = sHTMLBody

Cleanup_BQRestoreHTML = True
End Function

'   BQRestore - BQRestoreDoc - Cleanup the Doc after an HTML Restore
'
'   SPOS - On a Response Stupid may have added a SpaceAfter = 12 to all the para
'   above the RepSep. (Never bothered to figure out why).
'
Private Function Cleanup_BQRestoreDoc() As Boolean
Cleanup_BQRestoreDoc = False

    '   If not a Responce - Done
    '
    If Not IsResponse Then Cleanup_BQRestoreDoc = True:  Exit Function
    
    '   Find My Word RepSep Para
    '
    Dim wPara As Word.Paragraph
    Set wPara = Cleanup_RepSepFind()
    If wPara Is Nothing Then Stop: Exit Function
    
    '   wRange = Start of the doc to the Para just above My RepSep para
    '
    Dim wRange As Word.Range
    Set wRange = wPara.Range.Duplicate
    If wRange.MoveEnd(Unit:=wdParagraph, Count:=-1) = 0 Then Stop: Exit Function
    wRange.Start = wDoc.Content.Start
    
    '   Reset SpaceAfter in that range
    '
    '       If .SpaceAfter  = 12 then it's one of his - reset it.
    '       If anything else (like 8 I use for indent change) is one of mine - leave it alone.
    '
    For Each wPara In wRange.Paragraphs
        If wPara.SpaceAfter = 12 Then
            With wRange.ParagraphFormat
                .SpaceAfterAuto = False
                .SpaceAfter = 0
            End With
        End If
    Next wPara

Cleanup_BQRestoreDoc = True
End Function

' ---------------------------------------------------------------------
'   RepSep
' ---------------------------------------------------------------------

'   RepSep - Make the 3rd para of a Response My Word RepSep Para
'
Private Function Cleanup_RepSepAdd() As Boolean
Const ThisProc = "Cleanup_RepSepAdd"
Cleanup_RepSepAdd = False

    If Cleanup_TypeSend() Then Cleanup_RepSepAdd = True:  Exit Function
    If CleanupType = CleanupTypes.Manual Then Cleanup_RepSepAdd = True:  Exit Function
    
    '   If Doc already has My Word RepSep Para - done
    '
    Dim RSPara As Word.Paragraph
    Set RSPara = Cleanup_RepSepFind()
    If Not (RSPara Is Nothing) Then Cleanup_RepSepAdd = True:  Exit Function
    
    '   Get the 3rd para
    '
    If wDoc.Content.Paragraphs.Count < 3 Then Stop: Exit Function
    Dim wPara As Word.Paragraph
    Set wPara = wDoc.Content.Paragraphs.Item(3)
    
    '   If it still has an HTML top border - remove it
    '
    If Cleanup_RepSepType(wPara) = RepSepType_HTML Then
        If Not Cleanup_RepSepHTMLDelete(wPara) Then Stop: Exit Function
    End If

    '   Insert the Markers
    '
    With wPara.Range.Characters
    
        .First.InsertBefore RepSepMarkChr
        .Last.InsertBefore RepSepMarkChr
        
    End With

    '   Clean up the formating
    '
    With wPara.Range.ParagraphFormat

        .SpaceBeforeAuto = False
        .SpaceBefore = 5
        .SpaceAfterAuto = False
        .SpaceAfter = 0

    End With

    '   Set the Top Border
    '
    
    '   Color used for My Word RepSep Para Top Border
    '
    '       When convered to a "Border Level N" hDiv Border on Send Outlook will set the
    '       Border.ColorIndex to wdColorBrightGreen but the .Color, and the color
    '       it actually displays as will remain mine.
    '
    Const RepSepColor As Long = 51200           '  = 0xC800
    
    With wPara.Borders.Item(wdBorderTop)

        .LineStyle = WdLineStyle.wdLineStyleSingle
        .LineWidth = WdLineWidth.wdLineWidth100pt
        .Color = RepSepColor

    End With

    ' - - - - - - - - - - - - - - - - - - - - - -
    #If DEBUG_RepSep_SaveHTML Then
        If Not File_SaveHTML(Item.HTMLBody, ThisProc & "_AfterWordRepSepAdd") Then Stop: Exit Function
    #End If
    ' - - - - - - - - - - - - - - - - - - - - - -

Cleanup_RepSepAdd = True
End Function

'   RepSep - Remove the RepSep Marks from My Word RepSep Para
'
Private Function Cleanup_RepSepRemove() As Boolean
Cleanup_RepSepRemove = False

    If Not (Cleanup_TypeSend() And IsResponse) Then Cleanup_RepSepRemove = True: Exit Function

    '   Find the Para with the RepSep Marks
    '
    Dim wPara As Paragraph
    Set wPara = Cleanup_RepSepFind()
    If wPara Is Nothing Then Exit Function
    
    '   Remove the RepSep Marks
    '
    With wPara.Range
    
        .Characters.First.Delete                            ' Delete the first char of the para
        .Characters.Item(.Characters.Count - 1).Delete      ' Delete the char just infront of the ^p
    
    End With
    
Cleanup_RepSepRemove = True
End Function

'   RepSep - Return the Paragraph that is My Word RepSep Para or Nothing
'
Private Function Cleanup_RepSepFind() As Word.Paragraph

    Set Cleanup_RepSepFind = Nothing
    
    '   Find my RepSep Mark
    '
    Dim wRange As Word.Range
    Set wRange = Word_FindRange(wDoc.Content, RepSepMarkChr)
    If wRange Is Nothing Then Exit Function
    
    '   Return the range of the para it is inside
    '
    Set Cleanup_RepSepFind = wRange.Paragraphs.Item(1)

End Function

'   RepSep - Is a paragraph a Reply Seperator? If so, what type?
'
Public Function Cleanup_RepSepType(ByVal wPara As Word.Paragraph) As Integer

    Cleanup_RepSepType = RepSepType_None
    
    '   If it is My Word RepSep Para -> Word Para
    '
    If wPara.Range.Characters.First.Text = RepSepMarkChr Then
        Cleanup_RepSepType = RepSepType_Mine
        Exit Function
    End If
    
    '   If it has an hDiv Top Border -> HTML
    '
    With wPara.Range.HTMLDivisions
    
        If .Count > 0 Then
            If .Item(1).Borders.Item(wdBorderTop).Visible Then
                Cleanup_RepSepType = RepSepType_HTML
                Exit Function
            End If
        End If
    
    End With
    
End Function

'   RepSep - Remove a paragraph's hDiv Top Border
'
Private Function Cleanup_RepSepHTMLDelete(ByVal wPara As Word.Paragraph) As Boolean
Cleanup_RepSepHTMLDelete = False

    With wPara.Range.HTMLDivisions
    
        If .Count < 1 Then Stop: Exit Function
        .Item(1).Borders.Item(wdBorderTop).LineStyle = wdLineStyleNone
    
    End With

Cleanup_RepSepHTMLDelete = True
End Function

' ---------------------------------------------------------------------
'   Helpers
' ---------------------------------------------------------------------

Private Function Cleanup_SpellingIgnore() As Boolean
Cleanup_SpellingIgnore = False

    If Cleanup_TypeSend() Then Cleanup_SpellingIgnore = True:  Exit Function
    If CleanupType = CleanupTypes.Manual Then Cleanup_SpellingIgnore = True:  Exit Function

    '   Find the Para with the RepSep Mark
    '
    Dim wPara As Paragraph
    Set wPara = Cleanup_RepSepFind()
    If wPara Is Nothing Then Stop: Exit Function
    
    '   Range = RepSep para to the end of the doc
    '
    Dim wRange As Word.Range
    Set wRange = wPara.Range.Duplicate
    
    '   Mark the Range No Proofing
    '
    wRange.End = wDoc.Content.End
    wRange.NoProofing = True

Cleanup_SpellingIgnore = True
End Function

'   CleanupType Helper - True if CleanupType is either .SendNew or .SendResponse
'
Private Function Cleanup_TypeSend() As Boolean
    Cleanup_TypeSend = (CleanupType = CleanupTypes.SendNew) Or (CleanupType = CleanupTypes.SendResponse)
End Function

'   Marks Helper - Convert a Marker Asc value to it's name
'
Private Function Cleanup_MarkerAscToName(ByVal AscValue As Long) As String
    
    Dim AscName As String:  AscName = ""
    
    Select Case AscValue
        Case glbUnicode_BQStartMark:    AscName = glbUnicode_BQStartMarkName
        Case glbUnicode_BQEndMark:      AscName = glbUnicode_BQEndMarkName
        Case glbUnicode_RepSepMark:     AscName = glbUnicode_RepSepMarkName
        Case Else
            Stop: Exit Function
    End Select
    
    Cleanup_MarkerAscToName = AscName

End Function

'   Para Helper - Is a paragraph Special? (BQ Start/End or RepSep)
'
Private Function Cleanup_ParaAnySpecial(ByVal wPara As Word.Paragraph) As Boolean
Cleanup_ParaAnySpecial = True

    '   If the para begins with either BQMark - it's special
    '
    Dim SearchChr As Variant
    For Each SearchChr In Array(BQStartMarkChr, BQEndMarkChr)

        If wPara.Range.Characters.First.Text = SearchChr Then Exit Function

    Next SearchChr
    
    '   If RepSepType <> None - it's special
    '
    If Cleanup_RepSepType(wPara) <> RepSepType_None Then Exit Function

Cleanup_ParaAnySpecial = False
End Function

