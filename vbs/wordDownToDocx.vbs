Option Explicit
'' debug
Dim isLogInfo '' As Boolean
isLogInfo = True
Dim isMath
Dim isDebug
Dim isResumeNext
isResumeNext = False
Dim logCycle '' As long
logCycle = 10
dim gLinesCount

Const Scripting_Dictionary = "Scripting.Dictionary"
Const Scripting_FileSystemObject = "Scripting.FileSystemObject"
Const String_Empty = ""
Const String_Space = " "

''on error resume next
Dim fso
set fso = createObject(Scripting_FileSystemObject)

Dim logWarns
set logWarns = new XFiles

'' styles
'' https://learn.microsoft.com/ja-jp/office/vba/api/word.wdbuiltinstyle
Const wdMainTextStory = 1
Const wdStyleNormal = -1 '' normal
Const wdStyleTitle = -63
Const wdStyleSubtitle = -75
Const wdStyleEmphasis = -89
Const wdStyleStrong = -88
Const wdTabLeaderDots = 1
Const wdPropertyTitle = 1
Const wdWithInTable = 12
Const wdAutoFitContent = 1	
Const wdAutoFitFixed = 0	''do not auto fit
Const wdAutoFitWindow =	2
Const wdRefTypeHeading = 1
Const wdContentText	= -1
Const wdNumberFullContext = -4
Const wdPageNumber = 7
Const wdIndexIndent = 0
Const wdAlignParagraphRight = 2
Const wdAlignParagraphCenter = 1
Const wdAlignParagraphLeft = 0
Const wdOMathInline = 1
Const wdWindowStateMaximize	= 1	
Const wdWindowStateMinimize	= 2	
Const wdWindowStateNormal	= 0
Const wdPageBreak = 7
Const wdOrientLandscape = 1
Const wdOrientPortrait = 0
Const msoPropertyTypeString = 4
Const wdStyleTableLightShading = -159
Const wdCaptionPositionAbove = 0
Const wdCaptionPositionBelow = 1


Dim vbTab, vbCrLf, vbLf
vbTab = Chr(9)
vbCrLf = Chr(13) & Chr(10)
vbLf = Chr(10)
'' wd command
Const wcTitle = "title"
Const wcSubTitle = "subTitle"
Const wcToc = "toc"
Const wcSection = "section"
Const wcOderList= "OderList"
Const wcNormalList= "NormalList"
Const wcCreateNote = "CreateNote"
Const wcCreateWarning= "CreateWarning"
Const wcLink= "link"
Const wcTableCreate = "tableCreate"
Const wcImage = "image"
Const wcNewLine = "newLine"
Const wcCode= "code"
Const wcText= "text"
Const wcEndParagraph= "endParagraph"
Const wcIndentPlus= "indentPlus"
Const wcHr= "hr"
Const wcIndentMinus= "indentMinus"
Const wcDate= "date"
Const wcAuthor= "author"
Const wcDivision= "division"
Const wcDocNumber= "docNumber"
Const wcNewPage= "newPage"
Const wcPageSetup= "pageSetup"
Const wcDocxTemplate= "docxTemplate"
Const wcDocxEngine= "docxEngine"
Const wcCrossRef = "crossRef"
Const wcClearContent = "clearContent"

Const wcStyleAuthor ="author1"
Const wcStyleDate ="date1"
Const wcStyleDivision ="division1"
Const wcStyleHeading5 ="wdHeading5"
Const wcStyleCode ="code"
Const wcStylePicture1 ="picture1"
Const wcStyleBody1 ="body1"
Const wcStyleCodeSpan ="codespan"
Const wcStyleTableN ="styleN"

Const wcEmphasisCodeSpan = "codespan"
Const wcEmphasisSubscript = "sub"
Const wcEmphasisSuperscript = "sup"
Const wcEmphasisBold = "b"
Const wcEmphasisItalic = "i"
Const wcEmphasisStrikeOut = "~~"
Const wcEmphasisUnderline = "u"

Const wcPropertyDate = "dDate"
Const wcPropertyDivision = "dDivision"
Const wcPropertyAuthor = "dAuthor"
Const wcPropertyNumber = "dNumber"
Const wcProperty = "property"


Call main
WScript.Quit(0)

'' functions
'arge(0) markdown file path
'arge(1) docx(docxTemplate) path
'arge(2) isMath 0 or 1(true)
'arge(3) isDebug 0 or 1(true)

Sub main
    Dim objArgs
    set objArgs = WScript.Arguments

    Dim wdFilePath
    If objArgs.Count = 0 then
        '' for debug
        wdFilePath =  fso.getParentFolderName(WScript.ScriptFullName) + "\demo-headings.wd"
        wdFilePath =  fso.getParentFolderName(WScript.ScriptFullName) + "\empty.wd"
        wdFilePath =  fso.getParentFolderName(WScript.ScriptFullName) + "\demo.wd"
        LogInfo "document: ", wdFilePath
    Else
        wdFilePath = objArgs(0)
    End If

    Dim templateDocxPath
    templateDocxPath = fso.getParentFolderName(WScript.ScriptFullName) & "\sample-heder_.docx" 
    If objArgs.Count > 1 Then
        If fso.FileExists(objArgs(1)) Then
            templateDocxPath = objArgs(1)
        End If
    End If

    isMath = False
    If objArgs.Count > 2 Then
        If objArgs(2) = "1" Then
            isMath = True
        End If
    End If

    isDebug = False
    If objArgs.Count > 3 Then
        If objArgs(3) = "1" Then
            isDebug = True
        End If
    End If

    If isDebug Then
        logCycle = 1
    End If

    '' load wd file
    Dim wdLines
    set wdLines = new XFiles
    Call wdLines.Load(wdFilePath)

    dim markdownDirPath
    markdownDirPath = fso.getParentFolderName(wdFilePath)
    
    '' detect docxTemplate in wd file.
    Call logWarns.Clear()
    Dim w
    set w = new XWord
    Call w.CreateDocX(wdLines, markdownDirPath, templateDocxPath)
    If logWarns.Count > 0 Then
        logWarns.WriteToFile(wdFilePath + ".warn.log")
    End If
End Sub

Class XWord
    Public WordApp ''As Word.Application
    Public WordDoc ''As Word.Document
    Public m_indent '' As long
    Private m_templateInsideWd
    Dim HeaderCollection '' As Collection
    Dim RefCollection '' As Collection
    Dim m_markdownDirPath '' markdown path now convert

    Dim rngCurrent ''As Range

    '' options
    Dim m_isToc
    Dim m_toToc '' as long '' table of contents, 1 to tocTo
    Dim m_tocCaption 'default TOC
    dim m_isTocSet '' If toc is set once
    Dim m_rngToc 
    Dim m_crossRef '' cref format [[$n $t (p.$p)]]

    Private Sub Class_Initialize()
        Set WordApp = CreateObject("Word.Application")
    End Sub

    '' open sample not used
    ' Sub OpenWord(path)
    '     WordApp.Visible = True
    '     Set WordDoc = WordApp.Documents.open(path)
    '     Dim NumberOfWords, i
    '     NumberOfWords = WordDoc.Sentences.Count
    '     For i = 1 to NumberOfWords
    '         ''WScript.Echo WordDoc.Sentences(i)
    '     Next
    ' End Sub

     '' main
    Sub CreateDocX(wdLines, markdownDirPath, templateDocxPath)
        Call logWarns.clear()

        gLinesCount = wdLines.Count
        m_isToc = False
        m_indent = 1
        m_markdownDirPath = markdownDirPath
        m_crossRef = "[[$n $t (p.$p)]]"
        
        Set HeaderCollection = New Collection
        Set RefCollection = New Collection

        '' get docxTemplate inside wd file
        Call doCommands(wdLines.Item(1), wdLines, 1, "docxTemplate") 
        Call doCommands(wdLines.Item(2), wdLines, 2, "docxTemplate")

        With WordApp
            .DisplayAlerts = False
            .ScreenUpdating = False
            .WindowState = wdWindowStateMinimize
            .Visible = True
            .WindowState = wdWindowStateMinimize
            '' get file As docxTemplate 
            If m_templateInsideWd <> String_Empty then
                templateDocxPath = m_templateInsideWd
            End If

            If fso.FileExists(templateDocxPath) Then
                Set WordDoc = .Documents.Add(templateDocxPath)
            Else
                Set WordDoc = .Documents.Add()
            End If

            '' delete document
            me.ClearContent "false"

            '' clear properties
            Call resetCustomProperties()

            ''Call SetCurrentPositionRangeTop
            
            '' add text to the template
            Call SetCurrentPositionRangeEnd
        End With

        Dim i ''As Long
        Dim params ''As String
        Dim wdCommandLine ''As String


        If isResumeNext then
            on error resume next
        End If

        For i = 1 To wdLines.Count
            If wdLines.item(i) <> String_Empty Then
                If i mod logCycle = 0 Then
                    Me.logProgress i
                End If
                wdCommandLine = wdLines.item(i)

                Call doCommands(wdCommandLine, wdLines, i, String_Empty)  
            End If
            Catch "CreateDocx, For loop", 1001
        Next

        Call Me.InsertToc()
        Catch "InsertToc, For loop", 1001

        Call AddXRef()
        Catch "AddXRef, For loop", 1001

        Call UpdateFields

        WordDoc.Saved = True
        With WordApp
            .Visible = True
            .ScreenUpdating = True
            .WindowState = wdWindowStateNormal
        End With
        Catch "CreateDocx End, For loop", 1001
    End Sub

    sub logProgress(i)
        LogInfo "wd to docx progress%(Line)", FormatNumber(i / gLinesCount * 100, 1, True) & "%(" & i & ")"
    End sub

    sub doCommands(ByVal wdCommandLine, ByRef wdLines, ByRef i, ByVal fixCommand)
        If wdCommandLine = String_Empty Then
            exit sub
        End If
        Dim params
        params = Split(wdCommandLine, vbTab)
        
        ''
        If fixCommand = String_Empty  Then
            '' continue
        else
            If fixCommand = wdCommandLine  Then
                '' continue
            else
                exit sub
            End If
        End If

        Select Case params(0)
            Case wcTitle '"title"
                Call Me.AddTitle(params(1))
            Case wcSubTitle
                Call Me.AddSubTitle(params(1))
            Case wcToc
                If m_isTocSet = false Then
                    m_tocCaption = params(2)
                    AddLineWithNewLine m_tocCaption, wdStyleNormal
                    set m_rngToc = GetCurrentRangeStart
                    AddNewLine("toc")
                    m_isToc = True
                    m_toToc = params(1)
                    If IsNumeric(m_toToc) Then
                        If m_toToc > 5 then
                            m_toToc = 5
                        End If
                    Else
                        m_toToc = 3
                    End If
                    m_isTocSet = True
                End If
            Case wcSection
                Call Me.AddHead(params(1), params(2),  params(3))
            Case wcOderList
                Call Me.AddOderList(params(1), params(2))
            Case wcNormalList
                Call Me.AddNormalList(params(1), params(2))
            case wcCreateNote
                Call me.AddNote(params(1))
            case wcCreateWarning
                Call me.AddWarning(params(1))
            case wcLink
                '' link, href, title(hover), text
                Call me.AddLink(params(1), params(2), params(3))
            Case wcTableCreate
                Dim columnWith ''As String()
                Dim arrayInfo ''As String()
                Dim mergeInfo ''As String()
                Dim alignInfo '' As String()
                arrayInfo = getTableInfoArrayEx(wdLines, i, columnWith, mergeInfo, alignInfo)
                Call Me.AddTable(arrayInfo, columnWith, mergeInfo, alignInfo)
            Case wcImage
                Call Me.AddImage(params(1), params(2), params(3))
            Case wcNewLine
                Call AddNewLine(params(1))
            Case wcCode
                Me.AddCodeText (params(1))
            Case wcText
                Me.AddText (params(1))
            Case wcEndParagraph
                Me.AddEndParagraph (String_Empty)
            Case wcIndentPlus
                m_indent = m_indent + 1
                Call Me.SetIndent
            Case wcHr
                Call Me.AddHr
            Case wcIndentMinus
                m_indent = m_indent - 1
                Call Me.SetIndent
            case wcDate
                Call me.AddDate(params(1))
            case wcAuthor
                Call me.AddAuthor(params(1))
            case wcProperty
                Call me.SetCustomDocumentProperty(params(1), params(2))
            case wcCrossRef
                m_crossRef = params(1)
            case wcDivision
                Call me.AddDivision(params(1))

            case wcDocNumber
                Call me.AddDocNumber(params(1))
            case wcClearContent
                Call me.ClearContent(params(1))
            Case wcNewPage
                Call Me.NewPage
            Case wcPageSetup
                Call Me.PageSetup(params(1), params(2))
            Case wcDocxTemplate
                m_templateInsideWd = params(1)
            Case wcDocxEngine
                '' no operation
            Case Else
                Me.AddText "No Command:" & params(0) & ": " & wdCommandLine
                Call AddNewLine("Else")
        End Select
    End sub


    sub ClearContent(ByVal isClear) '' isClear is string, not boolean
        If LCase(isClear) = "true" Then
            WordDoc.StoryRanges(wdMainTextStory).Delete
        End If
    End sub

    sub resetCustomProperties()
        WordDoc.BuiltInDocumentProperties.Item(wdPropertyTitle).Value = String_Empty
        SetCustomDocumentProperty "dAuthor", String_Empty
        SetCustomDocumentProperty "dDate", String_Empty
        SetCustomDocumentProperty "dNumber", String_Empty
        SetCustomDocumentProperty "dNumber", String_Empty
    End sub
    '' crate word parts
    Sub InsertToc()
        If m_isToc = False Then
            exit Sub
        End If
        Dim r
        With WordDoc
            set r = .TablesOfContents.Add(m_rngToc, True, 1, m_toToc)', '' True, String_Empty, True, True, True
            .TablesOfContents(1).TabLeader = wdTabLeaderDots
            .TablesOfContents.Format = wdIndexIndent
        End With
        SetCurrentByEnd r.Range
    End Sub

    ''Sub AddTitle(ByRef mainTitle As String, ByRef subTitle As String)
    Sub AddTitle(ByRef mainTitle)
        WordDoc.BuiltInDocumentProperties.Item(wdPropertyTitle).Value = mainTitle
        AddText(mainTitle).Style = wdStyleTitle 
    End Sub

    Sub AddSubTitle(ByRef subTitle)
        ''WordDoc.BuiltInDocumentProperties.Item().Value = mainTitle
        AddText(subTitle).Style = wdStyleSubtitle
    End Sub


    Sub AddAuthor(ByRef author)
        SetCustomDocumentProperty "dAuthor", author
        AddText(author).Style = GetStyle(wcStyleAuthor) 
        AddNewLine "AddAuthor" 
    End Sub

    Sub AddDate(ByRef vDate)
        SetCustomDocumentProperty "dDate", vDate
        AddText(vDate).Style = GetStyle(wcStyleDate) 
        AddNewLine "AddDate"
    End Sub

    Sub AddDivision(ByRef division)
        SetCustomDocumentProperty "dDivision", division
        AddText(division).Style = GetStyle(wcStyleDivision) 
        AddNewLine "AddDivision" 
    End Sub

    Sub AddDocNumber(ByRef docNumber)
        SetCustomDocumentProperty "dNumber", docNumber
    End Sub

    Sub SetIndent()
        If (m_indent < 1) then
            m_indent = 1
        End If
        If (m_indent > 3) then
            m_indent = 3
        End If
        Dim myStyle
        myStyle = "body" & CStr(m_indent)
        rngCurrent.Style = GetStyle(myStyle)
    End Sub

    sub NewPage()
        rngCurrent.InsertBreak wdPageBreak
        SetCurrentByEnd rngCurrent
    End sub

    sub PageSetup(Orientation, pageSize)
        WordDoc.PageSetup.Orientation = wdOrientPortrait

        If pageSize = "wdSizeA4" then
            WordDoc.PageSetup.PageHeight = WordApp.MillimetersToPoints(297)
            WordDoc.PageSetup.PageWidth = WordApp.MillimetersToPoints(210)
        elseif pageSize = "wdSizeA3" then
            WordDoc.PageSetup.PageHeight = WordApp.MillimetersToPoints(419.9)
            WordDoc.PageSetup.PageWidth = WordApp.MillimetersToPoints(297)
        End If

        If Orientation = "wdOrientationLandscape" then
            WordDoc.PageSetup.Orientation = wdOrientLandscape
        End If

        If Orientation = "wdOrientationPortrait" then
            WordDoc.PageSetup.Orientation = wdOrientPortrait
        End If
    End sub

    ' With Selection.PageSetup
    '     .LineNumbering.Active = False
    '     .Orientation = wdOrientPortrait
    '     .TopMargin = MillimetersToPoints(24.2)
    '     .BottomMargin = MillimetersToPoints(20)
    '     .LeftMargin = MillimetersToPoints(20)
    '     .RightMargin = MillimetersToPoints(20)
    '     .Gutter = MillimetersToPoints(0)
    '     .HeaderDistance = MillimetersToPoints(7.5)
    '     .FooterDistance = MillimetersToPoints(0)
    '     .PageWidth = MillimetersToPoints(297)
    '     .PageHeight = MillimetersToPoints(419.9)
        '     .PageWidth = MillimetersToPoints(297)
        '    .PageHeight = MillimetersToPoints(210)
    '     .FirstPageTray = wdPrinterDefaultBin
    '     .OtherPagesTray = wdPrinterDefaultBin
    '     .SectionStart = wdSectionNewPage
    '     .OddAndEvenPagesHeaderFooter = False
    '     .DifferentFirstPageHeaderFooter = False
    '     .VerticalAlignment = wdAlignVerticalTop
    '     .SuppressEndnotes = False
    '     .MirrorMargins = False
    '     .TwoPagesOnOne = False
    '     .BookFoldPrinting = False
    '     .BookFoldRevPrinting = False
    '     .BookFoldPrintingSheets = 1
    '     .GutterPos = wdGutterPosLeft
    '     .LinesPage = 39
    '     .LayoutMode = wdLayoutModeLineGrid
    ' End With

    ''Const wdStyleHeading1 = -2
    ''Const wdStyleHeading2 = -3
    ''Sub AddHead(ByVal head As Long, ByRef text As String)
    Sub AddHead(ByVal head, ByRef text, ByRef idTitle)
        If head = 5 or head = 6 Then
            AddHead5 head, text, idTitle
            exit sub
        End If
        Dim heading 'As long
        heading = -1 - CLng(head)
        Dim rngNew
        set rngNew = AddText(text)
        rngNew.Style = heading

        Dim myParagraph
        set myParagraph = rngNew.Paragraphs(1).Range
        Dim keyOfHeading
        keyOfHeading = myParagraph.ListFormat.ListString & myParagraph.Text
        Dim r
        r = HeaderCollection.AddKyeNumber(idTitle)
        SetCurrentByEnd rngCurrent
    End Sub

    Sub AddHead5(ByVal head, ByRef text, ByRef idTitle)
        Dim rngNew
        set rngNew = AddText(text)
        rngNew.Style = GetStyle(wcStyleHeading5)

        Dim myParagraph
        set myParagraph = rngNew.Paragraphs(1).Range
        SetCurrentByEnd rngCurrent
    End Sub

    sub showCurrentRange(info)
        LogInfo String_Empty,String_Empty
        LogInfo "----------->: ", "showCurrentRange: " & info
        LogInfo "rngCurrent.start, End: ", rngCurrent.Start & ", " & rngCurrent.End
        Dim x
        For Each x In rngCurrent.Characters
            LogInfo "rngCurrent.Characters AscW: ", AscW(x)
        Next
        LogInfo "<-----------: ", String_Empty
        LogInfo String_Empty,String_Empty                                                            
    End sub

    ' Const wdStyleHeading7 = -8 > 
    ' Const wdStyleHeading8 = -9 '' 
    ' Const wdStyleHeading9 = -10 ''
    '' Sub AddOderList(ByVal head As Long, ByRef text As String)
    Sub AddOderList(ByVal head, ByRef text)
        ''AddHead head + 6, text, String_Empty
        Dim strHead ''As String
        strHead = "numlist" & CStr(head)
        AddText(text).Style = GetStyle(strHead)
    End Sub

    ''Sub AddNormalList(ByVal list As Long, ByRef text As String)
    Sub AddNormalList(ByVal list, ByRef text)
        Dim strHead ''As String
        strHead = "nlist" & CStr(list)
        AddText(text).Style = GetStyle(strHead)
    End Sub

    ''Sub AddNormalList(ByVal list As Long, ByRef text As String)
    Sub AddNote(ByRef text)
        Dim strHead ''As String
        strHead = "note1"
        AddText(text).Style = GetStyle(strHead)
    End Sub
    
    Sub AddNoteN(ByRef text)
        AddNote text
        Call AddNewLine("AddNoteN")
    End Sub

    Sub AddWarning(ByRef text)
        Dim strHead ''As String
        strHead = "warn1"
        AddText(text).Style = GetStyle(strHead)
    End Sub
    Sub AddWarningN(ByRef text)
        AddWarning text
        Call AddNewLine("AddWarningN")
    End Sub

    Sub AddCodeText(ByRef text)
        AddRawText(text).Style = GetStyle(wcStyleCode)
    End Sub

    sub SetCustomDocumentProperty(byval propertyName, byval value)
        dim Properties
        set Properties = WordDoc.CustomDocumentProperties

        If HasCustomDocumentProperty(propertyName) Then
            Properties.Item(propertyName) = value
            exit sub
        End If
        Properties.Add CStr(propertyName), False, msoPropertyTypeString, CStr(value)
    End Sub

    Function HasCustomDocumentProperty(byval propertyName)
        dim Properties
        set Properties = WordDoc.CustomDocumentProperties
        Dim i
        For i = 1 To Properties.Count
            If Properties.Item(i).Name = propertyName Then
                HasCustomDocumentProperty = True
                Exit Function
            End If
        Next
        HasCustomDocumentProperty = False
    End Function

    ''Sub AddText(ByRef text As String)
    Function AddText(ByRef text)
        If left(text,6) = "NOTE: " Then
            AddNoteN mid(text,7)
            exit Function
        elseif  left(text,9) = "WARNING: " Then
            AddWarningN mid(text,10)
            exit Function
        End If

        dim rngReturn
        set rngReturn = GetCurrentRangeStart()
        rngReturn.InsertAfter text
        ''emphasis rngReturn
        set AddText = rngReturn
        SetCurrentByEnd rngReturn
    End Function

    Function AddLineWithNewLine(ByRef text, ByRef nStyle)
        dim rngReturn
        set rngReturn = GetCurrentRangeStart()
        rngReturn.InsertAfter text
        set AddLineWithNewLine = rngReturn
        SetCurrentByEnd rngReturn
        
        '' nStyle is wd... or GetStyle(myStyle)
        rngCurrent.Style = nStyle

        rngCurrent.InsertParagraphBefore
        emphasis rngCurrent.Paragraphs(1).range
        SetCurrentByEnd rngCurrent

    End Function

    Function AddRawText(ByRef text)
        dim rngReturn
        set rngReturn = GetCurrentRangeStart()
        rngReturn.InsertAfter text
        set AddRawText = rngReturn
        SetCurrentByEnd rngReturn
    End Function

    Function SetCurrentByEnd(rng)
        set rngCurrent = WordDoc.Range(rng.End, rng.End)
        set  SetCurrentByEnd = rngCurrent
    End Function

    Sub SetCurrentNext()
        set rngCurrent = WordDoc.Range(rngCurrent.Start + 1, rngCurrent.End + 1)
    End Sub

    Function SetCurrentByTop(rng)
        set rngCurrent = WordDoc.Range(rng.Start, rng.Start)
        set  SetCurrentByTop = rngCurrent
    End Function

    Function SetCurrentPositionRangeTop()
        set rngCurrent = WordDoc.Range(0, 0) 
        set SetCurrentPositionRangeTop = rngCurrent
    End Function

    Function SetCurrentPositionRangeEnd()
        set rngCurrent = WordDoc.Bookmarks("\EndOfDoc").Range
        set SetCurrentPositionRangeEnd = rngCurrent
    End Function

    Function GetCurrentRangeStart()
        set  GetCurrentRangeStart = WordDoc.Range(rngCurrent.Start, rngCurrent.Start)
    End Function

    Function GetCurrentRangeEnd()
        set  GetCurrentRangeEnd = WordDoc.Range(rngCurrent.End, rngCurrent.End)  
    End Function

    sub AddHr()
        dim r
        set r = GetCurrentRangeStart()
        r.InlineShapes.AddHorizontalLineStandard
        SetCurrentNext
        SetCurrentNext
    End sub

    Sub AddNewLine(command)
        If (m_indent < 1) then
            m_indent = 1
        End If
        If (m_indent > 3) then
            m_indent = 3
        End If

        Dim myStyle
        myStyle = "body" & CStr(m_indent)
        rngCurrent.InsertParagraphBefore
        If command <> "convertCode" Then
            emphasis rngCurrent.Paragraphs(1).range
            math rngCurrent.Paragraphs(1).range
        End If
        SetCurrentByEnd rngCurrent
        If command <> "wd0NewLine" Then
            rngCurrent.Style = GetStyle(myStyle)
        End If
    End Sub

    Sub AddEndParagraph(command)
        If (m_indent < 1) then
            m_indent = 1
        End If
        If (m_indent > 3) then
            m_indent = 3
        End If

        If command <> "convertCode" Then
            emphasis rngCurrent.Paragraphs(1).range
            math rngCurrent.Paragraphs(1).range
        End If
    End Sub

    '' [text](ref "hover")
    Sub AddLink(ByRef ref, ByRef hover, ByRef text)

        '' If ref = String_Empty, for index. now set only text.
        '' what is index?
        If ref = String_Empty then
            AddText  "["
            AddText text
            AddText  "]" 
            exit sub
        End If

        '' normal link  etc. web site.
        If text <> String_Empty Then
            AddText  "["
            AddHyperLink ref, hover, text
            AddText  "]" 
            exit sub
        End If

        '' word xref
        'AddText "[["
        '' only add to ref collection, and set docx later
        dim rngRef
        set rngRef = addText(ref)
        ''Call RefCollection.AddRangeRefTitle(GetCurrentRangeStart, ref)
        Call RefCollection.AddRangeRefTitle(rngRef, ref, m_crossRef)
        'AddText "]]" 
    End Sub


    Sub AddHyperLink(ByRef ref, ByRef hover, ByRef text)
        ' WordDoc.Hyperlinks.Add Anchor:=.Range, Address:= "http:", SubAddress:=String_Empty, ScreenTip:=String_Empty, TextToDisplay:="disp"
        ' Anchor	Required	Object	The anchor for the hyperlink. Can be either a Range or Shape object.
        ' Address	Required	String	The address of the hyperlink.
        ' SubAddress	Optional	Variant	The SubAddress of the hyperlink.
        ' ScreenTip	Optional	Variant	The screen tip to be displayed when the mouse pointer is paused over the hyperlink.
        ' TextToDisplay	Optional	Variant	The text to be displayed for the hyperlink.
        dim hyperlink
        set hyperlink = WordDoc.Hyperlinks.Add(GetCurrentRangeStart, ref, String_Empty, hover, text)
        SetCurrentByEnd hyperlink.Range
    End Sub

    '' Sub AddImage(ByRef imagePath As String)
    '' mark downPath  to detect image path
    Sub AddImage(ByRef imagePathAsParam, text, title)
        Dim imagePath ''
        Dim imagePathR

        imagePathAsParam = replace(imagePathAsParam, "/","\")
        imagePathR = m_markdownDirPath & "\" & imagePathAsParam
        LogInfo "imagePathAsParam: ", imagePathAsParam
        LogInfo "imagePathR: ", imagePathR

        If fso.FileExists(imagePathAsParam) Then
            imagePath = imagePathAsParam
        elseif fso.FileExists(imagePathR) Then
            imagePath = imagePathR
        End If

        Dim thisShape ''As InlineShape

        ''WordApp..InlineShapes.AddPicture fileName:= imagePath, LinkToFile:=False, SaveWithDocument:= True
        If fso.FileExists(imagePath) Then
            set thisShape = GetCurrentRangeStart.InlineShapes.AddPicture(imagePath)
            Call shapeMatch(thisShape.Range)

            thisShape.Range.Style = GetStyle(wcStylePicture1)
            ''thisShape.Range.Style = GetStyle(wcStyleBody1)
            SetCurrentByEnd thisShape.Range
        Else
            AddText "Err: No image: " & imagePathAsParam
            LogWarn "No Image", imagePathAsParam
        End If
    End Sub

    '' // https://koukimra.com/archives/86
    Sub shapeMatch(rngPicture)
        Dim indentForShape '' single
        Dim i ''As Long
        Dim widthWord ''As Single
        DIm widthPictureWithIndent
        Dim widFull ''As Single 'point

        indentForShape = 50
        With rngPicture
            For i = 1 To .ShapeRange.Count
                With .ShapeRange(i)
                    If .WrapFormat.Type = wdWrapInline Then
                        .Width = GetMaxWidthInRange(.Anchor)
                    End If
                End With
            Next
        End With

        With rngPicture
            For i = 1 To .InlineShapes.Count
                With .InlineShapes(i)
                    widthWord = GetMaxWidthInRange(.Range)
                    widthPictureWithIndent = .Width + indentForShape

                    If (widthPictureWithIndent - widthWord) > 0 Then 'for width over
                        .Width = widthWord - indentForShape '  - .Left
                    Else
                        '' do nothing
                    End If
                    ' arrange ratio height to width
                    If .ScaleHeight <> .ScaleWidth Then
                        .ScaleHeight = .ScaleWidth
                    End If
                End With
            Next
        End With
    End Sub

    ''Function GetMaxWidthInRange(RNG As Range) As Single
    Function GetMaxWidthInRange(RNG) ' as single
        Dim widMax ''As Single
        With RNG
            If .Information(wdWithInTable) Then
                With .Cells(1)
                    widMax = .Width - (.LeftPadding + .RightPadding)
                End With
            Else
                'with of section
                widMax = .Sections(1).PageSetup.TextColumns(1).Width
            End If
            'width of paragraph
            With .Paragraphs(1)
                widMax = widMax - (.LeftIndent + .RightIndent)
            End With
        End With
        GetMaxWidthInRange = widMax
    End Function

    ''Sub AddTable(ByRef table() As String, columnWith() As String, mergeInfo() As String)
    Sub AddTable(table(), columnWith(), mergeInfo(), alignInfo())
        '' todo ??
        '' WordApp.ActiveDocument.Range.InsertParagraphAfter
        Dim oTable ''As Word.table
        'create table and assign it to variable
        Dim x ''As Long
        Dim y ''As Long
        Dim k ''As Long
        Dim tableRows ''As Long
        Dim tableColumns ''As Long
        tableRows = UBound(table, 1) + 1
        tableColumns = UBound(table, 2) + 1

        Dim tablePosition
        Set tablePosition = GetCurrentRangeEnd

        dim documentWidth
        documentWidth = GetMaxWidthInRange(tablePosition)
        tablePosition.Select

        Set oTable = WordApp.ActiveDocument.tables.Add(WordApp.ActiveDocument.Paragraphs.Last.Range, tableRows, tableColumns, 1)
        ''tablePosition.InsertCaption "•\", "InsertCaption2", "", wdCaptionPositionAbove, 0
        
        ' wdAutoFitContent	1	
        ' wdAutoFitFixed	0
        ' wdAutoFitWindow	2
        oTable.Style = GetTableStyle(wcStyleTableN)
        oTable.AutoFitBehavior wdAutoFitFixed
        
        Dim tableWidth ''As Single
        Dim tableWidthSettings ''As Single
        tableWidth = 0
        tableWidthSettings = 0
        For y = 1 To tableColumns
            tableWidth = tableWidth + oTable.columns(y).Width
            tableWidthSettings = tableWidthSettings + CSng(columnWith(y - 1))
        Next
        
        For y = 1 To tableColumns
        '' word table 1.., vbs array 0..
            oTable.columns(y).Width = (documentWidth * 0.99) * columnWith(y - 1) / tableWidthSettings
        Next
        
        oTable.Style = GetTableStyle(wcStyleTableN)
        Dim MergeEnd
        Dim cellRange
       
        '' insert value
        For x = 1 To tableRows
            For y = 1 To tableColumns
                Dim align
                If alignInfo(x-1, y-1) = "right" Then
                    align = wdAlignParagraphRight
                Elseif alignInfo(x-1, y-1) = "center" then
                    align = wdAlignParagraphCenter
                Else    
                    align = wdAlignParagraphLeft
                End If
                oTable.Cell(x, y).Range.ParagraphFormat.Style = wdStyleNormal
                oTable.Cell(x, y).Range.ParagraphFormat.Alignment = align

                set rngCurrent = oTable.Cell(x, y).Range
                Dim tableItem
                set tableItem = table(x - 1, y - 1)
                for k = 1 to tableItem.Count
                    '' todo markdown path for images
                    Call doCommands(tableItem.Item(k), String_Empty, String_Empty, String_Empty)
                next
            Next
        Next

        ' merge cells
        For x = tableRows To 1 step -1
            For y = 1 To tableColumns
                If  mergeInfo(x - 1, y - 1) <> String_Empty Then
                    '' mergeInfo  End row, End column(same)
                    MergeEnd = split(mergeInfo(x-1, y-1), ",")
                    If (CStr(x-1) <> MergeEnd(0)) and (CStr(y-1) = MergeEnd(1)) Then
                        oTable.Cell(x, y).Merge oTable.Cell(MergeEnd(0)+1, MergeEnd(1)+1)
                    Elseif (CStr(x-1) = MergeEnd(0)) and ( CStr(y-1) <> MergeEnd(1)) Then
                        oTable.Cell(x, y).Merge oTable.Cell(MergeEnd(0)+1, MergeEnd(1)+1)
                    Else
                        ''log "no merge", String_Empty, String_Empty
                    End If
                End If
            Next
        Next
        oTable.Style = GetTableStyle(wcStyleTableN)

        Dim r
        Set r = oTable.Range
        r.SetRange oTable.Range.End + 1, oTable.Range.End + 1
        SetCurrentByEnd r
        Call AddNewLine("AddTable")
    End Sub

    Sub AddMath(rng)
        Dim objRange ''As Range
        Dim objEq ''As OMath
        Set objRange = rng.OMaths.Add(rng)
        Set objEq = objRange.OMaths(1)

        'objEq.ParentOMath.Type = wdOMathInline '   wdOMathInline 1, wdOMathDisplay 0
        objEq.BuildUp
    End Sub


    ''Function getTableInfoArray(wdLines As XFiles, iCurrent As Long, columnWith() As String) As String()
    Function getTableInfoArrayEx(wdLines, iCurrent, columnWith, mergeInfo, alignInfo) ''As String()
        Dim i ''As Long
        Dim iInfo ''As Long

        Dim strSplit ''As String
        strSplit = Split(wdLines.Item(iCurrent), vbTab)

        Dim collectionValue
        Set collectionValue = New Collection

        '' get create table and rows , columns
        Dim tableInfo() ''As String()
        ReDim tableInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim mergeInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim alignInfo(strSplit(1) - 1, strSplit(2) - 1)
        ReDim columnWith(strSplit(2) - 1)
        Dim cellCount ''As Long
        cellCount = CLng(strSplit(1)) * CLng(strSplit(2))

        iCurrent = iCurrent +1

        Dim strCommand
        Dim p1
        Dim p2
        Dim p3
        Dim p4
        Dim currentLine
        Dim strColumnWidth ''As String
        strColumnWidth = "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"

        for i = iCurrent To 32000
            If i mod logCycle = 0 Then
                Me.logProgress i
            End If
            currentLine = Trim(wdLines.Item(i))
            If currentLine <> String_Empty Then
                strSplit = Split(wdLines.Item(i), vbTab)
                strCommand = LCase(strSplit(0))
                Select Case strCommand
                    Case "tablecontents"
                        '' tablecontents, row, col, value, align
                        Set collectionValue = New Collection
                        set tableInfo(CLng(strSplit(1)), CLng(strSplit(2))) = collectionValue
                        alignInfo(CLng(strSplit(1)), CLng(strSplit(2))) = strSplit(4)
                    Case "tablecontentslist"
                        collectionValue.AddAutoNumKey mid(currentLine, len("tablecontentslist") + 2)
                    Case "tablewidthinfo"
                        strColumnWidth = strSplit(1) & "," & strColumnWidth
                    Case "tablemarge"
                        '' merge rows
                        '' tableMarge, start row, column, End row, column, value(start row, column)
                        If ubound(strSplit) > 4 then
                            If (strSplit(5)) <> String_Empty or True then
                                If strSplit(1) = strSplit(3) Then
                                    '
                                Else
                                    mergeInfo(CLng(strSplit(1)), CLng(strSplit(2))) = strSplit(3) & "," &  strSplit(4)
                                End If
                            End If
                        End If
                    Case Else
                        iCurrent = i - 1
                        exit for
                End Select
            Else
                iCurrent = i
                exit for
            End If
        next

        '' get columns size
        For i = 0 To UBound(columnWith)
            '' column info is separate by comma
            columnWith(i) = Split(strColumnWidth, ",")(i)
        Next

        '' cell merge info
        iCurrent = iCurrent
        getTableInfoArrayEx = tableInfo
    End Function


    '' todo emphasis 
    ''  https://www.wordvbalab.com/code/7996/
    '' Sub enf(myRange As range)
    Sub emphasis(tRange)
        Dim myRange
        Dim myTempRange ''As Range
        Dim myChr(7) '(1 To 6) ''As String
        Dim i ''As Integer

        If (instr(tRange.Text,"<")) = 0 Then
            exit Sub
        End If

        myChr(1) = wcEmphasisCodeSpan
        myChr(2) = wcEmphasisSubscript
        myChr(3) = wcEmphasisSuperscript
        myChr(4) = wcEmphasisBold
        myChr(5) = wcEmphasisItalic
        myChr(6) = wcEmphasisStrikeOut
        myChr(7) = wcEmphasisUnderline

        For i = 1 To 6
            '' If tRange.End, Find can not detect targets.
            set myRange = WordDoc.range(tRange.Start, tRange.End + 1)
            With myRange.Find
                .Text = "\<" & myChr(i) & "\>*\</" & myChr(i) & "\>"
                .Wrap = 0 ''wdFindStop: 0
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchByte = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchFuzzy = False
                .MatchWildcards = True
                Do While .Execute = True
                    With myRange
                        '' If the style codeSpan is set, no other emphasis does not need.
                        If .Style <> GetStyle(wcStyleCodeSpan) Then
                            Select Case myChr(i)
                                Case wcEmphasisSubscript
                                    .Font.Subscript = True
                                Case wcEmphasisSuperscript
                                    .Font.Superscript = True
                                Case wcEmphasisBold
                                    ''.Font.Bold = True
                                    .Style = wdStyleStrong
                                Case wcEmphasisItalic
                                    .Style = wdStyleEmphasis
                                    ''.Font.Italic = True
                                Case wcEmphasisUnderline
                                    .Font.Underline = 1 '' wdUnderlineSingle: 1
                                Case wcEmphasisStrikeOut
                                    .Font.StrikeThrough = True
                                case wcEmphasisCodeSpan
                                    .Style = GetStyle(wcStyleCodeSpan)
                                    ' With .Font.Shading
                                    '     .Texture = 150 ''wdTexture15Percent: 150
                                    '     .ForegroundPatternColor = 0 ''wdColorBlack: 0
                                    '     .BackgroundPatternColor = 16777215 ''wdColorWhite: 16777215
                                    ' End With
                            End Select
                            
                            'delete End tag
                            Set myTempRange =  WordDoc.Range(.End - Len(myChr(i)) - 3, .End)
                            myTempRange.Delete
                            
                            'delete start tag
                            Set myTempRange =  WordDoc.Range(.Start, .Start + Len(myChr(i)) + 2)
                            myTempRange.Delete
                            .Collapse 0 ''wdCollapseEnd: 0
                        End If

                        ''wdCollapseEnd	0	
                        ''wdCollapseStart	1	
                    End With
                Loop
            End With
        Next        
    End Sub

    private Function GetStyle(strStyle)
        dim rStyle
        on error resume next
        rStyle = WordDoc.Styles(strStyle)

        If Err.Number <> 0 Then
            LogWarn "No Style", strStyle
            rStyle = WordDoc.Styles(wdStyleNormal)
        End If
        on error goto 0
        GetStyle = rStyle
    End Function

    private Function GetTableStyle(strStyle)
        dim rStyle
        on error resume next
        rStyle = WordDoc.Styles(strStyle)

        If Err.Number <> 0 Then
            LogWarn "No Style", strStyle
            rStyle = wdStyleTableLightShading
        End If
        on error goto 0
        GetTableStyle = rStyle
    End Function


    Sub Math(tRange)
        If isMath = False Then
            Exit Sub
        End IF
        Dim myRange
        Dim myTempRange ''As Range
        Dim myChr(7) '(1 To 6) ''As String
        Dim i ''As Integer

        If (instr(tRange.Text,"$")) = 0 Then
            exit Sub
        End If

        ''LogInfo "math", tRange.Text

        myChr(1) = "$$"
        myChr(2) = "$"

        Dim rngMath

        For i = 1 To 2
            '' If tRange.End, Find can not detect targets.
            set myRange = WordDoc.range(tRange.Start, tRange.End + 1)
            With myRange.Find
                .Text = myChr(i) & "*" & myChr(i)
                .Wrap = 0 ''wdFindStop: 0
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchByte = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchFuzzy = False
                .MatchWildcards = True
                Do While .Execute = True
                    With myRange
                        If .Style <> GetStyle(wcStyleCodeSpan) Then
                            Select Case myChr(i)
                                Case "$"
                                    .Font.Subscript = True
                                Case "$$"
                                    .Font.Superscript = True
                            End Select

                            '' add math
                            
                            'delete End tag
                            Set myTempRange =  WordDoc.Range(.End - Len(myChr(i)), .End)
                            myTempRange.Delete


                            
                            'delete start tagmamath
                            Set myTempRange =  WordDoc.Range(.Start, .Start + Len(myChr(i)))
                            myTempRange.Delete
                            AddMath myRange

                            .Collapse 0 ''wdCollapseEnd: 0
                            ''wdCollapseEnd	0	
                            ''wdCollapseStart	1	
                        End If
                    End With
                Loop
            End With
        Next        
    End Sub

    Sub AddXRef()
        Dim Items
        Items = RefCollection.Items()
        Dim pos
        Dim HeadingNo
        Dim rngInsert
        Dim fmt
        Dim ii
        for ii = 0 to ubound(Items)
            fmt = Items(ii).RefFormat
            pos =  Items(ii).RefPosition.Start
            If HeaderCollection.Exists(Items(ii).RefTitle) Then
                Items(ii).RefPosition.Text = String_Empty
                HeadingNo = HeaderCollection.Item(Items(ii).RefTitle) 
                AddXrefCore fmt, pos, HeadingNo
            else
                LogWarn "No xref", Items(ii).RefTitle
            End If
        Next
    End Sub 

    sub AddXrefCore(sRef, pos, HeadingNo)
        sRef = sRef + " "
        Dim sArrRef(100)
        Dim i
        dim t
        dim c
        c = 0
        for i = 1 to len(sRef) - 1
            t = mid(sRef,i,2)
            If t ="$n" then
                sArrRef(c) = t
                i = i + 1
            elseif t = "$t" then
                sArrRef(c) = t
                i = i + 1
            elseif t = "$p" then
                sArrRef(c) = t
                i = i + 1
            else
                sArrRef(c) = left(t, 1)
            End If
            c = c + 1
        next

        for i = c to 0 step -1
            t = sArrRef(i)
            If t ="$n" then
                WordDoc.Range(pos, pos).InsertCrossReference wdRefTypeHeading, wdNumberFullContext, HeadingNo, True, False, False, String_Space 
            elseif t = "$t" then
                WordDoc.Range(pos, pos).InsertCrossReference wdRefTypeHeading, wdContentText, HeadingNo, True, False, False, String_Space
            elseif t = "$p" then
                WordDoc.Range(pos, pos).InsertCrossReference wdRefTypeHeading, wdPageNumber, HeadingNo, True, False, False, String_Space 
            else
                WordDoc.Range(pos, pos).InsertBefore t
            End If
        next
    End sub


    ' Name	Value	Description
    ' wdCommentsStory	4	Comments story.
    ' wdEndnoteContinuationNoticeStory	17	Endnote continuation notice story.
    ' wdEndnoteContinuationSeparatorStory	16	Endnote continuation separator story.
    ' wdEndnoteSeparatorStory	15	Endnote separator story.
    ' wdEndnotesStory	3	Endnotes story.
    ' wdEvenPagesFooterStory	8	Even pages footer story.
    ' wdEvenPagesHeaderStory	6	Even pages header story.
    ' wdFirstPageFooterStory	11	First page footer story.
    ' wdFirstPageHeaderStory	10	First page header story.
    ' wdFootnoteContinuationNoticeStory	14	Footnote continuation notice story.
    ' wdFootnoteContinuationSeparatorStory	13	Footnote continuation separator story.
    ' wdFootnoteSeparatorStory	12	Footnote separator story.
    ' wdFootnotesStory	2	Footnotes story.
    ' wdMainTextStory	1	Main text story.
    ' wdPrimaryFooterStory	9	Primary footer story.
    ' wdPrimaryHeaderStory	7	Primary header story.
    ' wdTextFrameStory	5	Text frame story.
    Sub UpdateFields()
        Dim aStory ''As Range
        Dim aField ''As Field

        For Each aStory In WordDoc.StoryRanges
            For Each aField In aStory.Fields
                aField.Update
            Next
        Next
    End Sub


'' WordApp..InsertCrossReference()
' ReferenceType	Required	Variant	The type of item for which a cross-reference is to be inserted. 
'                           Can be any WdReferenceType or WdCaptionLabelID Constant or a user defined caption label.
' ReferenceKind	Required	WdReferenceKind	The information to be included in the cross-reference.
' ReferenceItem	Required	Variant	If ReferenceType is wdRefTypeBookmark, this argument specifies a bookmark name. 
'                           For all other ReferenceType values, this argument specifies the item number or name 
'                           in the Reference type option in the Cross-reference dialog box. 
'                           Use the GetCrossReferenceItems method to return a list of item names that can be used with this argument.
' InsertAsHyperlink	Optional	Variant	True to insert the cross-reference as a hyperlink to the referenced item.
' IncludePosition	Optional	Variant	True to insert "above" or "below," depending on the location of the reference item 
'                               in relation to the cross-reference.
' SeparateNumbers	Optional	Variant	True to use a separator to separate the numbers from the associated text.
'                               (Use only If the ReferenceType parameter is set to wdRefTypeNumberedItem 
'                               and the ReferenceKind parameter is set to wdNumberFullContext.)
' SeparatorString	Optional	Variant	Specifies the string to use as a separator If the SeparateNumbers parameter is set to True.

'' ReferenceKind
' wdContentText	-1	Insert text value of the specified item. For example, insert text of the specified heading.
' wdEndnoteNumber	6	Insert endnote reference mark.
' wdEndnoteNumberFormatted	17	Insert formatted endnote reference mark.
' wdEntireCaption	2	Insert label, number, and any additional caption of specified equation, figure, or table.
' wdFootnoteNumber	5	Insert footnote reference mark.
' wdFootnoteNumberFormatted	16	Insert formatted footnote reference mark.
' wdNumberFullContext	-4	Insert complete heading or paragraph number.
' wdNumberNoContext	-3	Insert heading or paragraph without its relative position in the outline numbered list.
' wdNumberRelativeContext	-2	Insert heading or paragraph with as much of its relative position 
'                               in the outline numbered list as necessary to identify the item.
' wdNumberFullContext wdNumberNoContext wdNumberRelativeContext	-4	Insert complete heading or paragraph number.
' wdNumberNoContext	-3	Insert heading or paragraph without its relative position in the outline numbered list.
' wdNumberRelativeContext
' wdOnlyCaptionText	4	Insert only the caption text of the specified equation, figure, or table.
' wdOnlyLabelAndNumber	3	Insert only the label and number of the specified equation, figure, or table.
' wdPageNumber	7	Insert page number of specified item.
' wdPosition	15	Insert the word "Above" or the word "Below" as appropriate.
End Class

Class XFiles
    Dim m_Files
    Dim m_FSO

    Private Sub Class_Initialize
        Set m_Files = CreateObject(Scripting_Dictionary)
        Set m_FSO = CreateObject(Scripting_FileSystemObject)
    End Sub

    Public Function Files()
        set Files = m_Files
    End Function

    Public Sub Clear()
        Set m_Files = CreateObject(Scripting_Dictionary)
    End Sub

    Public Sub Add(key, value)
        m_Files.Add key, value
    End Sub

    Public Sub AddAuto(value)
        m_Files.Add m_Files.Count + 1, value
    End Sub

    Public Function Count()
        Count = m_Files.Count
    End Function

    Public Function Item(i)
        Item = m_Files.Item(i)
    End Function

    Public Function ActiveWdItem(i)
        Dim r
        r =  m_Files.Item(i)
        If Len(r) > 1 then
            If left(r, 2) = "//" then
                ActiveWdItem = String_Empty
                exit function
            End If
        End If
        ActiveWdItem = r
    End Function

    Public Sub Remove(key)
        m_Files.remove key
    End Sub

    Public Sub Load(filePath)
        Dim pathFile
        Dim ppath
        Dim currentLine, tmpSplitLine, comment
        Dim firstChar
        Dim currentLineNo
        Set m_Files = CreateObject(Scripting_Dictionary)

        If m_FSO.FileExists(filePath) = false Then
            exit Sub
        End If
        With CreateObject("ADODB.Stream")
            .Charset = "UTF-8"
            .Open
            .LineSeparator = 10
            .LoadFromFile filePath
            Do Until .EOS
                ' adReadAll ' -1
                ' for byte
                ' adReadLine' -2
                ' real next line (LineSeparator property)B
                '[; ] relative path | comment | info
                currentLine = .ReadText(-2)
                currentLineNo = currentLineNo + 1
                m_Files.add currentLineNo, currentLine
            Loop
            .Close
        End with
    End Sub

    Sub WriteToFile(filename)
        'Dim writeStream As ADODB.Stream
        'Microsoft ActiveX Data Objects 2.5 Library
        ' WriteText str, 1 => add a newline
        ' WriteText str, 0 => add no newline
        Dim writeStream

        Set writeStream = CreateObject("ADODB.Stream")
        writeStream.Charset = "UTF-8"
        writeStream.Open

        'write to stream
        Dim i, items
        items = m_Files.items
        For i = 0 to m_Files.Count -1
            writeStream.WriteText m_Files.item(i)
            writeStream.WriteText String_Empty, 1
        next

        ' write to a file
        writeStream.SaveToFile filename, 2 'adSaveCreateOverWrite:2

        writeStream.Close
        Set writeStream = Nothing
    End Sub
End Class


Class Collection
    Dim m_collection

    Private Sub Class_Initialize
        Set m_collection = CreateObject(Scripting_Dictionary)
    End Sub

    Public Sub AddNumberValue(value)
        m_collection.Add CStr(m_collection.Count + 1), value
    End Sub

    Public Function AddKyeNumber(Key)
        Dim rKey
        rKey = TrimNewLine(Key)
        
        Dim numberForDuplicated
        numberForDuplicated = 1

        Do While m_collection.Exists(rKey)
            rKey = rKey & "-" & CStr(numberForDuplicated)
        Loop

        m_collection.Add rKey, CStr(m_collection.Count + 1)
        AddKyeNumber = rKey
    End Function

    Public Sub AddRangeRefTitle(rng, title, fmt)
        dim ref '' XRef
        set ref = new XRef
        Call ref.SetRef(rng, title, fmt)
        m_collection.Add CStr(m_collection.Count + 1), ref
        set ref = Nothing
    End Sub

    Public Function TrimNewLine(s)
        Dim tmp
        tmp = s
        tmp = replace(tmp, vbCr, String_Empty) 
        tmp = replace(tmp, vbLf, String_Empty) 
        TrimNewLine = tmp
    End Function

    Public Sub Clear()
        Set m_collection = CreateObject(Scripting_Dictionary)
    End Sub

    Public Sub AddAutoNumKey(value)
        dim key
        key = m_collection.Count + 1
        m_collection.Add key, value
    End Sub

    Public Sub Add(key, value)
        m_collection.Add key, value
    End Sub

    Public Function Exists(Key)
        Exists = m_collection.Exists(Key)
    End Function

    Public Function Keys()
        Keys = m_collection.Keys()
    End Function

    Public Function Items()
        Items = m_collection.Items()
    End Function

    Public Function Count()
        Count = m_collection.Count
    End Function

    Public Function Item(i)
        Item = m_collection.Item(i)
    End Function

    Public Sub Remove(key)
        m_collection.remove key
    End Sub

    ' vbUseCompareOption	-1	
    ' vbBinaryCompare	0	
    ' vbTextCompare	1	
    Public Sub SetCompareMode(cMode)
        m_collection.CompareMode = cMode
    End Sub
End Class

Class XRef
    Public RefPosition ''As Range
    Public RefTitle ''As String
    Public RefFormat '' as string 

    Sub SetRef(position, title, fmt)
        Set RefPosition = position
        RefTitle = title
        RefFormat = fmt
    End Sub
End Class

Sub LogDebug(ByVal title, Byval value)
  Call LogCore("DBG", title, value)
End Sub

Sub LogInfo(title, value)
  If isLogInfo Then
  Call LogCore("INF", title, value)
  End If
End Sub

Sub LogWarn(title, value)
  Call LogCore("WRN", title, value)
  Call logWarns.AddAuto("WRN" & ":" & title & " : " & value)
End Sub

Sub LogError(title, value)
  Call LogCore("ERR", title, value)
End Sub

Sub LogCore(messageType, title, value)
    ''exit Sub
    Dim outTitle
    Dim outValue
    outTitle = title
    outValue = value
    If outTitle = String_Empty Then
        outTitle = "(_empty_)"
    End If
    If outValue = String_Empty Then
        outValue = "(_empty_)"
    End If
    on error resume next
    WScript.StdOut.WriteLine messageType & ":" & outTitle & " : " & outValue
    If err.number <> 0 Then
        WScript.StdErr.WriteLine "ERR" & ":" & Err.Description & " : " & Err.Number 
    End If
    on error goto 0
End Sub

Function Catch(source, errCodeExit)
    exit function
    If Err.Number <> 0 Then
        WScript.StdErr.WriteLine (source)
        WScript.StdErr.WriteLine (errCodeExit)
        WScript.StdErr.WriteLine Err.Description
        LogError "Catch", source & "(" & Err.Description & ")"
    End If

    ''fatal
    If errCodeExit > 1000 Then
        WordApp.Visible = True
        WScript.Quit(errCodeExit)
    End If

    '' If predict error
    '' do error treat after this
    If Err.Number = errCodeExit Then
        Catch = True
        On Error Goto 0
    Else
        If isResumeNext then
            On Error Goto 0
            on error resume next
        End If
    End If
End Function






