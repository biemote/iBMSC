Imports iBMSC.Editor.Functions

Partial Public Class MainWindow

    Private Sub XMLWriteColumn(ByVal w As XmlTextWriter, ByVal I As Integer)
        w.WriteStartElement("Column")
        w.WriteAttributeString("Index", I.ToString())
        With column(I)
            'w.WriteAttributeString("Left", .Left)
            w.WriteAttributeString("Width", .Width.ToString())
            w.WriteAttributeString("Title", .Title)
            'w.WriteAttributeString("Text", .Text)
            'w.WriteAttributeString("Enabled", .Enabled)
            'w.WriteAttributeString("isNumeric", .isNumeric)
            'w.WriteAttributeString("Visible", .Visible)
            'w.WriteAttributeString("Identifier", .Identifier)
            w.WriteAttributeString("NoteColor", .cNote.ToString())
            w.WriteAttributeString("TextColor", .cText.ToArgb.ToString())
            w.WriteAttributeString("LongNoteColor", .cLNote.ToString())
            w.WriteAttributeString("LongTextColor", .cLText.ToArgb.ToString())
            w.WriteAttributeString("BG", .cBG.ToArgb.ToString())
        End With
        w.WriteEndElement()
    End Sub

    Private Sub XMLWriteFont(ByVal w As XmlTextWriter, ByVal local As String, ByVal f As Font)
        w.WriteStartElement(local)
        w.WriteAttributeString("Name", f.FontFamily.Name)
        w.WriteAttributeString("Size", f.SizeInPoints.ToString())
        w.WriteAttributeString("Style", CInt(f.Style).ToString())
        w.WriteEndElement()
    End Sub

    Private Sub XMLWriteKeybindings(ByVal w As XmlTextWriter, ByVal I As Integer)
        w.WriteStartElement("Option")
        w.WriteAttributeString("OpVar", Keybindings(I).OpVar.ToString())
        w.WriteAttributeString("Name", Keybindings(I).OpName)
        ' w.WriteAttributeString("Description", Keybindings(I).Description)
        w.WriteAttributeString("Combos", Join(Keybindings(I).Combo, ", "))
        ' w.WriteAttributeString("Category", Keybindings(I).Category.ToString())
        w.WriteEndElement()
    End Sub

    Private Sub XMLWriteOpenedFiles(ByVal w As XmlTextWriter, ByVal I As Integer)
        w.WriteStartElement("File")
        w.WriteAttributeString("Index", I.ToString())
        w.WriteAttributeString("File", BMSFiles(I).Filename)
        w.WriteAttributeString("Color", BMSFiles(I).TabColor.ToArgb.ToString())
        w.WriteEndElement()
    End Sub

    Private Sub XMLWritePlayerArguments(ByVal w As XmlTextWriter, ByVal I As Integer)
        w.WriteStartElement("Player")
        w.WriteAttributeString("Index", I.ToString())
        w.WriteAttributeString("Path", pArgs(I).Path)
        w.WriteAttributeString("FromBeginning", pArgs(I).aBegin)
        w.WriteAttributeString("FromHere", pArgs(I).aHere)
        w.WriteAttributeString("Stop", pArgs(I).aStop)
        w.WriteEndElement()
    End Sub

    Private Sub XMLWriteToolbarLocation(ByVal w As XmlTextWriter, ByVal TB As ToolStrip)
        w.WriteStartElement(TB.Name)
        If ToolStripContainer1.TopToolStripPanel.Controls.Contains(TB) Then
            w.WriteAttributeString("Location", "Top")
        ElseIf ToolStripContainer1.LeftToolStripPanel.Controls.Contains(TB) Then
            w.WriteAttributeString("Location", "Left")
        ElseIf ToolStripContainer1.BottomToolStripPanel.Controls.Contains(TB) Then
            w.WriteAttributeString("Location", "Bottom")
        Else
            w.WriteAttributeString("Location", "Right")
        End If
        w.WriteAttributeString("X", TB.Location.X.ToString())
        w.WriteAttributeString("Y", TB.Location.Y.ToString())
        w.WriteEndElement()
    End Sub

    Private Sub SaveSettings(ByVal Path As String, ByVal ThemeOnly As Boolean)
        Dim w As New XmlTextWriter(Path, System.Text.Encoding.Unicode)
        With w
            .WriteStartDocument()
            .Formatting = Formatting.Indented
            .Indentation = 4

            .WriteStartElement("iBMSC")
            .WriteAttributeString("Major", My.Application.Info.Version.Major.ToString())
            .WriteAttributeString("Minor", My.Application.Info.Version.Minor.ToString())
            .WriteAttributeString("Build", My.Application.Info.Version.Build.ToString())

            If Not ThemeOnly Then

                .WriteStartElement("Form")
                .WriteAttributeString("WindowState", IIf(isFullScreen, previousWindowState, Me.WindowState).ToString())
                .WriteAttributeString("Width", IIf(isFullScreen, previousWindowPosition.Width, Me.Width).ToString())
                .WriteAttributeString("Height", IIf(isFullScreen, previousWindowPosition.Height, Me.Height).ToString())
                .WriteAttributeString("Top", IIf(isFullScreen, previousWindowPosition.Top, Me.Top).ToString())
                .WriteAttributeString("Left", IIf(isFullScreen, previousWindowPosition.Left, Me.Left).ToString())
                .WriteAttributeString(POptionsScroll.Name, POptionsScroll.Width.ToString())
                For Each CB In {POHeaderSwitch, POGridSwitch, POWaveFormSwitch, POWAVSwitch, POBMPSwitch, POBeatSwitch, POExpansionSwitch, POHeaderExpander, POGridExpander, POWaveFormExpander, POWAVExpander, POBMPExpander, POBeatExpander}
                    .WriteAttributeString(CB.Name, CB.Checked.ToString())
                Next
                For Each P In {POHeaderInner, POGridInner, POWaveFormInner, POWAVInner, POBMPInner, POBeatInner, POExpansionInner}
                    .WriteAttributeString(P.Name, P.Height.ToString())
                Next

                For Each TB As ToolStrip In {TBMain, TBTab}
                    XMLWriteToolbarLocation(w, TB)
                Next
                .WriteEndElement()

                .WriteStartElement("Recent")
                .WriteAttributeString("Recent0", Recent(0))
                .WriteAttributeString("Recent1", Recent(1))
                .WriteAttributeString("Recent2", Recent(2))
                .WriteAttributeString("Recent3", Recent(3))
                .WriteAttributeString("Recent4", Recent(4))
                .WriteEndElement()

                .WriteStartElement("OpenedFiles")
                .WriteAttributeString("Count", UBound(BMSFiles).ToString())
                .WriteAttributeString("BMSFileIndex", BMSFileIndex.ToString())
                For i = 0 To UBound(BMSFiles)
                    XMLWriteOpenedFiles(w, i) : Next
                .WriteEndElement()


                .WriteStartElement("Edit")
                .WriteAttributeString("NTInput", NTInput.ToString())
                .WriteAttributeString("Language", DispLang)
                '.WriteAttributeString("SortingMethod", SortingMethod)
                .WriteAttributeString("ErrorCheck", ErrorCheck.ToString())
                .WriteAttributeString("AutoFocusMouseEnter", AutoFocusMouseEnter.ToString())
                .WriteAttributeString("FirstClickDisabled", FirstClickDisabled.ToString())
                .WriteAttributeString("ShowFileName", ShowFileName.ToString())
                .WriteAttributeString("ShowWaveform", ShowWaveform.ToString())
                .WriteAttributeString("MiddleButtonMoveMethod", MiddleButtonMoveMethod.ToString())
                .WriteAttributeString("AutoSaveInterval", AutoSaveInterval.ToString())
                .WriteAttributeString("PreviewOnClick", PreviewOnClick.ToString())
                '.WriteAttributeString("PreviewErrorCheck", PreviewErrorCheck)
                .WriteAttributeString("ClickStopPreview", ClickStopPreview.ToString())
                .WriteAttributeString("JackBPM", ErrorJackBPM.ToString())
                .WriteAttributeString("JackTH", ErrorJackTH.ToString())
                .WriteAttributeString("gLNGap", gLNGap.ToString())
                .WriteAttributeString("TempFileName", TempFileName.ToString())
                .WriteAttributeString("COverridesSaveOption", COverridesSaveOption.ToString())
                .WriteAttributeString("TemplateSnapToVPosition", TemplateSnapToVPosition.ToString())
                .WriteAttributeString("PastePatternToVPosition", PastePatternToVPosition.ToString())
                .WriteEndElement()

                .WriteStartElement("Save")
                .WriteAttributeString("TextEncoding", EncodingToString(TextEncoding))
                .WriteAttributeString("BMSGridLimit", BMSGridLimit.ToString())
                .WriteAttributeString("BeepWhileSaved", BeepWhileSaved.ToString())
                .WriteAttributeString("PreloadBMSStruct", PreloadBMSStruct.ToString())
                .WriteAttributeString("BPMx1296", BPMx1296.ToString())
                .WriteAttributeString("STOPx1296", STOPx1296.ToString())
                .WriteAttributeString("AudioLine", AudioLine.ToString())
                .WriteEndElement()

                .WriteStartElement("WAV")
                .WriteAttributeString("WAVMultiSelect", WAVMultiSelect.ToString())
                .WriteAttributeString("WAVChangeLabel", WAVChangeLabel.ToString())
                .WriteAttributeString("BeatChangeMode", BeatChangeMode.ToString())
                .WriteEndElement()

                .WriteStartElement("BMP")
                .WriteAttributeString("BMPMultiSelect", BMPMultiSelect.ToString())
                .WriteAttributeString("BMPChangeLabel", BMPChangeLabel.ToString())
                .WriteEndElement()

                .WriteStartElement("ShowHide")
                .WriteAttributeString("showMenu", mnSMenu.Checked.ToString())
                .WriteAttributeString("showTB", mnSTB.Checked.ToString())
                .WriteAttributeString("showOpPanel", mnSOP.Checked.ToString())
                .WriteAttributeString("showStatus", mnSStatus.Checked.ToString())
                .WriteAttributeString("showLSplit", mnSLSplitter.Checked.ToString())
                .WriteAttributeString("showRSplit", mnSRSplitter.Checked.ToString())
                .WriteEndElement()

                .WriteStartElement("Grid")
                .WriteAttributeString("gSnap", gSnap.ToString())
                .WriteAttributeString("xmlDisableVertical", DisableVerticalMove.ToString())
                .WriteAttributeString("gWheel", gWheel.ToString())
                .WriteAttributeString("gPgUpDn", gPgUpDn.ToString())
                .WriteAttributeString("gShow", gShowGrid.ToString())
                .WriteAttributeString("gShowS", gShowSubGrid.ToString())
                .WriteAttributeString("gShowBG", gShowBG.ToString())
                .WriteAttributeString("gShowM", gShowMeasureNumber.ToString())
                .WriteAttributeString("gShowV", gShowVerticalLine.ToString())
                .WriteAttributeString("gShowMB", gShowMeasureBar.ToString())
                .WriteAttributeString("gShowC", gShowC.ToString())
                .WriteAttributeString("gBPM", gBPM.ToString())
                .WriteAttributeString("gSTOP", gSTOP.ToString())
                .WriteAttributeString("gSCROLL", gSCROLL.ToString())
                .WriteAttributeString("gBLP", gDisplayBGAColumn.ToString())
                .WriteAttributeString("gP2", CHPlayer.SelectedIndex.ToString())
                .WriteAttributeString("gCol", CGB.Value.ToString())
                .WriteAttributeString("gDivide", gDivide.ToString())
                .WriteAttributeString("gSub", gSub.ToString())
                .WriteAttributeString("gSlash", gSlash.ToString())
                .WriteAttributeString("gxHeight", gxHeight.ToString())
                .WriteAttributeString("gxWidth", gxWidth.ToString())
                .WriteEndElement()

                .WriteStartElement("WaveForm")
                .WriteAttributeString("wLock", wLock.ToString())
                .WriteAttributeString("wPosition", wPosition.ToString())
                .WriteAttributeString("wLeft", wLeft.ToString())
                .WriteAttributeString("wWidth", wWidth.ToString())
                .WriteAttributeString("wPrecision", wPrecision.ToString())
                .WriteEndElement()

                .WriteStartElement("Player")
                .WriteAttributeString("Count", pArgs.Length.ToString())
                .WriteAttributeString("CurrentPlayer", CurrentPlayer.ToString())
                For i As Integer = 0 To UBound(pArgs)
                    XMLWritePlayerArguments(w, i) : Next
                .WriteEndElement()

                .WriteStartElement("Total")
                .WriteAttributeString("Option", TotalOption.ToString())
                .WriteAttributeString("Multiplier", TotalMultiplier.ToString())
                .WriteAttributeString("GlobalMultiplier", TotalGlobalMultiplier.ToString())
                .WriteAttributeString("Decimal", TotalDecimal.ToString())
                .WriteAttributeString("DisplayValue", TotalDisplayValue.ToString())
                .WriteAttributeString("DisplayText", TotalDisplayText.ToString())
                .WriteAttributeString("Autofill", TotalAutofill.ToString())
                .WriteEndElement()

                .WriteStartElement("KeyBindings")
                .WriteAttributeString("Count", UBound(Keybindings).ToString())
                For i As Integer = 0 To UBound(Keybindings)
                    XMLWriteKeybindings(w, i) : Next
                .WriteEndElement()

            End If
            .WriteStartElement("Columns")
            '.WriteAttributeString("Count", col.Length)
            For i As Integer = 0 To UBound(column)
                XMLWriteColumn(w, i) : Next
            .WriteEndElement()

            .WriteStartElement("VisualSettings")
            XMLWriteValue(w, "ColumnTitle", vo.ColumnTitle.Color.ToArgb.ToString())
            XMLWriteFont(w, "ColumnTitleFont", vo.ColumnTitleFont)
            XMLWriteValue(w, "Bg", vo.Bg.Color.ToArgb.ToString())
            XMLWriteValue(w, "Grid", vo.pGrid.Color.ToArgb.ToString())
            XMLWriteValue(w, "Sub", vo.pSub.Color.ToArgb.ToString())
            XMLWriteValue(w, "VLine", vo.pVLine.Color.ToArgb.ToString())
            XMLWriteValue(w, "MLine", vo.pMLine.Color.ToArgb.ToString())
            XMLWriteValue(w, "BGMWav", vo.pBGMWav.Color.ToArgb.ToString())
            XMLWriteValue(w, "SelBox", vo.SelBox.Color.ToArgb.ToString())
            XMLWriteValue(w, "TSCursor", vo.PECursor.Color.ToArgb.ToString())
            XMLWriteValue(w, "TSHalf", vo.PEHalf.Color.ToArgb.ToString())
            XMLWriteValue(w, "TSDeltaMouseOver", vo.PEDeltaMouseOver.ToString())
            XMLWriteValue(w, "TSMouseOver", vo.PEMouseOver.Color.ToArgb.ToString())
            XMLWriteValue(w, "TSSel", vo.PESel.Color.ToArgb.ToString())
            XMLWriteValue(w, "TSBPM", vo.PEBPM.Color.ToArgb.ToString())
            XMLWriteFont(w, "TSBPMFont", vo.PEBPMFont)
            XMLWriteValue(w, "MiddleDeltaRelease", vo.MiddleDeltaRelease.ToString())
            XMLWriteValue(w, "kHeight", vo.kHeight.ToString())
            XMLWriteFont(w, "kFont", vo.kFont)
            XMLWriteFont(w, "kMFont", vo.kMFont)
            XMLWriteValue(w, "kLabelVShift", vo.kLabelVShift.ToString())
            XMLWriteValue(w, "kLabelHShift", vo.kLabelHShift.ToString())
            XMLWriteValue(w, "kLabelHShiftL", vo.kLabelHShiftL.ToString())
            XMLWriteValue(w, "kMouseOver", vo.kMouseOver.Color.ToArgb.ToString())
            XMLWriteValue(w, "kMouseOverE", vo.kMouseOverE.Color.ToArgb.ToString())
            XMLWriteValue(w, "kSelected", vo.kSelected.Color.ToArgb.ToString())
            XMLWriteValue(w, "kOpacity", vo.kOpacity.ToString())
            .WriteEndElement()

            .WriteEndElement()
            .WriteEndDocument()
            .Close()
        End With
    End Sub

    Private Sub WriteColorOverride(ByVal Path As String, ByVal Warning As Boolean)
        Dim F As String = ColorOverridePath(Path)
        If My.Computer.FileSystem.FileExists(F) AndAlso Warning Then
            Dim xDiag = MsgBox(Strings.fopVisualOverride.OverwriteExistingSettings & vbCrLf & F, MsgBoxStyle.YesNo)
            If xDiag = DialogResult.No Then Exit Sub
        End If

        Try
            Dim w As New XmlTextWriter(F, System.Text.Encoding.Unicode)
            With w
                .WriteStartDocument()
                .Formatting = Formatting.Indented
                .Indentation = 4

                .WriteStartElement("ColorOverride")
                .WriteAttributeString("Count", UBound(COverrides).ToString())

                For i = 0 To UBound(COverrides)
                    .WriteStartElement("Color")
                    .WriteAttributeString("Index", i.ToString())
                    .WriteAttributeString("Name", COverrides(i).Name.ToString())
                    .WriteAttributeString("Enabled", COverrides(i).Enabled.ToString())
                    .WriteAttributeString("ColorOption", COverrides(i).ColorOption.ToString())
                    .WriteAttributeString("RangeL", COverrides(i).RangeL.ToString())
                    .WriteAttributeString("RangeU", COverrides(i).RangeU.ToString())
                    .WriteAttributeString("NoteColor", COverrides(i).NoteColor.ToString())
                    .WriteAttributeString("NoteColorU", COverrides(i).NoteColorU.ToString())
                    .WriteEndElement()
                Next

                .WriteEndElement()
                .WriteEndDocument()
                .Close()
            End With
        Catch ex As Exception
            MsgBox(Strings.Messages.SaveWarning & ex.ToString())
        End Try
    End Sub

    Private Function ColorOverridePath(ByVal Path As String) As String
        Select Case COverridesSaveOption
            Case 0
                If Not System.IO.Directory.Exists("Colors") Then My.Computer.FileSystem.CreateDirectory("Colors")
                Return "Colors\Untitled.bmsc.xml"
            Case 1
                If Not System.IO.Directory.Exists("Colors") Then My.Computer.FileSystem.CreateDirectory("Colors")
                Return "Colors\" + GetFileName(Path) + ".xml"
            Case 2
                Return ExcludeFileName(Path) & "\_Colors.xml"
            Case Else
                Return ""
        End Select
    End Function

    Private Sub XMLLoadColumn(ByVal n As XmlElement)
        Dim i As Integer = -1
        XMLLoadAttribute(n.GetAttribute("Index"), i)
        If i < 0 Or i > UBound(column) Then Exit Sub

        With column(i)
            'XMLLoadAttribute(n.GetAttribute("Left"), .Left)
            XMLLoadAttribute(n.GetAttribute("Width"), .Width)
            XMLLoadAttribute(n.GetAttribute("Title"), .Title)
            'XMLLoadAttribute(n.GetAttribute("Text"), .Text)
            Dim Display As Boolean
            Dim attr = n.GetAttribute("Display")
            XMLLoadAttribute(attr, Display)
            .isVisible = CBool(IIf(String.IsNullOrEmpty(attr), .isVisible, Display))

            'XMLLoadAttribute(n.GetAttribute("isNumeric"), .isNumeric)
            'XMLLoadAttribute(n.GetAttribute("Visible"), .Visible)
            'XMLLoadAttribute(n.GetAttribute("Identifier"), .Identifier)
            XMLLoadAttribute(n.GetAttribute("NoteColor"), .cNote)
            .setNoteColor(.cNote)
            XMLLoadAttribute(n.GetAttribute("TextColor"), .cText)
            XMLLoadAttribute(n.GetAttribute("LongNoteColor"), .cLNote)
            .setLNoteColor(.cLNote)
            XMLLoadAttribute(n.GetAttribute("LongTextColor"), .cLText)
            XMLLoadAttribute(n.GetAttribute("BG"), .cBG)
        End With
    End Sub

    Private Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Integer)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
    Private Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Single)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
    Private Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Color)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Private Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Font)
        If n Is Nothing Then Exit Sub

        Dim xName As String = Me.Font.FontFamily.Name
        Dim xSize As Integer = CInt(Me.Font.Size)
        Dim xStyle As Integer = Me.Font.Style
        XMLLoadAttribute(n.GetAttribute("Name"), xName)
        XMLLoadAttribute(n.GetAttribute("Size"), xSize)
        XMLLoadAttribute(n.GetAttribute("Style"), xStyle)
        v = New Font(xName, xSize, CType(xStyle, System.Drawing.FontStyle))
    End Sub

    Private Sub XMLLoadKeybinding(ByVal n As XmlElement)
        For xI = 0 To UBound(Keybindings)
            If n.GetAttribute("OpVar") <> "" Then
                If Keybindings(xI).OpVar = CInt(n.GetAttribute("OpVar")) Then
                    ' XMLLoadAttribute(n.GetAttribute("Name"), Keybindings(xI).OpName)
                    ' XMLLoadAttribute(n.GetAttribute("Description"), Keybindings(xI).Description)
                    Keybindings(xI).Combo = Split(n.GetAttribute("Combos"), ", ")
                    ' XMLLoadAttribute(n.GetAttribute("Category"), Keybindings(xI).Category)

                    RenameShortcut(Keybindings(xI))
                    Exit Sub
                End If
            Else ' Keybinding migration from 3.5.5.10
                Dim OpVar As Integer
                Select Case n.GetAttribute("Name")
                    Case "Move to A2" : OpVar = 0
                    Case "Move to A3" : OpVar = 1
                    Case "Move to A4" : OpVar = 2
                    Case "Move to A5" : OpVar = 3
                    Case "Move to A6" : OpVar = 4
                    Case "Move to A7" : OpVar = 5
                    Case "Move to A8" : OpVar = 6
                    Case "Move to A1" : OpVar = 7
                    Case "Move to D1" : OpVar = 10
                    Case "Move to D2" : OpVar = 11
                    Case "Move to D3" : OpVar = 12
                    Case "Move to D4" : OpVar = 13
                    Case "Move to D5" : OpVar = 14
                    Case "Move to D6" : OpVar = 15
                    Case "Move to D7" : OpVar = 16
                    Case "Move to D8" : OpVar = 17
                    Case "Move to P1" : OpVar = 20
                    Case "Move to P2" : OpVar = 21
                    Case "Move to P3" : OpVar = 22
                    Case "Move to P4" : OpVar = 23
                    Case "Move to P5" : OpVar = 24
                    Case "Move to P6" : OpVar = 25
                    Case "Move to P7" : OpVar = 26
                    Case "Move to P8" : OpVar = 27
                    Case "Move to P9" : OpVar = 28
                    Case "Move to BGM" : OpVar = 100
                    Case "Move to Template Position" : OpVar = 101
                    Case "Disable vertical moves" : OpVar = 102
                    Case "Snap to grid" : OpVar = 103
                    Case "→ Long Note" : OpVar = 104
                    Case "→ Short Note" : OpVar = 105
                    Case "Long Note ↔ Short Note" : OpVar = 106
                    Case "Auto Long Note" : OpVar = 107
                    Case "Auto Long Note" : OpVar = 108
                    Case "Check Technical Error" : OpVar = 109
                    Case "#RANDOM Editor" : OpVar = 110
                    Case "Undo" : OpVar = 111
                    Case "Redo" : OpVar = 112
                    Case "Cut" : OpVar = 113
                    Case "Copy" : OpVar = 114
                    Case "Paste" : OpVar = 115
                    Case "Paste Pattern" : OpVar = 116
                    Case "Select All" : OpVar = 117
                    Case "Select all with hovered note label" : OpVar = 118
                    Case "Move Note Up" : OpVar = 200
                    Case "Move Note Down" : OpVar = 201
                    Case "Move Note Left" : OpVar = 202
                    Case "Move Note Right" : OpVar = 203
                    Case "Insert Space/Define Measure" : OpVar = 204
                    Case "Delete" : OpVar = 205
                    Case "Home" : OpVar = 206
                    Case "End" : OpVar = 207
                    Case "PageUp" : OpVar = 208
                    Case "PageDown" : OpVar = 209
                    Case "Next" : OpVar = 210
                    Case "TabBetweenFiles" : OpVar = 211
                    Case "TabBetweenNotes" : OpVar = 212
                    Case "Decrease Division" : OpVar = 213
                    Case "Increase Division" : OpVar = 214
                    Case "Set CGDivision" : OpVar = 215
                    Case "Decrease CGHeight" : OpVar = 216
                    Case "Increase CGHeight" : OpVar = 217
                    Case "DecreaseCurrentWav" : OpVar = 218
                    Case "IncreaseCurrentWav" : OpVar = 219
                    Case "TBPreviewHighlighted_Click" : OpVar = 220
                End Select
                If Keybindings(xI).OpVar = OpVar Then Keybindings(xI).Combo = Split(n.GetAttribute("Combos"), ", ")
            End If
        Next
    End Sub

    Private Sub XMLLoadOpenedFiles(ByVal n As XmlElement)
        Dim i As Integer
        XMLLoadAttribute(n.GetAttribute("Index"), i)
        XMLLoadAttribute(n.GetAttribute("File"), BMSFiles(i).Filename)
        XMLLoadAttribute(n.GetAttribute("Color"), BMSFiles(i).TabColor)
        BMSFiles(i).TSB = NewBMSTab(BMSFiles(i).Filename)
        ColorTSBChange(BMSFiles(i).TSB, BMSFiles(i).TabColor)
    End Sub

    Private Sub XMLLoadPlayer(ByVal n As XmlElement)
        Dim i As Integer = -1
        XMLLoadAttribute(n.GetAttribute("Index"), i)
        If i < 0 Or i > UBound(pArgs) Then Exit Sub

        XMLLoadAttribute(n.GetAttribute("Path"), pArgs(i).Path)
        XMLLoadAttribute(n.GetAttribute("FromBeginning"), pArgs(i).aBegin)
        XMLLoadAttribute(n.GetAttribute("FromHere"), pArgs(i).aHere)
        XMLLoadAttribute(n.GetAttribute("Stop"), pArgs(i).aStop)
    End Sub

    Private Sub XMLLoadToolbarLocation(ByVal n As XmlElement, ByVal TB As ToolStrip)
        If n Is Nothing Then Exit Sub

        XMLLoadAttribute(n.GetAttribute("X"), TB.Location.X)
        XMLLoadAttribute(n.GetAttribute("Y"), TB.Location.Y)
        Select Case n.GetAttribute("Location")
            Case "Top"
                ToolStripContainer1.TopToolStripPanel.Controls.Add(TB)
            Case "Left"
                ToolStripContainer1.LeftToolStripPanel.Controls.Add(TB)
            Case "Bottom"
                ToolStripContainer1.BottomToolStripPanel.Controls.Add(TB)
            Case "Right"
                ToolStripContainer1.RightToolStripPanel.Controls.Add(TB)
        End Select
    End Sub

    Private Sub LoadSettings(ByVal Path As String)
        If Not My.Computer.FileSystem.FileExists(Path) Then Return

        'Dim xTempFileName As String = ""
        'Do
        'Try
        'xTempFileName = Me.RandomFileName(".xml")
        'File.Copy(Path, xTempFileName)
        'Catch
        'Continue Do
        'End Try
        'Exit Do
        'Loop
        Dim Doc As New XmlDocument
        Dim FileStream As New IO.FileStream(Path, FileMode.Open, FileAccess.Read)
        Doc.Load(FileStream)

        Dim Root As XmlElement = Doc.Item("iBMSC")
        If Root IsNot Nothing Then

            'version
            Dim Major As Integer = My.Application.Info.Version.Major
            Dim Minor As Integer = My.Application.Info.Version.Minor
            Dim Build As Integer = My.Application.Info.Version.Build
            Try
                Dim xMajor As Integer = CInt(Root.Attributes("Major").Value)
                Dim xMinor As Integer = CInt(Root.Attributes("Minor").Value)
                Dim xBuild As Integer = CInt(Root.Attributes("Build").Value)
                Major = xMajor
                Minor = xMinor
                Build = xBuild
            Catch ex As Exception
            End Try

            'form
            Dim eForm As XmlElement = Root.Item("Form")
            If eForm IsNot Nothing Then
                With eForm
                    Select Case .GetAttribute("WindowState")
                        Case FormWindowState.Normal.ToString()
                            Me.WindowState = FormWindowState.Normal
                            XMLLoadAttribute(.GetAttribute("Width"), Me.Width)
                            XMLLoadAttribute(.GetAttribute("Height"), Me.Height)
                            XMLLoadAttribute(.GetAttribute("Top"), Me.Top)
                            XMLLoadAttribute(.GetAttribute("Left"), Me.Left)
                        Case FormWindowState.Maximized.ToString()
                            Me.WindowState = FormWindowState.Maximized
                    End Select

                    XMLLoadAttribute(.GetAttribute(POptionsScroll.Name), POptionsScroll.Width)
                    For Each CB In {POHeaderSwitch, POGridSwitch, POWaveFormSwitch, POWAVSwitch, POBMPSwitch, POBeatSwitch, POExpansionSwitch, POHeaderExpander, POGridExpander, POWaveFormExpander, POWAVExpander, POBMPExpander, POBeatExpander}
                        XMLLoadAttribute(.GetAttribute(CB.Name), CB.Checked)
                    Next
                    For Each P In {POHeaderInner, POGridInner, POWaveFormInner, POWAVInner, POBMPInner, POBeatInner, POExpansionInner}
                        XMLLoadAttribute(.GetAttribute(P.Name), P.Height)
                    Next

                    For Each TB In {TBTab, TBMain}
                        XMLLoadToolbarLocation(.Item(TB.Name), TB)
                    Next
                End With
            End If

            'recent
            Dim eRecent As XmlElement = Root.Item("Recent")
            If eRecent IsNot Nothing Then
                With eRecent
                    XMLLoadAttribute(.GetAttribute("Recent0"), Recent(0)) : SetRecent(0, Recent(0))
                    XMLLoadAttribute(.GetAttribute("Recent1"), Recent(1)) : SetRecent(1, Recent(1))
                    XMLLoadAttribute(.GetAttribute("Recent2"), Recent(2)) : SetRecent(2, Recent(2))
                    XMLLoadAttribute(.GetAttribute("Recent3"), Recent(3)) : SetRecent(3, Recent(3))
                    XMLLoadAttribute(.GetAttribute("Recent4"), Recent(4)) : SetRecent(4, Recent(4))
                End With
            End If

            Dim eOpenedFiles As XmlElement = Root.Item("OpenedFiles")
            If eOpenedFiles IsNot Nothing Then
                With eOpenedFiles
                    Dim iL As Integer
                    XMLLoadAttribute(.GetAttribute("Count"), iL)
                    XMLLoadAttribute(.GetAttribute("BMSFileIndex"), BMSFileIndex)
                    ReDim Preserve BMSFiles(iL)

                    For Each eeFile As XmlElement In .ChildNodes
                        XMLLoadOpenedFiles(eeFile)
                    Next
                End With
            End If

            'edit
            Dim eEdit As XmlElement = Root.Item("Edit")
            If eEdit IsNot Nothing Then
                With eEdit
                    XMLLoadAttribute(.GetAttribute("NTInput"), NTInput)
                    TBNTInput.Checked = NTInput
                    TBNTInput_Click(TBNTInput, New EventArgs)

                    LoadLocale(My.Application.Info.DirectoryPath & "\" & .GetAttribute("Language"))

                    'XMLLoadAttribute(.GetAttribute("SortingMethod"), SortingMethod)

                    XMLLoadAttribute(.GetAttribute("ErrorCheck"), ErrorCheck)
                    TBErrorCheck.Checked = ErrorCheck
                    TBErrorCheck_Click(TBErrorCheck, New System.EventArgs)

                    XMLLoadAttribute(.GetAttribute("ShowFileName"), ShowFileName)
                    TBShowFileName.Checked = ShowFileName
                    TBShowFileName_Click(TBShowFileName, New System.EventArgs)

                    XMLLoadAttribute(.GetAttribute("ShowWaveform"), ShowWaveform)
                    TBShowWaveform.Checked = ShowWaveform
                    TBShowWaveform_Click(TBShowWaveform, New System.EventArgs)

                    XMLLoadAttribute(.GetAttribute("MiddleButtonMoveMethod"), MiddleButtonMoveMethod)
                    XMLLoadAttribute(.GetAttribute("AutoFocusMouseEnter"), AutoFocusMouseEnter)
                    XMLLoadAttribute(.GetAttribute("FirstClickDisabled"), FirstClickDisabled)

                    XMLLoadAttribute(.GetAttribute("AutoSaveInterval"), AutoSaveInterval)
                    If AutoSaveInterval > 0 Then AutoSaveTimer.Interval = AutoSaveInterval Else AutoSaveTimer.Enabled = False

                    XMLLoadAttribute(.GetAttribute("PreviewOnClick"), PreviewOnClick)
                    TBPreviewOnClick.Checked = PreviewOnClick
                    TBPreviewOnClick_Click(TBPreviewOnClick, New System.EventArgs)

                    XMLLoadAttribute(.GetAttribute("ClickStopPreview"), ClickStopPreview)

                    XMLLoadAttribute(.GetAttribute("JackBPM"), ErrorJackBPM)
                    XMLLoadAttribute(.GetAttribute("JackTH"), ErrorJackTH)
                    XMLLoadAttribute(.GetAttribute("gLNGap"), gLNGap)
                    XMLLoadAttribute(.GetAttribute("TempFileName"), TempFileName)
                    XMLLoadAttribute(.GetAttribute("COverridesSaveOption"), COverridesSaveOption)
                    XMLLoadAttribute(.GetAttribute("TemplateSnapToVPosition"), TemplateSnapToVPosition)
                    XMLLoadAttribute(.GetAttribute("PastePatternToVPosition"), PastePatternToVPosition)
                End With
            End If

            'save
            Dim eSave As XmlElement = Root.Item("Save")
            If eSave IsNot Nothing Then
                With eSave
                    Select Case UCase(.GetAttribute("TextEncoding"))
                        Case "SYSTEM ANSI" : TextEncoding = System.Text.Encoding.Default
                        Case "LITTLE ENDIAN UTF16" : TextEncoding = System.Text.Encoding.Unicode
                        Case "ASCII" : TextEncoding = System.Text.Encoding.ASCII
                        Case "BIG ENDIAN UTF16" : TextEncoding = System.Text.Encoding.BigEndianUnicode
                        Case "LITTLE ENDIAN UTF32" : TextEncoding = System.Text.Encoding.UTF32
                        Case "UTF7" : TextEncoding = System.Text.Encoding.UTF7
                        Case "UTF8" : TextEncoding = System.Text.Encoding.UTF8
                        Case "SJIS" : TextEncoding = System.Text.Encoding.GetEncoding(932)
                        Case "EUC-KR" : TextEncoding = System.Text.Encoding.GetEncoding(51949)
                            ' leave current encoding
                            ' Case Else 
                    End Select

                    XMLLoadAttribute(.GetAttribute("BMSGridLimit"), BMSGridLimit)
                    XMLLoadAttribute(.GetAttribute("BeepWhileSaved"), BeepWhileSaved)
                    XMLLoadAttribute(.GetAttribute("PreloadBMSStruct"), PreloadBMSStruct)
                    XMLLoadAttribute(.GetAttribute("BPMx1296"), BPMx1296)
                    XMLLoadAttribute(.GetAttribute("STOPx1296"), STOPx1296)
                    XMLLoadAttribute(.GetAttribute("AudioLine"), AudioLine)
                End With
            End If

            'WAV
            Dim eWAV As XmlElement = Root.Item("WAV")
            If eWAV IsNot Nothing Then
                With eWAV
                    XMLLoadAttribute(.GetAttribute("WAVMultiSelect"), WAVMultiSelect)
                    CWAVMultiSelect.Checked = WAVMultiSelect
                    CWAVMultiSelect_CheckedChanged(CWAVMultiSelect, New EventArgs)

                    XMLLoadAttribute(.GetAttribute("WAVChangeLabel"), WAVChangeLabel)
                    CWAVChangeLabel.Checked = WAVChangeLabel
                    CWAVChangeLabel_CheckedChanged(CWAVChangeLabel, New EventArgs)

                    Dim xInt As Integer = CInt(.GetAttribute("BeatChangeMode"))
                    Dim xBeatOpList As RadioButton() = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
                    If xInt >= 0 And xInt < xBeatOpList.Length Then
                        xBeatOpList(xInt).Checked = True
                        CBeatPreserve_Click(xBeatOpList(xInt), New System.EventArgs)
                    End If
                End With
            End If

            'BMP
            Dim eBMP As XmlElement = Root.Item("BMP")
            If eBMP IsNot Nothing Then
                With eBMP
                    XMLLoadAttribute(.GetAttribute("BMPMultiSelect"), BMPMultiSelect)
                    CBMPMultiSelect.Checked = BMPMultiSelect
                    CBMPMultiSelect_CheckedChanged(CBMPMultiSelect, New EventArgs)

                    XMLLoadAttribute(.GetAttribute("BMPChangeLabel"), BMPChangeLabel)
                    CBMPChangeLabel.Checked = BMPChangeLabel
                    CBMPChangeLabel_CheckedChanged(CBMPChangeLabel, New EventArgs)
                End With
            End If

            'ShowHide
            Dim eShowHide As XmlElement = Root.Item("ShowHide")
            If eShowHide IsNot Nothing Then
                With eShowHide
                    XMLLoadAttribute(.GetAttribute("showMenu"), mnSMenu.Checked)
                    XMLLoadAttribute(.GetAttribute("showTB"), mnSTB.Checked)
                    XMLLoadAttribute(.GetAttribute("showOpPanel"), mnSOP.Checked)
                    XMLLoadAttribute(.GetAttribute("showStatus"), mnSStatus.Checked)
                    XMLLoadAttribute(.GetAttribute("showLSplit"), mnSLSplitter.Checked)
                    XMLLoadAttribute(.GetAttribute("showRSplit"), mnSRSplitter.Checked)
                End With
            End If

            'Grid
            Dim eGrid As XmlElement = Root.Item("Grid")
            If eGrid IsNot Nothing Then
                With eGrid
                    XMLLoadAttribute(.GetAttribute("gSnap"), CGSnap.Checked)
                    XMLLoadAttribute(.GetAttribute("xmlDisableVertical"), CGDisableVertical.Checked)
                    XMLLoadAttribute(.GetAttribute("gWheel"), gWheel)
                    XMLLoadAttribute(.GetAttribute("gPgUpDn"), gPgUpDn)
                    XMLLoadAttribute(.GetAttribute("gShow"), CGShow.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowS"), CGShowS.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowBG"), CGShowBG.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowM"), CGShowM.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowV"), CGShowV.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowMB"), CGShowMB.Checked)
                    XMLLoadAttribute(.GetAttribute("gShowC"), CGShowC.Checked)
                    XMLLoadAttribute(.GetAttribute("gBPM"), CGBPM.Checked)
                    XMLLoadAttribute(.GetAttribute("gSTOP"), CGSTOP.Checked)
                    XMLLoadAttribute(.GetAttribute("gSCROLL"), CGSCROLL.Checked)
                    XMLLoadAttribute(.GetAttribute("gBLP"), CGBLP.Checked)
                    XMLLoadAttribute(.GetAttribute("gP2"), CHPlayer.SelectedIndex)
                    XMLLoadAttribute(.GetAttribute("gCol"), CGB.Value)
                    XMLLoadAttribute(.GetAttribute("gxHeight"), CGHeight.Value)
                    XMLLoadAttribute(.GetAttribute("gxWidth"), CGWidth.Value)
                    XMLLoadAttribute(.GetAttribute("gSlash"), gSlash)

                    Dim xgDivide As Integer = CInt(.GetAttribute("gDivide"))
                    If xgDivide >= CGDivide.Minimum And xgDivide <= CGDivide.Maximum Then CGDivide.Value = xgDivide

                    Dim xgSub As Integer = CInt(.GetAttribute("gSub"))
                    If xgSub >= CGSub.Minimum And xgSub <= CGSub.Maximum Then CGSub.Value = xgSub
                End With
            End If

            'WaveForm
            Dim eWaveForm As XmlElement = Root.Item("WaveForm")
            If eWaveForm IsNot Nothing Then
                With eWaveForm
                    XMLLoadAttribute(.GetAttribute("wLock"), BWLock.Checked)
                    XMLLoadAttribute(.GetAttribute("wPosition"), TWPosition.Value)
                    XMLLoadAttribute(.GetAttribute("wLeft"), TWLeft.Value)
                    XMLLoadAttribute(.GetAttribute("wWidth"), TWWidth.Value)
                    XMLLoadAttribute(.GetAttribute("wPrecision"), TWPrecision.Value)
                End With
            End If

            'Player
            Dim ePlayer As XmlElement = Root.Item("Player")
            If ePlayer IsNot Nothing Then
                With ePlayer
                    XMLLoadAttribute(.GetAttribute("CurrentPlayer"), CurrentPlayer)

                    Dim xCount As Integer = CInt(.GetAttribute("Count"))
                    If xCount > 0 Then ReDim Preserve pArgs(xCount - 1)
                End With

                For Each eePlayer As XmlElement In ePlayer.ChildNodes
                    Me.XMLLoadPlayer(eePlayer)
                Next
            End If

            Dim eTotal As XmlElement = Root.Item("Total")
            If eTotal IsNot Nothing Then
                With eTotal
                    XMLLoadAttribute(.GetAttribute("Option"), TotalOption)
                    XMLLoadAttribute(.GetAttribute("Multiplier"), TotalMultiplier)
                    XMLLoadAttribute(.GetAttribute("GlobalMultiplier"), TotalGlobalMultiplier)
                    XMLLoadAttribute(.GetAttribute("Decimal"), TotalDecimal)
                    XMLLoadAttribute(.GetAttribute("DisplayValue"), TotalDisplayValue)
                    XMLLoadAttribute(.GetAttribute("DisplayText"), TotalDisplayText)
                    XMLLoadAttribute(.GetAttribute("Autofill"), TotalAutofill)
                End With
            End If

            Dim eKeybindings As XmlElement = Root.Item("KeyBindings")
            If eKeybindings IsNot Nothing Then
                With eKeybindings
                    Dim xBound = CInt(.GetAttribute("Count"))
                    If xBound > UBound(Keybindings) Then ReDim Preserve Keybindings(xBound)

                    For Each eeKeybindings As XmlElement In .ChildNodes
                        XMLLoadKeybinding(eeKeybindings)
                    Next
                End With
            End If

            'Columns
            Dim eColumns As XmlElement = Root.Item("Columns")
            If eColumns IsNot Nothing Then
                For Each eeCol As XmlElement In eColumns.ChildNodes
                    Me.XMLLoadColumn(eeCol)
                Next
            End If

            'VisualSettings
            Dim eVisualSettings As XmlElement = Root.Item("VisualSettings")
            If eVisualSettings IsNot Nothing Then
                With eVisualSettings
                    XMLLoadElementValue(.Item("ColumnTitle"), vo.ColumnTitle.Color)
                    XMLLoadElementValue(.Item("ColumnTitleFont"), vo.ColumnTitleFont)
                    XMLLoadElementValue(.Item("Bg"), vo.Bg.Color)
                    XMLLoadElementValue(.Item("Grid"), vo.pGrid.Color)
                    XMLLoadElementValue(.Item("Sub"), vo.pSub.Color)
                    XMLLoadElementValue(.Item("VLine"), vo.pVLine.Color)
                    XMLLoadElementValue(.Item("MLine"), vo.pMLine.Color)

                    XMLLoadElementValue(.Item("BGMWav"), vo.pBGMWav.Color)
                    TWTransparency.Value = vo.pBGMWav.Color.A
                    TWTransparency2.Value = vo.pBGMWav.Color.A
                    TWSaturation.Value = CDec(vo.pBGMWav.Color.GetSaturation * 1000)
                    TWSaturation2.Value = CInt(vo.pBGMWav.Color.GetSaturation * 1000)

                    XMLLoadElementValue(.Item("SelBox"), vo.SelBox.Color)
                    XMLLoadElementValue(.Item("TSCursor"), vo.PECursor.Color)
                    XMLLoadElementValue(.Item("TSHalf"), vo.PEHalf.Color)
                    XMLLoadElementValue(.Item("TSDeltaMouseOver"), vo.PEDeltaMouseOver)
                    XMLLoadElementValue(.Item("TSMouseOver"), vo.PEMouseOver.Color)
                    XMLLoadElementValue(.Item("TSSel"), vo.PESel.Color)
                    XMLLoadElementValue(.Item("TSBPM"), vo.PEBPM.Color)
                    XMLLoadElementValue(.Item("TSBPMFont"), vo.PEBPMFont)
                    XMLLoadElementValue(.Item("MiddleDeltaRelease"), vo.MiddleDeltaRelease)
                    XMLLoadElementValue(.Item("kHeight"), vo.kHeight)
                    XMLLoadElementValue(.Item("kFont"), vo.kFont)
                    XMLLoadElementValue(.Item("kMFont"), vo.kMFont)
                    XMLLoadElementValue(.Item("kLabelVShift"), vo.kLabelVShift)
                    XMLLoadElementValue(.Item("kLabelHShift"), vo.kLabelHShift)
                    XMLLoadElementValue(.Item("kLabelHShiftL"), vo.kLabelHShiftL)
                    XMLLoadElementValue(.Item("kMouseOver"), vo.kMouseOver.Color)
                    XMLLoadElementValue(.Item("kMouseOverE"), vo.kMouseOverE.Color)
                    XMLLoadElementValue(.Item("kSelected"), vo.kSelected.Color)
                    XMLLoadElementValue(.Item("kOpacity"), vo.kOpacity)
                End With
            End If
        End If

        UpdateColumnsX()
        FileStream.Close()
        'File.Delete(xTempFileName)
    End Sub

    Private Sub XMLLoadLocaleMenu(ByVal n As XmlElement, ByRef target As String)
        If n Is Nothing Then Exit Sub
        If n.HasAttribute("amp") Then
            target = n.InnerText.Insert(Integer.Parse(n.GetAttribute("amp")), "&")
        Else
            target = n.InnerText
        End If
    End Sub

    Private Sub XMLLoadLocale(ByVal n As XmlElement, ByRef target As String)
        If n IsNot Nothing Then target = n.InnerText
    End Sub

    Private Sub XMLLoadLocaleToolTipUniversal(ByVal n As XmlElement, ByVal target As Control)
        If n Is Nothing Then Exit Sub
        ToolTipUniversal.SetToolTip(target, n.InnerText)
    End Sub

    Private Sub LoadLocale(ByVal Path As String)
        If Not My.Computer.FileSystem.FileExists(Path) Then Return

        Dim Doc As XmlDocument = Nothing
        Dim FileStream As IO.FileStream = Nothing

        Dim xPOHeaderPart2 As Boolean = POHeaderPart2.Visible
        Dim xPOGridPart2 As Boolean = POGridPart2.Visible
        Dim xPOWaveFormPart2 As Boolean = POWaveFormPart2.Visible
        POHeaderPart2.Visible = True
        POGridPart2.Visible = True
        POWaveFormPart2.Visible = True

        Try
            Doc = New XmlDocument
            FileStream = New IO.FileStream(Path, FileMode.Open, FileAccess.Read)
            Doc.Load(FileStream)

            Dim Root As XmlElement = Doc.Item("iBMSC.Locale")
            If Root Is Nothing Then Throw New NullReferenceException

            XMLLoadLocale(Root.Item("OK"), Strings.OK)
            XMLLoadLocale(Root.Item("Cancel"), Strings.Cancel)
            XMLLoadLocale(Root.Item("None"), Strings.None)
            XMLLoadLocale(Root.Item("Beat"), Strings.Beat)
            XMLLoadLocale(Root.Item("Beats"), Strings.Beats)
            XMLLoadLocale(Root.Item("st"), Strings.st)
            XMLLoadLocale(Root.Item("nd"), Strings.nd)
            XMLLoadLocale(Root.Item("rd"), Strings.rd)
            XMLLoadLocale(Root.Item("th"), Strings.th)

            Dim eFont As XmlElement = Root.Item("Font")
            If eFont IsNot Nothing Then
                Dim xSize As Integer = 9
                If eFont.HasAttribute("Size") Then xSize = CInt(eFont.GetAttribute("Size"))

                Dim fRegular As New Font(Me.Font.FontFamily, xSize, FontStyle.Regular)
                Dim xChildNode As XmlNode = eFont.LastChild
                Do While xChildNode IsNot Nothing
                    If xChildNode.LocalName <> "Family" Then Continue Do
                    If isFontInstalled(xChildNode.InnerText) Then
                        fRegular = New Font(xChildNode.InnerText, xSize)
                    End If
                    xChildNode = xChildNode.PreviousSibling
                Loop

                Font = fRegular
                Dim TMIList() As ToolStripMenuItem = {mnSys}
                Dim CMSList() As ContextMenuStrip = {Menu1, cmnLanguage, cmnTheme, cmnConversion}
                Dim MSList() As MenuStrip = {mnMain}
                Dim TSList() As ToolStrip = {TBMain}
                Dim SSList() As StatusStrip = {FStatus, FStatus2}
                ' Not sure if a function would work using object type as a variable
                For Each c In TMIList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next
                For Each c In CMSList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next
                For Each c In MSList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next
                For Each c In TSList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next
                For Each c In SSList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next

                Dim fBold As New Font(fRegular, FontStyle.Bold)

                Dim TSBList() As ToolStripButton = {TBStatistics, FSSS, FSSL, FSSH}
                Dim TSTBList() As ToolStripTextBox = {TVCM, TVCD, TVCBPM}
                Dim TSSLList() As ToolStripStatusLabel = {FSP1, FSP3, FSP2}
                Dim PList() As Panel = {PMain, PMainIn, PMainR, PMainInR, PMainL, PMainInL}
                For Each c In TSBList
                    Try
                        c.Font = fBold
                    Catch ex As Exception
                    End Try
                Next
                For Each c In TSTBList
                    Try
                        c.Font = fBold
                    Catch ex As Exception
                    End Try
                Next
                For Each c In TSSLList
                    Try
                        c.Font = fBold
                    Catch ex As Exception
                    End Try
                Next
            End If

            Dim eMonoFont As XmlElement = Root.Item("MonoFont")
            If eMonoFont IsNot Nothing Then
                Dim xSize As Integer = 9
                If eMonoFont.HasAttribute("Size") Then xSize = CInt(eMonoFont.GetAttribute("Size"))

                Dim fMono As New Font(POWAVInner.Font.FontFamily, xSize)
                Dim xChildNode As XmlNode = eMonoFont.LastChild
                Do While xChildNode IsNot Nothing
                    If xChildNode.LocalName <> "Family" Then Continue Do
                    If isFontInstalled(xChildNode.InnerText) Then
                        fMono = New Font(xChildNode.InnerText, xSize)
                    End If
                    xChildNode = xChildNode.PreviousSibling
                Loop

                Dim LBList() As ListBox = {LWAV, LBeat}
                Dim TBList() As TextBox = {TExpansion}
                For Each c In LBList
                    Try
                        c.Font = fMono
                    Catch ex As Exception
                    End Try
                Next
                For Each c In TBList
                    Try
                        c.Font = fMono
                    Catch ex As Exception
                    End Try
                Next
            End If

            Dim eMenu As XmlElement = Root.Item("Menu")
            If eMenu IsNot Nothing Then

                Dim eFile As XmlElement = eMenu.Item("File")
                If eFile IsNot Nothing Then
                    XMLLoadLocaleMenu(eFile.Item("Title"), mnFile.Text)
                    XMLLoadLocaleMenu(eFile.Item("New"), mnNew.Text)
                    XMLLoadLocaleMenu(eFile.Item("Open"), mnOpen.Text)
                    XMLLoadLocaleMenu(eFile.Item("OpenTemplate"), mnOpenTemplate.Text)
                    XMLLoadLocaleMenu(eFile.Item("ImportSM"), mnImportSM.Text)
                    XMLLoadLocaleMenu(eFile.Item("ImportIBMSC"), mnImportIBMSC.Text)
                    XMLLoadLocaleMenu(eFile.Item("Save"), mnSave.Text)
                    XMLLoadLocaleMenu(eFile.Item("SaveAs"), mnSaveAs.Text)
                    XMLLoadLocaleMenu(eFile.Item("ExportIBMSC"), mnExport.Text)
                    XMLLoadLocaleMenu(eFile.Item("Close"), mnClose.Text)
                    If Recent(0) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent0"), mnOpenR0.Text)
                    If Recent(1) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent1"), mnOpenR1.Text)
                    If Recent(2) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent2"), mnOpenR2.Text)
                    If Recent(3) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent3"), mnOpenR3.Text)
                    If Recent(4) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent4"), mnOpenR4.Text)
                    XMLLoadLocaleMenu(eFile.Item("Quit"), mnQuit.Text)
                End If

                Dim eEdit As XmlElement = eMenu.Item("Edit")
                If eEdit IsNot Nothing Then
                    XMLLoadLocaleMenu(eEdit.Item("Title"), mnEdit.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Undo"), mnUndo.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Redo"), mnRedo.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Cut"), mnCut.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Copy"), mnCopy.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Paste"), mnPaste.Text)
                    XMLLoadLocaleMenu(eEdit.Item("PastePattern"), mnPastePattern.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Delete"), mnDelete.Text)
                    XMLLoadLocaleMenu(eEdit.Item("SelectAll"), mnSelectAll.Text)
                    XMLLoadLocaleMenu(eEdit.Item("GoToMeasure"), mnGotoMeasure.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Find"), mnFind.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Stat"), mnStatistics.Text)
                    XMLLoadLocaleMenu(eEdit.Item("StatAdv"), mnStatisticsAdvanced.Text)
                    XMLLoadLocaleMenu(eEdit.Item("TimeSelectionTool"), mnTimeSelect.Text)
                    XMLLoadLocaleMenu(eEdit.Item("SelectTool"), mnSelect.Text)
                    XMLLoadLocaleMenu(eEdit.Item("WriteTool"), mnWrite.Text)
                    XMLLoadLocaleMenu(eEdit.Item("MyO2"), mnMyO2.Text)
                End If

                Dim eView As XmlElement = eMenu.Item("View")
                If eView IsNot Nothing Then
                    XMLLoadLocaleMenu(eView.Item("Title"), mnSys.Text)
                End If

                Dim eOptions As XmlElement = eMenu.Item("Options")
                If eOptions IsNot Nothing Then
                    XMLLoadLocaleMenu(eOptions.Item("Title"), mnOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("NT"), mnNTInput.Text)
                    XMLLoadLocaleMenu(eOptions.Item("ErrorCheck"), mnErrorCheck.Text)
                    XMLLoadLocaleMenu(eOptions.Item("PreviewOnClick"), mnPreviewOnClick.Text)
                    XMLLoadLocaleMenu(eOptions.Item("ShowFileName"), mnShowFileName.Text)
                    XMLLoadLocaleMenu(eOptions.Item("ShowWaveform"), mnShowWaveform.Text)
                    XMLLoadLocaleMenu(eOptions.Item("KeybindingOptions"), mnKOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("GeneralOptions"), mnGOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("VisualOptions"), mnVOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("VisualOverrideOptions"), mnVCOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("PlayerOptions"), mnPOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("Language"), mnLanguage.Text)
                    XMLLoadLocaleMenu(eOptions.Item("Theme"), mnTheme.Text)
                End If

                XMLLoadLocaleMenu(eMenu.Item("Conversion"), mnConversion.Text)

                Dim ePreview As XmlElement = eMenu.Item("Preview")
                If ePreview IsNot Nothing Then
                    XMLLoadLocaleMenu(ePreview.Item("Title"), mnPreview.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayBegin"), mnPlayB.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayHere"), mnPlay.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayStop"), mnStop.Text)
                End If

                Dim eTools As XmlElement = eMenu.Item("Tools")
                If eTools IsNot Nothing Then
                    XMLLoadLocaleMenu(eTools.Item("Title"), mnTools.Text)
                    XMLLoadLocaleMenu(eTools.Item("TechnicalErrorCheck"), mnTechnicalErrorCheck.Text)
                    XMLLoadLocaleMenu(eTools.Item("TOTAL"), mnTOTAL.Text)
                    XMLLoadLocaleMenu(eTools.Item("MyO2"), mnMyO2.Text)
                End If

                ' Dim eAbout As XmlElement = eMenu.Item("About")
                ' If eAbout IsNot Nothing Then
                '     XMLLoadLocaleMenu(eAbout.Item("Title"), mnAbout.Text)
                '     XMLLoadLocaleMenu(eAbout.Item("About"), mnAbout1.Text)
                '     XMLLoadLocaleMenu(eAbout.Item("CheckUpdates"), mnUpdate.Text)
                '     XMLLoadLocaleMenu(eAbout.Item("CheckUpdatesC"), mnUpdateC.Text)
                ' End If
            End If

            Dim eToolBar As XmlElement = Root.Item("ToolBar")
            If eToolBar IsNot Nothing Then
                XMLLoadLocale(eToolBar.Item("New"), TBNew.Text)
                XMLLoadLocale(eToolBar.Item("Open"), TBOpen.Text)
                XMLLoadLocale(eToolBar.Item("Save"), TBSave.Text)
                XMLLoadLocale(eToolBar.Item("Cut"), TBCut.Text)
                XMLLoadLocale(eToolBar.Item("Copy"), TBCopy.Text)
                XMLLoadLocale(eToolBar.Item("Paste"), TBPaste.Text)
                XMLLoadLocale(eToolBar.Item("PastePattern"), TBPastePattern.Text)
                XMLLoadLocale(eToolBar.Item("Find"), TBFind.Text)
                XMLLoadLocale(eToolBar.Item("Stat"), TBStatistics.ToolTipText)
                XMLLoadLocale(eToolBar.Item("Conversion"), POConvert.Text)
                XMLLoadLocale(eToolBar.Item("MyO2"), TBMyO2.Text)
                XMLLoadLocale(eToolBar.Item("ErrorCheck"), TBErrorCheck.Text)
                XMLLoadLocale(eToolBar.Item("PreviewOnClick"), TBPreviewOnClick.Text)
                XMLLoadLocale(eToolBar.Item("ShowFileName"), TBShowFileName.Text)
                XMLLoadLocale(eToolBar.Item("ShowWaveform"), TBShowWaveform.Text)
                XMLLoadLocale(eToolBar.Item("Undo"), TBUndo.Text)
                XMLLoadLocale(eToolBar.Item("Redo"), TBRedo.Text)
                XMLLoadLocale(eToolBar.Item("NT"), TBNTInput.Text)
                XMLLoadLocale(eToolBar.Item("TimeSelectionTool"), TBTimeSelect.Text)
                XMLLoadLocale(eToolBar.Item("SelectTool"), TBSelect.Text)
                XMLLoadLocale(eToolBar.Item("WriteTool"), TBWrite.Text)
                XMLLoadLocale(eToolBar.Item("PlayBegin"), TBPlayB.Text)
                XMLLoadLocale(eToolBar.Item("PlayHere"), TBPlay.Text)
                XMLLoadLocale(eToolBar.Item("PlayStop"), TBStop.Text)
                XMLLoadLocale(eToolBar.Item("PlayerOptions"), TBPOptions.Text)
                XMLLoadLocale(eToolBar.Item("VisualOptions"), TBVOptions.Text)
                XMLLoadLocale(eToolBar.Item("GeneralOptions"), TBGOptions.Text)
                XMLLoadLocale(eToolBar.Item("Language"), TBLanguage.Text)
                XMLLoadLocale(eToolBar.Item("Theme"), TBTheme.Text)
                XMLLoadLocale(eToolBar.Item("RecommendedTOTAL"), Strings.StatusBar.RecommendedTotal)
                ' XMLLoadLocale(eToolBar.Item("About"), TBAbout.Text)
            End If

            Dim eStatusBar As XmlElement = Root.Item("StatusBar")
            If eStatusBar IsNot Nothing Then
                XMLLoadLocale(eStatusBar.Item("ColumnCaption"), FSC.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("NoteIndex"), FSW.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("MeasureIndex"), FSM.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("GridResolution"), FSP1.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("ReducedResolution"), FSP3.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("MeasureResolution"), FSP2.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("AbsolutePosition"), FSP4.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("Length"), Strings.StatusBar.Length)
                XMLLoadLocale(eStatusBar.Item("Note"), Strings.StatusBar.Note)
                XMLLoadLocale(eStatusBar.Item("LongNote"), Strings.StatusBar.LongNote)
                XMLLoadLocale(eStatusBar.Item("Hidden"), Strings.StatusBar.Hidden)
                XMLLoadLocale(eStatusBar.Item("Landmine"), Strings.StatusBar.Landmine)
                XMLLoadLocale(eStatusBar.Item("Comment"), Strings.StatusBar.Comment)
                XMLLoadLocale(eStatusBar.Item("Approximate"), Strings.StatusBar.Approximate)
                XMLLoadLocale(eStatusBar.Item("Error"), Strings.StatusBar.Err)
                XMLLoadLocale(eStatusBar.Item("ErrorTechnical"), Strings.StatusBar.ErrTechnical)
                XMLLoadLocale(eStatusBar.Item("SelStart"), FSSS.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("SelLength"), FSSL.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("SelSplit"), FSSH.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("Reverse"), BVCReverse.Text)
                XMLLoadLocale(eStatusBar.Item("ByMultiple"), BVCApply.Text)
                XMLLoadLocale(eStatusBar.Item("ByValue"), BVCCalculate.Text)
            End If

            Dim eSubMenu As XmlElement = Root.Item("SubMenu")
            If eSubMenu IsNot Nothing Then

                Dim eShowHide As XmlElement = eSubMenu.Item("ShowHide")
                If eShowHide IsNot Nothing Then
                    'Dim xToolTip As String = ToolTipUniversal.GetToolTip(ttlIcon)
                    'XMLLoadLocaleMenu(eShowHide.Item("ToolTip"), xToolTip)
                    'ToolTipUniversal.SetToolTip(ttlIcon, xToolTip)

                    XMLLoadLocaleMenu(eShowHide.Item("Menu"), mnSMenu.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("ToolBar"), mnSTB.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("OptionsPanel"), mnSOP.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("StatusBar"), mnSStatus.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("LSplit"), mnSLSplitter.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("RSplit"), mnSRSplitter.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Grid"), CGShow.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Sub"), CGShowS.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BG"), CGShowBG.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("MeasureIndex"), CGShowM.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("MeasureLine"), CGShowMB.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Vertical"), CGShowV.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("ColumnCaption"), CGShowC.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BPM"), CGBPM.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("STOP"), CGSTOP.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("SCROLL"), CGSCROLL.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BLP"), CGBLP.Text)
                End If

                Dim eInsertMeasure As XmlElement = eSubMenu.Item("InsertMeasure")
                If eInsertMeasure IsNot Nothing Then
                    XMLLoadLocaleMenu(eInsertMeasure.Item("Insert"), MInsert.Text)
                    XMLLoadLocaleMenu(eInsertMeasure.Item("Remove"), MRemove.Text)
                End If

                Dim eLanguage As XmlElement = eSubMenu.Item("Language")
                If eLanguage IsNot Nothing Then
                    XMLLoadLocaleMenu(eLanguage.Item("Default"), TBLangDef.Text)
                    XMLLoadLocaleMenu(eLanguage.Item("Refresh"), TBLangRefresh.Text)
                End If

                Dim eTheme As XmlElement = eSubMenu.Item("Theme")
                If eTheme IsNot Nothing Then
                    XMLLoadLocaleMenu(eTheme.Item("Default"), TBThemeDef.Text)
                    XMLLoadLocaleMenu(eTheme.Item("Save"), TBThemeSave.Text)
                    XMLLoadLocaleMenu(eTheme.Item("Refresh"), TBThemeRefresh.Text)
                    XMLLoadLocaleMenu(eTheme.Item("LoadComptability"), TBThemeLoadComptability.Text)
                End If

                Dim eConvert As XmlElement = eSubMenu.Item("Convert")
                If eConvert IsNot Nothing Then
                    XMLLoadLocaleMenu(eConvert.Item("Long"), POBLong.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Short"), POBShort.Text)
                    XMLLoadLocaleMenu(eConvert.Item("LongShort"), POBLongShort.Text)
                    XMLLoadLocaleMenu(eConvert.Item("LongObjNT"), POBLongObjNT.Text)
                    XMLLoadLocaleMenu(eConvert.Item("LongNTObj"), POBLongNTObj.Text)
                    XMLLoadLocaleMenu(eConvert.Item("AutoLongV"), POBAutoLongVPosition.Text)
                    XMLLoadLocaleMenu(eConvert.Item("AutoLongC"), POBAutoLongColumn.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Hidden"), POBHidden.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Visible"), POBVisible.Text)
                    XMLLoadLocaleMenu(eConvert.Item("HiddenVisible"), POBHiddenVisible.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Relabel"), POBModify.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Mirror"), POBMirror.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Flip"), POBFlip.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Random"), POBRandom.Text)
                    XMLLoadLocaleMenu(eConvert.Item("R-Random"), POBRRandom.Text)
                    XMLLoadLocaleMenu(eConvert.Item("S-Random"), POBSRandom.Text)
                    XMLLoadLocaleMenu(eConvert.Item("H-Random"), POBHRandom.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Sort"), POBSort.Text)
                End If

                Dim eWAV As XmlElement = eSubMenu.Item("WAV")
                If eWAV IsNot Nothing Then
                    XMLLoadLocaleMenu(eWAV.Item("MultiSelection"), CWAVMultiSelect.Text)
                    XMLLoadLocaleMenu(eWAV.Item("MultiSelection"), CBMPMultiSelect.Text)
                    XMLLoadLocaleMenu(eWAV.Item("Synchronize"), CWAVChangeLabel.Text)
                    XMLLoadLocaleMenu(eWAV.Item("Synchronize"), CBMPChangeLabel.Text)
                End If

                Dim eBeat As XmlElement = eSubMenu.Item("Beat")
                If eBeat IsNot Nothing Then
                    XMLLoadLocaleMenu(eBeat.Item("Absolute"), CBeatPreserve.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Measure"), CBeatMeasure.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Cut"), CBeatCut.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Scale"), CBeatScale.Text)
                End If
            End If

            Dim eOptionsPanel As XmlElement = Root.Item("OptionsPanel")
            If eOptionsPanel IsNot Nothing Then

                Dim eHeader As XmlElement = eOptionsPanel.Item("Header")
                If eHeader IsNot Nothing Then
                    XMLLoadLocale(eHeader.Item("Header"), POHeaderSwitch.Text)
                    XMLLoadLocale(eHeader.Item("Title"), Label3.Text)
                    XMLLoadLocale(eHeader.Item("Artist"), Label4.Text)
                    XMLLoadLocale(eHeader.Item("Genre"), Label2.Text)
                    XMLLoadLocale(eHeader.Item("BPM"), Label9.Text)
                    XMLLoadLocale(eHeader.Item("Player"), Label8.Text)
                    XMLLoadLocale(eHeader.Item("Rank"), Label10.Text)
                    XMLLoadLocale(eHeader.Item("PlayLevel"), Label6.Text)
                    XMLLoadLocale(eHeader.Item("SubTitle"), Label15.Text)
                    XMLLoadLocale(eHeader.Item("SubArtist"), Label17.Text)
                    'XMLLoadLocale(eHeader.Item("Maker"), Label14.Text)
                    XMLLoadLocale(eHeader.Item("StageFile"), Label16.Text)
                    XMLLoadLocale(eHeader.Item("Banner"), Label12.Text)
                    XMLLoadLocale(eHeader.Item("BackBMP"), Label11.Text)
                    'XMLLoadLocale(eHeader.Item("MidiFile"), Label18.Text)
                    XMLLoadLocale(eHeader.Item("Difficulty"), Label21.Text)
                    XMLLoadLocale(eHeader.Item("ExRank"), Label23.Text)
                    XMLLoadLocale(eHeader.Item("Total"), Label20.Text)
                    'XMLLoadLocale(eHeader.Item("VolWav"), Label22.Text)
                    XMLLoadLocale(eHeader.Item("Comment"), Label19.Text)
                    'XMLLoadLocale(eHeader.Item("LnType"), Label13.Text)
                    XMLLoadLocale(eHeader.Item("LnObj"), Label24.Text)

                    RemoveHandler CHPlayer.SelectedIndexChanged, AddressOf CHPlayer_SelectedIndexChanged
                    XMLLoadLocale(eHeader.Item("Player1"), CHPlayer.Items.Item(0).ToString())
                    XMLLoadLocale(eHeader.Item("Player2"), CHPlayer.Items.Item(1).ToString())
                    XMLLoadLocale(eHeader.Item("Player3"), CHPlayer.Items.Item(2).ToString())
                    AddHandler CHPlayer.SelectedIndexChanged, AddressOf CHPlayer_SelectedIndexChanged

                    RemoveHandler CHRank.SelectedIndexChanged, AddressOf THGenre_TextChanged
                    XMLLoadLocale(eHeader.Item("Rank0"), CHRank.Items.Item(0).ToString())
                    XMLLoadLocale(eHeader.Item("Rank1"), CHRank.Items.Item(1).ToString())
                    XMLLoadLocale(eHeader.Item("Rank2"), CHRank.Items.Item(2).ToString())
                    XMLLoadLocale(eHeader.Item("Rank3"), CHRank.Items.Item(3).ToString())
                    XMLLoadLocale(eHeader.Item("Rank4"), CHRank.Items.Item(4).ToString())
                    AddHandler CHRank.SelectedIndexChanged, AddressOf THGenre_TextChanged

                    RemoveHandler CHDifficulty.SelectedIndexChanged, AddressOf THGenre_TextChanged
                    XMLLoadLocale(eHeader.Item("Difficulty0"), CHDifficulty.Items.Item(0).ToString())
                    XMLLoadLocale(eHeader.Item("Difficulty1"), CHDifficulty.Items.Item(1).ToString())
                    XMLLoadLocale(eHeader.Item("Difficulty2"), CHDifficulty.Items.Item(2).ToString())
                    XMLLoadLocale(eHeader.Item("Difficulty3"), CHDifficulty.Items.Item(3).ToString())
                    XMLLoadLocale(eHeader.Item("Difficulty4"), CHDifficulty.Items.Item(4).ToString())
                    XMLLoadLocale(eHeader.Item("Difficulty5"), CHDifficulty.Items.Item(5).ToString())
                    AddHandler CHDifficulty.SelectedIndexChanged, AddressOf THGenre_TextChanged
                End If

                Dim eGrid As XmlElement = eOptionsPanel.Item("Grid")
                If eGrid IsNot Nothing Then
                    XMLLoadLocale(eGrid.Item("Title"), POGridSwitch.Text)
                    XMLLoadLocale(eGrid.Item("Snap"), CGSnap.Text)
                    XMLLoadLocale(eGrid.Item("BCols"), Label1.Text)
                    XMLLoadLocale(eGrid.Item("DisableVertical"), CGDisableVertical.Text)
                    XMLLoadLocale(eGrid.Item("Scroll"), Label5.Text)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockLeft"), cVSLockL)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockMiddle"), cVSLock)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockRight"), cVSLockR)
                End If

                Dim eWaveForm As XmlElement = eOptionsPanel.Item("WaveForm")
                If eWaveForm IsNot Nothing Then
                    XMLLoadLocale(eWaveForm.Item("Title"), POWaveFormSwitch.Text)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Load"), BWLoad)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Clear"), BWClear)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Lock"), BWLock)
                End If

                Dim eWAV As XmlElement = eOptionsPanel.Item("WAV")
                If eWAV IsNot Nothing Then
                    XMLLoadLocale(eWAV.Item("Title"), POWAVSwitch.Text)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveUp"), BWAVUp)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveDown"), BWAVDown)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Browse"), BWAVBrowse)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Remove"), BWAVRemove)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Duplicate"), BWAVDuplicate)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("OverrideColor"), BWAVColorOverride)
                End If

                Dim eBMP As XmlElement = eOptionsPanel.Item("BMP")
                If eBMP IsNot Nothing Then
                    XMLLoadLocale(eBMP.Item("Title"), POBMPSwitch.Text)
                    XMLLoadLocaleToolTipUniversal(eBMP.Item("MoveUp"), BBMPUp)
                    XMLLoadLocaleToolTipUniversal(eBMP.Item("MoveDown"), BBMPDown)
                    XMLLoadLocaleToolTipUniversal(eBMP.Item("Browse"), BBMPBrowse)
                    XMLLoadLocaleToolTipUniversal(eBMP.Item("Remove"), BBMPRemove)
                End If

                XMLLoadLocale(eOptionsPanel.Item("Beat"), POBeatSwitch.Text)
                XMLLoadLocale(eOptionsPanel.Item("Beat.Apply"), BBeatApply.Text)
                XMLLoadLocale(eOptionsPanel.Item("Beat.Apply"), BBeatApplyV.Text)
                XMLLoadLocale(eOptionsPanel.Item("Expansion"), POExpansionSwitch.Text)
                XMLLoadLocale(eOptionsPanel.Item("ExpansionRandom"), BExpansion.Text)
            End If

            Dim eMessages As XmlElement = Root.Item("Messages")
            If eMessages IsNot Nothing Then
                XMLLoadLocale(eMessages.Item("Err"), Strings.Messages.Err)
                XMLLoadLocale(eMessages.Item("SaveOnExit"), Strings.Messages.SaveOnExit)
                XMLLoadLocale(eMessages.Item("SaveOnExit1"), Strings.Messages.SaveOnExit1)
                XMLLoadLocale(eMessages.Item("SaveOnExit2"), Strings.Messages.SaveOnExit2)
                XMLLoadLocale(eMessages.Item("SaveOnExitOther"), Strings.Messages.SaveOnExitOther)
                XMLLoadLocale(eMessages.Item("PromptEnter"), Strings.Messages.PromptEnter)
                XMLLoadLocale(eMessages.Item("PromptEnterNumeric"), Strings.Messages.PromptEnterNumeric)
                XMLLoadLocale(eMessages.Item("PromptEnterBPM"), Strings.Messages.PromptEnterBPM)
                XMLLoadLocale(eMessages.Item("PromptEnterSTOP"), Strings.Messages.PromptEnterSTOP)
                XMLLoadLocale(eMessages.Item("PromptEnterSCROLL"), Strings.Messages.PromptEnterSCROLL)
                XMLLoadLocale(eMessages.Item("PromptSlashValue"), Strings.Messages.PromptSlashValue)
                XMLLoadLocale(eMessages.Item("InvalidLabel"), Strings.Messages.InvalidLabel)
                XMLLoadLocale(eMessages.Item("CannotFind"), Strings.Messages.CannotFind)
                XMLLoadLocale(eMessages.Item("PleaseRespecifyPath"), Strings.Messages.PleaseRespecifyPath)
                XMLLoadLocale(eMessages.Item("PlayerNotFound"), Strings.Messages.PlayerNotFound)
                XMLLoadLocale(eMessages.Item("PreviewDelError"), Strings.Messages.PreviewDelError)
                XMLLoadLocale(eMessages.Item("NegativeFactorError"), Strings.Messages.NegativeFactorError)
                XMLLoadLocale(eMessages.Item("NegativeDivisorError"), Strings.Messages.NegativeDivisorError)
                XMLLoadLocale(eMessages.Item("PreferencePostpone"), Strings.Messages.PreferencePostpone)
                XMLLoadLocale(eMessages.Item("EraserObsolete"), Strings.Messages.EraserObsolete)
                XMLLoadLocale(eMessages.Item("SaveWarning"), Strings.Messages.SaveWarning)
                XMLLoadLocale(eMessages.Item("NoteOverlapError"), Strings.Messages.NoteOverlapError)
                XMLLoadLocale(eMessages.Item("BPMOverflowError"), Strings.Messages.BPMOverflowError)
                XMLLoadLocale(eMessages.Item("STOPOverflowError"), Strings.Messages.STOPOverflowError)
                XMLLoadLocale(eMessages.Item("SCROLLOverflowError"), Strings.Messages.SCROLLOverflowError)
                XMLLoadLocale(eMessages.Item("SavedFileWillContainErrors"), Strings.Messages.SavedFileWillContainErrors)
                XMLLoadLocale(eMessages.Item("FileAssociationPrompt"), Strings.Messages.FileAssociationPrompt)
                XMLLoadLocale(eMessages.Item("FileAssociationError"), Strings.Messages.FileAssociationError)
                XMLLoadLocale(eMessages.Item("RestoreDefaultSettings"), Strings.Messages.RestoreDefaultSettings)
                XMLLoadLocale(eMessages.Item("RestoreAutosavedFile"), Strings.Messages.RestoreAutosavedFile)
                XMLLoadLocale(eMessages.Item("GhostNotesShowMain"), Strings.Messages.GhostNotesShowMain)
                XMLLoadLocale(eMessages.Item("GhostNotesModifyExpansion1"), Strings.Messages.GhostNotesModifyExpansion1)
                XMLLoadLocale(eMessages.Item("GhostNotesModifyExpansion2"), Strings.Messages.GhostNotesModifyExpansion2)
            End If

            Dim eFileType As XmlElement = Root.Item("FileType")
            If eFileType IsNot Nothing Then
                XMLLoadLocale(eFileType.Item("_all"), Strings.FileType._all)
                XMLLoadLocale(eFileType.Item("_bms"), Strings.FileType._bms)
                XMLLoadLocale(eFileType.Item("BMS"), Strings.FileType.BMS)
                XMLLoadLocale(eFileType.Item("BME"), Strings.FileType.BME)
                XMLLoadLocale(eFileType.Item("BML"), Strings.FileType.BML)
                XMLLoadLocale(eFileType.Item("PMS"), Strings.FileType.PMS)
                XMLLoadLocale(eFileType.Item("TXT"), Strings.FileType.TXT)
                XMLLoadLocale(eFileType.Item("SM"), Strings.FileType.SM)
                XMLLoadLocale(eFileType.Item("IBMSC"), Strings.FileType.IBMSC)
                XMLLoadLocale(eFileType.Item("XML"), Strings.FileType.XML)
                XMLLoadLocale(eFileType.Item("THEME_XML"), Strings.FileType.THEME_XML)
                XMLLoadLocale(eFileType.Item("TH"), Strings.FileType.TH)
                XMLLoadLocale(eFileType.Item("_audio"), Strings.FileType._audio)
                XMLLoadLocale(eFileType.Item("_wave"), Strings.FileType._wave)
                XMLLoadLocale(eFileType.Item("WAV"), Strings.FileType.WAV)
                XMLLoadLocale(eFileType.Item("OGG"), Strings.FileType.OGG)
                XMLLoadLocale(eFileType.Item("MP3"), Strings.FileType.MP3)
                XMLLoadLocale(eFileType.Item("MID"), Strings.FileType.MID)
                XMLLoadLocale(eFileType.Item("_im"), Strings.FileType._im)
                XMLLoadLocale(eFileType.Item("_image"), Strings.FileType._image)
                XMLLoadLocale(eFileType.Item("_movie"), Strings.FileType._movie)
                XMLLoadLocale(eFileType.Item("EXE"), Strings.FileType.EXE)
            End If

            Dim eStatistics As XmlElement = Root.Item("Statistics")
            If eStatistics IsNot Nothing Then
                XMLLoadLocale(eStatistics.Item("Title"), Strings.fStatistics.Title)
                XMLLoadLocale(eStatistics.Item("lBPM"), Strings.fStatistics.lBPM)
                XMLLoadLocale(eStatistics.Item("lSTOP"), Strings.fStatistics.lSTOP)
                XMLLoadLocale(eStatistics.Item("lSCROLL"), Strings.fStatistics.lSCROLL)
                XMLLoadLocale(eStatistics.Item("lA"), Strings.fStatistics.lA)
                XMLLoadLocale(eStatistics.Item("lD"), Strings.fStatistics.lD)
                XMLLoadLocale(eStatistics.Item("lBGM"), Strings.fStatistics.lBGM)
                XMLLoadLocale(eStatistics.Item("lTotal"), Strings.fStatistics.lTotal)
                XMLLoadLocale(eStatistics.Item("lShort"), Strings.fStatistics.lShort)
                XMLLoadLocale(eStatistics.Item("lLong"), Strings.fStatistics.lLong)
                XMLLoadLocale(eStatistics.Item("lLnObj"), Strings.fStatistics.lLnObj)
                XMLLoadLocale(eStatistics.Item("lHidden"), Strings.fStatistics.lHidden)
                XMLLoadLocale(eStatistics.Item("lErrors"), Strings.fStatistics.lErrors)
            End If

            Dim ePlayerOptions As XmlElement = Root.Item("PlayerOptions")
            If ePlayerOptions IsNot Nothing Then
                XMLLoadLocale(ePlayerOptions.Item("Title"), Strings.fopPlayer.Title)
                XMLLoadLocale(ePlayerOptions.Item("Add"), Strings.fopPlayer.Add)
                XMLLoadLocale(ePlayerOptions.Item("Remove"), Strings.fopPlayer.Remove)
                XMLLoadLocale(ePlayerOptions.Item("Path"), Strings.fopPlayer.Path)
                XMLLoadLocale(ePlayerOptions.Item("PlayFromBeginning"), Strings.fopPlayer.PlayFromBeginning)
                XMLLoadLocale(ePlayerOptions.Item("PlayFromHere"), Strings.fopPlayer.PlayFromHere)
                XMLLoadLocale(ePlayerOptions.Item("StopPlaying"), Strings.fopPlayer.StopPlaying)
                XMLLoadLocale(ePlayerOptions.Item("References"), Strings.fopPlayer.References)
                XMLLoadLocale(ePlayerOptions.Item("DirectoryOfApp"), Strings.fopPlayer.DirectoryOfApp)
                XMLLoadLocale(ePlayerOptions.Item("CurrMeasure"), Strings.fopPlayer.CurrMeasure)
                XMLLoadLocale(ePlayerOptions.Item("FileName"), Strings.fopPlayer.FileName)
                XMLLoadLocale(ePlayerOptions.Item("FileNameTemplate"), Strings.fopPlayer.FileNameTemplate)
                XMLLoadLocale(ePlayerOptions.Item("RestoreDefault"), Strings.fopPlayer.RestoreDefault)
            End If

            Dim eVisualOptions As XmlElement = Root.Item("VisualOptions")
            If eVisualOptions IsNot Nothing Then
                XMLLoadLocale(eVisualOptions.Item("Title"), Strings.fopVisual.Title)
                XMLLoadLocale(eVisualOptions.Item("Width"), Strings.fopVisual.Width)
                XMLLoadLocale(eVisualOptions.Item("Caption"), Strings.fopVisual.Caption)
                XMLLoadLocale(eVisualOptions.Item("Note"), Strings.fopVisual.Note)
                XMLLoadLocale(eVisualOptions.Item("Label"), Strings.fopVisual.Label)
                XMLLoadLocale(eVisualOptions.Item("LongNote"), Strings.fopVisual.LongNote)
                XMLLoadLocale(eVisualOptions.Item("LongNoteLabel"), Strings.fopVisual.LongNoteLabel)
                XMLLoadLocale(eVisualOptions.Item("Bg"), Strings.fopVisual.Bg)
                XMLLoadLocale(eVisualOptions.Item("ColumnCaption"), Strings.fopVisual.ColumnCaption)
                XMLLoadLocale(eVisualOptions.Item("ColumnCaptionFont"), Strings.fopVisual.ColumnCaptionFont)
                XMLLoadLocale(eVisualOptions.Item("Background"), Strings.fopVisual.Background)
                XMLLoadLocale(eVisualOptions.Item("Grid"), Strings.fopVisual.Grid)
                XMLLoadLocale(eVisualOptions.Item("SubGrid"), Strings.fopVisual.SubGrid)
                XMLLoadLocale(eVisualOptions.Item("VerticalLine"), Strings.fopVisual.VerticalLine)
                XMLLoadLocale(eVisualOptions.Item("MeasureBarLine"), Strings.fopVisual.MeasureBarLine)
                XMLLoadLocale(eVisualOptions.Item("BGMWaveform"), Strings.fopVisual.BGMWaveform)
                XMLLoadLocale(eVisualOptions.Item("NoteHeight"), Strings.fopVisual.NoteHeight)
                XMLLoadLocale(eVisualOptions.Item("NoteLabel"), Strings.fopVisual.NoteLabel)
                XMLLoadLocale(eVisualOptions.Item("MeasureLabel"), Strings.fopVisual.MeasureLabel)
                XMLLoadLocale(eVisualOptions.Item("LabelVerticalShift"), Strings.fopVisual.LabelVerticalShift)
                XMLLoadLocale(eVisualOptions.Item("LabelHorizontalShift"), Strings.fopVisual.LabelHorizontalShift)
                XMLLoadLocale(eVisualOptions.Item("LongNoteLabelHorizontalShift"), Strings.fopVisual.LongNoteLabelHorizontalShift)
                XMLLoadLocale(eVisualOptions.Item("HiddenNoteOpacity"), Strings.fopVisual.HiddenNoteOpacity)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnMouseOver"), Strings.fopVisual.NoteBorderOnMouseOver)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnSelection"), Strings.fopVisual.NoteBorderOnSelection)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnAdjustingLength"), Strings.fopVisual.NoteBorderOnAdjustingLength)
                XMLLoadLocale(eVisualOptions.Item("SelectionBoxBorder"), Strings.fopVisual.SelectionBoxBorder)
                XMLLoadLocale(eVisualOptions.Item("TSCursor"), Strings.fopVisual.TSCursor)
                XMLLoadLocale(eVisualOptions.Item("TSSplitter"), Strings.fopVisual.TSSplitter)
                XMLLoadLocale(eVisualOptions.Item("TSCursorSensitivity"), Strings.fopVisual.TSCursorSensitivity)
                XMLLoadLocale(eVisualOptions.Item("TSMouseOverBorder"), Strings.fopVisual.TSMouseOverBorder)
                XMLLoadLocale(eVisualOptions.Item("TSFill"), Strings.fopVisual.TSFill)
                XMLLoadLocale(eVisualOptions.Item("TSBPM"), Strings.fopVisual.TSBPM)
                XMLLoadLocale(eVisualOptions.Item("TSBPMFont"), Strings.fopVisual.TSBPMFont)
                XMLLoadLocale(eVisualOptions.Item("MiddleSensitivity"), Strings.fopVisual.MiddleSensitivity)
            End If

            Dim eGeneralOptions As XmlElement = Root.Item("GeneralOptions")
            If eGeneralOptions IsNot Nothing Then
                XMLLoadLocale(eGeneralOptions.Item("Title"), Strings.fopGeneral.Title)
                XMLLoadLocale(eGeneralOptions.Item("AssociateFileType"), Strings.fopGeneral.AssociateFileType)
                XMLLoadLocale(eGeneralOptions.Item("MaxGridPartition"), Strings.fopGeneral.MaxGridPartition)
                XMLLoadLocale(eGeneralOptions.Item("MinimumJackThreshold"), Strings.fopGeneral.MinimumJackThreshold)
                XMLLoadLocale(eGeneralOptions.Item("AutoLongNoteGap"), Strings.fopGeneral.AutoLongNoteGap)
                XMLLoadLocale(eGeneralOptions.Item("TemporaryBMSFilename"), Strings.fopGeneral.TemporaryBMSFilename)

                XMLLoadLocale(eGeneralOptions.Item("AutoSave"), Strings.fopGeneral.AutoSave)
                XMLLoadLocale(eGeneralOptions.Item("minutes"), Strings.fopGeneral.minutes)
                XMLLoadLocale(eGeneralOptions.Item("BeepWhileSaved"), Strings.fopGeneral.BeepWhileSaved)
                XMLLoadLocale(eGeneralOptions.Item("Preload"), Strings.fopGeneral.Preload)
                XMLLoadLocale(eGeneralOptions.Item("ExtendBPM"), Strings.fopGeneral.ExtendBPM)
                XMLLoadLocale(eGeneralOptions.Item("ExtendSTOP"), Strings.fopGeneral.ExtendSTOP)
                XMLLoadLocale(eGeneralOptions.Item("DisplayAudioLines"), Strings.fopGeneral.DisplayAudioLines)
                XMLLoadLocale(eGeneralOptions.Item("SnapToVerticalPositionTemplate"), Strings.fopGeneral.SnapToVerticalPositionTemplate)
                XMLLoadLocale(eGeneralOptions.Item("SnapToVerticalPositionPattern"), Strings.fopGeneral.SnapToVerticalPositionPattern)

                XMLLoadLocale(eGeneralOptions.Item("MouseWheel"), Strings.fopGeneral.MouseWheel)
                XMLLoadLocale(eGeneralOptions.Item("TextEncoding"), Strings.fopGeneral.TextEncoding)
                XMLLoadLocale(eGeneralOptions.Item("PageUpDown"), Strings.fopGeneral.PageUpDown)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButton"), Strings.fopGeneral.MiddleButton)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButtonAuto"), Strings.fopGeneral.MiddleButtonAuto)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButtonDrag"), Strings.fopGeneral.MiddleButtonDrag)
                XMLLoadLocale(eGeneralOptions.Item("AutoFocusOnMouseEnter"), Strings.fopGeneral.AutoFocusOnMouseEnter)
                XMLLoadLocale(eGeneralOptions.Item("DisableFirstClick"), Strings.fopGeneral.DisableFirstClick)
                XMLLoadLocale(eGeneralOptions.Item("StopPreviewOnClick"), Strings.fopGeneral.StopPreviewOnClick)
            End If

            Dim eVisualOverrideOptions As XmlElement = Root.Item("VisualOverrideOptions")
            If eVisualOverrideOptions IsNot Nothing Then
                XMLLoadLocale(eVisualOverrideOptions.Item("Title"), Strings.fopVisualOverride.Title)
                XMLLoadLocale(eVisualOverrideOptions.Item("Add"), Strings.fopVisualOverride.Add)
                XMLLoadLocale(eVisualOverrideOptions.Item("Up"), Strings.fopVisualOverride.Up)
                XMLLoadLocale(eVisualOverrideOptions.Item("Down"), Strings.fopVisualOverride.Down)
                XMLLoadLocale(eVisualOverrideOptions.Item("Duplicate"), Strings.fopVisualOverride.Duplicate)
                XMLLoadLocale(eVisualOverrideOptions.Item("Split"), Strings.fopVisualOverride.Split)
                XMLLoadLocale(eVisualOverrideOptions.Item("SemiAutoAssign"), Strings.fopVisualOverride.SemiAutoAssign)
                XMLLoadLocale(eVisualOverrideOptions.Item("Remove"), Strings.fopVisualOverride.Remove)
                XMLLoadLocale(eVisualOverrideOptions.Item("LoadSettingsFrom"), Strings.fopVisualOverride.LoadSettingsFrom)
                XMLLoadLocale(eVisualOverrideOptions.Item("SaveSettingsTo"), Strings.fopVisualOverride.SaveSettingsTo)
                XMLLoadLocale(eVisualOverrideOptions.Item("EnableItem"), Strings.fopVisualOverride.EnableItem)
                XMLLoadLocale(eVisualOverrideOptions.Item("ColorS"), Strings.fopVisualOverride.ColorS)
                XMLLoadLocale(eVisualOverrideOptions.Item("ColorG"), Strings.fopVisualOverride.ColorG)
                XMLLoadLocale(eVisualOverrideOptions.Item("ColorGHSLU"), Strings.fopVisualOverride.ColorGHSLU)
                XMLLoadLocale(eVisualOverrideOptions.Item("ColorGHSLD"), Strings.fopVisualOverride.ColorGHSLD)
                XMLLoadLocale(eVisualOverrideOptions.Item("OptionName"), Strings.fopVisualOverride.OptionName)
                XMLLoadLocale(eVisualOverrideOptions.Item("LabelRange"), Strings.fopVisualOverride.LabelRange)
                XMLLoadLocale(eVisualOverrideOptions.Item("ToText"), Strings.fopVisualOverride.ToText)
                XMLLoadLocale(eVisualOverrideOptions.Item("NoteColor"), Strings.fopVisualOverride.NoteColor)
                XMLLoadLocale(eVisualOverrideOptions.Item("NoteColorRange"), Strings.fopVisualOverride.NoteColorRange)
                XMLLoadLocale(eVisualOverrideOptions.Item("Preview"), Strings.fopVisualOverride.Preview)

                XMLLoadLocale(eVisualOverrideOptions.Item("Chart"), Strings.fopVisualOverride.Chart)
                XMLLoadLocale(eVisualOverrideOptions.Item("Song"), Strings.fopVisualOverride.Song)
                XMLLoadLocale(eVisualOverrideOptions.Item("Editor"), Strings.fopVisualOverride.Editor)

                XMLLoadLocale(eVisualOverrideOptions.Item("SplitMsgCannot"), Strings.fopVisualOverride.SplitMsgCannot)
                XMLLoadLocale(eVisualOverrideOptions.Item("SplitMsgUpper"), Strings.fopVisualOverride.SplitMsgUpper)
                XMLLoadLocale(eVisualOverrideOptions.Item("SplitMsgNotBetweenRange"), Strings.fopVisualOverride.SplitMsgNotBetweenRange)

                XMLLoadLocale(eVisualOverrideOptions.Item("SemiAutoMsgAssign"), Strings.fopVisualOverride.SemiAutoMsgAssign)
                XMLLoadLocale(eVisualOverrideOptions.Item("SemiAutoMsgNone"), Strings.fopVisualOverride.SemiAutoMsgNone)
                XMLLoadLocale(eVisualOverrideOptions.Item("SemiAutoMsgAssignMore"), Strings.fopVisualOverride.SemiAutoMsgAssignMore)

                XMLLoadLocale(eVisualOverrideOptions.Item("SaveCurrentSettings"), Strings.fopVisualOverride.SaveCurrentSettings)
                XMLLoadLocale(eVisualOverrideOptions.Item("OverwriteExistingSettings"), Strings.fopVisualOverride.OverwriteExistingSettings)
            End If

            Dim eKeybindingOptions As XmlElement = Root.Item("KeybindingOptions")
            If eKeybindingOptions IsNot Nothing Then
                XMLLoadLocale(eKeybindingOptions.Item("Title"), Strings.fopKeybinding.Title)
                XMLLoadLocale(eKeybindingOptions.Item("BindingName"), Strings.fopKeybinding.BindingName)
                XMLLoadLocale(eKeybindingOptions.Item("Description"), Strings.fopKeybinding.Description)
                XMLLoadLocale(eKeybindingOptions.Item("Keybindings"), Strings.fopKeybinding.Keybindings)

                XMLLoadLocale(eKeybindingOptions.Item("MoveTo"), Strings.fopKeybinding.MoveTo)
                XMLLoadLocale(eKeybindingOptions.Item("MoveToDescription"), Strings.fopKeybinding.MoveToDescription)
                XMLLoadLocale(eKeybindingOptions.Item("MoveToScratchDescription"), Strings.fopKeybinding.MoveToScratchDescription)
                XMLLoadLocale(eKeybindingOptions.Item("MoveToBGMDescription"), Strings.fopKeybinding.MoveToBGMDescription)
                XMLLoadLocale(eKeybindingOptions.Item("MoveToTemplate"), Strings.fopKeybinding.MoveToTemplate)
                XMLLoadLocale(eKeybindingOptions.Item("MoveToTemplateDescription"), Strings.fopKeybinding.MoveToTemplateDescription)
                XMLLoadLocale(eKeybindingOptions.Item("CheckTechnicalError"), Strings.fopKeybinding.CheckTechnicalError)
                XMLLoadLocale(eKeybindingOptions.Item("SelectExpansionSection"), Strings.fopKeybinding.SelectExpansionSection)
                XMLLoadLocale(eKeybindingOptions.Item("PastePattern"), Strings.fopKeybinding.PastePattern)
                XMLLoadLocale(eKeybindingOptions.Item("SelectHovered"), Strings.fopKeybinding.SelectHovered)

                XMLLoadLocale(eKeybindingOptions.Item("ErrorUnavailable"), Strings.fopKeybinding.ErrorUnavailable)
                XMLLoadLocale(eKeybindingOptions.Item("ErrorAssigned"), Strings.fopKeybinding.ErrorAssigned)
                XMLLoadLocale(eKeybindingOptions.Item("ErrorNoteAssignment"), Strings.fopKeybinding.ErrorNoteAssignment)
                ReloadKeybindingNames(Keybindings)
            End If

            Dim eExpandOptions As XmlElement = Root.Item("ExpandOptions")
            If eExpandOptions IsNot Nothing Then
                XMLLoadLocale(eExpandOptions.Item("Title"), Strings.fopExpand.Title)
                XMLLoadLocale(eExpandOptions.Item("SelectExpansionCode"), Strings.fopExpand.SelectExpansionCode)
                XMLLoadLocale(eExpandOptions.Item("DisplayGhostNotes"), Strings.fopExpand.DisplayGhostNotes)
                XMLLoadLocale(eExpandOptions.Item("DisplayGhostNotesAll"), Strings.fopExpand.DisplayGhostNotesAll)
                XMLLoadLocale(eExpandOptions.Item("ModifyNotes"), Strings.fopExpand.ModifyNotes)
                XMLLoadLocale(eExpandOptions.Item("ModifySection"), Strings.fopExpand.ModifySection)
                XMLLoadLocale(eExpandOptions.Item("RemoveGhostNotes"), Strings.fopExpand.RemoveGhostNotes)
                XMLLoadLocale(eExpandOptions.Item("RemoveGhostNotes"), Strings.fopExpand.RemoveGhostNotes)
                XMLLoadLocale(eExpandOptions.Item("ErrorEmpty"), Strings.fopExpand.ErrorEmpty)
                XMLLoadLocale(eExpandOptions.Item("ErrorNoLineSelected"), Strings.fopExpand.ErrorNoLineSelected)
                XMLLoadLocale(eExpandOptions.Item("ErrorNotDetected"), Strings.fopExpand.ErrorNotDetected)
            End If

            Dim eTOTALTool As XmlElement = Root.Item("TOTALTool")
            If eTOTALTool IsNot Nothing Then
                XMLLoadLocale(eTOTALTool.Item("Title"), Strings.fopTOTAL.Title)
                XMLLoadLocale(eTOTALTool.Item("CalculationOption"), Strings.fopTOTAL.CalculationOption)
                XMLLoadLocale(eTOTALTool.Item("Parameters"), Strings.fopTOTAL.Parameters)
                XMLLoadLocale(eTOTALTool.Item("TOTAL"), Strings.fopTOTAL.TOTAL)
                XMLLoadLocale(eTOTALTool.Item("CalculationFormula"), Strings.fopTOTAL.CalculationFormula)
                XMLLoadLocale(eTOTALTool.Item("IIDXTOTALSupposition1"), Strings.fopTOTAL.IIDXTOTALSupposition1)
                XMLLoadLocale(eTOTALTool.Item("IIDXTOTALSupposition2"), Strings.fopTOTAL.IIDXTOTALSupposition2)
                XMLLoadLocale(eTOTALTool.Item("Multiplier"), Strings.fopTOTAL.Multiplier)
                XMLLoadLocale(eTOTALTool.Item("Notes"), Strings.fopTOTAL.Notes)
                XMLLoadLocale(eTOTALTool.Item("GlobalMultiplier"), Strings.fopTOTAL.GlobalMultiplier)
                XMLLoadLocale(eTOTALTool.Item("DecimalPlaces"), Strings.fopTOTAL.DecimalPlaces)
                XMLLoadLocale(eTOTALTool.Item("DisplayValue"), Strings.fopTOTAL.DisplayValue)
                XMLLoadLocale(eTOTALTool.Item("DisplayRecommendedTOTAL"), Strings.fopTOTAL.DisplayRecommendedTOTAL)
                XMLLoadLocale(eTOTALTool.Item("AutofillTOTAL"), Strings.fopTOTAL.AutofillTOTAL)
            End If

            Dim eFind As XmlElement = Root.Item("Find")
            If eFind IsNot Nothing Then
                XMLLoadLocale(eFind.Item("NoteRange"), Strings.fFind.NoteRange)
                XMLLoadLocale(eFind.Item("MeasureRange"), Strings.fFind.MeasureRange)
                XMLLoadLocale(eFind.Item("LabelRange"), Strings.fFind.LabelRange)
                XMLLoadLocale(eFind.Item("ValueRange"), Strings.fFind.ValueRange)
                XMLLoadLocale(eFind.Item("to"), Strings.fFind.to_)
                XMLLoadLocale(eFind.Item("Selected"), Strings.fFind.Selected)
                XMLLoadLocale(eFind.Item("UnSelected"), Strings.fFind.UnSelected)
                XMLLoadLocale(eFind.Item("ShortNote"), Strings.fFind.ShortNote)
                XMLLoadLocale(eFind.Item("LongNote"), Strings.fFind.LongNote)
                XMLLoadLocale(eFind.Item("Hidden"), Strings.fFind.Hidden)
                XMLLoadLocale(eFind.Item("Visible"), Strings.fFind.Visible)
                XMLLoadLocale(eFind.Item("Column"), Strings.fFind.Column)
                XMLLoadLocale(eFind.Item("SelectAll"), Strings.fFind.SelectAll)
                XMLLoadLocale(eFind.Item("SelectInverse"), Strings.fFind.SelectInverse)
                XMLLoadLocale(eFind.Item("UnselectAll"), Strings.fFind.UnselectAll)
                XMLLoadLocale(eFind.Item("Operation"), Strings.fFind.Operation)
                XMLLoadLocale(eFind.Item("ReplaceWithLabel"), Strings.fFind.ReplaceWithLabel)
                XMLLoadLocale(eFind.Item("ReplaceWithValue"), Strings.fFind.ReplaceWithValue)
                XMLLoadLocale(eFind.Item("Select"), Strings.fFind.Select_)
                XMLLoadLocale(eFind.Item("Unselect"), Strings.fFind.Unselect_)
                XMLLoadLocale(eFind.Item("Delete"), Strings.fFind.Delete_)
                XMLLoadLocale(eFind.Item("Close"), Strings.fFind.Close_)
            End If

            Dim eImportSM As XmlElement = Root.Item("ImportSM")
            If eImportSM IsNot Nothing Then
                XMLLoadLocale(eImportSM.Item("Title"), Strings.fImportSM.Title)
                XMLLoadLocale(eImportSM.Item("Difficulty"), Strings.fImportSM.Difficulty)
                XMLLoadLocale(eImportSM.Item("Note"), Strings.fImportSM.Note)
            End If

            Dim eFileAssociation As XmlElement = Root.Item("FileAssociation")
            If eFileAssociation IsNot Nothing Then
                XMLLoadLocale(eFileAssociation.Item("BMS"), Strings.FileAssociation.BMS)
                XMLLoadLocale(eFileAssociation.Item("BME"), Strings.FileAssociation.BME)
                XMLLoadLocale(eFileAssociation.Item("BML"), Strings.FileAssociation.BML)
                XMLLoadLocale(eFileAssociation.Item("PMS"), Strings.FileAssociation.PMS)
                XMLLoadLocale(eFileAssociation.Item("IBMSC"), Strings.FileAssociation.IBMSC)
                XMLLoadLocale(eFileAssociation.Item("Open"), Strings.FileAssociation.Open)
                XMLLoadLocale(eFileAssociation.Item("Preview"), Strings.FileAssociation.Preview)
                XMLLoadLocale(eFileAssociation.Item("ViewCode"), Strings.FileAssociation.ViewCode)
            End If

            DispLang = Path.Replace(My.Application.Info.DirectoryPath & "\", "")

        Catch ex As Exception
            MsgBox(Path & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Exclamation)

        Finally
            If FileStream IsNot Nothing Then FileStream.Close()

            POHeaderPart2.Visible = xPOHeaderPart2
            POGridPart2.Visible = xPOGridPart2
            POWaveFormPart2.Visible = xPOWaveFormPart2
        End Try

        For Each keybind In Keybindings
            RenameShortcut(keybind)
        Next

        'File.Delete(xTempFileName)
    End Sub

    Private Sub LoadThemeComptability(ByVal xPath As String)
        Try
            Dim xStrLine() As String = Split(My.Computer.FileSystem.ReadAllText(xPath), vbCrLf)
            If xStrLine(0).Trim <> "iBMSC Configuration Settings Format" And xStrLine(0).Trim <> "iBMSC Theme Format" Then Exit Sub

            Dim xW1 As String = ""
            Dim xW2 As String = ""

            For Each xLine As String In xStrLine
                xW1 = UCase(Mid(xLine, 1, InStr(xLine, "=") - 1))
                xW2 = Mid(xLine, InStr(xLine, "=") + 1)

                Dim xW2Int As Integer = CInt(xW2)

                Select Case xW1
                    Case "VOTITLE" : vo.ColumnTitle.Color = Color.FromArgb(xW2Int)
                    Case "VOTITLEFONT" : vo.ColumnTitleFont = StringToFont(xW2, Me.Font)
                    Case "VOBG" : vo.Bg.Color = Color.FromArgb(xW2Int)
                    Case "VOGRID" : vo.pGrid.Color = Color.FromArgb(xW2Int)
                    Case "VOSUB" : vo.pSub.Color = Color.FromArgb(xW2Int)
                    Case "VOVLINE" : vo.pVLine.Color = Color.FromArgb(xW2Int)
                    Case "VOMLINE" : vo.pMLine.Color = Color.FromArgb(xW2Int)
                    Case "VOBGMWAV"
                        vo.pBGMWav.Color = Color.FromArgb(xW2Int)
                        TWTransparency.Value = vo.pBGMWav.Color.A
                        TWTransparency2.Value = vo.pBGMWav.Color.A
                        TWSaturation.Value = CDec(vo.pBGMWav.Color.GetSaturation * 1000)
                        TWSaturation2.Value = CInt(vo.pBGMWav.Color.GetSaturation * 1000)
                    Case "VOSELBOX" : vo.SelBox.Color = Color.FromArgb(xW2Int)
                    Case "VOPECURSOR" : vo.PECursor.Color = Color.FromArgb(xW2Int)
                    Case "VOPEHALF" : vo.PEHalf.Color = Color.FromArgb(xW2Int)
                    Case "VOPEDELTAMOUSEOVER" : vo.PEDeltaMouseOver = xW2Int
                    Case "VOPEMOUSEOVER" : vo.PEMouseOver.Color = Color.FromArgb(xW2Int)
                    Case "VOPESEL" : vo.PESel.Color = Color.FromArgb(xW2Int)
                    Case "VOPEBPM" : vo.PEBPM.Color = Color.FromArgb(xW2Int)
                    Case "VOPEBPMFONT" : vo.PEBPMFont = StringToFont(xW2, Me.Font)
                    Case "VKHEIGHT" : vo.kHeight = xW2Int
                    Case "VKFONT" : vo.kFont = StringToFont(xW2, Me.Font)
                    Case "VKMFONT" : vo.kMFont = StringToFont(xW2, Me.Font)
                    Case "VKLABELVSHIFT" : vo.kLabelVShift = xW2Int
                    Case "VKLABELHSHIFT" : vo.kLabelHShift = xW2Int
                    Case "VKLABELHSHIFTL" : vo.kLabelHShiftL = xW2Int
                    Case "VKMOUSEOVER" : vo.kMouseOver.Color = Color.FromArgb(xW2Int)
                    Case "VKMOUSEOVERE " : vo.kMouseOverE.Color = Color.FromArgb(xW2Int)
                    Case "VKSELECTED" : vo.kSelected.Color = Color.FromArgb(xW2Int)
                        'Case "VKHIDTRANSPARENCY" : vo.kOpacity = xW2Int

                    Case "KLENGTH"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).Width = CInt(xE(i))
                        Next

                    Case "KTITLE"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).Title = xE(i)
                        Next

                    Case "KCOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).setNoteColor(CInt(xE(i)))
                        Next

                    Case "KCOLORL"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).setLNoteColor(CInt(xE(i)))
                        Next

                    Case "KFORECOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).cText = Color.FromArgb(CInt(xE(i)))
                        Next

                    Case "KFORECOLORL"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).cLText = Color.FromArgb(CInt(xE(i)))
                        Next

                    Case "KBGCOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i As Integer = 0 To 26
                            column(i).cBG = Color.FromArgb(CInt(xE(i)))
                        Next

                End Select

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, Strings.Messages.Err)

        Finally
            UpdateColumnsX()

        End Try
    End Sub

    Private Function LoadThemeComptability_SplitStringInto26Parts(ByVal xLine As String) As String()
        Dim xE() As String = Split(xLine, ",")
        ReDim Preserve xE(26)
        Return xE
    End Function

    Private Sub LoadLang(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim LangS As ToolStripItem = CType(sender, ToolStripItem)
        Dim xFN2 As String = LangS.ToolTipText
        'ReadLanguagePack(xFN2)
        LoadLocale(xFN2)
    End Sub

    Private Sub LoadLocaleXML(xStr As FileInfo)
        Dim d As New XmlDocument
        Dim fs As New FileStream(xStr.FullName, FileMode.Open, FileAccess.Read)

        Try
            d.Load(fs)
            Dim xName As String = d.Item("iBMSC.Locale").GetAttribute("Name")
            If xName = "" Then xName = xStr.Name

            cmnLanguage.Items.Add(xName, Nothing, AddressOf LoadLang)
            cmnLanguage.Items(cmnLanguage.Items.Count - 1).ToolTipText = xStr.FullName

        Catch ex As Exception
            MsgBox(xStr.FullName & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Exclamation)

        End Try

        fs.Close()
    End Sub

    Private Sub LoadTheme(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If Not File.Exists(My.Application.Info.DirectoryPath & "\Data\" & sender.Text) Then Exit Sub
        'SaveTheme = True
        'LoadCFF(My.Computer.FileSystem.ReadAllText(My.Application.Info.DirectoryPath & "\Theme\" & sender.Text, System.Text.Encoding.Unicode))
        Dim ThemeS As ToolStripItem = CType(sender, ToolStripItem)
        LoadSettings(My.Application.Info.DirectoryPath & "\Data\" & ThemeS.Text)
        RefreshPanelAll()
    End Sub

    Private Sub LoadColorOverride(ByVal Path As String)
        Dim F As String = ColorOverridePath(Path)

        If My.Computer.FileSystem.FileExists(F) Then

            Dim Doc As New XmlDocument
            Dim FileStream As New IO.FileStream(F, FileMode.Open, FileAccess.Read)
            Doc.Load(FileStream)

            Dim Root As XmlElement = Doc.Item("ColorOverride")

            Dim n As Integer = CInt(Root.GetAttribute("Count"))
            ReDim COverrides(n)
            Dim i As Integer
            For Each eColor As XmlElement In Root.ChildNodes
                With eColor
                    XMLLoadAttribute(.GetAttribute("Index"), i)
                    XMLLoadAttribute(.GetAttribute("Name"), COverrides(i).Name)
                    XMLLoadAttribute(.GetAttribute("Enabled"), COverrides(i).Enabled)
                    XMLLoadAttribute(.GetAttribute("ColorOption"), COverrides(i).ColorOption)
                    XMLLoadAttribute(.GetAttribute("RangeL"), COverrides(i).RangeL)
                    XMLLoadAttribute(.GetAttribute("RangeU"), COverrides(i).RangeU)
                    XMLLoadAttribute(.GetAttribute("NoteColor"), COverrides(i).NoteColor)
                    XMLLoadAttribute(.GetAttribute("NoteColorU"), COverrides(i).NoteColorU)
                End With
            Next
            FileStream.Close()
        Else
            ReDim COverrides(-1)
        End If
        LoadColorOverrideActive()
    End Sub

    Private Sub LoadColorOverrideActive()
        For xIN = 1 To UBound(COverridesColors)
            COverridesColors(xIN) = Color.Empty
            For xICO = 0 To UBound(COverrides)
                If Not COverrides(xICO).Enabled Then Continue For
                Dim CORangeL = COverrides(xICO).RangeL
                Dim CORangeU = COverrides(xICO).RangeU
                Dim CONoteColor = COverrides(xICO).NoteColor
                Dim CONoteColorU = COverrides(xICO).NoteColorU
                If xIN >= CORangeL AndAlso xIN <= CORangeU Then
                    Select Case COverrides(xICO).ColorOption
                        Case 0
                            COverridesColors(xIN) = Color.FromArgb(CONoteColor)
                            Exit For
                        Case 1
                            Dim CORange = Math.Max(1, CORangeU - CORangeL)
                            COverridesColors(xIN) = InterpolateColorARGB(Color.FromArgb(CONoteColor), Color.FromArgb(CONoteColorU), (xIN - CORangeL) / CORange)
                            Exit For
                        Case 2
                            Dim CORange = Math.Max(1, CORangeU - CORangeL)
                            COverridesColors(xIN) = InterpolateColorAHSL(Color.FromArgb(CONoteColor), Color.FromArgb(CONoteColorU), (xIN - CORangeL) / CORange)
                            Exit For
                        Case 3
                            Dim CORange = Math.Max(1, CORangeU - CORangeL)
                            COverridesColors(xIN) = InterpolateColorAHSL(Color.FromArgb(CONoteColor), Color.FromArgb(CONoteColorU), (xIN - CORangeL) / CORange, , 0)
                            Exit For
                    End Select
                End If
            Next
        Next
    End Sub
End Class
