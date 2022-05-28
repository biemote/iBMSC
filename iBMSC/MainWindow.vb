Imports System.Linq
Imports iBMSC.Editor


Public Class MainWindow


    'Public Structure MARGINS
    '    Public Left As Integer
    '    Public Right As Integer
    '    Public Top As Integer
    '    Public Bottom As Integer
    'End Structure

    '<System.Runtime.InteropServices.DllImport("dwmapi.dll")> _
    'Public Shared Function DwmIsCompositionEnabled(ByRef en As Integer) As Integer
    'End Function
    '<System.Runtime.InteropServices.DllImport("dwmapi.dll")> _
    'Public Shared Function DwmExtendFrameIntoClientArea(ByVal hwnd As IntPtr, ByRef margin As MARGINS) As Integer
    'End Function
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture" () As Integer

    'Private Declare Auto Function GetWindowLong Lib "user32" (ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    'Private Declare Auto Function SetWindowLong Lib "user32" (ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    '<DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    'Private Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean
    'End Function

    'Private Const GWL_STYLE As Integer = -16
    'Private Const WS_CAPTION As Integer = &HC00000
    'Private Const SWP_NOSIZE As Integer = &H1
    'Private Const SWP_NOMOVE As Integer = &H2
    'Private Const SWP_NOZORDER As Integer = &H4
    'Private Const SWP_NOACTIVATE As Integer = &H10
    'Private Const SWP_FRAMECHANGED As Integer = &H20
    'Private Const SWP_REFRESH As Integer = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED

    Public Sub New()
        InitializeComponent()
        Audio.Initialize()
    End Sub

    Dim MeasureLength(999) As Double
    Dim MeasureBottom(999) As Double

    Public Function MeasureUpper(idx As Integer) As Double
        Return MeasureBottom(idx) + MeasureLength(idx)
    End Function


    Dim Notes() As Note = {New Note(niBPM, -1, 1200000, 0, False)}
    Dim NotesTemplate() As Note = {New Note(niBPM, -1, 1200000, 0, False)}
    Dim mColumn(999) As Integer  '0 = no column, 1 = 1 column, etc.
    Dim GreatestVPosition As Double    '+ 2000 = -VS.Minimum

    Dim VSValue As Integer = 0 'Store value before ValueChange event
    Dim HSValue As Integer = 0 'Store value before ValueChange event

    'Dim SortingMethod As Integer = 1
    Dim MiddleButtonMoveMethod As Integer = 0
    Dim TextEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
    Dim DispLang As String = ""     'Display Language
    Dim Recent() As String = {"", "", "", "", ""}
    Dim NTInput As Boolean = True
    Dim ShowFileName As Boolean = False
    Dim ShowWaveform As Boolean = False
    Dim ShowBMP As Boolean = False

    Dim BeepWhileSaved As Boolean = True
    Dim PreloadBMSStruct As Boolean = False
    Dim BPMx1296 As Boolean = False
    Dim STOPx1296 As Boolean = False
    Dim AudioLine As Boolean = True
    Dim TemplateSnapToVPosition As Boolean = False
    Dim PastePatternToVPosition As Boolean = False

    Dim IsApplicationInitializing As Boolean = True
    Dim FirstMouseEnter As Boolean = True

    Dim WAVMultiSelect As Boolean = True
    Dim WAVChangeLabel As Boolean = True
    Dim BMPMultiSelect As Boolean = True
    Dim BMPChangeLabel As Boolean = True
    Dim BeatChangeMode As Integer = 0

    'Dim FloatTolerance As Double = 0.0001R
    Dim BMSGridLimit As Double = 1.0R

    Dim LnObj As Integer = 0    '0 for none, 1-1295 for 01-ZZ

    'IO
    Dim FileNameInit As String = "Untitled.bms"
    Dim FileName As String = FileNameInit
    Dim TempFileName As String = "___TempBMS.bmsc"
    Dim RandomTempFileName As String = "___TempRandom" & GenerateRandomString(6, False) & ".bmsc"
    Public ExpansionSplit(2) As String
    'Dim TitlePath As New Drawing2D.GraphicsPath
    Dim InitPath As String = ""
    Dim IsSaved As Boolean = True
    Dim GhostMode As Integer = 0 ' 0 - Default, ghost notes entirely uneditable, 1 - Ghost notes loaded with expectation of editing them, 2 - Ghost notes loaded as main notes and main notes temporarily changed to ghost notes
    Dim GhostEdit As Boolean = False
    Dim FileNameTemplate As String = ""

    'Variables for Drag/Drop
    Dim DDFileName() As String = {}
    Dim SupportedFileExtension() As String = {".bms", ".bme", ".bml", ".pms", ".txt", ".sm", ".ibmsc"}
    Dim SupportedAudioExtension() As String = {}
    Dim SupportedImageExtension() As String = {".bmp", ".png", ".jpg", ".jpeg", ".gif", ".mpg", ".mpeg", ".avi", ".m1v", ".m2v", ".m4v", ".mp4", ".webm", ".wmv"}

    'Variables for theme
    'Dim SaveTheme As Boolean = True

    'Variables for undo/redo
    Dim UndoRedoCount As Integer = 200
    Dim sUndo(UndoRedoCount) As UndoRedo.LinkedURCmd
    Dim sRedo(UndoRedoCount) As UndoRedo.LinkedURCmd
    Dim sI As Integer = 0

    'Variables for select tool
    Dim DisableVerticalMove As Boolean = False
    Dim KMouseOver As Integer = -1   'Mouse is on which note (for drawing green outline)
    Dim LastMouseDownLocation As PointF = New Point(-1, -1)          'Mouse is clicked on which point (location for display) (for selection box)
    Dim pMouseMove As PointF = New Point(-1, -1)          'Mouse is moved to which point   (location for display) (for selection box)
    'Dim KMouseDown As Integer = -1   'Mouse is clicked on which note (for moving)
    Dim deltaVPosition As Double = 0   'difference between mouse and VPosition of K
    Dim bAdjustLength As Boolean     'If adjusting note length instead of moving it
    Dim bAdjustUpper As Boolean      'true = Adjusting upper end, false = adjusting lower end
    Dim bAdjustSingle As Boolean     'true if there is only one note to be adjusted
    Dim tempX As Integer
    Dim tempY As Integer
    Dim tempV As Integer
    Dim tempH As Integer
    Dim MiddleButtonLocation As New Point(0, 0)
    Dim MiddleButtonClicked As Boolean = False
    Dim MouseMoveStatus As Point = New Point(0, 0)  'mouse is moved to which point (For Status Panel)
    'Dim uCol As Integer         'temp variables for undo, original enabled columnindex
    'Dim uVPos As Double         'temp variables for undo, original vposition
    'Dim uPairWithI As Double    'temp variables for undo, original note length
    Dim uAdded As Boolean       'temp variables for undo, if undo command is added
    'Dim uNote As Note           'temp variables for undo, original note
    Dim SelectedNotes(-1) As Note        'temp notes for undo
    Dim ctrlPressed As Boolean = False          'Indicates if the CTRL key is pressed while mousedown
    Dim DuplicatedSelectedNotes As Boolean = False     'Indicates if duplicate notes of select/unselect note

    'Variables for write tool
    Dim ShouldDrawTempNote As Boolean = False
    Dim SelectedColumn As Integer = -1
    Dim TempVPosition As Double = -1.0#
    Dim TempLength As Double = 0.0#

    'Variables for post effects tool
    Dim vSelStart As Double = 192.0#
    Dim vSelLength As Double = 0.0#
    Dim vSelHalf As Double = 0.0#
    Dim vSelMouseOverLine As Integer = 0  '0 = nothing, 1 = start, 2 = half, 3 = end
    Dim vSelAdjust As Boolean = False
    Dim vSelK() As Note = {}
    Dim vSelPStart As Double = 192.0#
    Dim vSelPLength As Double = 0.0#
    Dim vSelPHalf As Double = 0.0#

    'Variables for Full-Screen Mode
    Dim isFullScreen As Boolean = False
    Dim previousWindowState As FormWindowState = FormWindowState.Normal
    Dim previousWindowPosition As New Rectangle(0, 0, 0, 0)

    'Variables misc
    Dim menuVPosition As Double = 0.0#
    Dim tempResize As Integer = 0
    Dim hCOM(1295) As String '
    Dim hCOMNum As Integer = 0
    Dim gXKeyMode As String = "SP" ' Determines from column width 7key mode, 9key mode or 14key mode
    Dim gXKeyCol() As Integer
    Dim gLNGap As Double = 16
    Dim wLWAV(1295) As WavSample
    Dim WaveformLoaded As Boolean = False
    Dim WaveformLoadId As Integer = 1
    Dim NoteWVPosEnd() As Double
    Dim LWAVRefreshId As Integer = 1
    Dim LBMPRefreshId As Integer = 1

    '----AutoSave Options
    Dim PreviousAutoSavedFileName As String = ""
    Dim AutoSaveInterval As Integer = 120000

    '----ErrorCheck Options
    Dim ErrorCheck As Boolean = True
    Dim ErrorJackBPM As Double = 255
    Dim ErrorJackTH As Double = 16
    Dim ErrorJackSpeed As Double = 60 * 4 / ErrorJackBPM / ErrorJackTH

    '----Header Options
    Dim hBMP(1295) As String
    Dim hWAV(1295) As String
    Dim hBMSCROLL(1295) As Long
    Dim hBPM(1295) As Long   'x10000
    Dim hSTOP(1295) As Long

    '----Grid Options
    Dim gSnap As Boolean = True
    Dim gShowGrid As Boolean = True 'Grid
    Dim gShowSubGrid As Boolean = True 'Sub
    Dim gShowBG As Boolean = True 'BG Color
    Dim gShowMeasureNumber As Boolean = True 'Measure Label
    Dim gShowVerticalLine As Boolean = True 'Vertical
    Dim gShowMeasureBar As Boolean = True 'Measure Barline
    Dim gShowC As Boolean = True 'Column Caption
    Dim gDivide As Integer = 16
    Dim gSub As Integer = 4
    Dim gSlash As Integer = 192
    Dim gxHeight As Single = 1.0!
    Dim gxWidth As Single = 1.0!
    Dim gWheel As Integer = 96
    Dim gPgUpDn As Integer = 384

    Dim gDisplayBGAColumn As Boolean = True
    Dim gSCROLL As Boolean = True
    Dim gSTOP As Boolean = True
    Dim gBPM As Boolean = True
    'Dim gA8 As Boolean = False
    Dim iPlayer As Integer = 0
    Dim gColumns As Integer = 46

    '----Visual Options
    Dim vo As New visualSettings()

    Public Sub setVO(ByVal xvo As visualSettings)
        vo = xvo
    End Sub

    '----Note Waveforms
    Structure WavSample
        Public WavL() As Single
        Public WavR() As Single
        Public SampleRate As Integer
        Public Duration As Single

        Public Sub New(xWavL() As Single,
                       xWavR() As Single,
                       xSampleRate As Integer,
                       xDuration As Single)
            WavL = xWavL
            WavR = xWavR
            SampleRate = xSampleRate
            Duration = xDuration
        End Sub
    End Structure

    '----Visual Override Options
    Structure ColorOverride
        Public Name As String
        Public Enabled As Boolean
        Public ColorOption As Integer
        Public RangeL As Integer
        Public RangeU As Integer
        Public NoteColor As Integer
        Public NoteColorU As Integer

        Public Sub New(ByVal xName As String,
                       ByVal xEnabled As Boolean,
                       ByVal xColorOption As Integer,
                       ByVal xRangeL As Integer,
                       ByVal xRangeU As Integer,
                       ByVal xNoteColor As Integer,
                       ByVal xNoteColorU As Integer)

            Name = xName
            Enabled = xEnabled
            ColorOption = xColorOption
            RangeL = xRangeL
            RangeU = xRangeU
            NoteColor = xNoteColor
            NoteColorU = xNoteColorU
        End Sub
    End Structure
    Dim COverrides(-1) As ColorOverride
    Dim COverridesColors(1295) As Color
    Dim COverridesSaveOption As Integer = 1

    '----Keybinding Options
    Structure Keybinding
        Public OpVar As Integer
        Public OpName As String
        Public Description As String
        Public Combo() As String
        Public Category As Integer

        Public Sub New(xOpVar As Integer, ByVal xOpName As String, Optional xDescription As String = "", Optional xKey() As String = Nothing, Optional xCategory As Integer = -1)
            OpVar = xOpVar
            OpName = xOpName
            Description = xDescription
            Combo = xKey
            Category = xCategory
        End Sub
    End Structure

    Public KbCategorySP As Integer = 1
    Public KbCategoryPMS As Integer = 2
    Public KbCategoryDP As Integer = 3
    Public KbCategoryBGM As Integer = 4
    Public KbCategoryAllMod As Integer = 10 ' AllMod meaning all modifiers included
    Public KbCategoryHidden As Integer = 0
    Public KbCategory() As Integer = {KbCategoryPMS, KbCategoryDP, KbCategorySP, KbCategoryBGM, KbCategoryAllMod, KbCategoryHidden, -1} ' Order matters
    Public KeybindingsInit() As Keybinding = { ' SP Note Assignments
                                       New Keybinding(0, "Move to A2", "Move note to 1P Lane 1", {"D1", "NumPad1"}, KbCategorySP),
                                       New Keybinding(1, "Move to A3", "Move note to 1P Lane 2", {"D2", "NumPad2"}, KbCategorySP),
                                       New Keybinding(2, "Move to A4", "Move note to 1P Lane 3", {"D3", "NumPad3"}, KbCategorySP),
                                       New Keybinding(3, "Move to A5", "Move note to 1P Lane 4", {"D4", "NumPad4"}, KbCategorySP),
                                       New Keybinding(4, "Move to A6", "Move note to 1P Lane 5", {"D5", "NumPad5"}, KbCategorySP),
                                       New Keybinding(5, "Move to A7", "Move note to 1P Lane 6", {"D6", "NumPad6"}, KbCategorySP),
                                       New Keybinding(6, "Move to A8", "Move note to 1P Lane 7", {"D7", "NumPad7"}, KbCategorySP),
                                       New Keybinding(7, "Move to A1", "Move note to 1P Scratch Lane", {"D8", "NumPad8"}, KbCategorySP),
                                                                                                                                        _ ' DP Note Assignments
                                       New Keybinding(10, "Move to D1", "Move note to 2P Lane 1", {"Q", "Ctrl+D1", "NumPad1"}, KbCategoryDP),
                                       New Keybinding(11, "Move to D2", "Move note to 2P Lane 2", {"W", "Ctrl+D2", "NumPad2"}, KbCategoryDP),
                                       New Keybinding(12, "Move to D3", "Move note to 2P Lane 3", {"E", "Ctrl+D3", "NumPad3"}, KbCategoryDP),
                                       New Keybinding(13, "Move to D4", "Move note to 2P Lane 4", {"R", "Ctrl+D4", "NumPad4"}, KbCategoryDP),
                                       New Keybinding(14, "Move to D5", "Move note to 2P Lane 5", {"T", "Ctrl+D5", "NumPad5"}, KbCategoryDP),
                                       New Keybinding(15, "Move to D6", "Move note to 2P Lane 6", {"Y", "Ctrl+D6", "NumPad6"}, KbCategoryDP),
                                       New Keybinding(16, "Move to D7", "Move note to 2P Lane 7", {"U", "Ctrl+D7", "NumPad7"}, KbCategoryDP),
                                       New Keybinding(17, "Move to D8", "Move note to 2P Scratch Lane", {"I", "Ctrl+D8", "NumPad8"}, KbCategoryDP),
                                                                                                                                                   _ ' PMS Note Assignments
                                       New Keybinding(20, "Move to P1", "Move note to PMS Lane 1", {"D1", "NumPad1"}, KbCategoryPMS),
                                       New Keybinding(21, "Move to P2", "Move note to PMS Lane 2", {"D2", "NumPad2"}, KbCategoryPMS),
                                       New Keybinding(22, "Move to P3", "Move note to PMS Lane 3", {"D3", "NumPad3"}, KbCategoryPMS),
                                       New Keybinding(23, "Move to P4", "Move note to PMS Lane 4", {"D4", "NumPad4"}, KbCategoryPMS),
                                       New Keybinding(24, "Move to P5", "Move note to PMS Lane 5", {"D5", "NumPad5"}, KbCategoryPMS),
                                       New Keybinding(25, "Move to P6", "Move note to PMS Lane 6", {"D6", "NumPad6"}, KbCategoryPMS),
                                       New Keybinding(26, "Move to P7", "Move note to PMS Lane 7", {"D7", "NumPad7"}, KbCategoryPMS),
                                       New Keybinding(27, "Move to P8", "Move note to PMS Lane 8", {"D8", "NumPad8"}, KbCategoryPMS),
                                       New Keybinding(28, "Move to P9", "Move note to PMS Lane 9", {"D9", "NumPad9"}, KbCategoryPMS),
                                                                                                                                     _ ' Miscellaneous BMS
                                       New Keybinding(100, "Move to BGM", "Move note to BGM Lane", {"D0", "NumPad0"}, KbCategoryBGM),
                                       New Keybinding(101, "Move to Template Position", "Move note to Template Position if available", {"P"}),
                                       New Keybinding(102, "Disable Vertical Moves", "Disable vertical moves", {"D"}),
                                       New Keybinding(103, "Snap to Grid", "Snap to grid", {"G"}),
                                                                                                  _
                                       New Keybinding(104, "Convert to Long Note", "→ Long Note", {"L"}),
                                       New Keybinding(105, "Convert to Short Note", "→ Short Note", {"S"}),
                                       New Keybinding(106, "Convert between Long and Short Note", "Long Note ↔ Short Note", {""}),
                                       New Keybinding(107, "Auto Long Note (By VPosition)", "Auto Long Note (By VPosition)", {""}),
                                       New Keybinding(108, "Auto Long Note (By Column)", "Auto Long Note (By Column)", {""}),
                                                                                                                             _
                                       New Keybinding(109, "Check Technical Error", "Check for technical errors such as impossible scratches in DP or impossible chords in PMS", {"Ctrl+Alt+E"}),
                                       New Keybinding(110, "Select Expansion Section", "Select #IF sections in the Expansion field", {"Ctrl+Alt+R"}),
                                                                                                                                                     _ ' Miscellaneous Editor
                                       New Keybinding(111, "Undo", "", {"Ctrl+Z"}),
                                       New Keybinding(112, "Redo", "", {"Ctrl+Y", "Ctrl+Alt+Z"}),
                                       New Keybinding(113, "Cut", "", {"Ctrl+X"}),
                                       New Keybinding(114, "Copy", "", {"Ctrl+C"}),
                                       New Keybinding(115, "Paste", "", {"Ctrl+V"}),
                                       New Keybinding(116, "Paste Pattern", "Apply pattern of the notes on the clipboard to the highlighted notes.", {"Ctrl+Shift+V"}),
                                       New Keybinding(117, "Select All", "", {"Ctrl+A"}),
                                       New Keybinding(118, "Select All with Hovered Note Label", "", {"Ctrl+Shift+A"}),
                                                                                                                       _ ' All Modifiers
                                       New Keybinding(200, "Move Note Up", "*HIDDEN*", {"Up"}, KbCategoryAllMod),
                                       New Keybinding(201, "Move Note Down", "*HIDDEN*", {"Down"}, KbCategoryAllMod),
                                       New Keybinding(202, "Move Note Left", "*HIDDEN*", {"Left"}, KbCategoryAllMod),
                                       New Keybinding(203, "Move Note Right", "*HIDDEN*", {"Right"}, KbCategoryAllMod),
                                       New Keybinding(204, "Insert Space/Define Measure", "*HIDDEN*", {"Insert"}, KbCategoryAllMod),
                                       New Keybinding(205, "Delete", "*HIDDEN*", {"Delete"}, KbCategoryAllMod),
                                       New Keybinding(206, "Home", "*HIDDEN*", {"Home"}, KbCategoryAllMod),
                                       New Keybinding(207, "End", "*HIDDEN*", {"End"}, KbCategoryAllMod),
                                       New Keybinding(208, "PageUp", "*HIDDEN*", {"PageUp"}, KbCategoryAllMod),
                                       New Keybinding(209, "PageDown", "*HIDDEN*", {"PageDown"}, KbCategoryAllMod),
                                       New Keybinding(210, "Next", "*HIDDEN*", {"Next"}, KbCategoryAllMod), ' ???
                                       New Keybinding(211, "TabBetweenFiles", "*HIDDEN*", {"Tab"}, KbCategoryAllMod),
                                       New Keybinding(212, "TabBetweenNotes", "*HIDDEN*", {"Capital"}, KbCategoryAllMod),
                                       New Keybinding(213, "Decrease Division", "*HIDDEN*", {"Oemcomma"}, KbCategoryAllMod),
                                       New Keybinding(214, "Increase Division", "*HIDDEN*", {"OemPeriod"}, KbCategoryAllMod),
                                                                                                                             _ ' Hidden / Experimental
                                       New Keybinding(215, "Set CGDivision", "*HIDDEN*", {"OemQuestion"}, KbCategoryHidden),
                                       New Keybinding(216, "Decrease CGHeight", "*HIDDEN*", {"OemMinus"}, KbCategoryHidden),
                                       New Keybinding(217, "Increase CGHeight", "*HIDDEN*", {"Oemplus"}, KbCategoryHidden),
                                       New Keybinding(218, "DecreaseCurrentWav", "*HIDDEN*", {"Subtract"}, KbCategoryHidden),
                                       New Keybinding(219, "IncreaseCurrentWav", "*HIDDEN*", {"Add"}, KbCategoryHidden),
                                       New Keybinding(220, "TBPreviewHighlighted_Click", "*EXPERIMENTAL*", {"Shift+F4"}, KbCategoryHidden)
                                       }

    Dim Keybindings() As Keybinding = CType(KeybindingsInit, Keybinding())

    Private Function ExcludeShortcut(ByVal xStr As String) As String
        Dim IParenthesis = xStr.LastIndexOf("(")
        If IParenthesis = -1 Then Return xStr.Replace("&", "")
        Return xStr.Substring(0, IParenthesis - 1).Replace("&", "")
    End Function

    Public Sub ReloadKeybindingNames(ByRef K() As Keybinding)
        For i = 0 To UBound(K)
            With K(i)
                Select Case .OpVar
                    Case 0 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A2") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "1")
                    Case 1 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A3") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "2")
                    Case 2 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A4") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "3")
                    Case 3 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A5") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "4")
                    Case 4 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A6") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "5")
                    Case 5 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A7") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "6")
                    Case 6 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A8") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "1P").Replace("{2}", "7")
                    Case 7 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "A1") : .Description = Strings.fopKeybinding.MoveToScratchDescription.Replace("{}", "1P")

                    Case 10 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D1") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "1")
                    Case 11 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D2") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "2")
                    Case 12 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D3") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "3")
                    Case 13 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D4") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "4")
                    Case 14 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D5") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "5")
                    Case 15 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D6") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "6")
                    Case 16 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D7") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "2P").Replace("{2}", "7")
                    Case 17 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "D8") : .Description = Strings.fopKeybinding.MoveToScratchDescription.Replace("{}", "2P")

                    Case 20 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P1") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "1")
                    Case 21 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P2") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "2")
                    Case 22 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P3") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "3")
                    Case 23 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P4") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "4")
                    Case 24 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P5") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "5")
                    Case 25 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P6") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "6")
                    Case 26 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P7") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "7")
                    Case 27 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P8") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "8")
                    Case 28 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "P9") : .Description = Strings.fopKeybinding.MoveToDescription.Replace("{1}", "PMS").Replace("{2}", "9")

                    Case 100 : .OpName = Strings.fopKeybinding.MoveTo.Replace("{}", "BGM") : .Description = Strings.fopKeybinding.MoveToBGMDescription
                    Case 101 : .OpName = Strings.fopKeybinding.MoveToTemplate : .Description = Strings.fopKeybinding.MoveToTemplateDescription
                    Case 102 : .OpName = ExcludeShortcut(CGDisableVertical.Text) : .Description = ""
                    Case 103 : .OpName = ExcludeShortcut(CGSnap.Text) : .Description = ""

                    Case 104 : .OpName = ExcludeShortcut(POBLong.Text) : .Description = ""
                    Case 105 : .OpName = ExcludeShortcut(POBShort.Text) : .Description = ""
                    Case 106 : .OpName = ExcludeShortcut(POBLongShort.Text) : .Description = ""
                    Case 107 : .OpName = ExcludeShortcut(POBAutoLongVPosition.Text) : .Description = ""
                    Case 108 : .OpName = ExcludeShortcut(POBAutoLongColumn.Text) : .Description = ""

                    Case 109 : .OpName = ExcludeShortcut(mnTechnicalErrorCheck.Text) : .Description = Strings.fopKeybinding.CheckTechnicalError
                    Case 110 : .OpName = ExcludeShortcut(BExpansion.Text) : .Description = Strings.fopKeybinding.SelectExpansionSection

                    Case 111 : .OpName = ExcludeShortcut(mnUndo.Text) : .Description = ""
                    Case 112 : .OpName = ExcludeShortcut(mnRedo.Text) : .Description = ""
                    Case 113 : .OpName = ExcludeShortcut(mnCut.Text) : .Description = ""
                    Case 114 : .OpName = ExcludeShortcut(mnCopy.Text) : .Description = ""
                    Case 115 : .OpName = ExcludeShortcut(mnPaste.Text) : .Description = ""
                    Case 116 : .OpName = ExcludeShortcut(mnPastePattern.Text) : .Description = Strings.fopKeybinding.PastePattern
                    Case 117 : .OpName = ExcludeShortcut(mnSelectAll.Text) : .Description = ""
                    Case 118 : .OpName = Strings.fopKeybinding.SelectHovered : .Description = ""

                End Select
            End With
        Next
    End Sub

    '----Preview Options
    Structure PlayerArguments
        Public Path As String
        Public aBegin As String
        Public aHere As String
        Public aStop As String
        Public Sub New(ByVal xPath As String, ByVal xBegin As String, ByVal xHere As String, ByVal xStop As String)
            Path = xPath
            aBegin = xBegin
            aHere = xHere
            aStop = xStop
        End Sub
    End Structure

    Public pArgsInit() As PlayerArguments = {New PlayerArguments("<apppath>\mBMplay.exe",
                                                             """<filename>""",
                                                             "-s <measure> ""<filename>""",
                                                             "-t"),
                                         New PlayerArguments("<apppath>\uBMplay.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S"),
                                         New PlayerArguments("<apppath>\o2play.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S")}
    Public pArgs() As PlayerArguments = CType(pArgsInit.Clone(), PlayerArguments())
    Public CurrentPlayer As Integer = 0
    Dim PreviewOnClick As Boolean = True
    Dim PreviewErrorCheck As Boolean = False
    Dim ClickStopPreview As Boolean = True
    Dim pTempFileNames() As String = {}

    Dim InternalPlayNotes() As Note
    Dim InternalPlayNoteIndex As Integer
    Dim InternalPlayTimerStart As Long = 0
    Dim InternalPlayTimerEnd As Long = 0
    Dim InternalPlayTimerCount As Long = 0
    Dim InternalPlayWav(1295) As AudioC

    '----Split Panel Options
    Dim PanelWidth() As Single = {0, 100, 0}
    Dim PanelHScroll() As Integer = {0, 0, 0}
    Dim PanelVScroll() As Double = {0, 0, 0}
    Dim spLock() As Boolean = {False, False, False}
    Dim spDiff() As Integer = {0, 0, 0}
    Dim PanelFocus As Integer = 1 '0 = Left, 1 = Middle, 2 = Right
    Dim spMouseOver As Integer = 1

    Dim AutoFocusMouseEnter As Boolean = False
    Dim FirstClickDisabled As Boolean = True
    Dim tempFirstMouseDown As Boolean = False

    Dim spMain() As Panel = {}

    '----#TOTAL Options
    Dim TotalOption As Integer = 0
    Dim TotalMultiplier As Double = 0.25
    Dim TotalGlobalMultiplier As Double = 1
    Dim TotalDecimal As Integer = 0
    Dim TotalDisplayValue As Boolean = True
    Dim TotalDisplayText As Boolean = True
    Dim TotalAutofill As Boolean = True

    '----Find Delete Replace Options
    Dim fdriMesL As Integer
    Dim fdriMesU As Integer
    Dim fdriLblL As Integer
    Dim fdriLblU As Integer
    Dim fdriValL As Integer
    Dim fdriValU As Integer
    Dim fdriCol() As Integer

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="xHPosition">Original horizontal position.</param>
    ''' <param name="xHSVal">HS.Value</param>


    Private Function HorizontalPositiontoDisplay(ByVal xHPosition As Integer, ByVal xHSVal As Integer) As Integer
        Return CInt(xHPosition * gxWidth - xHSVal * gxWidth)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="xVPosition">Original vertical position.</param>
    ''' <param name="xVSVal">VS.Value</param>
    ''' <param name="xTHeight">Height of the panel. (DisplayRectangle, but not ClipRectangle)</param>


    Private Function NoteRowToPanelHeight(ByVal xVPosition As Double, ByVal xVSVal As Integer, ByVal xTHeight As Integer) As Integer
        Return xTHeight - CInt((xVPosition + xVSVal) * gxHeight) - 1
    End Function

    Public Function MeasureAtDisplacement(ByVal xVPos As Double) As Integer
        'Return Math.Floor((xVPos + FloatTolerance) / 192)
        'Return Math.Floor(xVPos / 192)
        Dim xI1 As Integer
        For xI1 = 1 To 999
            If xVPos < MeasureBottom(xI1) Then Exit For
        Next
        Return xI1 - 1
    End Function

    Private Function GetMaxVPosition() As Double
        Return MeasureUpper(999)
    End Function

    Private Function SnapToGrid(ByVal xVPos As Double) As Double
        Dim xOffset As Double = MeasureBottom(MeasureAtDisplacement(xVPos))
        Dim xRatio As Double = 192.0R / gDivide
        Return Math.Floor((xVPos - xOffset) / xRatio) * xRatio + xOffset
    End Function

    Private Sub CalculateGreatestVPosition()
        'If K Is Nothing Then Exit Sub
        Dim xI1 As Integer
        GreatestVPosition = 0

        If NTInput Then
            For xI1 = UBound(Notes) To 0 Step -1
                If Notes(xI1).VPosition + Notes(xI1).Length > GreatestVPosition Then GreatestVPosition = Notes(xI1).VPosition + Notes(xI1).Length
            Next
        Else
            For xI1 = UBound(Notes) To 0 Step -1
                If Notes(xI1).VPosition > GreatestVPosition Then GreatestVPosition = Notes(xI1).VPosition
            Next
        End If

        Dim xI2 As Integer = -CInt(IIf(GreatestVPosition + 2000 > GetMaxVPosition(), GetMaxVPosition, GreatestVPosition + 2000))
        MainPanelScroll.Minimum = xI2
        LeftPanelScroll.Minimum = xI2
        RightPanelScroll.Minimum = xI2
    End Sub


    Private Sub SortByVPositionInsertion() 'Insertion Sort
        If UBound(Notes) <= 0 Then Exit Sub
        Dim xNote As Note
        Dim xI1 As Integer
        Dim xI2 As Integer
        For xI1 = 2 To UBound(Notes)
            xNote = Notes(xI1)
            For xI2 = xI1 - 1 To 1 Step -1
                If Notes(xI2).VPosition > xNote.VPosition Then
                    Notes(xI2 + 1) = Notes(xI2)
                    '                    If KMouseDown = xI2 Then KMouseDown += 1
                    If xI2 = 1 Then
                        Notes(xI2) = xNote
                        '                       If KMouseDown = xI1 Then KMouseDown = xI2
                        Exit For
                    End If
                Else
                    Notes(xI2 + 1) = xNote
                    '                    If KMouseDown = xI1 Then KMouseDown = xI2 + 1
                    Exit For
                End If
            Next
        Next

    End Sub

    Private Sub SortByVPositionQuick(ByVal xMin As Integer, ByVal xMax As Integer) 'Quick Sort
        Dim xNote As Note
        Dim iHi As Integer
        Dim iLo As Integer
        Dim xI1 As Integer

        ' If min >= max, the list contains 0 or 1 items so it is sorted.
        If xMin >= xMax Then Exit Sub

        ' Pick the dividing value.
        xI1 = CInt((xMax + xMin) / 2)
        xNote = Notes(xI1)

        ' Swap it to the front.
        Notes(xI1) = Notes(xMin)

        iLo = xMin
        iHi = xMax
        Do
            ' Look down from hi for a value < med_value.
            Do While Notes(iHi).VPosition > xNote.VPosition OrElse (Notes(iHi).VPosition = xNote.VPosition AndAlso Notes(iHi).ColumnIndex > xNote.ColumnIndex)
                iHi = iHi - 1
                If iHi <= iLo Then Exit Do
            Loop
            If iHi <= iLo Then
                Notes(iLo) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iLo) = Notes(iHi)

            ' Look up from lo for a value >= med_value.
            iLo = iLo + 1
            Do While Notes(iLo).VPosition < xNote.VPosition
                iLo = iLo + 1
                If iLo >= iHi Then Exit Do
            Loop
            If iLo >= iHi Then
                iLo = iHi
                Notes(iHi) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iHi) = Notes(iLo)
        Loop

        ' Sort the two sublists.
        SortByVPositionQuick(xMin, iLo - 1)
        SortByVPositionQuick(iLo + 1, xMax)
    End Sub

    Private Sub SortByVPositionQuick3(ByVal xMin As Integer, ByVal xMax As Integer)
        Dim xxMin As Integer
        Dim xxMax As Integer
        Dim xxMid As Integer
        Dim xNote As Note
        Dim xNoteMid As Note
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        'If xMax = 0 Then
        '    xMin = LBound(K1)
        '    xMax = UBound(K1)
        'End If
        xxMin = xMin
        xxMax = xMax
        xxMid = xMax - xMin + 1
        xI1 = CInt(Int(xxMid * Rnd())) + xMin
        xI2 = CInt(Int(xxMid * Rnd())) + xMin
        xI3 = CInt(Int(xxMid * Rnd())) + xMin
        If Notes(xI1).VPosition <= Notes(xI2).VPosition And Notes(xI2).VPosition <= Notes(xI3).VPosition Then
            xxMid = xI2
        Else
            If Notes(xI2).VPosition <= Notes(xI1).VPosition And Notes(xI1).VPosition <= Notes(xI3).VPosition Then
                xxMid = xI1
            Else
                xxMid = xI3
            End If
        End If
        xNoteMid = Notes(xxMid)
        Do
            Do While Notes(xxMin).VPosition < xNoteMid.VPosition And xxMin < xMax
                xxMin = xxMin + 1
            Loop
            Do While xNoteMid.VPosition < Notes(xxMax).VPosition And xxMax > xMin
                xxMax = xxMax - 1
            Loop
            If xxMin <= xxMax Then
                xNote = Notes(xxMin)
                Notes(xxMin) = Notes(xxMax)
                Notes(xxMax) = xNote
                xxMin = xxMin + 1
                xxMax = xxMax - 1
            End If
        Loop Until xxMin > xxMax
        If xxMax - xMin < xMax - xxMin Then
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
        Else
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
        End If
    End Sub


    Private Sub UpdateMeasureBottom()
        MeasureBottom(0) = 0.0#
        For xI1 As Integer = 0 To 998
            MeasureBottom(xI1 + 1) = MeasureBottom(xI1) + MeasureLength(xI1)
        Next
    End Sub

    Private Function PathIsValid(ByVal sPath As String) As Boolean
        Return File.Exists(sPath) Or Directory.Exists(sPath)
    End Function

    Public Function PrevCodeToReal(ByVal InitStr As String) As String
        Dim xPath As String = IIf(Not PathIsValid(FileName),
                                        IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                        ExcludeFileName(FileName)).ToString() _
                                        & "\"
        Dim xFileName As String = xPath & TempFileName
        Dim xFileNameT As String = xPath & FileNameTemplate
        Dim xMeasure As Integer = MeasureAtDisplacement(Math.Abs(PanelVScroll(PanelFocus)))
        Dim xStr As String = Replace(InitStr, "<apppath>", My.Application.Info.DirectoryPath)
        xStr = Replace(xStr, "<measure>", xMeasure.ToString())
        xStr = Replace(xStr, "<filename>", xFileName)
        xStr = Replace(xStr, "<filenameT>", xFileNameT)
        Return xStr
    End Function

    Private Sub SetFileName(ByVal xFileName As String)
        FileName = xFileName.Trim
        InitPath = ExcludeFileName(FileName)
        SetIsSaved(IsSaved)
    End Sub

    Private Sub SetIsSaved(ByVal isSaved As Boolean)
        'pttl.Refresh()
        'pIsSaved.Visible = Not xBool
        Dim xVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor &
                             IIf(My.Application.Info.Version.Build = 0, "", "." & My.Application.Info.Version.Build).ToString()
        Text = IIf(isSaved, "", "*").ToString() & GetFileName(FileName) & " - " & My.Application.Info.Title & " " & xVersion
        Me.IsSaved = isSaved

        If BMSFiles IsNot Nothing Then
            If Not isSaved Then
                BMSFiles(BMSFileIndex).TSB.Image = My.Resources.x16New
            Else
                BMSFiles(BMSFileIndex).TSB.Image = My.Resources.x16Blank
            End If
        End If
    End Sub

    Private Sub PreviewNote(ByVal xFileLocation As String, ByVal bStop As Boolean)
        If bStop Then
            Audio.StopPlaying()
            TimerPreviewNote.Enabled = False
        End If
        Audio.Play(xFileLocation)
    End Sub

    Private Sub AddNote(note As Note,
               Optional ByVal xSelected As Boolean = False,
               Optional ByVal OverWrite As Boolean = True,
               Optional ByVal SortAndUpdatePairing As Boolean = True)

        If note.VPosition < 0 Or note.VPosition >= GetMaxVPosition() Then Exit Sub

        Dim xI1 As Integer = 1

        If OverWrite Then
            Do While xI1 <= UBound(Notes)
                If Notes(xI1).VPosition = note.VPosition And
                    Notes(xI1).ColumnIndex = note.ColumnIndex Then
                    RemoveNote(xI1)
                Else
                    xI1 += 1
                End If
            Loop
        End If

        ReDim Preserve Notes(UBound(Notes) + 1)
        note.Selected = note.Selected And nEnabled(note.ColumnIndex)
        Notes(UBound(Notes)) = note

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub RemoveNote(ByVal I As Integer, Optional ByVal SortAndUpdatePairing As Boolean = True)
        KMouseOver = -1
        Dim xI2 As Integer

        If TBWavIncrease.Checked Then
            If Notes(I).Value = LWAV.SelectedIndex * 10000 Then
                DecreaseCurrentWAV()
            End If
        End If

        For xI2 = I + 1 To UBound(Notes)
            Notes(xI2 - 1) = Notes(xI2)
        Next
        ReDim Preserve Notes(UBound(Notes) - 1)
        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()

    End Sub

    Private Sub AddNotesFromClipboard(Optional ByVal xSelected As Boolean = True, Optional ByVal SortAndUpdatePairing As Boolean = True)
        Dim xStrLine() As String = Split(Clipboard.GetText, vbCrLf)

        Dim xI1 As Integer
        For xI1 = 0 To UBound(Notes)
            Notes(xI1).Selected = False
        Next

        Dim xVS As Integer = CInt(PanelVScroll(PanelFocus))
        Dim xTempVP As Double
        Dim NoteLen As Integer = Notes.Length

        If xStrLine(0) = "iBMSC Clipboard Data" Then
            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = CDbl(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(Notes.Length)
                    With Notes(UBound(Notes))
                        .ColumnIndex = CInt(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = CLng(xStrSub(2))
                        .LongNote = CBool(xStrSub(3))
                        .Hidden = CBool(xStrSub(4))
                        .Landmine = CBool(xStrSub(5))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If NTInput Then ConvertBMSE2NT(NoteLen)
        ElseIf xStrLine(0) = "iBMSC Clipboard Data xNT" Then
            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = CDbl(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(Notes.Length)
                    With Notes(UBound(Notes))
                        .ColumnIndex = CInt(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = CLng(xStrSub(2))
                        .Length = CDbl(xStrSub(3))
                        .Hidden = CBool(xStrSub(4))
                        .Landmine = CBool(xStrSub(5))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If Not NTInput Then ConvertNT2BMSE(NoteLen)
        ElseIf xStrLine(0) = "BMSE ClipBoard Object Data Format" Then
            'paste
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                ' zdr: holy crap this is obtuse
                Dim posStr = Mid(xStrLine(xI1), 5, 7)
                Dim vPos = CDbl(posStr) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)

                Dim bmsIdent = Mid(xStrLine(xI1), 1, 3)
                Dim lineCol = BMSEChannelToColumnIndex(bmsIdent)

                Dim Value = CDbl(Mid(xStrLine(xI1), 12)) * 10000

                Dim attribute = Mid(xStrLine(xI1), 4, 1)

                Dim validCol = Len(xStrLine(xI1)) > 11 And lineCol > 0
                Dim inRange = vPos >= 0 And vPos < GetMaxVPosition()
                If validCol And inRange Then
                    ReDim Preserve Notes(Notes.Length)

                    With Notes(UBound(Notes))
                        .ColumnIndex = lineCol
                        .VPosition = vPos
                        .Value = CLng(Value)
                        .LongNote = attribute = "2"
                        .Hidden = attribute = "1"
                        .Selected = xSelected And nEnabled(.ColumnIndex)
                    End With
                End If
            Next

            'convert
            If NTInput Then ConvertBMSE2NT(NoteLen)
        End If

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Function GetNotesFromClipboard() As Note()
        Dim xStrLine() As String = Split(Clipboard.GetText, vbCrLf)

        Dim xVS As Integer = CInt(PanelVScroll(PanelFocus))
        Dim xTempVP As Double
        Dim NoteLen As Integer = Notes.Length
        Dim NoteCB As Note()

        If xStrLine(0) = "iBMSC Clipboard Data" Then
            Dim NotesBK As Note() = CType(Notes.Clone(), Note())
            ReDim Preserve Notes(0)

            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = CDbl(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(Notes.Length)
                    With Notes(UBound(Notes))
                        .ColumnIndex = CInt(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = CLng(xStrSub(2))
                        .LongNote = CBool(xStrSub(3))
                        .Hidden = CBool(xStrSub(4))
                        .Landmine = CBool(xStrSub(5))
                    End With
                End If
            Next

            'convert
            ' Wow this code is fucking shit
            UpdatePairing()
            If NTInput Then ConvertBMSE2NT()
            ReDim NoteCB(UBound(Notes) - 1)

            For xI1 = 1 To UBound(Notes)
                NoteCB(xI1 - 1) = Notes(xI1)
            Next
            Notes = NotesBK
            Return NoteCB
        ElseIf xStrLine(0) = "iBMSC Clipboard Data xNT" Then
            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = CDbl(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(Notes.Length)
                    With Notes(UBound(Notes))
                        .ColumnIndex = CInt(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = CLng(xStrSub(2))
                        .Length = CDbl(xStrSub(3))
                        .Hidden = CBool(xStrSub(4))
                        .Landmine = CBool(xStrSub(5))
                    End With
                End If
            Next

            'convert
            If Not NTInput Then ConvertNT2BMSE(NoteLen)
            ReDim NoteCB(UBound(Notes) - NoteLen)
            For xI1 = NoteLen To UBound(Notes)
                NoteCB(xI1 - NoteLen) = Notes(xI1)
            Next
            ReDim Preserve Notes(NoteLen - 1)
            Return NoteCB
        ElseIf xStrLine(0) = "BMSE ClipBoard Object Data Format" Then
            Dim NotesBK As Note() = CType(Notes.Clone(), Note())
            If Not NTInput Then ReDim Preserve Notes(0)

            'paste
            For xI1 = 1 To UBound(xStrLine)
                ' zdr: holy crap this is obtuse
                Dim posStr = Mid(xStrLine(xI1), 5, 7)
                Dim vPos = CDbl(posStr) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)

                Dim bmsIdent = Mid(xStrLine(xI1), 1, 3)
                Dim lineCol = BMSEChannelToColumnIndex(bmsIdent)

                Dim Value = CDbl(Mid(xStrLine(xI1), 12)) * 10000

                Dim attribute = Mid(xStrLine(xI1), 4, 1)

                Dim validCol = Len(xStrLine(xI1)) > 11 And lineCol > 0
                Dim inRange = vPos >= 0 And vPos < GetMaxVPosition()
                If validCol And inRange Then
                    ReDim Preserve Notes(Notes.Length)

                    With Notes(UBound(Notes))
                        .ColumnIndex = lineCol
                        .VPosition = vPos
                        .Value = CLng(Value)
                        .LongNote = attribute = "2"
                        .Hidden = attribute = "1"
                    End With
                End If
            Next

            'convert
            UpdatePairing()
            If NTInput Then ConvertBMSE2NT()
            ReDim NoteCB(UBound(Notes) - 1)

            For xI1 = 1 To UBound(Notes)
                NoteCB(xI1 - 1) = Notes(xI1)
            Next
            Notes = NotesBK
            Return NoteCB
        End If

        Return Nothing
    End Function

    Private Sub CopyNotes(Optional ByVal Unselect As Boolean = True)
        Dim xStrAll As String = "iBMSC Clipboard Data" & IIf(NTInput, " xNT", "").ToString()
        Dim xI1 As Integer
        Dim MinMeasure As Double = 999

        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected And MeasureAtDisplacement(Notes(xI1).VPosition) < MinMeasure Then MinMeasure = MeasureAtDisplacement(Notes(xI1).VPosition)
        Next
        MinMeasure = MeasureBottom(CInt(MinMeasure))

        If Not NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Selected Then
                    xStrAll &= vbCrLf & Notes(xI1).ColumnIndex.ToString & " " &
                                       (Notes(xI1).VPosition - MinMeasure).ToString & " " &
                                        Notes(xI1).Value.ToString & " " &
                                   CInt(Notes(xI1).LongNote).ToString & " " &
                                   CInt(Notes(xI1).Hidden).ToString & " " &
                                   CInt(Notes(xI1).Landmine).ToString
                    Notes(xI1).Selected = Not Unselect
                End If
            Next

        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Selected Then
                    xStrAll &= vbCrLf & Notes(xI1).ColumnIndex.ToString & " " &
                                       (Notes(xI1).VPosition - MinMeasure).ToString & " " &
                                        Notes(xI1).Value.ToString & " " &
                                        Notes(xI1).Length.ToString & " " &
                                   CInt(Notes(xI1).Hidden).ToString & " " &
                                   CInt(Notes(xI1).Landmine).ToString
                    Notes(xI1).Selected = Not Unselect
                End If
            Next
        End If

        Clipboard.SetText(xStrAll)
    End Sub

    Private Sub RemoveNotes(Optional ByVal SortAndUpdatePairing As Boolean = True)
        If UBound(Notes) = 0 Then Exit Sub

        KMouseOver = -1
        Dim xI1 As Integer = 1
        Dim xI2 As Integer
        Dim xIC As Integer
        Dim xComCount(1295) As Integer
        Do
            If Notes(xI1).Selected AndAlso Not Notes(xI1).Ghost Then
                For xI2 = xI1 + 1 To UBound(Notes)
                    Notes(xI2 - 1) = Notes(xI2)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
                xI1 = 0
            End If
            xI1 += 1
        Loop While xI1 < UBound(Notes) + 1

        ' Check and remove comments
        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Comment Then
                xIC = CInt(Notes(xI1).Value / 10000)
                xComCount(xIC) += 1
            End If
        Next
        xI1 = 1
        Do While xI1 <= hCOMNum
            If xComCount(xI1) = 0 Then
                RemoveCommentLine(xI1)
                If xI1 = xComCount.Length Then
                    xComCount(xI1) = 0
                    Exit Do
                Else
                    xComCount(xI1) = xComCount(xI1 + 1)
                End If
                For xI2 = 1 To UBound(Notes)
                    If Notes(xI2).Value / 10000 > xI1 Then Notes(xI2).Value -= 10000
                Next
            Else
                xI1 += 1
            End If
        Loop

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Function EnabledColumnIndexToColumnArrayIndex(ByVal cEnabled As Integer) As Integer
        Dim xI1 As Integer = 0
        Do
            If xI1 >= gColumns Then Exit Do
            If Not nEnabled(xI1) Then cEnabled += 1
            If xI1 >= cEnabled Then Exit Do
            xI1 += 1
        Loop
        Return cEnabled
    End Function

    Private Function ColumnArrayIndexToEnabledColumnIndex(ByVal cReal As Integer) As Integer
        Dim xI1 As Integer
        For xI1 = 0 To cReal - 1
            If Not nEnabled(xI1) Then cReal -= 1
        Next
        Return cReal
    End Function

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If pTempFileNames IsNot Nothing Then
            For Each xStr As String In pTempFileNames
                IO.File.Delete(xStr)
            Next
        End If
        If PreviousAutoSavedFileName <> "" Then IO.File.Delete(PreviousAutoSavedFileName)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' TODO: SaveAll
        Dim xStrAll As String
        Dim xRandomFile As Boolean = GetFileName(FileName).StartsWith("___TempRandom")
        If Not IsSaved Then
            Dim xStr As String = Strings.Messages.SaveOnExit
            If e.CloseReason = CloseReason.WindowsShutDown Then xStr = Strings.Messages.SaveOnExit1
            If e.CloseReason = CloseReason.TaskManagerClosing Then xStr = Strings.Messages.SaveOnExit2

            Dim xResult As MsgBoxResult = MsgBox(xStr, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes AndAlso Not xRandomFile Then
                If ExcludeFileName(FileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = InitPath

                    If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then e.Cancel = True : Exit Sub
                    SetFileName(xDSave.FileName)
                End If
                xStrAll = SaveBMS()
                My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
                NewRecent(FileName)
                If BeepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then e.Cancel = True
        End If

        If xRandomFile Then
            xStrAll = SaveBMS()
            My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
            If BeepWhileSaved Then Beep()
        End If

        For i = 0 To UBound(BMSFiles)
            If Not BMSFiles(i).IsSaved Then
                Dim xStr As String = Strings.Messages.SaveOnExitOther
                Dim xResult As MsgBoxResult = MsgBox(xStr, MsgBoxStyle.OkCancel Or MsgBoxStyle.Question, Me.Text)
                If xResult = MsgBoxResult.Cancel Then e.Cancel = True : Exit Sub
                Exit For
            End If
        Next

        If Not e.Cancel AndAlso Not xRandomFile Then
            'If SaveTheme Then
            '    My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Skin.cff", SaveSkinCFF, False, System.Text.Encoding.Unicode)
            'Else
            '    My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Skin.cff", "", False, System.Text.Encoding.Unicode)
            'End If
            '
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\PlayerArgs.cff", SavePlayerCFF, False, System.Text.Encoding.Unicode)
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Config.cff", SaveCFF, False, System.Text.Encoding.Unicode)
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\PreConfig.cff", "", False, System.Text.Encoding.Unicode)
            Me.SaveSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml", False)
        End If
    End Sub

    Private Function FilterFileBySupported(ByVal xFile() As String, ByVal xFilter() As String) As String()
        Dim xPath(-1) As String
        For xI1 As Integer = 0 To UBound(xFile)
            If My.Computer.FileSystem.FileExists(xFile(xI1)) And Array.IndexOf(xFilter, Path.GetExtension(xFile(xI1))) <> -1 Then
                ReDim Preserve xPath(UBound(xPath) + 1)
                xPath(UBound(xPath)) = xFile(xI1)
            End If

            If My.Computer.FileSystem.DirectoryExists(xFile(xI1)) Then
                Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(xFile(xI1)).GetFiles()
                For Each xStr As FileInfo In xFileNames
                    If Array.IndexOf(xFilter, xStr.Extension) = -1 Then Continue For
                    ReDim Preserve xPath(UBound(xPath) + 1)
                    xPath(UBound(xPath)) = xStr.FullName
                Next
            End If
        Next

        Return xPath
    End Function

    Private Sub InitializeNewBMS()
        'ReDim K(0)
        'With K(0)
        ' .ColumnIndex = niBPM
        ' .VPosition = -1
        ' .LongNote = False
        ' .Selected = False
        ' .Value = 1200000
        'End With

        THTitle.Text = ""
        THArtist.Text = ""
        THGenre.Text = ""
        THBPM.Value = 120
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0
        CHRank.SelectedIndex = 3
        THPlayLevel.Text = ""
        THSubTitle.Text = ""
        THSubArtist.Text = ""
        THStageFile.Text = ""
        THBanner.Text = ""
        THBackBMP.Text = ""
        CHDifficulty.SelectedIndex = 0
        THExRank.Text = ""
        THTotal.Text = ""
        THComment.Text = ""
        'THLnType.Text = "1"
        CHLnObj.SelectedIndex = 0
        GhostMode = 0
        FileNameTemplate = ""
        ReDim hCOM(1295)
        hCOMNum = 0
        COverrides = Nothing
        ReDim wLWAV(1295)
        WaveformLoaded = False

        TExpansion.Text = ""
        ReDim ExpansionSplit(2)


        ReDim MeasureLength(999)
        LBeat.Items.Clear()
        For xI1 As Integer = 0 To 999
            MeasureLength(xI1) = 192.0R
            MeasureBottom(xI1) = xI1 * 192.0R
            LBeat.Items.Add(Add3Zeros(xI1) & ":  1 ( " & CInt(nBeatD.Value) & " / " & CInt(nBeatD.Value) & " )")
        Next
    End Sub

    Private Sub InitializeOpenBMS()
        CHPlayer.SelectedIndex = 0
        'THLnType.Text = ""
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DDFileName = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedFileExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DragLeave
        ReDim DDFileName(-1)
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragDrop

        ' If ClosingPopSave() Then Exit Sub

        ReDim DDFileName(-1)
        ' If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        ' Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedFileExtension)
        ' If xPath.Length > 0 Then

        SaveBMSStruct()
        Dim xProg As New fLoadFileProgress(xOrigPath, IsSaved)
        xProg.ShowDialog(Me)
        ' End If

        RefreshPanelAll()
    End Sub

    Private Sub setFullScreen(ByVal value As Boolean)
        If value Then
            If Me.WindowState = FormWindowState.Minimized Then Exit Sub

            Me.SuspendLayout()
            previousWindowPosition.Location = Me.Location
            previousWindowPosition.Size = Me.Size
            previousWindowState = Me.WindowState

            Me.WindowState = FormWindowState.Normal
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Me.WindowState = FormWindowState.Maximized
            ToolStripContainer1.TopToolStripPanelVisible = False

            Me.ResumeLayout()
            isFullScreen = True
        Else
            Me.SuspendLayout()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            ToolStripContainer1.TopToolStripPanelVisible = True
            Me.WindowState = FormWindowState.Normal

            Me.WindowState = previousWindowState
            If Me.WindowState = FormWindowState.Normal Then
                Me.Location = previousWindowPosition.Location
                Me.Size = previousWindowPosition.Size
            End If

            Me.ResumeLayout()
            isFullScreen = False
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F11
                setFullScreen(Not isFullScreen)
        End Select
    End Sub

    Private Sub Form1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Friend Sub ReadFile(ByVal xPath As String, Optional xAddRecent As Boolean = True)
        Select Case LCase(Path.GetExtension(xPath))
            Case ".bms", ".bme", ".bml", ".pms", ".txt"
                SetFileName(xPath)
                OpenBMS(My.Computer.FileSystem.ReadAllText(xPath, TextEncoding))
                ClearUndo()
                If xAddRecent Then NewRecent(xPath)
                SetIsSaved(True)

            Case ".sm"
                SetFileName(FileNameInit)
                If OpenSM(My.Computer.FileSystem.ReadAllText(xPath, TextEncoding)) Then Return
                InitPath = ExcludeFileName(xPath)
                ClearUndo()
                SetIsSaved(False)

            Case ".ibmsc"
                SetFileName("Imported_" & GetFileName(xPath))
                OpeniBMSC(xPath)
                InitPath = ExcludeFileName(xPath)
                If xAddRecent Then NewRecent(xPath)
                SetIsSaved(False)

            Case Else
                SetFileName(xPath)
                OpenBMS(My.Computer.FileSystem.ReadAllText(xPath, TextEncoding))
                ClearUndo()
                If xAddRecent Then NewRecent(xPath)
                SetIsSaved(True)

        End Select
    End Sub


    Public Function GCD(ByVal NumA As Double, ByVal NumB As Double) As Double
        Dim xNMax As Double = NumA
        Dim xNMin As Double = NumB
        If NumA < NumB Then
            xNMax = NumB
            xNMin = NumA
        End If
        Do While xNMin >= BMSGridLimit
            GCD = xNMax - Math.Floor(xNMax / xNMin) * xNMin
            xNMax = xNMin
            xNMin = GCD
        Loop
        GCD = xNMax
    End Function

    <DllImport("user32.dll")> Private Shared Function LoadCursorFromFile(ByVal fileName As String) As IntPtr
    End Function
    Public Shared Function ActuallyLoadCursor(ByVal path As String) As Cursor
        Return New Cursor(LoadCursorFromFile(path))
    End Function

    Private Sub Unload() Handles MyBase.Disposed
        Audio.Finalize()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SupportedAudioExtension = GetSupportedExtensions()

        'On Error Resume Next
        Me.TopMost = True
        Me.SuspendLayout()
        Me.Visible = False

        'POBMP.Dispose()
        'POBGA.Dispose()

        'Me.MaximizedBounds = Screen.GetWorkingArea(Me)
        'Me.Visible = False
        SetFileName(FileName)
        'Me.ShowCaption = False
        'SetWindowText(Me.Handle.ToInt32, FileName)

        'nBeatD.SelectedIndex = 4

        Try
            Dim xTempFileName As String = RandomFileName(".cur")
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeDown, False)
            Dim xDownCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeLeft, False)
            Dim xLeftCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeRight, False)
            Dim xRightCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            File.Delete(xTempFileName)

            POWAVResizer.Cursor = xDownCursor
            POBMPResizer.Cursor = xDownCursor
            POBeatResizer.Cursor = xDownCursor
            POExpansionResizer.Cursor = xDownCursor

            POptionsResizer.Cursor = xLeftCursor

            SpL.Cursor = xRightCursor
            SpR.Cursor = xLeftCursor
        Catch ex As Exception

        End Try

        spMain = New Panel() {PMainInL, PMainIn, PMainInR}

        sUndo(0) = New UndoRedo.NoOperation
        sUndo(1) = New UndoRedo.NoOperation
        sRedo(0) = New UndoRedo.NoOperation
        sRedo(1) = New UndoRedo.NoOperation
        sI = 0


        LWAV.Items.Clear()
        LBMP.Items.Clear()
        For xI1 = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ": " & hWAV(xI1))
            LBMP.Items.Add(C10to36(xI1) & ": " & hBMP(xI1))
        Next

        LWAV.SelectedIndex = 0
        LBMP.SelectedIndex = 0
        CHPlayer.SelectedIndex = 0

        CalculateGreatestVPosition()
        TBLangRefresh_Click(TBLangRefresh, Nothing)
        TBThemeRefresh_Click(TBThemeRefresh, Nothing)

        POHeaderPart2.Visible = False
        POGridPart2.Visible = False
        POWaveFormPart2.Visible = False

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml") Then
            LoadSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml")
            'Else
            '---- Settings for first-time start-up ---------------------------------------------------------------------------
            'Me.LoadLocale(My.Application.Info.DirectoryPath & "\Data\chs.Lang.xml")
            '-----------------------------------------------------------------------------------------------------------------
        End If
        'On Error GoTo 0

        InitializeBMSFiles()

        For xI = 0 To UBound(BMSFiles)
            TBTab.Items.Add(BMSFiles(xI).TSB)
        Next
        SetBMSFileIndex(BMSFileIndex)

        'pIsSaved.Visible = Not IsSaved

        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length <= 1 Then

            Dim xFiles() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath).GetFiles("AutoSave_*.IBMSC")
            If xFiles IsNot Nothing AndAlso xFiles.Length > 0 Then

                'Me.TopMost = True
                If MsgBox(Replace(Strings.Messages.RestoreAutosavedFile, "{}", xFiles.Length.ToString()), MsgBoxStyle.YesNo Or MsgBoxStyle.MsgBoxSetForeground) = MsgBoxResult.Yes Then
                    Dim xFUB = UBound(xFiles)
                    Dim xFilesPaths(xFUB) As String
                    For xI = 0 To xFUB
                        xFilesPaths(xI) = xFiles(xI).FullName
                    Next
                    SetBMSFileIndex(UBound(BMSFiles) - 1)
                    AddBMSFiles(xFilesPaths)
                    ' For Each xF As FileInfo In xFiles
                    '     'MsgBox(xF.FullName)
                    '     System.Diagnostics.Process.Start(Application.ExecutablePath, """" & xF.FullName & """")
                    ' Next
                End If

                For Each xF As FileInfo In xFiles
                    ReDim Preserve pTempFileNames(UBound(pTempFileNames) + 1)
                    pTempFileNames(UBound(pTempFileNames)) = xF.FullName
                Next
            End If
        End If


        tempResize = Me.WindowState
        Me.TopMost = False
        Me.WindowState = CType(tempResize, FormWindowState)

        Me.Visible = True

        Dim xStr() As String = Environment.GetCommandLineArgs

        If xStr.Length > 1 Then
            Dim xStrFiles(UBound(xStr) - 1) As String
            For xI = 0 To UBound(xStrFiles)
                xStrFiles(xI) = xStr(xI + 1)
            Next
            SetBMSFileIndex(UBound(BMSFiles) - 1)
            AddBMSFiles(xStrFiles)
        End If

        SetIsSaved(True)
        If PreloadBMSStruct Then SaveAllBMSStruct()

        IsApplicationInitializing = False

        If BMSFileIndex <> UBound(BMSFiles) Then ReadFile(BMSFiles(BMSFileIndex).Filename) Else TBNew_Click(Nothing, Nothing)

        LoadColorOverride(FileName)

        POStatusRefresh()
        Me.ResumeLayout()
    End Sub

    Private Sub UpdatePairing()
        Dim i As Integer, j As Integer

        If NTInput Then
            For i = 0 To UBound(Notes)
                Notes(i).LNPair = 0
                If Notes(i).Length < 0 Then Notes(i).Length = 0
                If Notes(i).ErrorType = 1 Then Continue For
                Notes(i).HasError = False
                Notes(i).ErrorType = 0
            Next

            For i = 1 To UBound(Notes)
                If Notes(i).Length <> 0 Then
                    For j = i + 1 To UBound(Notes)
                        If Notes(j).VPosition > Notes(i).VPosition + Notes(i).Length Then Exit For
                        If Notes(j).ColumnIndex = Notes(i).ColumnIndex Then Notes(j).HasError = True
                    Next
                Else
                    For j = i + 1 To UBound(Notes)
                        If Notes(j).VPosition > Notes(i).VPosition Then Exit For
                        If Notes(j).ColumnIndex = Notes(i).ColumnIndex Then Notes(j).HasError = True
                    Next

                    If Notes(i).Value \ 10000 = LnObj AndAlso Not IsColumnNumeric(Notes(i).ColumnIndex) Then
                        For j = i - 1 To 1 Step -1
                            If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                            If Notes(j).Hidden Then Continue For

                            If Notes(j).Length <> 0 OrElse Notes(j).Value \ 10000 = LnObj Then
                                Notes(i).HasError = True
                            Else
                                Notes(i).LNPair = j
                                Notes(j).LNPair = i
                            End If
                            Exit For
                        Next
                        If j = 0 Then
                            Notes(i).HasError = True
                        End If
                    End If
                End If
            Next

        Else
            For i = 0 To UBound(Notes)
                If Notes(i).ErrorType = 1 Then Continue For
                Notes(i).HasError = False
                Notes(i).ErrorType = 0
                Notes(i).LNPair = 0
            Next

            For i = 1 To UBound(Notes)

                If Notes(i).LongNote Then
                    'LongNote: If overlapping a note, then error.
                    '          Else if already matched by a LongNote below, then match it.
                    '          Otherwise match anything above.
                    '              If ShortNote above then error on above.
                    '          If nothing above then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        If Notes(j).VPosition = Notes(i).VPosition Then
                            Notes(i).HasError = True
                            Exit For
                        ElseIf Notes(j).LongNote And Notes(j).LNPair = i Then
                            Notes(i).LNPair = j
                            Exit For
                        Else
                            Exit For
                        End If
                    Next

                    If Not Notes(i).HasError AndAlso Notes(i).LNPair = 0 Then

                        For j = i + 1 To UBound(Notes)
                            If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                            Notes(i).LNPair = j
                            Notes(j).LNPair = i
                            If Not Notes(j).LongNote AndAlso Notes(j).Value \ 10000 <> LnObj Then
                                Notes(j).HasError = True
                            End If
                            Exit For
                        Next

                        If j = UBound(Notes) + 1 Then
                            Notes(i).HasError = True
                        End If
                    End If

                ElseIf Notes(i).Value \ 10000 = LnObj And
                    Not IsColumnNumeric(Notes(i).ColumnIndex) Then
                    'LnObj: Match anything below.
                    '           If matching a LongNote not matching back, then error on below.
                    '           If overlapping a note, then error.
                    '           If mathcing a LnObj below, then error on below.
                    '       If nothing below, then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(i).ColumnIndex <> Notes(j).ColumnIndex Then Continue For
                        If Notes(j).LNPair <> 0 And Notes(j).LNPair <> i Then
                            Notes(j).HasError = True
                        End If
                        Notes(i).LNPair = j
                        Notes(j).LNPair = i
                        If Notes(i).VPosition = Notes(j).VPosition Then
                            Notes(i).HasError = True
                        End If
                        If Notes(j).Value \ 10000 = LnObj Then
                            Notes(j).HasError = True
                        End If
                        Exit For
                    Next

                    If j = 0 Then
                        Notes(i).HasError = True
                    End If

                Else
                    'ShortNote: If overlapping a note, then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(j).VPosition < Notes(i).VPosition Then Exit For
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        Notes(i).HasError = True
                        Exit For
                    Next

                End If
            Next


        End If

        Dim currentMS = 0.0#
        Dim currentBPM = Notes(0).Value / 10000
        Dim currentBPMVPosition = 0.0#
        For i = 1 To UBound(Notes)
            If Notes(i).ColumnIndex = niBPM Then
                currentMS += (Notes(i).VPosition - currentBPMVPosition) / currentBPM * 1250
                currentBPM = Notes(i).Value / 10000
                currentBPMVPosition = Notes(i).VPosition
            End If
            'K(i).TimeOffset = currentMS + (K(i).VPosition - currentBPMVPosition) / currentBPM * 1250
        Next
    End Sub

    Private Sub CheckTechnicalError(sender As Object, e As EventArgs)
        For xIN = 1 To UBound(Notes)
            If Notes(xIN).ErrorType = 1 Then
                Notes(xIN).ErrorType = 0
                Notes(xIN).HasError = False
            End If
        Next
        If gXKeyMode = "PMS" Then CheckErrorImpossibleChord()
        If gXKeyMode = "DP" Then CheckErrorImpossibleScratch()
        CheckErrorJack()
    End Sub

    Private Sub CheckErrorImpossibleChord()
        For xIN = 1 To UBound(Notes)
            If Not gXKeyCol.Contains(Notes(xIN).ColumnIndex) Or Not IsPlayableNote(xIN) Then Continue For
            Dim xIColArray() As Integer = {FindColumnNumber(Notes(xIN).ColumnIndex)} ' Array of columns
            Dim xINArray() As Integer = {xIN} ' Array of notes in the columns
            Dim xIComp = xIN - 1
            Do While xIComp > 0 ' If note is in range, add to xIColArray and xINArray
                If GetTimeFromVPosition(Notes(xIN).VPosition) - GetTimeFromVPosition(Notes(xIComp).VPosition) > ErrorJackSpeed Then Exit Do
                If Not gXKeyCol.Contains(Notes(xIComp).ColumnIndex) Or Not IsPlayableNote(xIComp) Then xIComp -= 1 : Continue Do
                ReDim Preserve xIColArray(xIColArray.Length)
                xIColArray(UBound(xIColArray)) = FindColumnNumber(Notes(xIComp).ColumnIndex)
                ReDim Preserve xINArray(xINArray.Length)
                xINArray(UBound(xINArray)) = xIComp
                xIComp -= 1
            Loop

            Dim xIColArray2 = FindLNColumnsAtVPosition(Notes(xIN).VPosition) ' Account for LNs
            ReDim Preserve xIColArray(UBound(xIColArray) + xIColArray2.Length)
            For xINCol = 0 To UBound(xIColArray2)
                xIColArray(UBound(xIColArray) - xINCol) = FindColumnNumber(xIColArray2(xINCol))
            Next
            Array.Sort(xIColArray)

            Dim posLHand As Integer = 0 ' Position of left hand based on the leftmost note, i.e. 1 - 4
            Dim posRHand As Integer = 0 ' Position of right hand based on the leftmost note, i.e. 4 - 7
            For Each xIN2 In xIColArray
                If posLHand = 0 Then ' Assign to left hand
                    posLHand = xIN2
                    Continue For
                ElseIf posLHand + 2 < xIN2 AndAlso posRHand = 0 Then ' Assign to right hand
                    posRHand = xIN2
                ElseIf posRHand <> 0 AndAlso posRHand + 2 < xIN2 Then ' If right hand is assigned and note is out of hand range
                    For Each xINAssign In xINArray ' Assign error type 1 to every note in xINArray
                        Notes(xINAssign).HasError = True
                        Notes(xINAssign).ErrorType = 1
                    Next
                End If
            Next
        Next
    End Sub

    Private Sub CheckErrorImpossibleScratch()
        Dim xKArrayLFull() As Integer = {niA2, niA3, niA4, niA5, niA6, niA7, niA8}
        Dim xKArrayRFull() As Integer = {niD1, niD2, niD3, niD4, niD5, niD6, niD7}
        Dim xKArrayL() As Integer = {niA5, niA6, niA7, niA8}
        Dim xKArrayR() As Integer = {niD1, niD2, niD3, niD4}
        Dim xScrL As Integer = niA1
        Dim xScrR As Integer = niD8
        For xIN = 1 To UBound(Notes) ' Check for notes near scratch notes
            If (Not Notes(xIN).ColumnIndex = xScrL AndAlso Not Notes(xIN).ColumnIndex = xScrR) Or Not IsPlayableNote(xIN) Then Continue For
            Dim xKArray() = xKArrayL
            Dim xKArrayFull() = xKArrayLFull
            If Notes(xIN).ColumnIndex = xScrR Then xKArray = xKArrayR : xKArrayFull = xKArrayRFull

            Dim xIColArray(-1) As Integer
            Dim xIComp = xIN - 1
            Do While xIComp > 0 ' If note is in range, add to xIColArray and xINArray
                If GetTimeFromVPosition(Notes(xIN).VPosition) - GetTimeFromVPosition(Notes(xIComp).VPosition) > ErrorJackSpeed Then Exit Do
                If Not xKArray.Contains(Notes(xIComp).ColumnIndex) Or Not IsPlayableNote(xIComp) Then xIComp -= 1 : Continue Do
                ReDim Preserve xIColArray(xIColArray.Length)
                xIComp -= 1
            Loop
            xIComp = xIN + 1
            Do While xIComp <= UBound(Notes) ' If note is in range, add to xIColArray and xINArray
                If GetTimeFromVPosition(Notes(xIComp).VPosition) - GetTimeFromVPosition(Notes(xIN).VPosition) > ErrorJackSpeed Then Exit Do
                If Not xKArray.Contains(Notes(xIComp).ColumnIndex) Or Not IsPlayableNote(xIComp) Then xIComp += 1 : Continue Do
                ReDim Preserve xIColArray(xIColArray.Length)
                xIComp += 1
            Loop

            Dim xIColArray2 = FindLNColumnsAtVPosition(Notes(xIN).VPosition) ' Account for LNs
            Dim xIColArray3(-1) As Integer
            For xIComp = 0 To UBound(xIColArray2)
                If xKArrayFull.Contains(xIColArray2(xIComp)) Then ReDim Preserve xIColArray3(xIColArray3.Length)
            Next

            If xIColArray.Length + xIColArray3.Length > 0 Then
                Notes(xIN).HasError = True
                Notes(xIN).ErrorType = 1
            End If
        Next
    End Sub

    Private Sub CheckErrorJack()
        For xIN = 1 To UBound(Notes)
            If Not IsPlayableNote(xIN) Then Continue For
            Dim xIComp = xIN - 1
            Do While xIComp > 0
                If Notes(xIComp).VPosition = Notes(xIN).VPosition Then xIComp -= 1 : Continue Do
                If GetTimeFromVPosition(Notes(xIN).VPosition) - GetTimeFromVPosition(Notes(xIComp).VPosition) > ErrorJackSpeed Then Exit Do
                If gXKeyCol.Contains(Notes(xIN).ColumnIndex) AndAlso Notes(xIN).ColumnIndex = Notes(xIComp).ColumnIndex Then
                    Notes(xIN).HasError = True
                    Notes(xIN).ErrorType = 1
                    Exit Do
                End If
                xIComp -= 1
            Loop
        Next
    End Sub

    Private Function FindColumnNumber(xI As Integer) As Integer
        For i = 0 To UBound(gXKeyCol)
            If xI = gXKeyCol(i) Then Return i + 1
        Next
        Return 0
    End Function

    Private Function FindLNColumnsAtVPosition(ByVal VPos As Double) As Integer()
        ' NTInput
        If Not NTInput Then ConvertBMSE2NT()

        Dim xN = From note In Notes
                 Where note.Length > 0 AndAlso note.VPosition <= VPos AndAlso VPos <= note.VPosition + note.Length AndAlso gXKeyCol.Contains(note.ColumnIndex)
                 Select note

        Dim col(xN.Count - 1) As Integer
        For i = 0 To UBound(col)
            col(i) = xN(i).ColumnIndex
        Next
        If Not NTInput Then ConvertNT2BMSE()
        Return col
    End Function

    Private Function IsPlayableNote(ByVal xI As Integer) As Boolean
        Return Not (Notes(xI).Hidden Or Notes(xI).Landmine Or Notes(xI).Comment)
    End Function

    Public Sub ExceptionSave(ByVal Path As String)
        SaveiBMSC(Path)
    End Sub

    ''' <summary>
    ''' True if pressed cancel. False elsewise.
    ''' </summary>
    ''' <returns>True if pressed cancel. False elsewise.</returns>

    Private Function ClosingPopSave() As Boolean
        If Not IsSaved Then
            Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveOnExit, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes Then
                If ExcludeFileName(FileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = InitPath

                    If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Return True
                    SetFileName(xDSave.FileName)
                End If
                Dim xStrAll As String = SaveBMS()
                My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
                NewRecent(FileName)
                If BeepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then Return True
        End If
        Return False
    End Function

    Private Sub TBNew_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles TBNew.Click, mnNew.Click
        SaveBMSStruct()
        If BMSFiles(UBound(BMSFiles)).Filename <> FileNameInit Then AddBMSFile(FileNameInit)
        SetBMSFileIndex(UBound(BMSFiles))

        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        ' If ClosingPopSave() Then Exit Sub

        ClearUndo()
        SetFileName(FileNameInit)
        InitializeNewBMS()

        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        ReDim hBMSCROLL(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With
        THBPM.Value = 120

        LWAVRefresh()
        LWAV.SelectedIndex = 0
        LBMPRefresh()
        LBMP.SelectedIndex = 0

        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        CalculateTotalPlayableNotes(False)
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBNewC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles TBNewC.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        ClearUndo()

        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        ReDim hBMSCROLL(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With
        THBPM.Value = 120

        SetFileName(FileNameInit)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        If MsgBox("Please copy your code to clipboard and click OK.", MsgBoxStyle.OkCancel, "Create from code") = MsgBoxResult.Cancel Then Exit Sub
        OpenBMS(Clipboard.GetText)
    End Sub

    Private Sub TBOpen_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpen.ButtonClick, mnOpen.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        ' If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                            Strings.FileType.BMS & "|*.bms|" &
                            Strings.FileType.BME & "|*.bme|" &
                            Strings.FileType.BML & "|*.bml|" &
                            Strings.FileType.PMS & "|*.pms|" &
                            Strings.FileType.TXT & "|*.txt|" &
                            Strings.FileType._all & "|*.*"
        xDOpen.DefaultExt = "bms"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()
        xDOpen.Multiselect = True

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        SaveBMSStruct()
        Dim xProg As New fLoadFileProgress(xDOpen.FileNames, IsSaved)
        xProg.ShowDialog(Me)

        RefreshPanelAll()
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBOpenTemplate_ButtonClick(sender As Object, e As EventArgs) Handles mnOpenTemplate.Click
        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                            Strings.FileType.BMS & "|*.bms|" &
                            Strings.FileType.BME & "|*.bme|" &
                            Strings.FileType.BML & "|*.bml|" &
                            Strings.FileType.PMS & "|*.pms|" &
                            Strings.FileType.TXT & "|*.txt|" &
                            Strings.FileType._all & "|*.*"
        xDOpen.DefaultExt = "bms"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        NotesTemplate = OpenBMSFunc(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, TextEncoding))
        FileNameTemplate = GetFileName(xDOpen.FileName)
    End Sub

    Private Sub TBImportIBMSC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBImportIBMSC.Click, mnImportIBMSC.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Return

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType.IBMSC & "|*.ibmsc"
        xDOpen.DefaultExt = "ibmsc"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Return
        SetFileName("Imported_" & GetFileName(xDOpen.FileName))
        InitPath = ExcludeFileName(xDOpen.FileName)
        OpeniBMSC(xDOpen.FileName)
        NewRecent(xDOpen.FileName)
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBImportSM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBImportSM.Click, mnImportSM.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType.SM & "|*.sm"
        xDOpen.DefaultExt = "sm"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        If OpenSM(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, TextEncoding)) Then Exit Sub
        SetFileName(FileNameInit)
        InitPath = ExcludeFileName(xDOpen.FileName)
        ClearUndo()
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBSave_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSave.ButtonClick, mnSave.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1

        If ExcludeFileName(FileName) = "" Then
            Dim xDSave As New SaveFileDialog
            xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                            Strings.FileType.BMS & "|*.bms|" &
                            Strings.FileType.BME & "|*.bme|" &
                            Strings.FileType.BML & "|*.bml|" &
                            Strings.FileType.PMS & "|*.pms|" &
                            Strings.FileType.TXT & "|*.txt|" &
                            Strings.FileType._all & "|*.*"
            xDSave.DefaultExt = "bms"
            xDSave.InitialDirectory = InitPath

            If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
            SetFileName(xDSave.FileName)
            NewRecent(FileName)

            BMSFiles(BMSFileIndex).Filename = FileName
            BMSFiles(BMSFileIndex).TSB.Text = GetFileName(FileName)
        End If
        Dim xStrAll As String = SaveBMS()
        My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
        If BeepWhileSaved Then Beep()
    End Sub

    Private Sub TBSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSaveAs.Click, mnSaveAs.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1

        Dim xDSave As New SaveFileDialog
        xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                        Strings.FileType.BMS & "|*.bms|" &
                        Strings.FileType.BME & "|*.bme|" &
                        Strings.FileType.BML & "|*.bml|" &
                        Strings.FileType.PMS & "|*.pms|" &
                        Strings.FileType.TXT & "|*.txt|" &
                        Strings.FileType._all & "|*.*"
        xDSave.DefaultExt = "bms"
        xDSave.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()

        If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        SetFileName(xDSave.FileName)
        Dim xStrAll As String = SaveBMS()
        My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        SetFileName(FileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        BMSFiles(BMSFileIndex).Filename = FileName
        BMSFiles(BMSFileIndex).TSB.Text = GetFileName(FileName)

        If BeepWhileSaved Then Beep()
    End Sub

    Private Sub TBExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBExport.Click, mnExport.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1

        Dim xDSave As New SaveFileDialog
        xDSave.Filter = Strings.FileType.IBMSC & "|*.ibmsc"
        xDSave.DefaultExt = "ibmsc"
        xDSave.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()
        If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        SaveiBMSC(xDSave.FileName)
        'My.Computer.FileSystem.WriteAllText(xDSave.FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        If BeepWhileSaved Then Beep()
    End Sub

    Private Sub VSGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPanelScroll.GotFocus, LeftPanelScroll.GotFocus, RightPanelScroll.GotFocus
        Dim VScrollS As VScrollBar = CType(sender, VScrollBar)
        PanelFocus = CInt(VScrollS.Tag)
        spMain(PanelFocus).Focus()
    End Sub

    Private Sub VSValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPanelScroll.ValueChanged, LeftPanelScroll.ValueChanged, RightPanelScroll.ValueChanged
        Dim VScrollS As VScrollBar = CType(sender, VScrollBar)

        Dim iI As Integer = CInt(VScrollS.Tag)

        ' az: We got a wheel event when we're zooming in/out
        If My.Computer.Keyboard.CtrlKeyDown Or My.Computer.Keyboard.ShiftKeyDown Then
            VScrollS.Value = VSValue ' Undo the scroll
            Exit Sub
        End If

        If iI = PanelFocus And Not LastMouseDownLocation = New Point(-1, -1) And Not VSValue = -1 Then LastMouseDownLocation.Y += (VSValue - VScrollS.Value) * gxHeight
        PanelVScroll(iI) = VScrollS.Value

        If spLock((iI + 1) Mod 3) Then
            Dim xVS As Integer = CInt(PanelVScroll(iI) + spDiff(iI))
            If xVS > 0 Then xVS = 0
            If xVS < MainPanelScroll.Minimum Then xVS = MainPanelScroll.Minimum
            Select Case iI
                Case 0 : MainPanelScroll.Value = xVS
                Case 1 : RightPanelScroll.Value = xVS
                Case 2 : LeftPanelScroll.Value = xVS
            End Select
        End If

        If spLock((iI + 2) Mod 3) Then
            Dim xVS As Integer = CInt(PanelVScroll(iI) - spDiff((iI + 2) Mod 3))
            If xVS > 0 Then xVS = 0
            If xVS < MainPanelScroll.Minimum Then xVS = MainPanelScroll.Minimum
            Select Case iI
                Case 0 : RightPanelScroll.Value = xVS
                Case 1 : LeftPanelScroll.Value = xVS
                Case 2 : MainPanelScroll.Value = xVS
            End Select
        End If

        spDiff(iI) = CInt(PanelVScroll((iI + 1) Mod 3) - PanelVScroll(iI))
        spDiff((iI + 2) Mod 3) = CInt(PanelVScroll(iI) - PanelVScroll((iI + 2) Mod 3))

        VSValue = VScrollS.Value
        RefreshPanel(iI, spMain(iI).DisplayRectangle)
    End Sub

    Private Sub cVSLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cVSLockL.CheckedChanged, cVSLock.CheckedChanged, cVSLockR.CheckedChanged
        Dim CBS As CheckBox = CType(sender, CheckBox)
        Dim iI As Integer = CInt(CBS.Tag)
        spLock(iI) = CBS.Checked
        If Not spLock(iI) Then Return

        spDiff(iI) = CInt(PanelVScroll((iI + 1) Mod 3) - PanelVScroll(iI))
        spDiff((iI + 2) Mod 3) = CInt(PanelVScroll(iI) - PanelVScroll((iI + 2) Mod 3))

        'POHeaderB.Text = spDiff(0) & "_" & spDiff(1) & "_" & spDiff(2)
    End Sub

    Private Sub HSGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles HS.GotFocus, HSL.GotFocus, HSR.GotFocus
        Dim HScrollS As HScrollBar = CType(sender, HScrollBar)
        PanelFocus = CInt(HScrollS.Tag)
        spMain(PanelFocus).Focus()
    End Sub

    Private Sub HSValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles HS.ValueChanged, HSL.ValueChanged, HSR.ValueChanged
        Dim HScrollS As HScrollBar = CType(sender, HScrollBar)
        Dim iI As Integer = CInt(HScrollS.Tag)
        If Not LastMouseDownLocation = New Point(-1, -1) And Not HSValue = -1 Then LastMouseDownLocation.X += (HSValue - HScrollS.Value) * gxWidth
        PanelHScroll(iI) = HScrollS.Value
        HSValue = HScrollS.Value
        RefreshPanel(iI, spMain(iI).DisplayRectangle)
    End Sub

    Private Sub TBSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSelect.Click, mnSelect.Click
        TBSelect.Checked = True
        TBWrite.Checked = False
        TBTimeSelect.Checked = False
        mnSelect.Checked = True
        mnWrite.Checked = False
        mnTimeSelect.Checked = False

        FStatus2.Visible = False
        FStatus.Visible = True

        ShouldDrawTempNote = False
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0

        vSelStart = MeasureBottom(MeasureAtDisplacement(-PanelVScroll(PanelFocus)) + 1)
        vSelLength = 0

        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBWrite.Click, mnWrite.Click
        TBSelect.Checked = False
        TBWrite.Checked = True
        TBTimeSelect.Checked = False
        mnSelect.Checked = False
        mnWrite.Checked = True
        mnTimeSelect.Checked = False

        FStatus2.Visible = False
        FStatus.Visible = True

        ShouldDrawTempNote = True
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0

        vSelStart = MeasureBottom(MeasureAtDisplacement(-PanelVScroll(PanelFocus)) + 1)
        vSelLength = 0

        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBPostEffects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBTimeSelect.Click, mnTimeSelect.Click
        TBSelect.Checked = False
        TBWrite.Checked = False
        TBTimeSelect.Checked = True
        mnSelect.Checked = False
        mnWrite.Checked = False
        mnTimeSelect.Checked = True

        FStatus.Visible = False
        FStatus2.Visible = True

        vSelMouseOverLine = 0
        ShouldDrawTempNote = False
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0
        ValidateSelection()

        Dim xI1 As Integer
        For xI1 = 0 To UBound(Notes)
            Notes(xI1).Selected = False
        Next
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub CGHeight_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight.ValueChanged
        gxHeight = CSng(CGHeight.Value)
        CGHeight2.Value = CInt(IIf(CGHeight.Value * 4 < CGHeight2.Maximum, CDec(CGHeight.Value * 4), CGHeight2.Maximum))
        RefreshPanelAll()
    End Sub

    Private Sub CGHeight2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight2.Scroll
        CGHeight.Value = CDec(CGHeight2.Value / 4)
    End Sub

    Private Sub CGWidth_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth.ValueChanged
        gxWidth = CSng(CGWidth.Value)
        CGWidth2.Value = CInt(IIf(CGWidth.Value * 4 < CGWidth2.Maximum, CDec(CGWidth.Value * 4), CGWidth2.Maximum))

        HS.LargeChange = CInt(PMainIn.Width / gxWidth)
        If HS.Value > HS.Maximum - HS.LargeChange + 1 Then HS.Value = HS.Maximum - HS.LargeChange + 1
        HSL.LargeChange = CInt(PMainInL.Width / gxWidth)
        If HSL.Value > HSL.Maximum - HSL.LargeChange + 1 Then HSL.Value = HSL.Maximum - HSL.LargeChange + 1
        HSR.LargeChange = CInt(PMainInR.Width / gxWidth)
        If HSR.Value > HSR.Maximum - HSR.LargeChange + 1 Then HSR.Value = HSR.Maximum - HSR.LargeChange + 1

        RefreshPanelAll()
    End Sub

    Private Sub CGWidth2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth2.Scroll
        CGWidth.Value = CDec(CGWidth2.Value / 4)
    End Sub

    Private Sub CGDivide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDivide.ValueChanged
        gDivide = CInt(CGDivide.Value)
        RefreshPanelAll()
    End Sub
    Private Sub CGSub_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSub.ValueChanged
        gSub = CInt(CGSub.Value)
        RefreshPanelAll()
    End Sub
    Private Sub BGSlash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BGSlash.Click
        Dim xd As Integer = CInt(InputBox(Strings.Messages.PromptSlashValue, , gSlash.ToString()))
        If xd = 0 Then Exit Sub
        If xd > CGDivide.Maximum Then xd = CInt(CGDivide.Maximum)
        If xd < CGDivide.Minimum Then xd = CInt(CGDivide.Minimum)
        gSlash = xd
    End Sub


    Private Sub CGSnap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSnap.CheckedChanged
        gSnap = CGSnap.Checked
        RefreshPanelAll()
    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim xI1 As Integer

        Select Case PanelFocus
            Case 0
                With LeftPanelScroll
                    xI1 = CInt(.Value + (tempY / 5) / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSL
                    xI1 = CInt(.Value + (tempX / 10) / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 1
                With MainPanelScroll
                    xI1 = CInt(.Value + (tempY / 5) / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HS
                    xI1 = CInt(.Value + (tempX / 10) / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 2
                With RightPanelScroll
                    xI1 = CInt(.Value + (tempY / 5) / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSR
                    xI1 = CInt(.Value + (tempX / 10) / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
        End Select

        Dim xMEArgs As New System.Windows.Forms.MouseEventArgs(Windows.Forms.MouseButtons.Left, 0, MouseMoveStatus.X, MouseMoveStatus.Y, 0)
        PMainInMouseMove(spMain(PanelFocus), xMEArgs)

    End Sub

    Private Sub TimerMiddle_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMiddle.Tick
        If Not MiddleButtonClicked Then TimerMiddle.Enabled = False : Return

        Dim xI1 As Integer

        Select Case PanelFocus
            Case 0
                With LeftPanelScroll
                    xI1 = CInt(.Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSL
                    xI1 = CInt(.Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 1
                With MainPanelScroll
                    xI1 = CInt(.Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HS
                    xI1 = CInt(.Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 2
                With RightPanelScroll
                    xI1 = CInt(.Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight)
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSR
                    xI1 = CInt(.Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth)
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
        End Select

        Dim xMEArgs As New System.Windows.Forms.MouseEventArgs(Windows.Forms.MouseButtons.Left, 0, MouseMoveStatus.X, MouseMoveStatus.Y, 0)
        PMainInMouseMove(spMain(PanelFocus), xMEArgs)
    End Sub

    Private Sub ValidateWAVListView()
        Try
            Dim xRect As Rectangle = LWAV.GetItemRectangle(LWAV.SelectedIndex)
            If xRect.Top + xRect.Height > LWAV.DisplayRectangle.Height Then SendMessage(LWAV.Handle, &H115, 1, 0)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ValidateBMPListView()
        Try
            Dim xRect As Rectangle = LBMP.GetItemRectangle(LBMP.SelectedIndex)
            If xRect.Top + xRect.Height > LBMP.DisplayRectangle.Height Then SendMessage(LBMP.Handle, &H115, 1, 0)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LWAV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LWAV.Click
        If TBWrite.Checked Then FSW.Text = C10to36(LWAV.SelectedIndex + 1)

        PreviewNote("", True)
        If Not PreviewOnClick Then Exit Sub
        If hWAV(LWAV.SelectedIndex + 1) = "" Then Exit Sub

        Dim xFileLocation As String = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString() & "\" & hWAV(LWAV.SelectedIndex + 1)
        PreviewNote(xFileLocation, False)
    End Sub

    Private Sub LWAV_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LWAV.DoubleClick
        Dim xDWAV As New OpenFileDialog
        xDWAV.DefaultExt = "wav"
        xDWAV.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDWAV.Multiselect = True
        xDWAV.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)

        ' Replace multiple
        AddToPOWAV(xDWAV.FileNames)
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub LWAV_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LWAV.KeyDown
        Select Case e.KeyCode
            Case Keys.Space
                LWAV_Click(sender, e)
            Case Keys.Enter
                LWAV_DoubleClick(sender, e)
            Case Keys.Delete
                Dim xLWAVIds(LWAV.SelectedIndices.Count - 1) As Integer
                For i = 0 To LWAV.SelectedIndices.Count - 1
                    xLWAVIds(i) = LWAV.SelectedIndices(i)
                Next
                For i = 0 To xLWAVIds.Count - 1
                    hWAV(xLWAVIds(i) + 1) = ""
                    wLWAV(xLWAVIds(i) + 1) = New WavSample({}, {}, 0, 0)
                    LWAV.Items.Item(xLWAVIds(i)) = C10to36(xLWAVIds(i) + 1) & ": "
                Next
                If IsSaved Then SetIsSaved(False)
        End Select
    End Sub

    Private Sub LWAVRefresh()
        LWAV.BeginUpdate()
        LWAVRefreshId = 1
        LWAV.Enabled = False
        TimerLWAVRefresh.Enabled = True
    End Sub

    Private Sub TimerLWAVRefresh_Tick(sender As Object, e As EventArgs) Handles TimerLWAVRefresh.Tick
        Dim xIL = LWAVRefreshId - 1
        LWAV.Items(xIL) = C10to36(LWAVRefreshId) & ": " & hWAV(LWAVRefreshId)
        ' Console.WriteLine(LWAVRefreshId)

        If LWAVRefreshId = 1295 Then
            LWAVRefreshId = 1
            LWAV.Enabled = True
            TimerLWAVRefresh.Enabled = False
            LWAV.EndUpdate()
            Exit Sub
        End If
        LWAVRefreshId += 1
    End Sub

    Private Sub LBMPRefresh()
        LBMP.BeginUpdate()
        LBMPRefreshId = 1
        LBMP.Enabled = False
        TimerLBMPRefresh.Enabled = True
    End Sub

    Private Sub TimerLBMPRefresh_Tick(sender As Object, e As EventArgs) Handles TimerLBMPRefresh.Tick
        Dim xIL = LBMPRefreshId - 1
        LBMP.Items(xIL) = C10to36(LBMPRefreshId) & ": " & hBMP(LBMPRefreshId)
        ' Console.WriteLine(LBMPRefreshId)

        If LBMPRefreshId = 1295 Then
            LBMPRefreshId = 1
            LBMP.Enabled = True
            TimerLBMPRefresh.Enabled = False
            LBMP.EndUpdate()
            Exit Sub
        End If
        LBMPRefreshId += 1
    End Sub

    Private Sub LBeatRefresh()
        For xILB = 0 To 999
            Dim a As Double = MeasureLength(xILB) / 192.0R
            Dim xxD = GetDenominator(a)
            LBeat.Items(xILB) = Add3Zeros(xILB) & ": " & a & IIf(xxD > 10000, "", " ( " & CLng(a * xxD) & " / " & xxD & " ) ").ToString()
        Next
    End Sub

    Private Sub TBErrorCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBErrorCheck.Click, mnErrorCheck.Click
        If TypeOf sender Is ToolStripButton Then
            Dim senderC As ToolStripButton = CType(sender, ToolStripButton)
            ErrorCheck = senderC.Checked
        ElseIf TypeOf sender Is ToolStripMenuItem Then
            Dim senderC As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
            ErrorCheck = senderC.Checked
        End If

        TBErrorCheck.Checked = ErrorCheck
        mnErrorCheck.Checked = ErrorCheck
        TBErrorCheck.Image = CType(IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN), Image)
        mnErrorCheck.Image = CType(IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN), Image)
        RefreshPanelAll()
    End Sub

    Private Sub TBPreviewOnClick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPreviewOnClick.Click, mnPreviewOnClick.Click
        PreviewNote("", True)
        If TypeOf sender Is ToolStripButton Then
            Dim senderC As ToolStripButton = CType(sender, ToolStripButton)
            PreviewOnClick = senderC.Checked
        ElseIf TypeOf sender Is ToolStripMenuItem Then
            Dim senderC As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
            PreviewOnClick = senderC.Checked
        End If
        TBPreviewOnClick.Checked = PreviewOnClick
        mnPreviewOnClick.Checked = PreviewOnClick
        TBPreviewOnClick.Image = CType(IIf(PreviewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN), Image)
        mnPreviewOnClick.Image = CType(IIf(PreviewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN), Image)
    End Sub

    'Private Sub TBPreviewErrorCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    PreviewErrorCheck = TBPreviewErrorCheck.Checked
    '    TBPreviewErrorCheck.Image = IIf(PreviewErrorCheck, My.Resources.x16PreviewCheck, My.Resources.x16PreviewCheckN)
    'End Sub

    Private Sub TBShowFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBShowFileName.Click, mnShowFileName.Click
        If TypeOf sender Is ToolStripButton Then
            Dim senderC As ToolStripButton = CType(sender, ToolStripButton)
            ShowFileName = senderC.Checked
        ElseIf TypeOf sender Is ToolStripMenuItem Then
            Dim senderC As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
            ShowFileName = senderC.Checked
        End If
        TBShowFileName.Checked = ShowFileName
        mnShowFileName.Checked = ShowFileName
        TBShowFileName.Image = CType(IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN), Image)
        mnShowFileName.Image = CType(IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN), Image)
        RefreshPanelAll()
    End Sub

    Private Sub TBShowWaveform_Click(sender As Object, e As EventArgs) Handles TBShowWaveform.Click, mnShowWaveform.Click
        If TypeOf sender Is ToolStripButton Then
            Dim senderC As ToolStripButton = CType(sender, ToolStripButton)
            ShowWaveform = senderC.Checked
        ElseIf TypeOf sender Is ToolStripMenuItem Then
            Dim senderC As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
            ShowWaveform = senderC.Checked
        End If
        TBShowWaveform.Checked = ShowWaveform
        mnShowWaveform.Checked = ShowWaveform
        TBShowWaveform.Image = CType(IIf(ShowWaveform, My.Resources.x16ShowWaveform, My.Resources.x16ShowWaveformN), Image)
        mnShowWaveform.Image = CType(IIf(ShowWaveform, My.Resources.x16ShowWaveform, My.Resources.x16ShowWaveformN), Image)

        If ShowWaveform Then
            If Not WaveformLoaded Then
                TimerLoadWaveform.Enabled = True
            Else
                LoadNoteWVPosEnd()
            End If
        Else
            TimerLoadWaveform.Enabled = False
        End If
    End Sub

    Private Sub TimerLoadWaveform_Tick(sender As Object, e As EventArgs) Handles TimerLoadWaveform.Tick
        Console.WriteLine(WaveformLoadId)
        If hWAV(WaveformLoadId) <> "" Then wLWAV(WaveformLoadId) = LoadWaveForm(ExcludeFileName(FileName) & "\" & hWAV(WaveformLoadId))
        If WaveformLoadId = UBound(wLWAV) Then
            WaveformLoadId = 1
            TimerLoadWaveform.Enabled = False
            WaveformLoaded = True

            LoadNoteWVPosEnd()
            RefreshPanelAll()
            Exit Sub
        End If

        WaveformLoadId += 1
    End Sub

    Private Sub LoadNoteWVPosEnd()
        ReDim NoteWVPosEnd(UBound(Notes))
        For i = 0 To UBound(Notes)
            If IsColumnSound(Notes(i).ColumnIndex) Then
                NoteWVPosEnd(i) = GetVPositionFromTime(GetTimeFromVPosition(Notes(i).VPosition) + wLWAV(CInt(Notes(i).Value / 10000)).Duration)
            Else
                NoteWVPosEnd(i) = -1
            End If
        Next
    End Sub

    Private Sub TBCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCut.Click, mnCut.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Me.RedoRemoveNoteSelected(True, xUndo, xRedo)
        'Dim xRedo As String = sCmdKDs()
        'Dim xUndo As String = sCmdKs(True)

        CopyNotes(False)
        RemoveNotes(False)
        AddUndo(xUndo, xBaseRedo.Next)

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()
    End Sub

    Private Sub TBCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCopy.Click, mnCopy.Click
        CopyNotes()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPaste.Click, mnPaste.Click
        AddNotesFromClipboard()

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Me.RedoAddNoteSelected(True, xUndo, xRedo)
        AddUndo(xUndo, xBaseRedo.Next)

        'AddUndo(sCmdKDs(), sCmdKs(True))

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()
    End Sub

    Private Sub TBPastePattern_Click(ByVal sender As Object, ByVal e As EventArgs) Handles mnPastePattern.Click, TBPastePattern.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' Preferable to work under NT settings
        Dim NTInputTemp As Boolean = NTInput
        If Not NTInputTemp Then
            NTInput = True
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertBMSE2NT()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        Dim NotesCB As Note() = GetNotesFromClipboard()
        If NotesCB Is Nothing Then
            If Not NTInputTemp Then ConvertNT2BMSE()
            Exit Sub
        End If

        ' Count selected notes
        Dim xLS As Integer = (From note In Notes Where note.Selected Select note).Count

        If NotesCB.Length <> xLS Then
            Dim xDiag = MsgBox("Warning: The clipboard note count is different from the highlighted note count. Continue?", MsgBoxStyle.YesNo)
            If xDiag = MsgBoxResult.No Then Exit Sub
        End If

        Dim xICB As Integer = -1
        For xIN = 1 To UBound(Notes)
            If Not Notes(xIN).Selected OrElse Notes(xIN).Ghost Then Continue For

            xICB = (xICB + 1) Mod NotesCB.Length
            Dim xCol = NotesCB(xICB).ColumnIndex
            Dim xVPos = CDbl(IIf(PastePatternToVPosition, NotesCB(xICB).VPosition, Notes(xIN).VPosition))
            Dim xLen = NotesCB(xICB).Length

            RedoMoveNote(Notes(xIN), xCol, xVPos, xUndo, xRedo)
            Notes(xIN).ColumnIndex = xCol
            Notes(xIN).VPosition = xVPos

            RedoLongNoteModify(Notes(xIN), Notes(xIN).VPosition, xLen, xUndo, xRedo)
            Notes(xIN).Length = xLen
        Next

        If Not NTInputTemp Then
            NTInput = False
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertNT2BMSE()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        AddUndo(xUndo, xBaseRedo.Next)

        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub LBeatCopyPaste(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles LBeat.PreviewKeyDown
        Select Case e.KeyCode
            Case Keys.C
                If My.Computer.Keyboard.CtrlKeyDown Then
                    Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer

                    LBeat.SelectedIndices.CopyTo(xIndices, 0)
                    Dim xMeasureLengthSelected(UBound(xIndices)) As String
                    For xIL = 0 To UBound(xIndices)
                        xMeasureLengthSelected(xIL) = (MeasureLength(xIL) / 192.0R).ToString
                    Next

                    Clipboard.SetText(Join(xMeasureLengthSelected, vbCrLf))
                End If
            Case Keys.V
                If My.Computer.Keyboard.CtrlKeyDown Then
                    If LBeat.SelectedIndex = -1 Then Exit Sub
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                    Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
                    Dim xMeasureLengthBefore = CType(MeasureLength.Clone, Double())

                    Dim xMeasureLengthClipboard() As String = Split(Clipboard.GetText, vbCrLf)

                    Dim xIL0 As Integer = LBeat.SelectedIndex
                    Dim xRatio As Double
                    For xIL = 0 To UBound(xMeasureLengthClipboard)
                        LBeat.SelectedIndices.Clear()
                        If Double.TryParse(xMeasureLengthClipboard(xIL), xRatio) Then
                            If xRatio <= 0.0# Or xRatio >= 1000.0# Then System.Media.SystemSounds.Hand.Play() : Exit Sub

                            LBeat.SelectedIndex = xIL0 + xIL
                            ApplyBeat(xRatio, xUndo, xRedo)
                        End If
                    Next

                    LBeat.SelectedIndices.Clear()
                    For xI1 As Integer = 0 To UBound(xMeasureLengthClipboard)
                        LBeat.SelectedIndices.Add(xIL0 + xI1)
                    Next

                    RedoChangeMeasure(xMeasureLengthBefore, CType(MeasureLength.Clone(), Double()), xUndo, xRedo)
                    AddUndo(xUndo, xBaseRedo.Next)

                    RefreshPanelAll()
                    POStatusRefresh()
                End If
            Case Keys.Z
                If My.Computer.Keyboard.CtrlKeyDown Then TBUndo_Click(TBUndo, New EventArgs)
            Case Keys.Y
                If My.Computer.Keyboard.CtrlKeyDown Then TBRedo_Click(TBRedo, New EventArgs)
        End Select
    End Sub

    'Private Function pArgPath(ByVal I As Integer)
    '    Return Mid(pArgs(I), 1, InStr(pArgs(I), vbCrLf) - 1)
    'End Function

    Private Sub TBPreviewHighlighted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Experimental feature. Not optimized, will eat up a lot of RAM if a lot of notes are played.

        If TimerInternalPlay.Enabled = True Then
            TimerInternalPlay.Enabled = False
            For i = 1 To UBound(InternalPlayWav)
                InternalPlayWav(i).Finalized()
            Next
            Exit Sub
        End If

        For i = 1 To UBound(InternalPlayWav)
            InternalPlayWav(i) = New AudioC
            InternalPlayWav(i).Initialize()
        Next

        ReDim InternalPlayNotes(UBound(Notes))
        Dim xI1 As Integer = -1
        For xI2 = 1 To UBound(Notes)
            With Notes(xI2)
                If .Selected AndAlso Not .Comment AndAlso IsColumnSound(.ColumnIndex) AndAlso hWAV(CInt(.Value / 10000)) <> "" Then
                    xI1 += 1
                    InternalPlayNotes(xI1) = Notes(xI2)
                End If
            End With
        Next
        ReDim Preserve InternalPlayNotes(xI1)
        ' InternalPlayNotes = From Note In Notes Where Note.Selected AndAlso Not Note.Comment AndAlso IsColumnSound(Note.ColumnIndex) AndAlso hWAV(Note.Value / 10000) <> ""
        '                     Select Note

        If InternalPlayNotes.Count >= 100 Then
            Dim xResult As MsgBoxResult = MsgBox("Warning: You're about to play a lot of notes." & vbCrLf & "This is not recommended as this function has not been fully developed." & vbCrLf & "Do you wish to continue?", MsgBoxStyle.YesNo)
            If xResult = MsgBoxResult.No Then Exit Sub
        End If
        InternalPlayTimerStart = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds
        InternalPlayTimerEnd = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds +
                               CLng((GetTimeFromVPosition(InternalPlayNotes(InternalPlayNotes.Count - 1).VPosition) - GetTimeFromVPosition(InternalPlayNotes(0).VPosition)) * 1000)
        InternalPlayNoteIndex = 0
        TimerInternalPlay.Enabled = True
    End Sub

    Private Sub TimerInternalPlay_Tick(sender As Object, e As EventArgs) Handles TimerInternalPlay.Tick
        InternalPlayTimerCount = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds - InternalPlayTimerStart
        InternalPlaySub()
    End Sub

    Private Sub InternalPlaySub()
        If InternalPlayNoteIndex > InternalPlayNotes.Length - 1 Then
            Dim xIWL = InternalPlayNoteIndex - 1
            If InternalPlayTimerCount > CLng((GetTimeFromVPosition(InternalPlayNotes(xIWL).VPosition) - GetTimeFromVPosition(InternalPlayNotes(0).VPosition)) * 1000) Then
                TimerInternalPlay.Enabled = False
                For i = 1 To UBound(InternalPlayWav)
                    InternalPlayWav(i).Finalized()
                Next
            End If

            Exit Sub
        End If

        Dim NoteTime = GetTimeFromVPosition(InternalPlayNotes(InternalPlayNoteIndex).VPosition) - GetTimeFromVPosition(InternalPlayNotes(0).VPosition)
        If InternalPlayTimerCount >= CLng(NoteTime * 1000) Then
            Dim xIW As Integer = CInt(InternalPlayNotes(InternalPlayNoteIndex).Value / 10000)
            If xIW <= 0 Then xIW = 1
            If xIW >= 1296 Then xIW = 1295

            Dim xFileLocation As String = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString() & "\" & hWAV(xIW)
            InternalPlayWav(xIW).Play(xFileLocation)

            InternalPlayNoteIndex += 1
        End If
    End Sub

    Private Function GetFileName(ByVal s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        Return Mid(s, CInt(IIf(fslash > bslash, fslash, bslash)) + 1)
    End Function

    Private Function ExcludeFileName(ByVal s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        If (bslash Or fslash) = 0 Then Return ""
        Return Mid(s, 1, CInt(IIf(fslash > bslash, fslash, bslash)) - 1)
    End Function

    Private Sub PlayerMissingPrompt()
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)
        MsgBox(Strings.Messages.CannotFind.Replace("{}", PrevCodeToReal(xArg.Path)) & vbCrLf &
               Strings.Messages.PleaseRespecifyPath, MsgBoxStyle.Critical, Strings.Messages.PlayerNotFound)

        Dim xDOpen As New OpenFileDialog
        xDOpen.InitialDirectory = IIf(ExcludeFileName(PrevCodeToReal(xArg.Path)) = "",
                                      My.Application.Info.DirectoryPath,
                                      ExcludeFileName(PrevCodeToReal(xArg.Path))).ToString()
        xDOpen.FileName = PrevCodeToReal(xArg.Path)
        xDOpen.Filter = Strings.FileType.EXE & "|*.exe"
        xDOpen.DefaultExt = "exe"
        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        'pArgs(CurrentPlayer) = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>") & _
        '                                           Mid(pArgs(CurrentPlayer), InStr(pArgs(CurrentPlayer), vbCrLf))
        'xStr = Split(pArgs(CurrentPlayer), vbCrLf)
        pArgs(CurrentPlayer).Path = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>")
        xArg = pArgs(CurrentPlayer)
    End Sub


    Private Sub TBPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPlay.Click, mnPlay.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        ' az: Treat it like we cancelled the operation
        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)).ToString() & "\" & TempFileName
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)
        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aHere))
    End Sub

    Private Sub TBPlayB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPlayB.Click, mnPlayB.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)).ToString() & "\" & TempFileName
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)

        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aBegin))
    End Sub

    Private Sub TBStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBStop.Click, mnStop.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aStop))
    End Sub

    Private Sub AddTempFileList(ByVal s As String)
        Dim xAdd As Boolean = True
        If pTempFileNames IsNot Nothing Then
            For Each xStr1 As String In pTempFileNames
                If xStr1 = s Then xAdd = False : Exit For
            Next
        End If

        If xAdd Then
            ReDim Preserve pTempFileNames(UBound(pTempFileNames) + 1)
            pTempFileNames(UBound(pTempFileNames)) = s
        End If
    End Sub

    Private Sub TBStatistics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBStatistics.Click, mnStatistics.Click
        SortByVPositionInsertion()
        UpdatePairing()

        Dim data(6, 5) As Integer
        Dim dataLNToggle As Boolean = True ' For halving long note counts in Not NTInput mode
        For i As Integer = 1 To UBound(Notes)
            With Notes(i)
                Dim row As Integer = -1
                Select Case .ColumnIndex
                    ' Case niSCROLL : row = 0
                    Case niBPM : row = 0
                    Case niSTOP : row = 1
                    Case niA1, niA2, niA3, niA4, niA5, niA6, niA7, niA8 : row = 2
                    Case niD1, niD2, niD3, niD4, niD5, niD6, niD7, niD8 : row = 3
                    Case Is >= niB : row = 4
                    Case Else : row = 5
                End Select

                Do While row <> 6
                    If Not NTInput Then
                        If Not .LongNote Then data(row, 0) += 1
                        If .LongNote Then
                            If dataLNToggle Then data(row, 1) += 1
                            dataLNToggle = Not dataLNToggle
                        End If
                        If .Value \ 10000 = LnObj Then data(row, 2) += 1
                        If .Hidden Then data(row, 3) += 1
                        If .HasError Then data(row, 4) += 1
                        data(row, 5) += 1

                    Else
                        Dim noteUnit As Integer = 1
                        If .Length = 0 Then data(row, 0) += 1
                        If .Length <> 0 Then data(row, 1) += 2 : noteUnit = 2

                        If .Value \ 10000 = LnObj Then data(row, 2) += noteUnit
                        If .Hidden Then data(row, 3) += noteUnit
                        If .HasError Then data(row, 4) += noteUnit
                        data(row, 5) += noteUnit

                    End If

                    row = 6
                Loop
            End With
        Next

        Dim dStat As New dgStatisticsLegacy(data)
        dStat.ShowDialog()
    End Sub

    Private Sub TBStatisticsAdvanced_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnStatisticsAdvanced.Click
        SortByVPositionInsertion()
        UpdatePairing()
        Dim rows = 26 - 2 ' TableLayoutPanel1.RowCount doesn't work
        Dim cols = 8 - 2  ' Size of tables including names sub 2
        Dim rowLabels = {Strings.fStatistics.lSCROLL, Strings.fStatistics.lBPM, Strings.fStatistics.lSTOP,
                         Strings.fStatistics.lA1, Strings.fStatistics.lA2, Strings.fStatistics.lA3, Strings.fStatistics.lA4, Strings.fStatistics.lA5, Strings.fStatistics.lA6, Strings.fStatistics.lA7, Strings.fStatistics.lA8,
                         Strings.fStatistics.lD1, Strings.fStatistics.lD2, Strings.fStatistics.lD3, Strings.fStatistics.lD4, Strings.fStatistics.lD5, Strings.fStatistics.lD6, Strings.fStatistics.lD7, Strings.fStatistics.lD8,
                         Strings.fStatistics.lA, Strings.fStatistics.lD, Strings.fStatistics.lBGA, Strings.fStatistics.lBGM, Strings.fStatistics.lNotes,
                         Strings.fStatistics.lTotal}
        Dim colLabels = {Strings.fStatistics.lShort, Strings.fStatistics.lLong, Strings.fStatistics.lLnObj, Strings.fStatistics.lHidden, Strings.fStatistics.lLandmines, Strings.fStatistics.lErrors,
                         Strings.fStatistics.lTotal}
        Dim data(rows, cols) As Integer
        Dim dataWAV(1295, 1) As Integer
        Dim dataLNToggle As Boolean = True ' For halving long note counts in Not NTInput mode

        Dim noteLanes = {niA1, niA2, niA3, niA4, niA5, niA6, niA7, niA8,
                         niD1, niD2, niD3, niD4, niD5, niD6, niD7, niD8}

        ' Check if #WAV has been assigned
        For i = 0 To dataWAV.GetUpperBound(0)
            If Not IsNothing(hWAV(i)) Then
                dataWAV(i, 0) = 1
            Else
                dataWAV(i, 0) = 0
            End If
        Next

        For i As Integer = 1 To UBound(Notes)
            With Notes(i)
                Dim row As Integer
                Select Case .ColumnIndex
                    Case niSCROLL : row = 0
                    Case niBPM : row = 1
                    Case niSTOP : row = 2
                    Case niA1 : row = 3
                    Case niA2 : row = 4
                    Case niA3 : row = 5
                    Case niA4 : row = 6
                    Case niA5 : row = 7
                    Case niA6 : row = 8
                    Case niA7 : row = 9
                    Case niA8 : row = 10
                    Case niD1 : row = 11
                    Case niD2 : row = 12
                    Case niD3 : row = 13
                    Case niD4 : row = 14
                    Case niD5 : row = 15
                    Case niD6 : row = 16
                    Case niD7 : row = 17
                    Case niD8 : row = 18
                    Case Is >= niB : row = 22
                    Case Else : row = 21
                End Select


                Dim idWAV As Integer = CInt(.Value / 10000)
                If Not NTInput Then
                    If Not (.LongNote Or .Hidden Or .Landmine Or .Hidden Or .Value \ 10000 = LnObj Or .Comment) Then
                        data(row, 0) += 1
                        If noteLanes.Contains(.ColumnIndex) Or .ColumnIndex >= niB Then
                            dataWAV(idWAV, 1) += 1
                        End If
                    End If
                    If .LongNote Then
                        If dataLNToggle Then data(row, 1) += 1 : dataWAV(idWAV, 1) += 1
                        dataLNToggle = Not dataLNToggle
                    End If
                    If .Value \ 10000 = LnObj Then data(row, 2) += 1
                    If .Hidden Then data(row, 3) += 1
                    If .Landmine Then data(row, 4) += 1
                    If .HasError Then data(row, 5) += 1

                Else
                    If Not (.LongNote Or .Hidden Or .Landmine Or .Hidden Or .Value \ 10000 = LnObj Or .Comment) Then
                        data(row, 0) += 1
                        If noteLanes.Contains(.ColumnIndex) Or .ColumnIndex >= niB Then
                            dataWAV(idWAV, 1) += 1
                        End If
                    End If
                    If .Length <> 0 Then data(row, 1) += 1
                    If .Value \ 10000 = LnObj Then data(row, 2) += 1
                    If .Hidden Then data(row, 3) += 1
                    If .Landmine Then data(row, 4) += 1
                    If .HasError Then data(row, 5) += 1

                End If

            End With
        Next

        ' Calculate Total notes on column and row
        For r As Integer = 0 To rows - 1
            For c As Integer = 0 To cols - 1
                data(r, cols) += data(r, c) ' Total in row

                ' Ignore rows "A1-A8" and "D1-D8"
                Select Case r
                    Case 19 To 20 : Continue For
                    Case Else : data(rows, c) += data(r, c) ' Total in column
                End Select

                Select Case r
                    Case 3 To 10 : data(19, c) += data(r, c) : data(23, c) += data(r, c)
                    Case 11 To 18 : data(20, c) += data(r, c) : data(23, c) += data(r, c)
                End Select
            Next
        Next
        data(rows, cols) = UBound(Notes)
        ' Change to the whole table cause more convenient
        Dim dStat As New dgStatistics(data, rowLabels, colLabels, dataWAV)
        dStat.ShowDialog()
    End Sub

    ''' <summary>
    ''' Remark: Pls sort and updatepairing before this process.
    ''' </summary>

    Private Sub CalculateTotalPlayableNotes(Optional ModifyTotal As Boolean = True)
        Dim xI1 As Integer
        Dim xIAll As Integer = 0
        Dim xITemp As Integer = 0

        If Not NTInput Then
            For xI1 = 1 To UBound(Notes)
                With Notes(xI1)
                    If .ColumnIndex >= niA1 And .ColumnIndex <= niD8 Then
                        If Not (.LongNote Or .Hidden Or .Landmine Or .Hidden Or .Value \ 10000 = LnObj Or .Comment) Then
                            xIAll += 1
                        ElseIf .LongNote AndAlso Not (.Hidden Or .Landmine Or .Hidden Or .Value \ 10000 = LnObj Or .Comment) Then
                            xITemp += 1
                        End If
                    End If
                End With
            Next
            xIAll += CInt(xITemp / 2)
        Else
            For xI1 = 1 To UBound(Notes)
                With Notes(xI1)
                    If .ColumnIndex >= niA1 And .ColumnIndex <= niD8 AndAlso
                       Not (.LongNote Or .Hidden Or .Landmine Or .Hidden Or .Value \ 10000 = LnObj Or .Comment) Then xIAll += 1
                End With
            Next
        End If

        Dim TotalValue As Double
        Select Case TotalOption
            Case 0
                TotalValue = xIAll * 7.605 / (0.01 * xIAll + 6.5) * TotalGlobalMultiplier
            Case 1
                TotalValue = CDbl(IIf(xIAll < 400, 200 + xIAll / 5, IIf(xIAll < 600, 280 + (xIAll - 400) / 2.5, 360 + (xIAll - 600) / 5))) * TotalGlobalMultiplier
            Case 2
                TotalValue = xIAll * TotalMultiplier * TotalGlobalMultiplier
        End Select
        If TotalDisplayValue Then
            TBTotalValue.Text = IIf(TotalDisplayText, Strings.StatusBar.RecommendedTotal & ": ", "").ToString() & Math.Round(TotalValue, TotalDecimal).ToString()
        Else
            TBTotalValue.Text = ""
        End If
        TBStatistics.Text = xIAll.ToString()
        If ModifyTotal AndAlso TotalAutofill Then THTotal.Text = Math.Round(TotalValue, TotalDecimal).ToString()
    End Sub

    Public Function GetMouseVPosition(Optional snap As Boolean = True) As Double
        Dim panHeight = spMain(PanelFocus).Height
        Dim panDisplacement = PanelVScroll(PanelFocus)
        Dim vpos = (panHeight - panDisplacement * gxHeight - MouseMoveStatus.Y - 1) / gxHeight
        If snap Then
            Return SnapToGrid(vpos)
        Else
            Return vpos
        End If
    End Function

    Private Sub POStatusRefresh()

        If TBSelect.Checked Then
            Dim xI1 As Integer = KMouseOver
            If xI1 < 0 Then

                TempVPosition = GetMouseVPosition(gSnap)

                SelectedColumn = GetColumnAtX(MouseMoveStatus.X, PanelHScroll(PanelFocus))

                Dim xMeasure As Integer = MeasureAtDisplacement(TempVPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = TempVPosition - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(CDbl(IIf(xVposMod = 0, xMLength, xVposMod)), xMLength)

                FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = TempVPosition.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(TempVPosition).ToString("F4")
                FSC.Text = nTitle(SelectedColumn)
                FSW.Text = ""
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = ""
                FSE.Text = ""

            ElseIf xI1 <= UBound(Notes) Then
                Dim xMeasure As Integer = MeasureAtDisplacement(Notes(xI1).VPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = Notes(xI1).VPosition - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(CDbl(IIf(xVposMod = 0, xMLength, xVposMod)), xMLength)

                FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = Notes(xI1).VPosition.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(Notes(xI1).VPosition).ToString("F4")
                FSC.Text = nTitle(Notes(xI1).ColumnIndex)
                FSW.Text = IIf(IsColumnNumeric(Notes(xI1).ColumnIndex),
                               Notes(xI1).Value / 10000,
                               C10to36(Notes(xI1).Value \ 10000)).ToString()
                FSM.Text = Add3Zeros(xMeasure)

                ' TODO: Count stops
                If Notes(xI1).Length > 0 Then
                    FST.ForeColor = System.Drawing.Color.Olive
                    FST.Text = Strings.StatusBar.LongNote & " " &
                               Notes(xI1).Length / 192.0R * 4 & " " & Strings.Beats & " " &
                               "(" & GetTimeFromVPosition(Notes(xI1).VPosition + Notes(xI1).Length) - GetTimeFromVPosition(Notes(xI1).VPosition) & "s)"
                ElseIf Notes(xI1).LNPair <> 0 Then
                    FST.ForeColor = System.Drawing.Color.Olive
                    FST.Text = Strings.StatusBar.LongNote & " " &
                               Math.Abs(Notes(xI1).VPosition - Notes(Notes(xI1).LNPair).VPosition) / 192.0R * 4 & " " & Strings.Beats & " " &
                               "(" & Math.Abs(GetTimeFromVPosition(Notes(Notes(xI1).LNPair).VPosition) - GetTimeFromVPosition(Notes(xI1).VPosition)) & "s)"
                ElseIf Notes(xI1).LongNote Then
                    FST.ForeColor = System.Drawing.Color.Olive
                    FST.Text = Strings.StatusBar.LongNote
                ElseIf Notes(xI1).Hidden Then
                    FST.ForeColor = System.Drawing.Color.Blue
                    FST.Text = Strings.StatusBar.Hidden
                ElseIf Notes(xI1).Landmine Then
                    FST.ForeColor = System.Drawing.Color.Red
                    FST.Text = Strings.StatusBar.Landmine
                ElseIf Notes(xI1).Comment Then
                    FST.ForeColor = System.Drawing.Color.Green
                    FST.Text = Strings.StatusBar.Comment
                Else
                    FST.ForeColor = System.Drawing.Color.Olive
                    FST.Text = Strings.StatusBar.Note
                End If

                If Notes(xI1).HasError Then
                    Select Case Notes(xI1).ErrorType
                        Case 0
                            FSE.Text = Strings.StatusBar.Err
                        Case 1
                            FSE.Text = Strings.StatusBar.ErrTechnical
                    End Select
                Else
                    FSE.Text = ""
                End If

            End If

        ElseIf TBWrite.Checked Then
            If SelectedColumn < 0 Then Exit Sub

            Dim xMeasure As Integer = MeasureAtDisplacement(TempVPosition)
            Dim xMLength As Double = MeasureLength(xMeasure)
            Dim xVposMod As Double = TempVPosition - MeasureBottom(xMeasure)
            Dim xGCD As Double = GCD(CDbl(IIf(xVposMod = 0, xMLength, xVposMod)), xMLength)

            FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
            FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
            FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
            FSP4.Text = TempVPosition.ToString() & "  "
            TimeStatusLabel.Text = GetTimeFromVPosition(TempVPosition).ToString("F4")
            FSC.Text = nTitle(SelectedColumn)
            FSW.Text = C10to36(LWAV.SelectedIndex + 1)
            FSM.Text = Add3Zeros(xMeasure)
            If TempLength > 0 Then
                FST.ForeColor = System.Drawing.Color.Olive
                FST.Text = Strings.StatusBar.LongNote & " " &
                           TempLength / 192.0R * 4 & " " & Strings.Beats & " " &
                           "(" & Strings.StatusBar.Approximate & " " & GetTimeFromVPosition(TempVPosition + TempLength) - GetTimeFromVPosition(TempVPosition) & "s)"
            ElseIf My.Computer.Keyboard.CtrlKeyDown AndAlso My.Computer.Keyboard.ShiftKeyDown Then
                FST.ForeColor = System.Drawing.Color.Red
                FST.Text = Strings.StatusBar.Landmine
            ElseIf My.Computer.Keyboard.ShiftKeyDown AndAlso Not NTInput Then
                FST.ForeColor = System.Drawing.Color.Olive
                FST.Text = Strings.StatusBar.LongNote
            ElseIf My.Computer.Keyboard.CtrlKeyDown Then
                FST.ForeColor = System.Drawing.Color.Blue
                FST.Text = Strings.StatusBar.Hidden
            End If

        ElseIf TBTimeSelect.Checked Then
            FSSS.Text = vSelStart.ToString()
            FSSL.Text = vSelLength.ToString()
            FSSH.Text = vSelHalf.ToString()

        End If
        FStatus.Invalidate()
    End Sub

    Private Function GetTimeFromVPosition(vpos As Double) As Double
        Dim timing_notes = (From note In Notes
                            Where note.ColumnIndex = niBPM Or note.ColumnIndex = niSTOP
                            Group By Column = note.ColumnIndex
                               Into NoteGroups = Group).ToDictionary(Function(x) x.Column, Function(x) x.NoteGroups)

        Dim bpm_notes = timing_notes.Item(niBPM)

        Dim stop_notes As IEnumerable(Of Note) = Nothing

        If timing_notes.ContainsKey(niSTOP) Then
            stop_notes = timing_notes.Item(niSTOP)
        End If


        Dim stop_contrib As Double
        Dim bpm_contrib As Double
        Dim duration = 0.0

        For i = 0 To bpm_notes.Count() - 1
            ' az: sum bpm contribution first
            ' P: Yeah but not all of them
            Dim current_note = bpm_notes.ElementAt(i)
            Dim notevpos = Math.Max(0, current_note.VPosition)
            If notevpos > vpos Then Exit For

            If i + 1 <> bpm_notes.Count() Then
                Dim next_note = bpm_notes.ElementAt(i + 1)
                duration = Math.Min(next_note.VPosition, vpos) - notevpos
            Else
                duration = vpos - notevpos
            End If

            Dim current_bps = 60 / (current_note.Value / 10000)
            bpm_contrib += current_bps * duration / 48

            If stop_notes Is Nothing Then Continue For

            Dim stops = From stp In stop_notes
                        Where stp.VPosition >= notevpos And
                            stp.VPosition < notevpos + duration

            Dim stop_beats = stops.Sum(Function(x) x.Value / 10000.0) / 48
            stop_contrib += current_bps * stop_beats

        Next

        Return stop_contrib + bpm_contrib
    End Function

    Private Function GetVPositionFromTime(ByVal Time As Double) As Double
        Dim timing_notes = (From note In Notes
                            Where note.ColumnIndex = niBPM Or note.ColumnIndex = niSTOP
                            Group By Column = note.ColumnIndex
        Into NoteGroups = Group).ToDictionary(Function(x) x.Column, Function(x) x.NoteGroups)

        Dim bpm_notes = timing_notes.Item(niBPM)

        Dim stop_notes As IEnumerable(Of Note) = Nothing

        If timing_notes.ContainsKey(niSTOP) Then
            stop_notes = timing_notes.Item(niSTOP)
        End If

        Dim stop_subtract_time As Double
        Dim bpm_contrib_time As Double
        Dim DurationVPos As Double
        Dim remaining_time As Double = Time

        Dim VPos As Double = 0

        For i = 0 To bpm_notes.Count() - 1
            If remaining_time = 0 Then Exit For

            Dim current_note = bpm_notes.ElementAt(i)
            Dim notevpos = Math.Max(0, current_note.VPosition)

            ' Beats per second
            Dim current_bps = 60 / (current_note.Value / 10000)

            ' Get duration from BPM notes first
            If i + 1 <> bpm_notes.Count() Then
                ' Get duration in seconds between this and next bpm notes
                Dim next_note = bpm_notes.ElementAt(i + 1)
                DurationVPos = next_note.VPosition - notevpos
                bpm_contrib_time = current_bps * DurationVPos / 48

                ' If remaining_time is out of range between current and next bpm_note
                If bpm_contrib_time >= remaining_time Then
                    VPos += 48 * remaining_time / current_bps
                    remaining_time = 0
                Else
                    VPos += DurationVPos
                    remaining_time -= bpm_contrib_time
                End If
            Else
                bpm_contrib_time = remaining_time
                VPos += 48 * remaining_time / current_bps
                remaining_time = 0
            End If

            If stop_notes Is Nothing Then Continue For

            ' Dim stops = From stp In stop_notes
            '             Where stp.VPosition >= notevpos And
            '                 stp.VPosition < VPos
            Dim stops(stop_notes.Count - 1) As Note
            Dim xIS = -1
            For Each stp In stop_notes
                If stp.VPosition >= notevpos AndAlso stp.VPosition < VPos Then
                    xIS += 1
                    stops(xIS) = stp
                End If
            Next
            ReDim Preserve stops(xIS)

            Dim stop_contrib As Double = 0

            For j = 0 To stops.Count() - 1
                ' Get negative duration from stop notes
                Dim current_stop_note = stops.ElementAt(j)
                If current_stop_note.VPosition >= VPos Then Exit For

                stop_subtract_time = current_bps * current_stop_note.Value / 10000.0 / 48

                ' If the stop note duration exceeds excess time from current calculation
                If stop_subtract_time > remaining_time Then
                    ' Recalculate VPos from stop note VPos
                    Dim remaining_time_stop As Double = remaining_time + bpm_contrib_time - current_bps * (current_stop_note.VPosition - notevpos) / 48 - stop_contrib
                    remaining_time = 0
                    VPos = current_stop_note.VPosition
                    ' If the stop note duration exceeds duration between stop note and specified time
                    If stop_subtract_time >= remaining_time_stop Then
                        Exit For
                    Else
                        remaining_time_stop -= stop_subtract_time
                        VPos += 48 * remaining_time_stop / current_bps
                    End If
                Else
                    remaining_time -= stop_subtract_time
                End If
                stop_contrib += stop_subtract_time
            Next
        Next
        Return VPos
    End Function


    Private Sub ValidateSelection()
        If vSelStart < 0 Then vSelLength += vSelStart : vSelHalf += vSelStart : vSelStart = 0
        If vSelStart > GetMaxVPosition() - 1 Then vSelLength += vSelStart - GetMaxVPosition() + 1 : vSelHalf += vSelStart - GetMaxVPosition() + 1 : vSelStart = GetMaxVPosition() - 1
        If vSelStart + vSelLength < 0 Then vSelLength = -vSelStart
        If vSelStart + vSelLength > GetMaxVPosition() - 1 Then vSelLength = GetMaxVPosition() - 1 - vSelStart

        If Math.Sign(vSelHalf) <> Math.Sign(vSelLength) Then vSelHalf = 0
        If Math.Abs(vSelHalf) > Math.Abs(vSelLength) Then vSelHalf = vSelLength
    End Sub



    Private Sub TVCM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCM.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim xTP As Double
            If Not Double.TryParse(TVCM.Text, xTP) Then
                TVCM.Text = 1.ToString()
            Else
                TVCM.Text = xTP.ToString()
            End If
            If CDbl(TVCM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCM.Text = 1.ToString()
                TVCM.Focus()
                TVCM.SelectAll()
            Else
                BVCApply_Click(BVCApply, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCM_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCM.LostFocus
        Dim xTP As Double
        If Not Double.TryParse(TVCM.Text, xTP) Then
            TVCM.Text = 0.ToString()
        Else
            TVCM.Text = xTP.ToString()
        End If
        If CDbl(TVCM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCM.Text = 1.ToString()
            TVCM.Focus()
            TVCM.SelectAll()
        End If
    End Sub

    Private Sub TVCD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCD.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim xTP As Double
            If Not Double.TryParse(TVCD.Text, xTP) Then
                TVCD.Text = 0.ToString()
            Else
                TVCD.Text = xTP.ToString()
            End If
            If CDbl(TVCD.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCD.Text = 1.ToString()
                TVCD.Focus()
                TVCD.SelectAll()
            Else
                BVCApply_Click(BVCApply, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCD.LostFocus
        Dim xTP As Double
        If Not Double.TryParse(TVCD.Text, xTP) Then
            TVCD.Text = 0.ToString()
        Else
            TVCD.Text = xTP.ToString()
        End If
        If CDbl(TVCD.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCD.Text = 1.ToString()
            TVCD.Focus()
            TVCD.SelectAll()
        End If
    End Sub

    Private Sub TVCBPM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCBPM.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim xTP As Double
            If Not Double.TryParse(TVCBPM.Text, xTP) Then
                TVCBPM.Text = 0.ToString()
            Else
                TVCBPM.Text = xTP.ToString()
            End If
            If CDbl(TVCBPM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCBPM.Text = (Notes(0).Value / 10000).ToString()
                TVCBPM.Focus()
                TVCBPM.SelectAll()
            Else
                BVCCalculate_Click(BVCCalculate, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCBPM_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCBPM.LostFocus
        Dim xTP As Double
        If Not Double.TryParse(TVCBPM.Text, xTP) Then
            TVCBPM.Text = 0.ToString()
        Else
            TVCBPM.Text = xTP.ToString()
        End If
        If CDbl(TVCBPM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCBPM.Text = (Notes(0).Value / 10000).ToString()
            TVCBPM.Focus()
            TVCBPM.SelectAll()
        End If
    End Sub

    Private Function FindNoteIndex(note As Note) As Integer
        Dim xI1 As Integer
        If NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).equalsNT(note) Then Return xI1
            Next
        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).equalsBMSE(note) Then Return xI1
            Next
        End If
        Return xI1
    End Function




    Private Function sIA() As Integer
        Return CInt(IIf(sI > UndoRedoCount - 1, 0, sI + 1))
    End Function

    Private Function sIM() As Integer
        Return CInt(IIf(sI < 1, UndoRedoCount, sI - 1))
    End Function



    Private Sub TBUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUndo.Click, mnUndo.Click
        KMouseOver = -1
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        If sUndo(sI).ofType = UndoRedo.opNoOperation Then Exit Sub
        PerformCommand(sUndo(sI))
        sI = sIM()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
    End Sub

    Private Sub TBRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRedo.Click, mnRedo.Click
        Console.WriteLine("Redo sI " + sI.ToString())
        KMouseOver = -1
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        If sRedo(sIA).ofType = UndoRedo.opNoOperation Then Exit Sub
        PerformCommand(sRedo(sIA))
        sI = sIA()

        Console.WriteLine("Redo sI " + sI.ToString())

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
    End Sub

    'Undo appends before, Redo appends after.
    'After a sequence of Commands, 
    '   Undo will be the first one to execute, 
    '   Redo will be the last one to execute.
    'Remember to save the first Redo.

    'In case where undo is Nothing: Dont worry.
    'In case where redo is Nothing: 
    '   If only one redo is in a sequence, put Nothing.
    '   If several redo are in a sequence, 
    '       Create Void first. 
    '       Record its reference into a seperate copy. (xBaseRedo = xRedo)
    '       Use this xRedo as the BaseRedo.
    '       When calling AddUndo subroutine, use xBaseRedo.Next as cRedo.

    'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
    'Dim xRedo As UndoRedo.LinkedURCmd = Nothing
    '... 'Me.RedoRemoveNote(K(xI1), True, xUndo, xRedo)
    'AddUndo(xUndo, xRedo)

    'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
    'Dim xRedo As New UndoRedo.Void
    'Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
    '... 'Me.RedoRemoveNote(K(xI1), True, xUndo, xRedo)
    'AddUndo(xUndo, xBaseRedo.Next)

    Private Sub TBAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Aboutboxx1 As New AboutBox1()
        'If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\About.png") Then
        Aboutboxx1.bBitmap = My.Resources.About0
        'Aboutboxx1.SelectBitmap()
        Aboutboxx1.ClientSize = New Size(1000, 500)
        Aboutboxx1.ClickToCopy.Visible = True
        Aboutboxx1.ShowDialog(Me)
        'Else
        '    MsgBox(locale.Messages.cannotfind & " ""About.png""", MsgBoxStyle.Critical, locale.Messages.err)
        'End If
    End Sub

    Private Sub TBVOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBVOptions.Click, mnVOptions.Click

        Dim xDiag As New OpVisual(vo, column, LWAV.Font)
        xDiag.ShowDialog(Me)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub TBVCOptions_Click(sender As Object, e As EventArgs) Handles mnVCOptions.Click, BWAVColorOverride.Click
        Dim xDiag As New OpVisualOverride(COverrides, hWAV, COverridesSaveOption)
        ' Save settings
        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            COverrides = CType(xDiag.COverrides.Clone(), ColorOverride())
            If COverridesSaveOption <> xDiag.CoBSave.SelectedIndex Then
                COverridesSaveOption = xDiag.CoBSave.SelectedIndex

                If Not IsNothing(COverrides) Then If COverrides.Length > 0 Then WriteColorOverride(FileName, True)
            Else
                If Not IsNothing(COverrides) Then If COverrides.Length > 0 Then WriteColorOverride(FileName, False)

            End If
        End If

        ' Load settings if chosen
        If xDiag.CoBLoad.SelectedIndex <> -1 Then
            COverridesSaveOption = xDiag.CoBLoad.SelectedIndex
            LoadColorOverride(FileName)

            TBVCOptions_Click(sender, New EventArgs)
        End If

        LoadColorOverrideActive()
        ' LWAV.Refresh()

        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    ' Private Sub LWAV_DrawItem(sender As Object, e As DrawItemEventArgs) Handles LWAV.DrawItem
    '     e.DrawBackground()
    '     Dim c As Color
    '     GetColor(New Note(0, 0, (e.Index + 1) * 10000, 0), c, c, 1, COverridesActive)
    '     e.Graphics.FillRectangle(New SolidBrush(c), e.Bounds)
    '     e.Graphics.DrawString(LWAV.Items.Item(e.Index).ToString(), e.Font, New SolidBrush(e.ForeColor), e.Bounds)
    '     e.DrawFocusRectangle()
    '     e.Graphics.Dispose()
    ' End Sub

    Private Sub AddToPOWAV(ByVal xPath() As String)
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        If xIndices.Length < xPath.Length Then
            Dim i As Integer = xIndices.Length
            Dim currWavIndex As Integer = xIndices(UBound(xIndices)) + 1
            ReDim Preserve xIndices(UBound(xPath))

            Do While i < xIndices.Length And currWavIndex <= 1294
                Do While currWavIndex <= 1294 AndAlso hWAV(currWavIndex + 1) <> ""
                    currWavIndex += 1
                Loop
                If currWavIndex > 1294 Then Exit Do

                xIndices(i) = currWavIndex
                currWavIndex += 1
                i += 1
            Loop

            If currWavIndex > 1294 Then
                ReDim Preserve xPath(i - 1)
                ReDim Preserve xIndices(i - 1)
            End If
        End If

        'Dim xI2 As Integer = 0
        For xI1 As Integer = 0 To UBound(xPath)
            'If xI2 > UBound(xIndices) Then Exit For
            'hWAV(xIndices(xI2) + 1) = GetFileName(xPath(xI1))
            'LWAV.Items.Item(xIndices(xI2)) = C10to36(xIndices(xI2) + 1) & ": " & GetFileName(xPath(xI1))
            hWAV(xIndices(xI1) + 1) = GetFileName(xPath(xI1))
            LWAV.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": " & GetFileName(xPath(xI1))
            'xI2 += 1
            ' Add waveforms to wLWAV
            If ShowWaveform Then wLWAV(xIndices(xI1) + 1) = LoadWaveForm(xPath(xI1))
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To CInt(IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath)))
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragDrop
        ReDim DDFileName(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedAudioExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPOWAV(xPath)
    End Sub

    Private Sub POWAV_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DDFileName = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedAudioExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.DragLeave
        ReDim DDFileName(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.Resize
        LWAV.Height = POWAV.Height - 25
    End Sub
    Private Sub POBeat_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBeat.Resize
        LBeat.Height = POBeat.Height - 25
    End Sub
    Private Sub POExpansion_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POExpansion.Resize
        TExpansion.Height = POExpansion.Height - 25 - BExpansion.Height
    End Sub

    Private Sub AddToPOBMP(ByVal xPath() As String)
        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        If xIndices.Length < xPath.Length Then
            Dim i As Integer = xIndices.Length
            Dim currBMPIndex As Integer = xIndices(UBound(xIndices)) + 1
            ReDim Preserve xIndices(UBound(xPath))

            Do While i < xIndices.Length And currBMPIndex <= 1294
                Do While currBMPIndex <= 1294 AndAlso hBMP(currBMPIndex + 1) <> ""
                    currBMPIndex += 1
                Loop
                If currBMPIndex > 1294 Then Exit Do

                xIndices(i) = currBMPIndex
                currBMPIndex += 1
                i += 1
            Loop

            If currBMPIndex > 1294 Then
                ReDim Preserve xPath(i - 1)
                ReDim Preserve xIndices(i - 1)
            End If
        End If

        'Dim xI2 As Integer = 0
        For xI1 As Integer = 0 To UBound(xPath)
            'If xI2 > UBound(xIndices) Then Exit For
            'hBMP(xIndices(xI2) + 1) = GetFileName(xPath(xI1))
            'LBMP.Items.Item(xIndices(xI2)) = C10to36(xIndices(xI2) + 1) & ": " & GetFileName(xPath(xI1))
            hBMP(xIndices(xI1) + 1) = GetFileName(xPath(xI1))
            LBMP.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": " & GetFileName(xPath(xI1))
            'xI2 += 1
            ' Add BMPeforms to wLBMP
        Next

        LBMP.SelectedIndices.Clear()
        For xI1 As Integer = 0 To CInt(IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath)))
            LBMP.SelectedIndices.Add(xIndices(xI1))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POBMP.DragDrop
        ReDim DDFileName(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedImageExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPOBMP(xPath)
    End Sub

    Private Sub POBMP_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POBMP.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DDFileName = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedImageExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBMP.DragLeave
        ReDim DDFileName(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBMP.Resize
        LBMP.Height = POBMP.Height - 25
    End Sub

    Private Sub mn_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim TSMIS As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        TSMIS.ForeColor = Color.White
    End Sub
    Private Sub mn_DropDownOpened(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim TSMIS As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        TSMIS.ForeColor = Color.Black
    End Sub
    Private Sub mn_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim TSMIS As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If TSMIS.Pressed Then Return
        TSMIS.ForeColor = Color.Black
    End Sub
    Private Sub mn_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim TSMIS As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If TSMIS.Pressed Then Return
        TSMIS.ForeColor = Color.White
    End Sub

    Private Sub TBPOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPOptions.Click, mnPOptions.Click
        Dim xDOp As New OpPlayer(CurrentPlayer)
        xDOp.ShowDialog(Me)
    End Sub

    Private Sub THGenre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    THGenre.TextChanged, THTitle.TextChanged, THArtist.TextChanged, THPlayLevel.TextChanged, CHRank.SelectedIndexChanged,
    THSubTitle.TextChanged, THSubArtist.TextChanged, THStageFile.TextChanged, THBanner.TextChanged, THBackBMP.TextChanged,
    CHDifficulty.SelectedIndexChanged, THExRank.TextChanged, THTotal.TextChanged, THComment.TextChanged, TExpansion.TextChanged
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub CHLnObj_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHLnObj.SelectedIndexChanged
        If IsSaved Then SetIsSaved(False)
        LnObj = CHLnObj.SelectedIndex
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub ConvertBMSE2NT(Optional xIFrom As Integer = 1)
        ReDim SelectedNotes(-1)
        If xIFrom = 1 Then SortByVPositionInsertion()

        For i2 As Integer = xIFrom To UBound(Notes)
            Notes(i2).Length = 0.0#
        Next

        Dim i As Integer = xIFrom
        Dim j As Integer
        Dim xUbound As Integer = UBound(Notes)

        Do While i <= xUbound
            If Not Notes(i).LongNote Then i += 1 : Continue Do

            For j = i + 1 To xUbound
                If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For

                If Notes(j).LongNote Then
                    Notes(i).Length = Notes(j).VPosition - Notes(i).VPosition
                    For j2 As Integer = j To xUbound - 1
                        Notes(j2) = Notes(j2 + 1)
                    Next
                    xUbound -= 1
                    Exit For

                ElseIf Notes(j).Value \ 10000 = LnObj Then
                    Exit For

                End If
            Next

            i += 1
        Loop

        ReDim Preserve Notes(xUbound)

        For i = xIFrom To xUbound
            Notes(i).LongNote = False
        Next

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub ConvertNT2BMSE(Optional xIFrom As Integer = 1)
        ReDim SelectedNotes(-1)
        Dim xK(xIFrom - 1) As Note
        For xI = 0 To xIFrom - 1
            xK(xI) = Notes(xI)
        Next

        For xI1 As Integer = xIFrom To UBound(Notes)
            ReDim Preserve xK(UBound(xK) + 1)
            With xK(UBound(xK))
                .ColumnIndex = Notes(xI1).ColumnIndex
                .LongNote = Notes(xI1).Length > 0
                .Landmine = Notes(xI1).Landmine
                .Value = Notes(xI1).Value
                .VPosition = Notes(xI1).VPosition
                .Selected = Notes(xI1).Selected
                .Hidden = Notes(xI1).Hidden
                .Ghost = Notes(xI1).Ghost
                .Comment = Notes(xI1).Comment
            End With

            If Notes(xI1).Length > 0 Then
                ReDim Preserve xK(UBound(xK) + 1)
                With xK(UBound(xK))
                    .ColumnIndex = Notes(xI1).ColumnIndex
                    .LongNote = True
                    .Landmine = False
                    .Value = Notes(xI1).Value
                    .VPosition = Notes(xI1).VPosition + Notes(xI1).Length
                    .Selected = Notes(xI1).Selected
                    .Hidden = Notes(xI1).Hidden
                    .Ghost = Notes(xI1).Ghost
                    .Comment = Notes(xI1).Comment
                End With
            End If
        Next

        Notes = xK

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub TBWavIncrease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBWavIncrease.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        TBWavIncrease.Checked = Not TBWavIncrease.Checked
        Me.RedoWavIncrease(TBWavIncrease.Checked, xUndo, xRedo)
        AddUndo(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub TBNTInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBNTInput.Click, mnNTInput.Click
        If TypeOf sender Is ToolStripButton Then
            Dim senderC As ToolStripButton = CType(sender, ToolStripButton)
            NTInput = senderC.Checked
        ElseIf TypeOf sender Is ToolStripMenuItem Then
            Dim senderC As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
            NTInput = senderC.Checked
        End If

        RefreshItemsByNTInput()

        bAdjustLength = False
        bAdjustUpper = False

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)
        Me.RedoNT(NTInput, False, xUndo, xRedo)
        If NTInput Then
            ConvertBMSE2NT()
        Else
            ConvertNT2BMSE()
        End If
        Me.RedoAddNoteAll(False, xUndo, xRedo)

        AddUndo(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
    End Sub

    Private Sub RefreshItemsByNTInput()
        TBNTInput.Checked = NTInput
        mnNTInput.Checked = NTInput
        POBLongObjNT.Visible = NTInput
        POBLongNTObj.Visible = NTInput
        POBLong.Visible = Not NTInput
        POBLongShort.Visible = Not NTInput
    End Sub

    Private Sub THBPM_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles THBPM.ValueChanged
        If Notes IsNot Nothing Then Notes(0).Value = CLng(THBPM.Value * 10000) : RefreshPanelAll()
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub TWPosition_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPosition.ValueChanged
        wPosition = TWPosition.Value
        TWPosition2.Value = CInt(IIf(wPosition > TWPosition2.Maximum, TWPosition2.Maximum, wPosition))
        RefreshPanelAll()
    End Sub

    Private Sub TWLeft_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWLeft.ValueChanged
        wLeft = CInt(TWLeft.Value)
        TWLeft2.Value = CInt(IIf(wLeft > TWLeft2.Maximum, TWLeft2.Maximum, wLeft))
        RefreshPanelAll()
    End Sub

    Private Sub TWWidth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWWidth.ValueChanged
        wWidth = CInt(TWWidth.Value)
        TWWidth2.Value = CInt(IIf(wWidth > TWWidth2.Maximum, TWWidth2.Maximum, wWidth))
        RefreshPanelAll()
    End Sub

    Private Sub TWPrecision_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPrecision.ValueChanged
        ' (0 - 100) → (0 - 20)
        wPrecision = TWPrecision.Value / 5
        TWPrecision2.Value = CInt(IIf(TWPrecision.Value > TWPrecision2.Maximum, TWPrecision2.Maximum, TWPrecision.Value))
        RefreshPanelAll()
    End Sub

    Private Sub TWTransparency_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWTransparency.ValueChanged
        TWTransparency2.Value = CInt(TWTransparency.Value)
        vo.pBGMWav.Color = Color.FromArgb(CInt(TWTransparency.Value), vo.pBGMWav.Color)
        RefreshPanelAll()
    End Sub

    Private Sub TWSaturation_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWSaturation.ValueChanged
        Dim xColor As Color = vo.pBGMWav.Color
        TWSaturation2.Value = CInt(TWSaturation.Value)
        vo.pBGMWav.Color = HSL2RGB(CInt(xColor.GetHue), CInt(TWSaturation.Value), CInt(xColor.GetBrightness * 1000), xColor.A)
        RefreshPanelAll()
    End Sub

    Private Sub TWPosition2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPosition2.Scroll
        TWPosition.Value = TWPosition2.Value
    End Sub

    Private Sub TWLeft2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWLeft2.Scroll
        TWLeft.Value = TWLeft2.Value
    End Sub

    Private Sub TWWidth2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWWidth2.Scroll
        TWWidth.Value = TWWidth2.Value
    End Sub

    Private Sub TWPrecision2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPrecision2.Scroll
        TWPrecision.Value = TWPrecision2.Value
    End Sub

    Private Sub TWTransparency2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWTransparency2.Scroll
        TWTransparency.Value = TWTransparency2.Value
    End Sub

    Private Sub TWSaturation2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWSaturation2.Scroll
        TWSaturation.Value = TWSaturation2.Value
    End Sub

    Private Sub TBLangDef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBLangDef.Click
        DispLang = ""
        MsgBox(Strings.Messages.PreferencePostpone, MsgBoxStyle.Information)
    End Sub

    Private Sub TBLangRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBLangRefresh.Click
        For xI1 As Integer = cmnLanguage.Items.Count - 1 To 3 Step -1
            Try
                cmnLanguage.Items.RemoveAt(xI1)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles("*.Lang.xml")

        For Each xStr As FileInfo In xFileNames
            LoadLocaleXML(xStr)
        Next
    End Sub

    Private Sub TBTotalValue_Click(sender As Object, e As EventArgs) Handles TBTotalValue.Click, mnTOTAL.Click
        Dim xDiag As New OpTotal(CInt(TBStatistics.Text), TotalOption, TotalMultiplier, TotalGlobalMultiplier, TotalDecimal, TotalDisplayValue, TotalDisplayText, TotalAutofill)
        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With xDiag
                TotalOption = .TotalOption
                TotalMultiplier = .NMultiplier.Value
                TotalGlobalMultiplier = .NGlobalMultiplier.Value
                TotalDecimal = CInt(.NDecimal.Value)
                TotalDisplayValue = .CBDisplayValue.Checked
                TotalDisplayText = .CBDisplayText.Checked
                TotalAutofill = .CBAutoFill.Checked
            End With

            CalculateTotalPlayableNotes()
        End If
    End Sub

    Private Sub THTotal_KeyDown(sender As Object, e As KeyEventArgs) Handles THTotal.KeyDown
        TotalAutofill = False
    End Sub

    Private Sub UpdateColumnsX()
        column(0).Left = 0
        'If col(0).Width = 0 Then col(0).Visible = False

        For xI1 As Integer = 1 To UBound(column)
            column(xI1).Left = column(xI1 - 1).Left + CInt(IIf(column(xI1 - 1).isVisible, column(xI1 - 1).Width, 0))
            'If col(xI1).Width = 0 Then col(xI1).Visible = False
        Next
        HSL.Maximum = nLeft(gColumns) + column(niB).Width
        HS.Maximum = nLeft(gColumns) + column(niB).Width
        HSR.Maximum = nLeft(gColumns) + column(niB).Width
    End Sub

    Private Sub CHPlayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHPlayer.SelectedIndexChanged
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0

        iPlayer = CHPlayer.SelectedIndex
        Dim xGP2 As Boolean = iPlayer <> 0
        column(niD1).isVisible = xGP2
        column(niD2).isVisible = xGP2
        column(niD3).isVisible = xGP2
        column(niD4).isVisible = xGP2
        column(niD5).isVisible = xGP2
        column(niD6).isVisible = xGP2
        column(niD7).isVisible = xGP2
        column(niD8).isVisible = xGP2
        column(niS3).isVisible = xGP2

        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()

        If IsApplicationInitializing Then Exit Sub
        RefreshPanelAll()
    End Sub

    Private Sub CGB_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGB.ValueChanged
        gColumns = CInt(niB + CGB.Value - 1)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub TBGOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBGOptions.Click, mnGOptions.Click
        Dim xTE As Integer
        Select Case UCase(EncodingToString(TextEncoding)) ' az: wow seriously? is there really no better way? 
            Case "SYSTEM ANSI" : xTE = 0
            Case "LITTLE ENDIAN UTF16" : xTE = 1
            Case "ASCII" : xTE = 2
            Case "BIG ENDIAN UTF16" : xTE = 3
            Case "LITTLE ENDIAN UTF32" : xTE = 4
            Case "UTF7" : xTE = 5
            Case "UTF8" : xTE = 6
            Case "SJIS" : xTE = 7
            Case "EUC-KR" : xTE = 8
            Case Else : xTE = 0
        End Select

        Dim xDiag As New OpGeneral(gWheel, gPgUpDn, MiddleButtonMoveMethod, xTE, CInt(192.0R / BMSGridLimit), ErrorJackBPM, ErrorJackTH, gLNGap, TempFileName,
            AutoSaveInterval, BeepWhileSaved, PreloadBMSStruct, BPMx1296, STOPx1296, AudioLine, TemplateSnapToVPosition, PastePatternToVPosition,
            AutoFocusMouseEnter, FirstClickDisabled, ClickStopPreview)

        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With xDiag
                gWheel = .zWheel
                gPgUpDn = .zPgUpDn
                TextEncoding = .zEncoding
                'SortingMethod = .zSort
                MiddleButtonMoveMethod = .zMiddle
                AutoSaveInterval = .zAutoSave
                BMSGridLimit = 192.0R / .zGridPartition
                ErrorJackBPM = .nJackBPM.Value
                ErrorJackTH = .nJackTH.Value
                ErrorJackSpeed = 60 * 4 / .nJackBPM.Value / .nJackTH.Value
                gLNGap = .NLNGap.Value
                TempFileName = .TTemp.Text
                BeepWhileSaved = .cBeep.Checked
                PreloadBMSStruct = .cPreloadBMSStruct.Checked
                BPMx1296 = .cBpm1296.Checked
                STOPx1296 = .cStop1296.Checked
                AudioLine = .cAudioLine.Checked
                TemplateSnapToVPosition = .cTemplateSnapToVPosition.Checked
                PastePatternToVPosition = .cPastePatternToVPosition.Checked
                AutoFocusMouseEnter = .cMEnterFocus.Checked
                FirstClickDisabled = .cMClickFocus.Checked
                ClickStopPreview = .cMStopPreview.Checked
            End With
            If AutoSaveInterval > 0 Then AutoSaveTimer.Interval = AutoSaveInterval
            AutoSaveTimer.Enabled = AutoSaveInterval > 0
        End If
    End Sub

    Private Sub TBKOptions_Click(sender As Object, e As EventArgs) Handles TBKOptions.Click, mnKOptions.Click
        Dim xDiag As New OpKeybinding(Keybindings)
        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With xDiag

                ' Rename shortcutstrings
                For Each keybind In .Keybinds
                    For i = 0 To UBound(Keybindings)
                        If Keybindings(i).OpName = keybind.OpName Then
                            Keybindings(i) = keybind
                            Exit For
                        End If
                    Next
                    RenameShortcut(keybind)
                Next
            End With
        End If
    End Sub

    Private Sub RenameShortcut(ByVal keybind As Keybinding)
        Dim Kb0 = keybind.Combo(0)
        Dim KbAll = Join(keybind.Combo, ", ")
        Select Case keybind.OpName
            Case "Snap to Grid"
                RenameShortcutKb0(CGSnap.Text, Kb0)
            Case "Disable Vertical Moves"
                RenameShortcutKb0(CGDisableVertical.Text, Kb0)
            Case "Convert to Long Note"
                POBLong.ShortcutKeyDisplayString = Kb0
            Case "Convert to Short Note"
                POBShort.ShortcutKeyDisplayString = Kb0
            Case "Convert between Long and Short Note"
                POBLongShort.ShortcutKeyDisplayString = Kb0
            Case "Auto Long Note (By VPosition)"
                POBAutoLongVPosition.ShortcutKeyDisplayString = Kb0
            Case "Auto Long Note (By Column)"
                POBAutoLongColumn.ShortcutKeyDisplayString = Kb0
            Case "Undo"
                mnUndo.ShortcutKeyDisplayString = Kb0
                RenameShortcutKb0(TBUndo.Text, KbAll)
            Case "Redo"
                mnRedo.ShortcutKeyDisplayString = Kb0
                RenameShortcutKb0(TBRedo.Text, KbAll)
            Case "Cut"
                mnCut.ShortcutKeyDisplayString = Kb0
                RenameShortcutKb0(TBCut.Text, KbAll)
            Case "Copy"
                mnCopy.ShortcutKeyDisplayString = Kb0
                RenameShortcutKb0(TBCopy.Text, KbAll)
            Case "Paste"
                mnPaste.ShortcutKeyDisplayString = Kb0
                RenameShortcutKb0(TBPaste.Text, KbAll)
            Case "Select All"
                mnSelectAll.ShortcutKeyDisplayString = Kb0
            Case "Check Technical Error"
                mnTechnicalErrorCheck.ShortcutKeyDisplayString = Kb0
        End Select
    End Sub

    Private Sub RenameShortcutKb0(ByRef xStr As String, ByVal Kb As String)
        Dim IParenthesis = xStr.LastIndexOf("(")
        If IParenthesis = -1 Then Return
        xStr = xStr.Substring(0, IParenthesis) & "(" & Kb & ")"
    End Sub

    Private Sub POBLongObjNT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLongObjNT.Click
        If Not NTInput OrElse LnObj = 0 Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Ghost OrElse LnObj <> Notes(xI1).Value / 10000 OrElse Notes(xI1).LNPair = 0 Then Notes(xI1).Selected = False : Continue For

            Dim LNPair As Integer = Notes(xI1).LNPair
            Dim xLen As Double = Notes(xI1).VPosition - Notes(LNPair).VPosition
            Me.RedoLongNoteModify(Notes(LNPair), Notes(LNPair).VPosition, xLen, xUndo, xRedo)
            Notes(LNPair).Length = Notes(xI1).VPosition - Notes(LNPair).VPosition
        Next

        Me.RedoRemoveNoteSelected(True, xUndo, xRedo)
        RemoveNotes(True)

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBLongNTObj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLongNTObj.Click
        If Not NTInput OrElse LnObj = 0 Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Ghost OrElse Notes(xI1).Length = 0 Then Continue For

            Dim vPos As Double = Notes(xI1).VPosition + Notes(xI1).Length
            Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, 0, xUndo, xRedo)
            Notes(xI1).Length = 0

            Dim NoteLNObj As Note = CType(Notes(xI1), Note)
            NoteLNObj.VPosition = vPos
            NoteLNObj.Value = LnObj * 10000
            RedoAddNote(NoteLNObj, xUndo, xRedo)
            AddNote(NoteLNObj)
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub



    Private Sub POBLong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLong.Click
        If NTInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

            Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, True, xUndo, xRedo)
            Notes(xI1).LongNote = True
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBShort.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not NTInput Then
            For xI1 As Integer = 1 To UBound(Notes)
                If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

                Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, 0, xUndo, xRedo)
                Notes(xI1).LongNote = False
            Next

        Else
            For xI1 As Integer = 1 To UBound(Notes)
                If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

                Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, 0, xUndo, xRedo)
                Notes(xI1).Length = 0
            Next
        End If

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormalLong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLongShort.Click
        If NTInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

            Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Not Notes(xI1).LongNote, xUndo, xRedo)
            Notes(xI1).LongNote = Not Notes(xI1).LongNote
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBAutoLongVPosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBAutoLongVPosition.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' Change to NTInput because easier to code
        If Not NTInput Then
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertBMSE2NT()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        For i = 1 To UBound(Notes)
            If Notes(i).Selected Then Console.WriteLine(i)
        Next

        Dim xIPrev(-1) As Integer
        Dim xI1 = 1
        Do While xI1 <= UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then xI1 += 1 : Continue Do

            ' Get xICurrCol to see if LN shortening is required due to being in the same column as the next note
            Dim xICurr = GetNotesBySelectedAndVPosition(Notes(xI1).VPosition)
            Dim xICurrCol(UBound(xICurr)) As Integer
            For i = 0 To UBound(xICurrCol)
                xICurrCol(i) = Notes(xICurr(i)).ColumnIndex
            Next
            If xIPrev.Length <> 0 Then
                For Each xIPrevIndv In xIPrev
                    Dim xLen As Double = Math.Max(Notes(xI1).VPosition - Notes(xIPrevIndv).VPosition -
                                                    CDbl(IIf(xICurrCol.Contains(Notes(xIPrevIndv).ColumnIndex), gLNGap * 192 / 4, 0)),
                                                  0)

                    RedoLongNoteModify(Notes(xIPrevIndv), Notes(xIPrevIndv).VPosition, xLen, xUndo, xRedo)
                    Notes(xIPrevIndv).Length = xLen

                Next
            End If
            xIPrev = xICurr
            xI1 = xIPrev(UBound(xIPrev)) + 1
        Loop

        ' If xIPrev.Length <> 0 Then
        '     For Each xIPrevIndv In xIPrev
        '         Dim xWLWAV As Integer = CInt(Notes(xIPrevIndv).Value / 10000)
        '         If wLWAV(xWLWAV).Duration = 0 Then wLWAV(xWLWAV) = LoadDuration(ExcludeFileName(FileName) & "\" & hWAV(xWLWAV))
        ' 
        '         Dim xLen As Double = Math.Max(GetVPositionFromTime(GetTimeFromVPosition(Notes(xIPrevIndv).VPosition) +
        '                                                            wLWAV(xWLWAV).Duration) -
        '                                       Notes(xIPrevIndv).VPosition - gLNGap * 192 / 4, 0)
        ' 
        '         RedoLongNoteModify(Notes(xIPrevIndv), Notes(xIPrevIndv).VPosition, xLen, xUndo, xRedo)
        '         Notes(xIPrevIndv).Length = xLen
        '     Next
        ' End If

        If Not NTInput Then
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertNT2BMSE()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBAutoLongColumn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBAutoLongColumn.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' Change to NTInput because easier to code
        If Not NTInput Then
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertBMSE2NT()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        Dim xIPrev(UBound(gXKeyCol)) As Integer
        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Ghost OrElse Not gXKeyCol.Contains(Notes(xI1).ColumnIndex) Then Continue For

            Dim xICol = Array.IndexOf(gXKeyCol, Notes(xI1).ColumnIndex)
            If xIPrev(xICol) <> 0 Then
                Dim xLen As Double = Math.Max(Notes(xI1).VPosition - Notes(xIPrev(xICol)).VPosition - gLNGap * 192 / 4, 0)

                RedoLongNoteModify(Notes(xIPrev(xICol)), Notes(xIPrev(xICol)).VPosition, xLen, xUndo, xRedo)
                Notes(xIPrev(xICol)).Length = xLen
            End If
            xIPrev(xICol) = xI1
        Next

        ' For Each xIPrevIndv In xIPrev
        '     If xIPrevIndv <> 0 Then
        '         Dim xWLWAV As Integer = CInt(Notes(xIPrevIndv).Value / 10000)
        '         If wLWAV(xWLWAV).Duration = 0 Then wLWAV(xWLWAV) = LoadDuration(ExcludeFileName(FileName) & "\" & hWAV(xWLWAV))
        ' 
        '         Dim xLen As Double = Math.Max(GetVPositionFromTime(GetTimeFromVPosition(Notes(xIPrevIndv).VPosition) +
        '                                                            wLWAV(xWLWAV).Duration) -
        '                                       Notes(xIPrevIndv).VPosition - gLNGap * 192 / 4, 0)
        ' 
        '         RedoLongNoteModify(Notes(xIPrevIndv), Notes(xIPrevIndv).VPosition, xLen, xUndo, xRedo)
        '         Notes(xIPrevIndv).Length = xLen
        '     End If
        ' Next

        If Not NTInput Then
            RedoRemoveNoteAll(False, xUndo, xRedo)
            ConvertNT2BMSE()
            RedoAddNoteAll(False, xUndo, xRedo)
        End If

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHidden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHidden.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Landmine OrElse Notes(xI1).Ghost Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), True, True, xUndo, xRedo)
            Notes(xI1).Hidden = True
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), False, True, xUndo, xRedo)
            Notes(xI1).Hidden = False
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHiddenVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHiddenVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Landmine OrElse Notes(xI1).Ghost Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), Not Notes(xI1).Hidden, True, xUndo, xRedo)
            Notes(xI1).Hidden = Not Notes(xI1).Hidden
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBModify_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles POBModify.Click
        Dim xNum As Boolean = False
        Dim xLbl As Boolean = False
        Dim xI1 As Integer

        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected AndAlso IsColumnNumeric(Notes(xI1).ColumnIndex) Then xNum = True : Exit For
        Next
        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected AndAlso Not IsColumnNumeric(Notes(xI1).ColumnIndex) Then xLbl = True : Exit For
        Next
        If Not (xNum Or xLbl) Then Exit Sub

        If xNum Then
            Dim xD1 As Long = CLng(InputBox(Strings.Messages.PromptEnterNumeric, Text)) * 10000
            If Not xD1 = 0 Then
                If xD1 <= 0 Then xD1 = 1

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                For xI1 = 1 To UBound(Notes)
                    If Not IsColumnNumeric(Notes(xI1).ColumnIndex) Or Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

                    Me.RedoRelabelNote(Notes(xI1), xD1, xUndo, xRedo)
                    Notes(xI1).Value = xD1
                Next
                AddUndo(xUndo, xBaseRedo.Next)
            End If
        End If

        If xLbl Then
            Dim xStr As String = UCase(Trim(InputBox(Strings.Messages.PromptEnter, Me.Text)))

            If Len(xStr) <> 0 Then
                Dim Valid As Boolean = True

                If xStr = "00" Or xStr = "0" Then Valid = False
                If Not Len(xStr) = 1 And Not Len(xStr) = 2 Then Valid = False

                Dim xI3 As Integer = Asc(Mid(xStr, 1, 1))
                If Not ((xI3 >= 48 And xI3 <= 57) Or (xI3 >= 65 And xI3 <= 90)) Then Valid = False
                If Len(xStr) = 2 Then
                    Dim xI4 As Integer = Asc(Mid(xStr, 2, 1))
                    If Not ((xI4 >= 48 And xI4 <= 57) Or (xI4 >= 65 And xI4 <= 90)) Then Valid = False
                End If

                If Valid Then
                    Dim xVal As Integer = C36to10(xStr) * 10000

                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                    Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                    For xI1 = 1 To UBound(Notes)
                        If IsColumnNumeric(Notes(xI1).ColumnIndex) Or Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

                        Me.RedoRelabelNote(Notes(xI1), xVal, xUndo, xRedo)
                        Notes(xI1).Value = xVal
                    Next
                    AddUndo(xUndo, xBaseRedo.Next)
                Else
                    MsgBox(Strings.Messages.InvalidLabel, MsgBoxStyle.Critical, Strings.Messages.Err)
                End If
            End If
        End If

        RefreshPanelAll()
    End Sub

    Private Sub POBMirror_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBMirror.Click
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        'xRedo &= sCmdKM(niA1, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(niA7) - RealColumnToEnabled(niA1), 0, True) & vbCrLf
        'xUndo &= sCmdKM(niA7, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(niA1) - RealColumnToEnabled(niA7), 0, True) & vbCrLf

        ' gXKeyCol: Unmodified array
        ' Array 1: Modified array based on range
        ' Array R: Array 1 reversed

        ' New function: Declare an array to see the range of selected notes. B columns ignored.
        Dim xRangeL As Integer = niB ' Big number
        Dim xRangeU As Integer = 0 ' Smol number

        ' Range finder
        For xI1 = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For
            If xRangeL > Notes(xI1).ColumnIndex Then xRangeL = Notes(xI1).ColumnIndex
            If xRangeU < Notes(xI1).ColumnIndex Then xRangeU = Notes(xI1).ColumnIndex
        Next

        Dim xniArray1(UBound(gXKeyCol)) As Integer
        Dim xIA1 As Integer = -1
        For Each ni In gXKeyCol
            If xRangeL <= ni AndAlso ni <= xRangeU Then
                xIA1 += 1
                xniArray1(xIA1) = ni
            End If
        Next
        If xIA1 <= 0 Then Exit Sub

        ReDim Preserve xniArray1(xIA1)

        Dim xniArrayR = xniArray1.Reverse()
        Dim xniArrayLen = xniArray1.Length

        Dim xCol As Integer
        For xI1 = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

            xCol = Notes(xI1).ColumnIndex
            ' MsgBox("Test" & "xCol: " & xCol & " xI1: " & xI1)
            For xI2 = 0 To xniArrayLen - 1
                ' MsgBox("Test 2 xI2: " & xI2)
                If xCol = xniArray1(xI2) Then
                    xCol = xniArrayR(xI2)
                    Exit For
                End If

            Next

            Me.RedoMoveNote(Notes(xI1), xCol, Notes(xI1).VPosition, xUndo, xRedo)
            Notes(xI1).ColumnIndex = xCol
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub


    Private Sub POBFlip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBFlip.Click
        If gXKeyMode <> "DP" Then Exit Sub
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' Array 1: Unmodified array
        ' Array R: Flipped array

        Dim xniArray1 = New Integer() {niA1, niA2, niA3, niA4, niA5, niA6, niA7, niA8, niD1, niD2, niD3, niD4, niD5, niD6, niD7, niD8}
        Dim xniArrayR = New Integer() {niD8, niD1, niD2, niD3, niD4, niD5, niD6, niD7, niA2, niA3, niA4, niA5, niA6, niA7, niA8, niA1}

        Dim xniArrayLen = xniArray1.Length

        Dim xCol As Integer
        For xI1 = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Or Notes(xI1).Ghost Then Continue For

            xCol = Notes(xI1).ColumnIndex
            For xI2 = 0 To xniArrayLen - 1
                If xCol = xniArray1(xI2) Then
                    xCol = xniArrayR(xI2)
                    Exit For
                End If

            Next

            Me.RedoMoveNote(Notes(xI1), xCol, Notes(xI1).VPosition, xUndo, xRedo)
            Notes(xI1).ColumnIndex = xCol
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBRandomAll(sender As Object, e As EventArgs) Handles POBRandom.Click, POBRRandom.Click, POBSRandom.Click, POBHRandom.Click
        Dim NameS As String = CType(sender, ToolStripMenuItem).Name

        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' gXKeyCol: Unmodified array
        ' Array 1: Modified array based on range
        ' Array R: Array 1 randomized

        ' New function: Declare an array to see the range of selected notes. B columns ignored.
        Dim xRangeL As Integer = niB ' Big number
        Dim xRangeU As Integer = 0 ' Smol number

        ' Range finder
        For xI1 = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For
            If xRangeL > Notes(xI1).ColumnIndex Then xRangeL = Notes(xI1).ColumnIndex
            If xRangeU < Notes(xI1).ColumnIndex Then xRangeU = Notes(xI1).ColumnIndex
        Next

        Dim xniArray1(UBound(gXKeyCol)) As Integer
        Dim xIA1 As Integer = -1
        For Each ni In gXKeyCol
            If xRangeL <= ni AndAlso ni <= xRangeU Then
                xIA1 += 1
                xniArray1(xIA1) = ni
            End If
        Next
        If xIA1 <= 0 Then Exit Sub

        ReDim Preserve xniArray1(xIA1)

        Dim xniArrayR() As Integer = CType(xniArray1.Clone(), Integer())

        Select Case NameS
            Case "POBRandom", "POBRRandom"
                ' Shuffle columns
                If NameS = "POBRandom" Then
                    Shuffle(xniArrayR)
                Else
                    Dim R As Integer = CInt(Math.Floor(xniArrayR.Length * Rnd()))
                    Dim M As Integer = CInt(Math.Floor(2 * Rnd()) * 2 - 1)
                    For i = 0 To UBound(xniArrayR)
                        xniArrayR(i) = xniArray1((i * M + R + xniArrayR.Length) Mod xniArrayR.Length)
                    Next
                End If

                ' Move notes
                Dim xCol As Integer
                For xI1 = 1 To UBound(Notes)
                    With Notes(xI1)
                        If Not .Selected Or .Ghost Then Continue For

                        xCol = .ColumnIndex
                        For xI2 = 0 To xniArray1.Length - 1
                            If xCol = xniArray1(xI2) Then
                                xCol = xniArrayR(xI2)
                                Exit For
                            End If

                        Next

                        Me.RedoMoveNote(Notes(xI1), xCol, .VPosition, xUndo, xRedo)
                        .ColumnIndex = xCol
                    End With
                Next
            Case "POBSRandom", "POBHRandom"
                ' Find array of indexes of selected notes in the same vPosition
                xI1 = 1
                Dim xI1Arr(-1) As Integer
                Dim vPos As Double
                Dim xI1ArrPrevUBound As Integer = -1 ' Used for HRandom
                ' Find the first index of selected notes
                Do While xI1 <= UBound(Notes)
                    If Not Notes(xI1).Selected Or Notes(xI1).Ghost Or Notes(xI1).ColumnIndex < xniArray1(0) Or Notes(xI1).ColumnIndex > xniArray1(UBound(xniArray1)) Then xI1 += 1 : Continue Do
                    ' Begin building array until vPosition changes
                    vPos = Notes(xI1).VPosition
                    Do While xI1 <= UBound(Notes) AndAlso Notes(xI1).VPosition = vPos
                        If Not Notes(xI1).Selected Or Notes(xI1).Ghost Or Notes(xI1).ColumnIndex < xniArray1(0) Or Notes(xI1).ColumnIndex > xniArray1(UBound(xniArray1)) Then xI1 += 1 : Continue Do
                        ReDim Preserve xI1Arr(xI1Arr.Length)
                        xI1Arr(UBound(xI1Arr)) = xI1
                        xI1 += 1
                    Loop

                    ' Shuffle columns
                    If NameS = "POBSRandom" Or xI1ArrPrevUBound = -1 Then ' If SRandom or for HRandom, if it's the first set of notes
                        Shuffle(xniArrayR)
                    Else
                        Dim xniArrayR1(xI1ArrPrevUBound) As Integer
                        Dim xniArrayR2(UBound(xniArray1) - xI1ArrPrevUBound - 1) As Integer
                        For i = 0 To UBound(xniArrayR1)
                            xniArrayR1(i) = xniArrayR(i)
                        Next
                        For i = 0 To UBound(xniArrayR2)
                            xniArrayR2(i) = xniArrayR(i + xI1ArrPrevUBound + 1)
                        Next
                        Shuffle(xniArrayR1)
                        Shuffle(xniArrayR2)
                        For i = 0 To UBound(xniArrayR2)
                            xniArrayR(i) = xniArrayR2(i)
                        Next
                        For i = UBound(xniArrayR2) + 1 To UBound(xniArrayR)
                            xniArrayR(i) = xniArrayR1(i - UBound(xniArrayR2) - 1)
                        Next
                    End If

                    ' Move notes
                    For xI2 = 0 To UBound(xI1Arr)
                        Dim xI2I = xI1Arr(xI2)
                        Me.RedoMoveNote(Notes(xI2I), xniArrayR(xI2), Notes(xI2I).VPosition, xUndo, xRedo)
                        Notes(xI2I).ColumnIndex = xniArrayR(xI2)
                    Next
                    xI1ArrPrevUBound = UBound(xI1Arr)
                    ReDim xI1Arr(-1)
                Loop
        End Select

        AddUndo(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBSort_Click(sender As Object, e As EventArgs) Handles POBSort.Click
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' Array 1: Unmodified array
        Dim xniArray1() As Integer = gXKeyCol
        Dim niBMax As Integer = niB

        Dim xniArrayLen = xniArray1.Length

        Dim vPos As Double
        Dim xIArray(0) As Integer
        Dim xValueArray(0) As Long
        Dim xITemp As Integer

        ' Find array of indexes of selected notes in the same vPosition
        xI1 = 1
        Dim xI1Arr(-1) As Integer
        ' Find the first index of selected notes
        Do While xI1 <= UBound(Notes)
            If Not Notes(xI1).Selected OrElse Notes(xI1).Ghost OrElse Not IsColumnSound(Notes(xI1).ColumnIndex) Then xI1 += 1 : Continue Do
            ' Begin building array until vPosition changes too much
            vPos = Notes(xI1).VPosition
            Do While xI1 <= UBound(Notes) AndAlso Math.Abs(GetTimeFromVPosition(Notes(xI1).VPosition) - GetTimeFromVPosition(vPos)) <= ErrorJackSpeed
                If Not Notes(xI1).Selected OrElse Notes(xI1).Ghost OrElse Not IsColumnSound(Notes(xI1).ColumnIndex) Then xI1 += 1 : Continue Do
                ReDim Preserve xI1Arr(xI1Arr.Length)
                xI1Arr(UBound(xI1Arr)) = xI1
                xI1 += 1
            Loop

            ' Extend xniArray1 if need be
            If xI1Arr.Length > xniArray1.Length Then
                Dim xUB0 As Integer = UBound(xniArray1)
                ReDim Preserve xniArray1(UBound(xI1Arr))
                For i = xUB0 + 1 To UBound(xniArray1)
                    xniArray1(i) = niBMax
                    niBMax += 1
                Next
            End If

            ' Sort columns, insertion sort
            For xI2 = 1 To UBound(xI1Arr)
                For xI3 = xI2 To 1 Step -1
                    If Notes(xI1Arr(xI3 - 1)).Value > Notes(xI1Arr(xI3)).Value Then
                        xITemp = xI1Arr(xI3 - 1)
                        xI1Arr(xI3 - 1) = xI1Arr(xI3)
                        xI1Arr(xI3) = xITemp
                    End If
                Next
            Next

            ' Move notes
            For xI2 = 0 To UBound(xI1Arr)
                Dim xI2I = xI1Arr(xI2)
                RedoMoveNote(Notes(xI2I), xniArray1(xI2), Notes(xI2I).VPosition, xUndo, xRedo)
                Notes(xI2I).ColumnIndex = xniArray1(xI2)
            Next
            ReDim xI1Arr(-1)
        Loop

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionQuick(0, UBound(Notes))
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Function GetNotesBySelectedAndVPosition(ByVal VPos As Double, Optional xI As Integer = 1) As Integer()
        Dim xIArr(-1) As Integer
        For xIN = xI To UBound(Notes)
            If Not Notes(xIN).Selected OrElse Notes(xIN).Ghost Then Continue For

            If Notes(xIN).VPosition = VPos Then
                ReDim Preserve xIArr(xIArr.Length)
                xIArr(UBound(xIArr)) = xIN
            ElseIf Notes(xIN).VPosition > VPos Then
                Exit For
            End If
        Next
        Return xIArr
    End Function

    Private Function GetNotesBySelectedAndVPosition(ByVal xIN As Integer, Optional xI As Integer = 1) As Integer()
        Return GetNotesBySelectedAndVPosition(Notes(xIN).VPosition, xI)
    End Function

    Private Sub TBMyO2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBMyO2.Click, mnMyO2.Click
        Dim xDiag As New dgMyO2
        xDiag.Show()
    End Sub


    Private Sub TBFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBFind.Click, mnFind.Click
        Dim xDiag As New dgFind(gColumns)
        xDiag.Show()
    End Sub

    ''' <Summary>
    ''' Checks if the note satisfies the input measure, label, and value ranges.
    ''' </Summary>

    Private Function fdrCheck(ByVal xNote As Note) As Boolean
        Return xNote.VPosition >= MeasureBottom(fdriMesL) And xNote.VPosition < MeasureBottom(fdriMesU) + MeasureLength(fdriMesU) AndAlso
               CBool(IIf(IsColumnNumeric(xNote.ColumnIndex),
                   xNote.Value >= fdriValL And xNote.Value <= fdriValU,
                   xNote.Value >= fdriLblL And xNote.Value <= fdriLblU)) AndAlso
               Array.IndexOf(fdriCol, xNote.ColumnIndex - 1) <> -1
        ' lol Fixed
    End Function

    Private Function fdrRangeS(ByVal xbLim1 As Boolean, ByVal xbLim2 As Boolean, ByVal xVal As Boolean) As Boolean
        Return (Not xbLim1 And xbLim2 And xVal) Or (xbLim1 And Not xbLim2 And Not xVal) Or (xbLim1 And xbLim2)
    End Function

    Public Sub fdrFind(ByVal iRange() As Boolean,
                         ByVal xMesL As Integer, ByVal xMesU As Integer,
                         ByVal xLblL As String, ByVal xLblU As String,
                         ByVal xValL As Integer, ByVal xValU As Integer,
                         ByVal iCol() As Integer, ByVal fdrFunction As String, Optional xReplaceLbl As String = Nothing, Optional xReplaceVal As Integer = Nothing)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange(0)
        Dim xbUnsel As Boolean = iRange(1)
        Dim xbShort As Boolean = iRange(2)
        Dim xbLong As Boolean = iRange(3)
        Dim xbHidden As Boolean = iRange(4)
        Dim xbVisible As Boolean = iRange(5)
        Dim xbNoError As Boolean = iRange(6)
        Dim xbError As Boolean = iRange(7)
        Dim xbNotComment As Boolean = iRange(8)
        Dim xbComment As Boolean = iRange(9)

        Select Case fdrFunction
            Case "TBSelect"
                Dim xSel(UBound(Notes)) As Boolean
                For xI1 As Integer = 1 To UBound(Notes)
                    xSel(xI1) = Notes(xI1).Selected
                Next

                'Main process
                For xI1 As Integer = 1 To UBound(Notes)
                    With Notes(xI1)
                        If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) Then

                            .Selected = fdrCheck(Notes(xI1))
                        End If
                    End With
                Next
            Case "TBUnselect"
                Dim xSel(UBound(Notes)) As Boolean
                For xI1 As Integer = 1 To UBound(Notes)
                    xSel(xI1) = Notes(xI1).Selected
                Next

                'Main process
                For xI1 As Integer = 1 To UBound(Notes)
                    With Notes(xI1)
                        If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) Then

                            .Selected = Not fdrCheck(Notes(xI1))
                        End If
                    End With
                Next
            Case "FNotePrevious"
                Dim xSel(UBound(Notes)) As Boolean
                For xI1 As Integer = 1 To UBound(Notes)
                    xSel(xI1) = Notes(xI1).Selected
                    Notes(xI1).Selected = False
                Next

                For xI1 = UBound(Notes) To 1 Step -1
                    With Notes(xI1)
                        If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) AndAlso
                            fdrCheck(Notes(xI1)) AndAlso
                            .VPosition < -PanelVScroll(PanelFocus) Then

                            PanelVScroll(PanelFocus) = - .VPosition
                            .Selected = True
                            Exit For
                        End If
                    End With
                Next
            Case "FNoteNext"
                Dim xSel(UBound(Notes)) As Boolean
                For xI1 As Integer = 1 To UBound(Notes)
                    xSel(xI1) = Notes(xI1).Selected
                    Notes(xI1).Selected = False
                Next

                For xI1 = 1 To UBound(Notes)
                    With Notes(xI1)
                        If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) AndAlso
                            fdrCheck(Notes(xI1)) AndAlso
                            .VPosition > -PanelVScroll(PanelFocus) Then

                            PanelVScroll(PanelFocus) = - .VPosition
                            .Selected = True
                            Exit For
                        End If
                    End With
                Next
            Case "TBDelete"
                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                'Main process
                Dim xI1 As Integer = 1
                Do While xI1 <= UBound(Notes)
                    With Notes(xI1)
                        If (Not .Ghost) AndAlso
                            ((xbSel And .Selected) Or (xbUnsel And Not .Selected)) AndAlso
                            fdrCheck(Notes(xI1)) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) Then

                            RedoRemoveNote(Notes(xI1), xUndo, xRedo)
                            RemoveNote(xI1, False)
                        Else
                            xI1 += 1
                        End If
                    End With
                Loop

                AddUndo(xUndo, xBaseRedo.Next)
                SortByVPositionInsertion()
                UpdatePairing()
                CalculateTotalPlayableNotes()
            Case "TBrl"
                Dim xxLbl As Integer = C36to10(xReplaceLbl) * 10000

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                'Main process
                For xI1 As Integer = 1 To UBound(Notes)
                    With Notes(xI1)
                        If (Not .Ghost) AndAlso
                            ((xbSel And .Selected) Or (xbUnsel And Not .Selected)) AndAlso
                            fdrCheck(Notes(xI1)) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            Not IsColumnNumeric(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) Then

                            Me.RedoRelabelNote(Notes(xI1), xxLbl, xUndo, xRedo)
                            .Value = xxLbl
                        End If
                    End With
                Next

                AddUndo(xUndo, xBaseRedo.Next)
            Case "TBrv"
                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                'Main process
                For xI1 As Integer = 1 To UBound(Notes)
                    With Notes(xI1)
                        If (Not .Ghost) AndAlso
                            ((xbSel And .Selected) Or (xbUnsel And Not .Selected)) AndAlso
                            fdrCheck(Notes(xI1)) AndAlso
                            nEnabled(.ColumnIndex) AndAlso
                            IsColumnNumeric(.ColumnIndex) AndAlso
                            fdrRangeS(xbShort, xbLong, CBool(IIf(NTInput, .Length, .LongNote))) AndAlso
                            fdrRangeS(xbVisible, xbHidden, .Hidden) AndAlso
                            fdrRangeS(xbNoError, xbError, .HasError) AndAlso
                            fdrRangeS(xbNotComment, xbComment, .Comment) Then

                            Me.RedoRelabelNote(Notes(xI1), xReplaceVal, xUndo, xRedo)
                            .Value = xReplaceVal
                        End If
                    End With
                Next

                AddUndo(xUndo, xBaseRedo.Next)
        End Select

        RefreshPanelAll()
        Beep()
    End Sub

    Private Sub MInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MInsert.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)
        Dim xVP As Double = MeasureBottom(xMeasure)

        If NTInput Then
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) >= 999 Then
                    Me.RedoRemoveNote(Notes(xI1), xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            Dim xdVP As Double
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP And Notes(xI1).VPosition + Notes(xI1).Length <= MeasureBottom(999) Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition + xMLength, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength

                ElseIf Notes(xI1).VPosition >= xVP Then
                    xdVP = MeasureBottom(999) - 1 - Notes(xI1).VPosition - Notes(xI1).Length
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition + xMLength, Notes(xI1).Length + xdVP, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength
                    Notes(xI1).Length += xdVP

                ElseIf Notes(xI1).VPosition + Notes(xI1).Length >= xVP Then
                    xdVP = CDbl(IIf(Notes(xI1).VPosition + Notes(xI1).Length > MeasureBottom(999) - 1, GetMaxVPosition() - 1 - Notes(xI1).VPosition - Notes(xI1).Length, xMLength))
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Notes(xI1).Length + xdVP, xUndo, xRedo)
                    Notes(xI1).Length += xdVP
                End If
            Next

        Else
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) >= 999 Then
                    Me.RedoRemoveNote(Notes(xI1), xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition + xMLength, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength
                End If
            Next
        End If

        For xI1 As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(xI1) = MeasureLength(xI1 - 1)
        Next
        UpdateMeasureBottom()

        AddUndo(xUndo, xBaseRedo.Next)
        UpdatePairing()
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
    End Sub

    Private Sub MRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MRemove.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)
        Dim xVP As Double = MeasureBottom(xMeasure)

        If NTInput Then
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) = xMeasure And MeasureAtDisplacement(Notes(xI1).VPosition + Notes(xI1).Length) = xMeasure Then
                    Me.RedoRemoveNote(Notes(xI1), xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            Dim xdVP As Double
            xVP = MeasureBottom(xMeasure)
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP + xMLength Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition - xMLength, xUndo, xRedo)
                    Notes(xI1).VPosition -= xMLength

                ElseIf Notes(xI1).VPosition >= xVP Then
                    xdVP = xMLength + xVP - Notes(xI1).VPosition
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition + xdVP - xMLength, Notes(xI1).Length - xdVP, xUndo, xRedo)
                    Notes(xI1).VPosition += xdVP - xMLength
                    Notes(xI1).Length -= xdVP

                ElseIf Notes(xI1).VPosition + Notes(xI1).Length >= xVP Then
                    xdVP = CDbl(IIf(Notes(xI1).VPosition + Notes(xI1).Length >= xVP + xMLength, xMLength, Notes(xI1).VPosition + Notes(xI1).Length - xVP + 1))
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Notes(xI1).Length - xdVP, xUndo, xRedo)
                    Notes(xI1).Length -= xdVP
                End If
            Next

        Else
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) = xMeasure Then
                    Me.RedoRemoveNote(Notes(xI1), xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            xVP = MeasureBottom(xMeasure)
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition - xMLength, xUndo, xRedo)
                    Notes(xI1).VPosition -= xMLength
                End If
            Next
        End If

        For xI1 As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(xI1 - 1) = MeasureLength(xI1)
        Next
        UpdateMeasureBottom()

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
    End Sub

    Private Sub TBThemeDef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeDef.Click
        Dim xTempFileName As String = My.Application.Info.DirectoryPath & "\____TempFile.Theme.xml"
        My.Computer.FileSystem.WriteAllText(xTempFileName, My.Resources.O2Mania_Theme, False, System.Text.Encoding.Unicode)
        LoadSettings(xTempFileName)
        System.IO.File.Delete(xTempFileName)

        RefreshPanelAll()
    End Sub

    Private Sub TBThemeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeSave.Click
        Dim xDiag As New SaveFileDialog
        xDiag.Filter = Strings.FileType.THEME_XML & "|*.Theme.xml"
        xDiag.DefaultExt = "Theme.xml"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Data"
        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        Me.SaveSettings(xDiag.FileName, True)
        If BeepWhileSaved Then Beep()
        TBThemeRefresh_Click(TBThemeRefresh, New System.EventArgs)
    End Sub

    Private Sub TBThemeRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeRefresh.Click
        For xI1 As Integer = cmnTheme.Items.Count - 1 To 5 Step -1
            Try
                cmnTheme.Items.RemoveAt(xI1)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles("*.Theme.xml")
        For Each xStr As FileInfo In xFileNames
            cmnTheme.Items.Add(xStr.Name, Nothing, AddressOf LoadTheme)
        Next
    End Sub

    Private Sub TBThemeLoadComptability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeLoadComptability.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType.TH & "|*.th"
        xDiag.DefaultExt = "th"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath
        If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\Theme") Then xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Theme"
        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        Me.LoadThemeComptability(xDiag.FileName)
        RefreshPanelAll()
    End Sub

    ''' <summary>
    ''' Will return Double.PositiveInfinity if canceled.
    ''' </summary>
    Private Function InputBoxDouble(ByVal Prompt As String, ByVal LBound As Double, ByVal UBound As Double, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As Double
        Dim xStr As String = InputBox(Prompt, Title, DefaultResponse)
        If xStr = "" Then Return Double.PositiveInfinity

        If Not Double.TryParse(xStr, InputBoxDouble) Then Return Double.PositiveInfinity
        If InputBoxDouble > UBound Then InputBoxDouble = UBound
        If InputBoxDouble < LBound Then InputBoxDouble = LBound
    End Function

    Private Sub FSSS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSS.Click
        Dim xMax As Double = CDbl(IIf(vSelLength > 0, GetMaxVPosition() - vSelLength, GetMaxVPosition))
        Dim xMin As Double = CDbl(IIf(vSelLength < 0, -vSelLength, 0))
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelStart.ToString())
        If xDouble = Double.PositiveInfinity Then Return

        vSelStart = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSL.Click
        Dim xMax As Double = GetMaxVPosition() - vSelStart
        Dim xMin As Double = -vSelStart
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelLength.ToString())
        If xDouble = Double.PositiveInfinity Then Return

        vSelLength = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSH.Click
        Dim xMax As Double = CDbl(IIf(vSelLength > 0, vSelLength, 0))
        Dim xMin As Double = CDbl(IIf(vSelLength > 0, 0, -vSelLength))
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelHalf.ToString())
        If xDouble = Double.PositiveInfinity Then Return

        vSelHalf = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BVCReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCReverse.Click
        vSelStart += vSelLength
        vSelHalf -= vSelLength
        vSelLength *= -1
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub AutoSaveTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoSaveTimer.Tick
        Dim xTime As Date = Now
        Dim xFileName As String
        With xTime
            xFileName = My.Application.Info.DirectoryPath & "\AutoSave_" &
                              .Year & "_" & .Month & "_" & .Day & "_" & .Hour & "_" & .Minute & "_" & .Second & "_" & .Millisecond & ".IBMSC"
        End With
        'My.Computer.FileSystem.WriteAllText(xFileName, SaveiBMSC, False, System.Text.Encoding.Unicode)
        SaveiBMSC(xFileName)

        On Error Resume Next
        If PreviousAutoSavedFileName <> "" Then IO.File.Delete(PreviousAutoSavedFileName)
        On Error GoTo 0

        PreviousAutoSavedFileName = xFileName
    End Sub

    Private Sub CWAVMultiSelect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CWAVMultiSelect.CheckedChanged
        WAVMultiSelect = CWAVMultiSelect.Checked
        LWAV.SelectionMode = CType(IIf(WAVMultiSelect, SelectionMode.MultiExtended, SelectionMode.One), SelectionMode)
    End Sub

    Private Sub CWAVChangeLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CWAVChangeLabel.CheckedChanged
        WAVChangeLabel = CWAVChangeLabel.Checked
    End Sub

    Private Sub BWAVUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVUp.Click
        If LWAV.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 0 To 1294
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For xI1 As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hWAV(xI1 + 1)
                hWAV(xI1 + 1) = hWAV(xI1)
                hWAV(xI1) = xStr

                LWAV.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hWAV(xI1 + 1)
                LWAV.Items.Item(xI1 - 1) = C10to36(xI1) & ": " & hWAV(xI1)

                If WAVChangeLabel Then
                    Dim xL1 As String = C10to36(xI1)
                    Dim xL2 As String = C10to36(xI1 + 1)
                    For xI2 As Integer = 1 To UBound(Notes)
                        If Not IsColumnSound(Notes(xI2).ColumnIndex) Then Continue For

                        If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 10000

                        ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000

                        End If
                    Next

                End If
                xIndices(xIndex) += -1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BWAVDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVDown.Click
        If LWAV.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 1294 To 0 Step -1
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For xI1 As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hWAV(xI1 + 1)
                hWAV(xI1 + 1) = hWAV(xI1 + 2)
                hWAV(xI1 + 2) = xStr

                LWAV.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hWAV(xI1 + 1)
                LWAV.Items.Item(xI1 + 1) = C10to36(xI1 + 2) & ": " & hWAV(xI1 + 2)

                If WAVChangeLabel Then
                    Dim xL1 As String = C10to36(xI1 + 2)
                    Dim xL2 As String = C10to36(xI1 + 1)
                    For xI2 As Integer = 1 To UBound(Notes)
                        If Not IsColumnSound(Notes(xI2).ColumnIndex) Then Continue For

                        If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 10000

                        ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 20000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 20000

                        End If
                    Next

                End If
                xIndices(xIndex) += 1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BWAVBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVBrowse.Click
        Dim xDWAV As New OpenFileDialog
        xDWAV.DefaultExt = "wav"
        xDWAV.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDWAV.InitialDirectory = CStr(IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)))
        xDWAV.Multiselect = WAVMultiSelect

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)

        AddToPOWAV(xDWAV.FileNames)
    End Sub

    Private Sub BWAVRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVRemove.Click
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)
        For xI1 As Integer = 0 To UBound(xIndices)
            hWAV(xIndices(xI1) + 1) = ""
            LWAV.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": "
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BWAVDuplicate_Click(sender As Object, e As EventArgs) Handles BWAVDuplicate.Click
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)

        ' Duplicate WAVs
        Dim xWAVDup(2 * (xIndices(UBound(xIndices)) - xIndices(0) + 1)) As String
        ' Duplicate first selected WAV
        xWAVDup(0) = hWAV(xIndices(0) + 1)
        xWAVDup(1) = hWAV(xIndices(0) + 1)
        Dim xID As Integer = 1
        For xI1 = 1 To UBound(xIndices)
            ' If skipped indices, add same amount of skipped indices
            If xIndices(xI1) - xIndices(xI1 - 1) > 1 Then
                For xISkip = 2 To xIndices(xI1) - xIndices(xI1 - 1)
                    xID += 1
                    xWAVDup(xID) = ""
                Next
            End If

            ' Continue duplicating selected WAVs
            For xI2 = 0 To 1
                xID += 1
                xWAVDup(xID) = hWAV(xIndices(xI1) + 1)
            Next
        Next
        ReDim Preserve xWAVDup(xID)

        ' Add duplicated WAV list to hWAV
        Dim xIL0 As Integer = LWAV.SelectedIndex
        Dim xIL1 As Integer
        For xI3 = 0 To xID
            xIL1 = xIL0 + xI3
            If xIL1 > LWAV.Items.Count - 1 Then xIL1 = LWAV.Items.Count - 1 : Exit For
            hWAV(xIL1 + 1) = xWAVDup(xI3)
            LWAV.Items.Item(xIL1) = C10to36(xIL1 + 1) & ": " & xWAVDup(xI3)
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 = xIL0 To xIL1
            LWAV.SelectedIndices.Add(xI1)
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub


    Private Sub CBMPMultiSelect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBMPMultiSelect.CheckedChanged
        BMPMultiSelect = CBMPMultiSelect.Checked
        LBMP.SelectionMode = CType(IIf(BMPMultiSelect, SelectionMode.MultiExtended, SelectionMode.One), SelectionMode)
    End Sub

    Private Sub CBMPChangeLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBMPChangeLabel.CheckedChanged
        BMPChangeLabel = CBMPChangeLabel.Checked
    End Sub

    Private Sub BBMPUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPUp.Click
        If LBMP.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 0 To 1294
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For xI1 As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hBMP(xI1 + 1)
                hBMP(xI1 + 1) = hBMP(xI1)
                hBMP(xI1) = xStr

                LBMP.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hBMP(xI1 + 1)
                LBMP.Items.Item(xI1 - 1) = C10to36(xI1) & ": " & hBMP(xI1)

                If BMPChangeLabel Then
                    Dim xL1 As String = C10to36(xI1)
                    Dim xL2 As String = C10to36(xI1 + 1)
                    For xI2 As Integer = 1 To UBound(Notes)
                        If IsColumnNumeric(Notes(xI2).ColumnIndex) Then Continue For

                        If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 10000

                        ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000

                        End If
                    Next

                End If
                xIndices(xIndex) += -1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBMPDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPDown.Click
        If LBMP.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 1294 To 0 Step -1
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For xI1 As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hBMP(xI1 + 1)
                hBMP(xI1 + 1) = hBMP(xI1 + 2)
                hBMP(xI1 + 2) = xStr

                LBMP.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hBMP(xI1 + 1)
                LBMP.Items.Item(xI1 + 1) = C10to36(xI1 + 2) & ": " & hBMP(xI1 + 2)

                If BMPChangeLabel Then


                    Dim xL1 As String = C10to36(xI1 + 2)
                    Dim xL2 As String = C10to36(xI1 + 1)
                    For xI2 As Integer = 1 To UBound(Notes)
                        If IsColumnNumeric(Notes(xI2).ColumnIndex) Then Continue For

                        If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 10000

                        ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                            Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 20000, xUndo, xRedo)
                            Notes(xI2).Value = xI1 * 10000 + 20000

                        End If
                    Next

                End If
                xIndices(xIndex) += 1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBMPBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPBrowse.Click
        Dim xDBMP As New OpenFileDialog
        xDBMP.Filter = Strings.FileType._im & "|*.bmp;*.png;*.jpg;*.gif;*.mpg;*.mpeg;*.avi;*.m1v;*.m2v;*.m4v;*.mp4;*.webm;*.wmv|" &
                       Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.gif|" &
                       Strings.FileType._movie & "|*.mpg;*.mpeg;*.avi;*.m1v;*.m2v;*.m4v;*.mp4;*.webm;*.wmv|" &
                       Strings.FileType._all & "|*.*"
        xDBMP.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()
        xDBMP.DefaultExt = "png"
        xDBMP.Multiselect = BMPMultiSelect

        If xDBMP.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDBMP.FileName)

        AddToPOBMP(xDBMP.FileNames)
    End Sub

    Private Sub BBMPRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPRemove.Click
        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)
        For xI1 As Integer = 0 To UBound(xIndices)
            hBMP(xIndices(xI1) + 1) = ""
            LBMP.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": "
        Next

        LBMP.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(xI1))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mnMain.MouseDown ', TBMain.MouseDown  ', pttl.MouseDown, pIsSaved.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, &H112, &HF012, 0)
            If e.Clicks = 2 Then
                If Me.WindowState = FormWindowState.Maximized Then Me.WindowState = FormWindowState.Normal Else Me.WindowState = FormWindowState.Maximized
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            'mnSys.Show(sender, e.Location)
        End If
    End Sub

    Private Sub mnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSelectAll.Click
        If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = nEnabled(Notes(xI1).ColumnIndex)
        Next
        If TBTimeSelect.Checked Then
            CalculateGreatestVPosition()
            vSelStart = 0
            vSelLength = MeasureBottom(MeasureAtDisplacement(GreatestVPosition)) + MeasureLength(MeasureAtDisplacement(GreatestVPosition))
        End If
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnDelete.Click
        If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Me.RedoRemoveNoteSelected(True, xUndo, xRedo)
        RemoveNotes(True)

        AddUndo(xUndo, xBaseRedo.Next)
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("http://www.cs.mcgill.ca/~ryang6/iBMSC/")
    End Sub

    Private Sub mnUpdateC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("http://bbs.rohome.net/thread-1074065-1-1.html")
    End Sub

    Private Sub mnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnQuit.Click
        Close()
    End Sub


    Private Sub EnableDWM()
        mnMain.BackColor = Color.Black
        'TBMain.BackColor = Color.FromArgb(64, 64, 64)

        For Each xmn As ToolStripMenuItem In mnMain.Items
            xmn.ForeColor = Color.White
            AddHandler xmn.DropDownClosed, AddressOf mn_DropDownClosed
            AddHandler xmn.DropDownOpened, AddressOf mn_DropDownOpened
            AddHandler xmn.MouseEnter, AddressOf mn_MouseEnter
            AddHandler xmn.MouseLeave, AddressOf mn_MouseLeave
        Next
    End Sub

    Private Sub DisableDWM()
        mnMain.BackColor = SystemColors.Control
        'TBMain.BackColor = SystemColors.Control

        For Each xmn As ToolStripMenuItem In mnMain.Items
            xmn.ForeColor = SystemColors.ControlText
            RemoveHandler xmn.DropDownClosed, AddressOf mn_DropDownClosed
            RemoveHandler xmn.DropDownOpened, AddressOf mn_DropDownOpened
            RemoveHandler xmn.MouseEnter, AddressOf mn_MouseEnter
            RemoveHandler xmn.MouseLeave, AddressOf mn_MouseLeave
        Next
    End Sub

    Private Sub ttlIcon_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'ttlIcon.Image = My.Resources.icon2_16
        'mnSys.Show(ttlIcon, 0, ttlIcon.Height)
    End Sub
    Private Sub ttlIcon_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        'ttlIcon.Image = My.Resources.icon2_16_highlight
    End Sub
    Private Sub ttlIcon_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        'ttlIcon.Image = My.Resources.icon2_16
    End Sub

    Private Sub mnSMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSMenu.CheckedChanged
        mnMain.Visible = mnSMenu.Checked
    End Sub
    Private Sub mnSTB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSTB.CheckedChanged
        TBMain.Visible = mnSTB.Checked
    End Sub
    Private Sub mnSOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSOP.CheckedChanged
        POptions.Visible = mnSOP.Checked
    End Sub
    Private Sub mnSStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSStatus.CheckedChanged
        pStatus.Visible = mnSStatus.Checked
    End Sub
    Private Sub mnSLSplitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSLSplitter.CheckedChanged
        SpL.Visible = mnSLSplitter.Checked
    End Sub
    Private Sub mnSRSplitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSRSplitter.CheckedChanged
        SpR.Visible = mnSRSplitter.Checked
    End Sub
    Private Sub CGShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShow.CheckedChanged
        gShowGrid = CGShow.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowS.CheckedChanged
        gShowSubGrid = CGShowS.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowBG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowBG.CheckedChanged
        gShowBG = CGShowBG.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowM.CheckedChanged
        gShowMeasureNumber = CGShowM.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowV.CheckedChanged
        gShowVerticalLine = CGShowV.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowMB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowMB.CheckedChanged
        gShowMeasureBar = CGShowMB.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowC.CheckedChanged
        gShowC = CGShowC.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGBLP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBLP.CheckedChanged
        gDisplayBGAColumn = CGBLP.Checked

        column(niBGA).isVisible = gDisplayBGAColumn
        column(niLAYER).isVisible = gDisplayBGAColumn
        column(niPOOR).isVisible = gDisplayBGAColumn
        column(niS4).isVisible = gDisplayBGAColumn

        If IsApplicationInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub
    Private Sub CGSCROLL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSCROLL.CheckedChanged
        gSCROLL = CGSCROLL.Checked

        column(niSCROLL).isVisible = gSCROLL

        If IsApplicationInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub
    Private Sub CGSTOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSTOP.CheckedChanged
        gSTOP = CGSTOP.Checked

        column(niSTOP).isVisible = gSTOP

        If IsApplicationInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub
    Private Sub CGBPM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBPM.CheckedChanged
        'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        'Dim xRedo As UndoRedo.LinkedURCmd = Nothing
        'Me.RedoChangeVisibleColumns(gBLP, gSTOP, iPlayer, gBLP, CGSTOP.Checked, iPlayer, xUndo, xRedo)
        gBPM = CGBPM.Checked

        column(niBPM).isVisible = gBPM

        If IsApplicationInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub CGDisableVertical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDisableVertical.CheckedChanged
        DisableVerticalMove = CGDisableVertical.Checked
    End Sub

    Private Sub CBeatPreserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBeatPreserve.Click, CBeatMeasure.Click, CBeatCut.Click, CBeatScale.Click
        'If Not sender.Checked Then Exit Sub
        Dim RadioS As RadioButton = CType(sender, RadioButton)
        Dim xBeatList() As RadioButton = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
        BeatChangeMode = Array.IndexOf(Of RadioButton)(xBeatList, RadioS)
        'For xI1 As Integer = 0 To mnBeat.Items.Count - 1
        'If xI1 <> BeatChangeMode Then CType(mnBeat.Items(xI1), ToolStripMenuItem).Checked = False
        'Next
        'sender.Checked = True
    End Sub


    Private Sub tBeatValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tBeatValue.LostFocus
        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then tBeatValue.BackColor = Color.FromArgb(&HFFFFC0C0) Else tBeatValue.BackColor = Nothing

            tBeatValue.Text = a.ToString()
        End If
    End Sub



    Private Sub ApplyBeat(ByVal xRatio As Double, ByRef xUndo As UndoRedo.LinkedURCmd, ByRef xRedo As UndoRedo.LinkedURCmd, Optional xN As Integer = -1, Optional xD As Integer = -1)
        SortByVPositionInsertion()
        If xN = -1 AndAlso xD = -1 Then
            If xRatio = 1 Then
                xD = CInt(CGSub.Value)
                xN = xD
            Else
                xD = CInt(GetDenominator(xRatio))
                xN = CInt(xRatio * xD)
            End If
        End If

        Me.RedoChangeMeasureLengthSelected(192 * xRatio, xUndo, xRedo)

        Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer
        LBeat.SelectedIndices.CopyTo(xIndices, 0)


        For Each xI1 As Integer In xIndices
            Dim dLength As Double = xRatio * 192.0R - MeasureLength(xI1)
            Dim dRatio As Double = xRatio * 192.0R / MeasureLength(xI1)

            Dim xBottom As Double = 0
            For xI2 As Integer = 0 To xI1 - 1
                xBottom += MeasureLength(xI2)
            Next
            Dim xUpBefore As Double = xBottom + MeasureLength(xI1)
            Dim xUpAfter As Double = xUpBefore + dLength

            Select Case BeatChangeMode ' CBeatPreserve => Case 0
                Case 1
case2:              Dim xI0 As Integer

                    If NTInput Then
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                            If Notes(xI0).VPosition + Notes(xI0).Length >= xUpBefore Then
                                Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, xUndo, xRedo)
                                Notes(xI0).Length += dLength
                            End If
                        Next
                    Else
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                        Next
                    End If

                    For xI9 As Integer = xI0 To UBound(Notes)
                        Me.RedoLongNoteModify(Notes(xI9), Notes(xI9).VPosition + dLength, Notes(xI9).Length, xUndo, xRedo)
                        Notes(xI9).VPosition += dLength
                    Next

                Case 2
                    If dLength < 0 Then
                        If NTInput Then
                            Dim xI0 As Integer = 1
                            Dim xU As Integer = UBound(Notes)
                            Do While xI0 <= xU
                                If Notes(xI0).VPosition < xUpAfter Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length >= xUpAfter And Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Dim nLen As Double = xUpAfter - Notes(xI0).VPosition - 1.0R
                                        Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                    End If
                                ElseIf Notes(xI0).VPosition < xUpBefore Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Me.RedoRemoveNote(Notes(xI0), xUndo, xRedo)
                                        RemoveNote(xI0)
                                        xI0 -= 1
                                        xU -= 1
                                    Else
                                        Dim nLen As Double = Notes(xI0).Length - xUpBefore + Notes(xI0).VPosition
                                        Me.RedoLongNoteModify(Notes(xI0), xUpBefore, nLen, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                        Notes(xI0).VPosition = xUpBefore
                                    End If
                                End If
                                xI0 += 1
                            Loop
                        Else
                            Dim xI0 As Integer
                            Dim xI9 As Integer
                            For xI0 = 1 To UBound(Notes)
                                If Notes(xI0).VPosition >= xUpAfter Then Exit For
                            Next
                            For xI9 = xI0 To UBound(Notes)
                                If Notes(xI9).VPosition >= xUpBefore Then Exit For
                            Next

                            For xI8 As Integer = xI0 To xI9 - 1
                                Me.RedoRemoveNote(Notes(xI8), xUndo, xRedo)
                            Next
                            For xI8 As Integer = xI9 To UBound(Notes)
                                Notes(xI8 - xI9 + xI0) = Notes(xI8)
                            Next
                            ReDim Preserve Notes(UBound(Notes) - xI9 + xI0)
                        End If
                    End If

                    GoTo case2

                Case 3
                    If NTInput Then
                        For xI0 As Integer = 1 To UBound(Notes)
                            If Notes(xI0).VPosition < xBottom Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, xUndo, xRedo)
                                    Notes(xI0).Length += dLength
                                ElseIf Notes(xI0).VPosition + Notes(xI0).Length > xBottom Then
                                    Dim nLen As Double = (Notes(xI0).Length + Notes(xI0).VPosition - xBottom) * dRatio + xBottom - Notes(xI0).VPosition
                                    Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                End If
                            ElseIf Notes(xI0).VPosition < xUpBefore Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Dim nLen As Double = (xUpBefore - Notes(xI0).VPosition) * dRatio + Notes(xI0).VPosition + Notes(xI0).Length - xUpBefore
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Me.RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                Else
                                    Dim nLen As Double = Notes(xI0).Length * dRatio
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Me.RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                End If
                            Else
                                Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition + dLength, Notes(xI0).Length, xUndo, xRedo)
                                Notes(xI0).VPosition += dLength
                            End If
                        Next
                    Else
                        Dim xI0 As Integer
                        Dim xI9 As Integer
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xBottom Then Exit For
                        Next
                        For xI9 = xI0 To UBound(Notes)
                            If Notes(xI9).VPosition >= xUpBefore Then Exit For
                        Next

                        For xI8 As Integer = xI0 To xI9 - 1
                            Dim nVP As Double = (Notes(xI8).VPosition - xBottom) * dRatio + xBottom
                            Me.RedoLongNoteModify(Notes(xI0), nVP, Notes(xI0).Length, xUndo, xRedo)
                            Notes(xI8).VPosition = nVP
                        Next

                        'GoTo case2

                        For xI8 As Integer = xI9 To UBound(Notes)
                            Me.RedoLongNoteModify(Notes(xI8), Notes(xI8).VPosition + dLength, Notes(xI8).Length, xUndo, xRedo)
                            Notes(xI8).VPosition += dLength
                        Next
                    End If

            End Select

            MeasureLength(xI1) = xRatio * 192.0R
            LBeat.Items(xI1) = Add3Zeros(xI1) & ": " & xRatio & " ( " & xN & " / " & xD & " ) "
        Next
        UpdateMeasureBottom()

        LBeat.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LBeat.SelectedIndices.Add(xIndices(xI1))
        Next

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
    End Sub

    Private Sub BBeatApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBeatApply.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xxD As Integer = CInt(nBeatD.Value)
        Dim xxN As Integer = CInt(nBeatN.Value)
        Dim xxRatio As Double = xxN / xxD

        ApplyBeat(xxRatio, xUndo, xRedo, xxN, xxD)
        AddUndo(xUndo, xBaseRedo.Next)

        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBeatApplyV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBeatApplyV.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then System.Media.SystemSounds.Hand.Play() : Exit Sub

            Dim xxD As Long = GetDenominator(a)

            ApplyBeat(a, xUndo, xRedo)
        End If
        AddUndo(xUndo, xBaseRedo.Next)

        RefreshPanelAll()
        POStatusRefresh()
    End Sub


    Private Sub BHStageFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BHStageFile.Click, BHBanner.Click, BHBackBMP.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.gif|" &
                       Strings.FileType._all & "|*.*"
        xDiag.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)).ToString()
        xDiag.DefaultExt = "png"

        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDiag.FileName)

        If [Object].ReferenceEquals(sender, BHStageFile) Then
            THStageFile.Text = GetFileName(xDiag.FileName)
        ElseIf [Object].ReferenceEquals(sender, BHBanner) Then
            THBanner.Text = GetFileName(xDiag.FileName)
        ElseIf [Object].ReferenceEquals(sender, BHBackBMP) Then
            THBackBMP.Text = GetFileName(xDiag.FileName)
        End If
    End Sub

    Private Sub Switches_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    POHeaderSwitch.CheckedChanged,
    POGridSwitch.CheckedChanged,
    POWaveFormSwitch.CheckedChanged,
    POWAVSwitch.CheckedChanged,
    POBMPSwitch.CheckedChanged,
    POBeatSwitch.CheckedChanged,
    POExpansionSwitch.CheckedChanged

        Try
            Dim Source As CheckBox = CType(sender, CheckBox)
            Dim Target As Panel = Nothing

            If Object.ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf Object.ReferenceEquals(sender, POHeaderSwitch) Then : Target = POHeaderInner
            ElseIf Object.ReferenceEquals(sender, POGridSwitch) Then : Target = POGridInner
            ElseIf Object.ReferenceEquals(sender, POWaveFormSwitch) Then : Target = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POWAVSwitch) Then : Target = POWAVInner
            ElseIf Object.ReferenceEquals(sender, POBMPSwitch) Then : Target = POBMPInner
            ElseIf Object.ReferenceEquals(sender, POBeatSwitch) Then : Target = POBeatInner
            ElseIf Object.ReferenceEquals(sender, POExpansionSwitch) Then : Target = POExpansionInner
            End If

            If Source.Checked Then
                Target.Visible = True
            Else
                Target.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Expanders_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    POHeaderExpander.CheckedChanged,
    POGridExpander.CheckedChanged,
    POWaveFormExpander.CheckedChanged,
    POWAVExpander.CheckedChanged,
    POBMPExpander.CheckedChanged,
    POBeatExpander.CheckedChanged

        Try
            Dim Source As CheckBox = CType(sender, CheckBox)
            Dim Target As Panel = Nothing
            'Dim TargetParent As Panel = Nothing

            If Object.ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf Object.ReferenceEquals(sender, POHeaderExpander) Then : Target = POHeaderPart2 ' : TargetParent = POHeaderInner
            ElseIf Object.ReferenceEquals(sender, POGridExpander) Then : Target = POGridPart2 ' : TargetParent = POGridInner
            ElseIf Object.ReferenceEquals(sender, POWaveFormExpander) Then : Target = POWaveFormPart2 ' : TargetParent = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POWAVExpander) Then : Target = POWAVPart2 ' : TargetParent = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POBMPExpander) Then : Target = POBMPPart2 ' : TargetParent = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POBeatExpander) Then : Target = POBeatPart2 ' : TargetParent = POWaveFormInner
            End If

            If Source.Checked Then
                Target.Visible = True
                'Source.Image = My.Resources.Collapse
            Else
                Target.Visible = False
                'Source.Image = My.Resources.Expand
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub VerticalResizer_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseDown, POBMPResizer.MouseDown, POBeatResizer.MouseDown, POExpansionResizer.MouseDown
        tempResize = e.Y
    End Sub

    Private Sub HorizontalResizer_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POptionsResizer.MouseDown, SpL.MouseDown, SpR.MouseDown
        tempResize = e.X
    End Sub

    Private Sub POResizer_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseMove, POBMPResizer.MouseMove, POBeatResizer.MouseMove, POExpansionResizer.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.Y = tempResize Then Exit Sub

        Try
            Dim Source As Button = CType(sender, Button)
            Dim Target As Panel = CType(Source.Parent, Panel)

            Dim xHeight As Integer = Target.Height + e.Y - tempResize
            If xHeight < 10 Then xHeight = 10
            Target.Height = xHeight

            Target.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub POptionsResizer_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POptionsResizer.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = POptionsScroll.Width - e.X + tempResize
            If xWidth < 25 Then xWidth = 25
            POptionsScroll.Width = xWidth

            Me.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SpR_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SpR.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = PMainR.Width - e.X + tempResize
            If xWidth < 0 Then xWidth = 0
            PMainR.Width = xWidth

            Me.ToolStripContainer1.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SpL_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SpL.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = PMainL.Width + e.X - tempResize
            If xWidth < 0 Then xWidth = 0
            PMainL.Width = xWidth

            Me.ToolStripContainer1.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub mnGotoMeasure_Click(sender As Object, e As EventArgs) Handles mnGotoMeasure.Click
        Dim s = InputBox(Strings.Messages.PromptEnterMeasure, "Enter Measure")

        ' Interpolation, Value = F + (C - F) * (i - Floor(i))
        Dim i As Double
        If Double.TryParse(s, i) AndAlso i > 0 AndAlso i < 999 Then
            Dim iCe As Integer = CInt(Math.Ceiling(i))
            Dim iFl As Integer = CInt(Math.Floor(i))
            PanelVScroll(PanelFocus) = -(MeasureBottom(iFl) + (MeasureBottom(iCe) - MeasureBottom(iFl)) * (i - iFl))
        End If
    End Sub


    ' Generic shuffle for basic type arrays
    Public Function Shuffle(Of T)(items As T()) As T()
        Dim temp As T
        Dim j As Int32

        For i = items.Count To 2 Step -1
            ' Pick an item for position i.
            j = CInt(Math.Floor(i * Rnd()))
            ' Swap 
            i -= 1
            temp = items(i)
            items(i) = items(j)
            items(j) = temp
        Next
        Return items
    End Function

    Private Sub Expand_Load(sender As Object, e As EventArgs) Handles BExpansion.Click
        If Not TExpansion.Enabled Then Exit Sub

        ReDim ExpansionSplit(2)
        Dim xDOp As New OpExpand()
        ExpansionSplit(1) = "-"
        Try
            xDOp.ShowDialog(Me)
        Catch
            Exit Sub
        End Try
    End Sub

    Public Sub Expand_DisplayGhostNotes(Optional xGhostMode As Integer = 0)
        Select Case GhostMode
            Case 0
                GhostMode = xGhostMode
            Case 1
                GhostMode = 0
            Case 2
                Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveWarning & Strings.Messages.GhostNotesShowMain, MsgBoxStyle.YesNo)
                If xResult = MsgBoxResult.No Then Exit Sub
                If xResult = MsgBoxResult.Yes Then SaveBMS()
                SwapGhostNotes()
                GhostMode = 0
        End Select
        OpenBMS(ExpansionSplit(1), True)
    End Sub

    Public Sub Expand_ModifyNotes()
        Select Case GhostMode
            Case 2
                SwapGhostNotes()
        End Select
        RemoveGhostNotes()
        GhostMode = 2
        OpenBMS(ExpansionSplit(1), True)
        SwapGhostNotes()
    End Sub

    Public Sub Expand_ModifySection()
        ' TODO: Revise to the now available tab style
        RemoveGhostNotes()
        GhostMode = 0
        Dim RandomTempFilePath = ExcludeFileName(FileName) & "\" & RandomTempFileName
        ' Picks another random filename because the programme somehow generated the same exact RandomFileName as a previous instance. 1 in 2-billion chance btw
        Do Until Not My.Computer.FileSystem.FileExists(RandomTempFilePath)
            RandomTempFileName = "___TempRandom" & GenerateRandomString(6, False) & ".bmsc"
            RandomTempFilePath = ExcludeFileName(FileName) & "\" & RandomTempFileName
        Loop
        ExpansionSplit(1) = GenerateHeaderMeta() & GenerateHeaderIndexedData() & ExpansionSplit(1)
        ' Dim xStrHeader As String = GenerateHeaderMeta()
        ' xStrHeader &= GenerateHeaderIndexedData()
        My.Computer.FileSystem.WriteAllText(RandomTempFilePath, ExpansionSplit(1), False, TextEncoding)
        Process.Start(My.Application.Info.DirectoryPath & "\" & My.Application.Info.ProductName & ".exe", RandomTempFilePath)
        TimerExternalExpansion.Enabled = True
        TExpansion.Enabled = False
    End Sub

    Private Sub TimerExternalExpansion_Tick(sender As Object, e As EventArgs) Handles TimerExternalExpansion.Tick
        Dim ReadText As String = Nothing
        Dim RandomTempFilePath = ExcludeFileName(FileName) & "\" & RandomTempFileName
        ReadText = My.Computer.FileSystem.ReadAllText(RandomTempFilePath, TextEncoding)
        If Not ReadText.EndsWith("*---------------------- RANDOM DATA FIELD") Then Exit Sub

        TExpansion.Enabled = True
        ExpansionSplit(1) = ""
        TExpansion.Text = ""
        Dim xStrCompare() As String = Split(Replace(Replace(Replace(SaveBMS(), vbLf, vbCr), vbCr & vbCr, vbCr), vbCr, vbCrLf), vbCrLf,, CompareMethod.Text)
        For Each xStrLine In Split(ReadText, vbCrLf)
            If (Not xStrCompare.Contains(xStrLine) AndAlso xStrLine <> "*---------------------- RANDOM DATA FIELD") Or
                        SWIC(xStrLine, "#RANDOM") Or SWIC(xStrLine, "#IF") Or SWIC(xStrLine, "#ENDIF") Then
                ExpansionSplit(1) &= xStrLine & vbCrLf
            End If
        Next
        TExpansion.Text = Join(ExpansionSplit, vbCrLf)
        AddTempFileList(RandomTempFilePath)
        TimerExternalExpansion.Enabled = False
    End Sub

    Public Sub Expand_RemoveGhostNotes()
        Select Case GhostMode
            Case 2
                SwapGhostNotes()
        End Select
        RemoveGhostNotes()
        GhostMode = 0
    End Sub

    Public Sub RemoveGhostNotes()
        Dim xI1 As Integer = 1
        Do While xI1 <= UBound(Notes)
            If Notes(xI1).Ghost Then
                RemoveNote(xI1)
            Else
                xI1 += 1
            End If
        Loop
    End Sub

    Public Sub RemoveCommentNotes()
        Dim xI1 As Integer = 1
        Do While xI1 <= UBound(Notes)
            If Notes(xI1).Comment Then
                RemoveNote(xI1)
            Else
                xI1 += 1
            End If
        Loop
    End Sub

    Public Sub SwapGhostNotes()
        For xI1 = 1 To UBound(Notes)
            Notes(xI1).Ghost = Not Notes(xI1).Ghost
        Next
    End Sub

    Public Function GenerateRandomString(ByRef len As Integer, ByRef upper As Boolean) As String
        Dim rand As New Random()
        Dim allowableChars() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
        Dim final As String = String.Empty
        For i As Integer = 0 To len - 1
            final += allowableChars(rand.Next(allowableChars.Length - 1))
        Next

        Return IIf(upper, final.ToUpper(), final).ToString()
    End Function

    Private Sub TExpansion_Click(sender As Object, e As EventArgs) Handles TExpansion.Click
        Select Case GhostMode
            Case 1
                Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveWarning & Strings.Messages.GhostNotesModifyExpansion1, MsgBoxStyle.YesNo)
                If xResult = MsgBoxResult.No Then PMain.Focus() : Exit Sub
                GhostMode = 0
            Case 2
                Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveWarning & Strings.Messages.GhostNotesModifyExpansion2, MsgBoxStyle.YesNo)
                If xResult = MsgBoxResult.No Then PMain.Focus() : Exit Sub
                SaveBMS()
                Expand_RemoveGhostNotes()
        End Select
    End Sub

    Public Sub AddCommentLine(ByVal xComment As String)
        For i = 1 To UBound(hCOM)
            If IsNothing(hCOM(i)) Then hCOM(i) = xComment : hCOMNum = i : Exit For
        Next
    End Sub

    Public Sub RemoveCommentLine(ByVal xI As Integer)
        Do While hCOM(xI) <> ""
            If xI = hCOM.Length Then
                hCOM(xI) = ""
                Exit Do
            Else
                hCOM(xI) = hCOM(xI + 1)
            End If
            xI += 1
        Loop
        hCOMNum = xI - 1
    End Sub

    Private Sub TimerPreviewNote_Tick(sender As Object, e As EventArgs) Handles TimerPreviewNote.Tick
        Dim TimeNow = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds
        If TimeNow > InternalPlayTimerEnd Then TimerPreviewNote.Enabled = False
        InternalPlayTimerCount = TimeNow - InternalPlayTimerStart
        RefreshPanelAll()
    End Sub

    Private Sub TimerRefreshPanel_Tick(sender As Object, e As EventArgs) Handles TimerRefreshPanel.Tick
        RefreshPanelAll()
    End Sub
End Class