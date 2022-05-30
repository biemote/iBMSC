Imports System.Linq
Imports iBMSC.Editor

Partial Public Class MainWindow

    ' BMS File tabs
    ''' <summary>
    ''' Everything related to file tabs are put here.
    ''' 
    ''' Currently, Untitled.bms must be at the last of the tab list.
    ''' </summary>

    Dim BMSFiles As BMSFile()
    Dim BMSFileIndex As Integer = 0

    ''' <summary>
    ''' Info about a BMS File.
    ''' </summary>

    Structure BMSFile
        Public Filename As String
        Public RandomSource As String
        Public Struct As BMSStruct

        Public TSB As ToolStripButton
        Public TabColor As Color

        ''' <param name="xFilename">Filename.</param>
        ''' <param name="xRandomSource">Source filename to edit #RANDOM section for.</param>
        ''' <param name="xStruct">BMS Structure with all information about a chart such as notes and #WAV list.</param>
        ''' <param name="xTSB">Tool strip button, or a tab.</param>
        ''' <param name="xTabColor">Color of the tab (when unchecked).</param>
        Public Sub New(xFilename As String, xTSB As ToolStripButton, xTabColor As Color, Optional xStruct As BMSStruct = Nothing)
            Filename = xFilename
            Struct = xStruct

            TSB = xTSB
            TabColor = xTabColor
        End Sub

        Public Sub AddRandomSource(xPath As String)
            RandomSource = xPath
        End Sub

        Public Function IsInitialized() As Boolean
            Return Struct.Notes IsNot Nothing
        End Function

        Public Function IsSaved() As Boolean
            If IsInitialized() Then
                Return Struct.IsSaved
            Else
                Return True
            End If
        End Function
    End Structure

    Structure BMSStruct
        Public Notes() As Note
        Public NotesTemplate() As Note
        Public hWAV() As String
        Public hBMP() As String
        Public hBPM() As Long
        Public hSTOP() As Long
        Public hBMSCROLL() As Long
        Public hCOM() As String
        Public wLWAV() As WavSample
        Public HeaderT() As String ' Text
        Public HeaderN() As Decimal ' Numeric
        Public HeaderI() As Integer ' Integer
        Public Expansion As String
        Public MeasureLength() As Double
        Public FileNameTemplate As String

        Public ExpansionSplit() As String
        Public GhostMode As Integer

        Public sUndo() As UndoRedo.LinkedURCmd
        Public sRedo() As UndoRedo.LinkedURCmd
        Public sI As Integer

        Public ExpansionEnabled As Boolean
        Public IsSaved As Boolean
        Public NTInput As Boolean
        Public WaveformLoaded As Boolean

        Public Sub New(xNotes() As Note, xNotesTemplate() As Note,
                       xWAV() As String, xBMP() As String, xBPM() As Long, xSTOP() As Long, xBMSCROLL() As Long, xCOM() As String, xLWAV() As WavSample,
                       xHeaderT() As String, xHeaderN() As Decimal, xHeaderI() As Integer, xExpansion As String, xMeasureLength() As Double, xFileNameTemplate As String,
                       xExpansionSplit() As String, xGhostMode As Integer,
                       xUndo() As UndoRedo.LinkedURCmd, xRedo() As UndoRedo.LinkedURCmd, xSI As Integer,
                       xExpansionEnabled As Boolean, xIsSaved As Boolean, xNTInput As Boolean, xWaveformLoaded As Boolean)

            Notes = xNotes
            NotesTemplate = xNotesTemplate
            hWAV = xWAV
            hBMP = xBMP
            hBPM = xBPM
            hSTOP = xSTOP
            hBMSCROLL = xBMSCROLL
            hCOM = xCOM
            wLWAV = xLWAV
            HeaderT = xHeaderT
            HeaderN = xHeaderN
            HeaderI = xHeaderI
            Expansion = xExpansion
            MeasureLength = xMeasureLength
            FileNameTemplate = xFileNameTemplate

            ExpansionSplit = xExpansionSplit
            GhostMode = xGhostMode

            sUndo = xUndo
            sRedo = xRedo
            sI = xSI

            ExpansionEnabled = xExpansionEnabled
            IsSaved = xIsSaved
            NTInput = xNTInput
            WaveformLoaded = xWaveformLoaded
        End Sub
    End Structure

    Private Function TBClose_Click(sender As Object, e As EventArgs) As Boolean Handles mnClose.Click
        If ClosingPopSave() Then Return False

        If BMSFileIndex = UBound(BMSFiles) Or BMSFileIndex = -1 Then TBNew_Click(Nothing, Nothing) : Return False

        If BMSFiles(BMSFileIndex).RandomSource IsNot Nothing Then
            CloseFileWithRandomSource()

        Else
            Dim xIRemove As Integer = BMSFileIndex
            TBTab_Click(BMSFiles(BMSFileIndex + 1).TSB, New EventArgs)
            RemoveBMSFile(xIRemove)
            SetBMSFileIndex(xIRemove)

        End If

        Return True
    End Function

    Private Sub TBTab_Click(sender As Object, e As EventArgs)
        Dim TSBS As ToolStripButton = CType(sender, ToolStripButton)
        If TSBS.Checked Then Exit Sub

        TimerPreviewNote.Enabled = False
        SaveBMSStruct()

        Dim i As Integer = FindBMSTabIndex(TSBS)
        SetBMSFileIndex(i)

        If BMSFiles(BMSFileIndex).IsInitialized Then
            SetFileName(BMSFiles(BMSFileIndex).Filename)
            LoadBMSStruct(BMSFileIndex)
        Else
            If BMSFiles(BMSFileIndex).Filename = FileNameInit Then
                TBNew_Click(Nothing, Nothing)
            Else
                ReadFile(BMSFiles(BMSFileIndex).Filename)
            End If
        End If
    End Sub

    Private Sub TBTab_MouseDown(sender As Object, e As MouseEventArgs)
        Dim TSBCurrent = BMSFiles(BMSFileIndex).TSB
        Dim TSBSender = CType(sender, ToolStripButton)
        Dim xIClicked = FindBMSTabIndex(TSBSender)

        If e.Button = MouseButtons.Middle Then
            If xIClicked = BMSFileIndex Then
                TBClose_Click(Nothing, Nothing)

            Else
                TBTab_Click(BMSFiles(xIClicked).TSB, New EventArgs)
                TBClose_Click(Nothing, Nothing)
                TBTab_Click(TSBCurrent, New EventArgs)

            End If

        ElseIf e.Button = MouseButtons.Right Then
            Dim xColorPicker As New ColorPicker
            xColorPicker.SetOrigColor(TSBSender.BackColor)
            If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
            ColorTSBChange(TSBSender, xColorPicker.NewColor)
            BMSFiles(xIClicked).TabColor = xColorPicker.NewColor
        End If
    End Sub

    Private Sub TBTab_MouseMove(sender As Object, e As MouseEventArgs)
        Dim TSBS = CType(sender, ToolStripButton)
        Dim xITab = FindBMSTabIndex(TSBS)
        If Not BMSFiles(xITab).IsInitialized Then Exit Sub
        Dim BannerDir = ExcludeFileName(BMSFiles(xITab).Filename) & "\" & BMSFiles(xITab).Struct.HeaderT(8)
        If Not My.Computer.FileSystem.FileExists(BannerDir) Then
            BannerDir = ExcludeFileName(BMSFiles(xITab).Filename) & "\" & BMSFiles(xITab).Struct.HeaderT(7)
            If Not My.Computer.FileSystem.FileExists(BannerDir) Then Exit Sub
        End If

        With PBOnTabHover
            .Image = New Bitmap(BannerDir)
            .Size = .Image.Size
            ' .Location = e.Location
            ' .Parent = Me
            .Visible = True
        End With
    End Sub

    Private Sub TBTab_MouseLeave(sender As Object, e As EventArgs)
        PBOnTabHover.Visible = False
    End Sub

    Public Sub AddBMSFiles(xPaths As String())
        For xI = 0 To UBound(xPaths)
            NewRecent(xPaths(xI))
            AddBMSFile(xPaths(xI))
        Next
    End Sub

    Public Sub AddBMSFiles(xPaths As String)
        NewRecent(xPaths)
        AddBMSFile(xPaths)
    End Sub

    Private Sub AddBMSFile(xPath As String)
        Dim i As Integer = FindBMSTabIndex(xPath)
        If i <> -1 Then
            SetBMSFileIndex(i)

        Else
            If BMSFileIndex = UBound(BMSFiles) AndAlso xPath <> FileNameInit Then BMSFileIndex -= 1
            ReDim Preserve BMSFiles(BMSFiles.Length)

            For xI = UBound(BMSFiles) - 1 To BMSFileIndex + 1 Step -1
                BMSFiles(xI + 1) = BMSFiles(xI)
            Next

            BMSFileIndex += 1
            With BMSFiles(BMSFileIndex)
                .Filename = xPath
                .TabColor = System.Drawing.SystemColors.Control
                .TSB = NewBMSTab(xPath)
                .Struct = Nothing
            End With

            ' Re-add buttons to TBTab
            For i = 1 To TBTab.Items.Count
                TBTab.Items.RemoveAt(0)
            Next
            For i = 0 To UBound(BMSFiles)
                TBTab.Items.Add(BMSFiles(i).TSB)
            Next
            SetBMSFileIndex(BMSFileIndex)
        End If
    End Sub

    Private Sub ColorTSBChange(ByVal xTSB As ToolStripButton, ByVal c As Color) ' Copied from OpPlayer
        xTSB.BackColor = c
        xTSB.ForeColor = CType(IIf(CInt(c.GetBrightness * 255) + 255 - c.A >= 128, Color.Black, Color.White), Color)
    End Sub

    Private Function FindBMSTabIndex(ByVal xTSB As ToolStripButton) As Integer
        For i = 0 To UBound(BMSFiles)
            If BMSFiles(i).TSB Is xTSB Then Return i
        Next
        Return -1
    End Function

    Private Function FindBMSTabIndex(ByVal xStr As String) As Integer
        For i = 0 To UBound(BMSFiles)
            If BMSFiles(i).Filename = xStr Then Return i
        Next
        Return -1
    End Function

    Private Sub InitializeBMSFiles()
        If BMSFiles IsNot Nothing Then
            Dim BMSFileListCheck(UBound(BMSFiles)) As BMSFile
            Dim i = -1
            For Each BMS In BMSFiles
                If My.Computer.FileSystem.FileExists(BMS.Filename) OrElse BMS.Filename = FileNameInit Then
                    i += 1
                    BMSFileListCheck(i) = BMS
                End If
            Next
            BMSFiles = CType(BMSFileListCheck.Clone(), BMSFile())
            ReDim Preserve BMSFiles(i)
        Else
            ReDim BMSFiles(0)
            BMSFiles(0) = New BMSFile(FileNameInit, NewBMSTab(FileNameInit), System.Drawing.SystemColors.Control)
        End If
    End Sub

    Private Function NewBMSTab(xPath As String) As ToolStripButton ' Optional xColor as Color = Color.Empty is not accepted
        Dim xTSB As New ToolStripButton
        With xTSB
            .Image = My.Resources.x16Blank
            .Name = GetFileName(xPath)
            .Text = GetFileName(xPath)
        End With
        AddHandler xTSB.Click, AddressOf TBTab_Click
        AddHandler xTSB.MouseDown, AddressOf TBTab_MouseDown
        AddHandler xTSB.MouseMove, AddressOf TBTab_MouseMove
        AddHandler xTSB.MouseLeave, AddressOf TBTab_MouseLeave
        Return xTSB
    End Function

    Private Sub RemoveBMSFile(xI As Integer)
        For i = xI To UBound(BMSFiles) - 1
            BMSFiles(i) = BMSFiles(i + 1)
        Next
        ReDim Preserve BMSFiles(UBound(BMSFiles) - 1)
        TBTab.Items.RemoveAt(xI)
    End Sub

    Private Sub SetBMSFileIndex(xI As Integer)
        BMSFileIndex = xI
        For i = 0 To UBound(BMSFiles)
            If i = BMSFileIndex Then
                BMSFiles(i).TSB.Checked = True
            Else
                BMSFiles(i).TSB.Checked = False
            End If
        Next
    End Sub

    Private Sub SaveBMSStruct(Optional xI As Integer = -1)
        If xI = -1 Then xI = BMSFileIndex
        ' Console.WriteLine(FileName)
        ' Console.WriteLine(MeasureLength(0))
        ' If BMSFileStructs(0).MeasureLength IsNot Nothing Then Console.WriteLine("BMSStruct 0, MeasureLength 0: " & BMSFileStructs(0).MeasureLength(0))
        Dim HeaderT() As String = {THTitle.Text, THArtist.Text, THGenre.Text, THPlayLevel.Text, THTotal.Text,
                                   THSubTitle.Text, THSubArtist.Text, THStageFile.Text, THBanner.Text, THBackBMP.Text, THExRank.Text, THComment.Text}
        Dim HeaderN() As Decimal = {THBPM.Value}
        Dim HeaderI() As Integer = {CHPlayer.SelectedIndex, CHRank.SelectedIndex, CHDifficulty.SelectedIndex, CHLnObj.SelectedIndex}

        BMSFiles(xI).Struct = New BMSStruct(Notes, NotesTemplate,
                                           hWAV, hBMP, hBPM, hSTOP, hBMSCROLL, hCOM, wLWAV,
                                           HeaderT, HeaderN, HeaderI, TExpansion.Text, MeasureLength, FileNameTemplate,
                                           ExpansionSplit, GhostMode,
                                           sUndo, sRedo, sI,
                                           TExpansion.Enabled, IsSaved, NTInput, WaveformLoaded)

    End Sub

    Private Sub SaveAllBMSStruct()
        For i = 0 To UBound(BMSFiles) - 1
            ReadFile(BMSFiles(i).Filename, False)
            SaveBMSStruct(i)
        Next
    End Sub

    Private Sub LoadBMSStruct(Optional xI As Integer = -1)
        If xI = -1 Then xI = BMSFileIndex

        With BMSFiles(xI).Struct
            Notes = .Notes
            NotesTemplate = .NotesTemplate

            hWAV = .hWAV
            hBMP = .hBMP
            hBPM = .hBPM
            hSTOP = .hSTOP
            hBMSCROLL = .hBMSCROLL
            hCOM = .hCOM
            wLWAV = .wLWAV

            THTitle.Text = .HeaderT(0)
            THArtist.Text = .HeaderT(1)
            THGenre.Text = .HeaderT(2)
            THPlayLevel.Text = .HeaderT(3)
            THTotal.Text = .HeaderT(4)
            THSubTitle.Text = .HeaderT(5)
            THSubArtist.Text = .HeaderT(6)
            THStageFile.Text = .HeaderT(7)
            THBanner.Text = .HeaderT(8)
            THBackBMP.Text = .HeaderT(9)
            THExRank.Text = .HeaderT(10)
            THComment.Text = .HeaderT(11)

            THBPM.Value = .HeaderN(0)

            CHPlayer.SelectedIndex = .HeaderI(0)
            CHRank.SelectedIndex = .HeaderI(1)
            CHDifficulty.SelectedIndex = .HeaderI(2)
            CHLnObj.SelectedIndex = .HeaderI(3)

            TExpansion.Text = .Expansion
            MeasureLength = .MeasureLength
            FileNameTemplate = .FileNameTemplate

            ExpansionSplit = .ExpansionSplit
            GhostMode = .GhostMode

            sUndo = .sUndo
            sRedo = .sRedo
            sI = .sI

            TExpansion.Enabled = .ExpansionEnabled
            IsSaved = .IsSaved
            NTInput = .NTInput
            WaveformLoaded = .WaveformLoaded
        End With

        If Not WaveformLoaded AndAlso ShowWaveform Then WaveformLoadId = 1 : TimerLoadWaveform.Enabled = True
        SetIsSaved(IsSaved)

        LWAVRefresh() ' P: Wow why does refreshing this list take so damn long
        LBMPRefresh() ' P: Likely this too
        LBeatRefresh()
        RefreshItemsByNTInput()

        LoadColorOverride(FileName)
        UpdateMeasureBottom()
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub
End Class
