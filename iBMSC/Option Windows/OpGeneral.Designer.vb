﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OpGeneral
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.CWheel = New System.Windows.Forms.ComboBox()
        Me.CTextEncoding = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TBAssociate = New System.Windows.Forms.Button()
        Me.cBeep = New System.Windows.Forms.CheckBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cBpm1296 = New System.Windows.Forms.CheckBox()
        Me.cStop1296 = New System.Windows.Forms.CheckBox()
        Me.cMEnterFocus = New System.Windows.Forms.CheckBox()
        Me.cMClickFocus = New System.Windows.Forms.CheckBox()
        Me.TBAssociatePMS = New System.Windows.Forms.Button()
        Me.TBAssociateIBMSC = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CPgUpDn = New System.Windows.Forms.ComboBox()
        Me.NAutoSave = New System.Windows.Forms.NumericUpDown()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cAutoSave = New System.Windows.Forms.CheckBox()
        Me.cMStopPreview = New System.Windows.Forms.CheckBox()
        Me.nGridPartition = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TBAssociateBME = New System.Windows.Forms.Button()
        Me.TBAssociateBML = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.rMiddleAuto = New System.Windows.Forms.RadioButton()
        Me.rMiddleDrag = New System.Windows.Forms.RadioButton()
        Me.nJackBPM = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.nJackTH = New System.Windows.Forms.NumericUpDown()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.LabelTH = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NAutoSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nGridPartition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.nJackBPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nJackTH, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(229, 500)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(170, 33)
        Me.TableLayoutPanel1.TabIndex = 109
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(78, 27)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(88, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(78, 27)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'CWheel
        '
        Me.CWheel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CWheel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CWheel.FormattingEnabled = True
        Me.CWheel.Items.AddRange(New Object() {"1", "1 / 2", "1 / 3", "1 / 4"})
        Me.CWheel.Location = New System.Drawing.Point(154, 306)
        Me.CWheel.Name = "CWheel"
        Me.CWheel.Size = New System.Drawing.Size(237, 23)
        Me.CWheel.TabIndex = 101
        '
        'CTextEncoding
        '
        Me.CTextEncoding.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CTextEncoding.FormattingEnabled = True
        Me.CTextEncoding.Items.AddRange(New Object() {"ANSI (Locale dependant)", "Little Endian UTF16", "ASCII", "Big Endian UTF16", "Little Endian UTF32", "UTF7", "UTF8", "Shift-JIS", "EUC-KR"})
        Me.CTextEncoding.Location = New System.Drawing.Point(137, 18)
        Me.CTextEncoding.Name = "CTextEncoding"
        Me.CTextEncoding.Size = New System.Drawing.Size(254, 23)
        Me.CTextEncoding.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(12, 308)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 17)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Mouse Wheel"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(-5, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Text Encoding"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(-5, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Associate Filetype"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TBAssociate
        '
        Me.TBAssociate.Location = New System.Drawing.Point(137, 47)
        Me.TBAssociate.Name = "TBAssociate"
        Me.TBAssociate.Size = New System.Drawing.Size(122, 23)
        Me.TBAssociate.TabIndex = 3
        Me.TBAssociate.Text = "*.bms"
        Me.TBAssociate.UseVisualStyleBackColor = True
        '
        'cBeep
        '
        Me.cBeep.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cBeep.AutoSize = True
        Me.cBeep.Checked = True
        Me.cBeep.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cBeep.Location = New System.Drawing.Point(32, 206)
        Me.cBeep.Name = "cBeep"
        Me.cBeep.Size = New System.Drawing.Size(116, 19)
        Me.cBeep.TabIndex = 12
        Me.cBeep.Text = "Beep while saved"
        Me.cBeep.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.PictureBox1.Location = New System.Drawing.Point(20, 290)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(375, 1)
        Me.PictureBox1.TabIndex = 47
        Me.PictureBox1.TabStop = False
        '
        'cBpm1296
        '
        Me.cBpm1296.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cBpm1296.AutoSize = True
        Me.cBpm1296.Location = New System.Drawing.Point(32, 231)
        Me.cBpm1296.Name = "cBpm1296"
        Me.cBpm1296.Size = New System.Drawing.Size(254, 19)
        Me.cBpm1296.TabIndex = 13
        Me.cBpm1296.Text = "Extend number of multi-byte BPMs to 1296"
        Me.cBpm1296.UseVisualStyleBackColor = True
        '
        'cStop1296
        '
        Me.cStop1296.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cStop1296.AutoSize = True
        Me.cStop1296.Location = New System.Drawing.Point(32, 256)
        Me.cStop1296.Name = "cStop1296"
        Me.cStop1296.Size = New System.Drawing.Size(197, 19)
        Me.cStop1296.TabIndex = 14
        Me.cStop1296.Text = "Extend number of STOPs to 1296"
        Me.cStop1296.UseVisualStyleBackColor = True
        '
        'cMEnterFocus
        '
        Me.cMEnterFocus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cMEnterFocus.AutoSize = True
        Me.cMEnterFocus.Location = New System.Drawing.Point(34, 416)
        Me.cMEnterFocus.Name = "cMEnterFocus"
        Me.cMEnterFocus.Size = New System.Drawing.Size(322, 19)
        Me.cMEnterFocus.TabIndex = 106
        Me.cMEnterFocus.Text = "Automatically set focus to editing panel on mouse enter"
        Me.cMEnterFocus.UseVisualStyleBackColor = True
        '
        'cMClickFocus
        '
        Me.cMClickFocus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cMClickFocus.AutoSize = True
        Me.cMClickFocus.Location = New System.Drawing.Point(34, 441)
        Me.cMClickFocus.Name = "cMClickFocus"
        Me.cMClickFocus.Size = New System.Drawing.Size(293, 19)
        Me.cMClickFocus.TabIndex = 107
        Me.cMClickFocus.Text = "Disable first click if the editing panel is not focused"
        Me.cMClickFocus.UseVisualStyleBackColor = True
        '
        'TBAssociatePMS
        '
        Me.TBAssociatePMS.Location = New System.Drawing.Point(310, 78)
        Me.TBAssociatePMS.Name = "TBAssociatePMS"
        Me.TBAssociatePMS.Size = New System.Drawing.Size(81, 23)
        Me.TBAssociatePMS.TabIndex = 7
        Me.TBAssociatePMS.Text = "*.pms"
        Me.TBAssociatePMS.UseVisualStyleBackColor = True
        '
        'TBAssociateIBMSC
        '
        Me.TBAssociateIBMSC.Location = New System.Drawing.Point(265, 47)
        Me.TBAssociateIBMSC.Name = "TBAssociateIBMSC"
        Me.TBAssociateIBMSC.Size = New System.Drawing.Size(127, 23)
        Me.TBAssociateIBMSC.TabIndex = 4
        Me.TBAssociateIBMSC.Text = "*.ibmsc"
        Me.TBAssociateIBMSC.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.Location = New System.Drawing.Point(12, 337)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 17)
        Me.Label5.TabIndex = 102
        Me.Label5.Text = "PageUp / PageDown"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CPgUpDn
        '
        Me.CPgUpDn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CPgUpDn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CPgUpDn.FormattingEnabled = True
        Me.CPgUpDn.Items.AddRange(New Object() {"8", "6", "4", "3", "2", "1", "1 / 2"})
        Me.CPgUpDn.Location = New System.Drawing.Point(154, 335)
        Me.CPgUpDn.Name = "CPgUpDn"
        Me.CPgUpDn.Size = New System.Drawing.Size(237, 23)
        Me.CPgUpDn.TabIndex = 103
        '
        'NAutoSave
        '
        Me.NAutoSave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NAutoSave.DecimalPlaces = 1
        Me.NAutoSave.Location = New System.Drawing.Point(171, 179)
        Me.NAutoSave.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.NAutoSave.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NAutoSave.Name = "NAutoSave"
        Me.NAutoSave.Size = New System.Drawing.Size(62, 23)
        Me.NAutoSave.TabIndex = 11
        Me.NAutoSave.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.Location = New System.Drawing.Point(239, 181)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(73, 17)
        Me.Label7.TabIndex = 59
        Me.Label7.Text = "minutes"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cAutoSave
        '
        Me.cAutoSave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cAutoSave.AutoSize = True
        Me.cAutoSave.Checked = True
        Me.cAutoSave.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cAutoSave.Location = New System.Drawing.Point(32, 181)
        Me.cAutoSave.Name = "cAutoSave"
        Me.cAutoSave.Size = New System.Drawing.Size(76, 19)
        Me.cAutoSave.TabIndex = 10
        Me.cAutoSave.Text = "AutoSave"
        Me.cAutoSave.UseVisualStyleBackColor = True
        '
        'cMStopPreview
        '
        Me.cMStopPreview.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cMStopPreview.AutoSize = True
        Me.cMStopPreview.Checked = True
        Me.cMStopPreview.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cMStopPreview.Location = New System.Drawing.Point(34, 466)
        Me.cMStopPreview.Name = "cMStopPreview"
        Me.cMStopPreview.Size = New System.Drawing.Size(253, 19)
        Me.cMStopPreview.TabIndex = 108
        Me.cMStopPreview.Text = "Stop preview if clicked on the editing panel"
        Me.cMStopPreview.UseVisualStyleBackColor = True
        '
        'nGridPartition
        '
        Me.nGridPartition.Location = New System.Drawing.Point(219, 107)
        Me.nGridPartition.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nGridPartition.Minimum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.nGridPartition.Name = "nGridPartition"
        Me.nGridPartition.Size = New System.Drawing.Size(79, 23)
        Me.nGridPartition.TabIndex = 9
        Me.nGridPartition.Value = New Decimal(New Integer() {192, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(44, 108)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(169, 17)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Max Grid Partition in BMS"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TBAssociateBME
        '
        Me.TBAssociateBME.Location = New System.Drawing.Point(137, 78)
        Me.TBAssociateBME.Name = "TBAssociateBME"
        Me.TBAssociateBME.Size = New System.Drawing.Size(76, 23)
        Me.TBAssociateBME.TabIndex = 5
        Me.TBAssociateBME.Text = "*.bme"
        Me.TBAssociateBME.UseVisualStyleBackColor = True
        '
        'TBAssociateBML
        '
        Me.TBAssociateBML.Location = New System.Drawing.Point(219, 78)
        Me.TBAssociateBML.Name = "TBAssociateBML"
        Me.TBAssociateBML.Size = New System.Drawing.Size(85, 23)
        Me.TBAssociateBML.TabIndex = 6
        Me.TBAssociateBML.Text = "*.bml"
        Me.TBAssociateBML.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.Location = New System.Drawing.Point(12, 365)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 17)
        Me.Label3.TabIndex = 104
        Me.Label3.Text = "Mouse Middle Button"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FlowLayoutPanel1.AutoSize = True
        Me.FlowLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.FlowLayoutPanel1.Controls.Add(Me.rMiddleAuto)
        Me.FlowLayoutPanel1.Controls.Add(Me.rMiddleDrag)
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(154, 364)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(141, 38)
        Me.FlowLayoutPanel1.TabIndex = 105
        '
        'rMiddleAuto
        '
        Me.rMiddleAuto.AutoSize = True
        Me.rMiddleAuto.Checked = True
        Me.rMiddleAuto.Location = New System.Drawing.Point(3, 0)
        Me.rMiddleAuto.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.rMiddleAuto.Name = "rMiddleAuto"
        Me.rMiddleAuto.Size = New System.Drawing.Size(135, 19)
        Me.rMiddleAuto.TabIndex = 0
        Me.rMiddleAuto.TabStop = True
        Me.rMiddleAuto.Text = "Click and Auto Scroll"
        Me.rMiddleAuto.UseVisualStyleBackColor = True
        '
        'rMiddleDrag
        '
        Me.rMiddleDrag.AutoSize = True
        Me.rMiddleDrag.Location = New System.Drawing.Point(3, 19)
        Me.rMiddleDrag.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.rMiddleDrag.Name = "rMiddleDrag"
        Me.rMiddleDrag.Size = New System.Drawing.Size(102, 19)
        Me.rMiddleDrag.TabIndex = 1
        Me.rMiddleDrag.TabStop = True
        Me.rMiddleDrag.Text = "Click and Drag"
        Me.rMiddleDrag.UseVisualStyleBackColor = True
        '
        'nJackBPM
        '
        Me.nJackBPM.Location = New System.Drawing.Point(3, 3)
        Me.nJackBPM.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nJackBPM.Minimum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.nJackBPM.Name = "nJackBPM"
        Me.nJackBPM.Size = New System.Drawing.Size(52, 23)
        Me.nJackBPM.TabIndex = 9
        Me.nJackBPM.Value = New Decimal(New Integer() {255, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(44, 141)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(169, 17)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Min Jack Speed"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'nJackTH
        '
        Me.nJackTH.Location = New System.Drawing.Point(105, 3)
        Me.nJackTH.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nJackTH.Minimum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.nJackTH.Name = "nJackTH"
        Me.nJackTH.Size = New System.Drawing.Size(37, 23)
        Me.nJackTH.TabIndex = 9
        Me.nJackTH.Value = New Decimal(New Integer() {16, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label9.Location = New System.Drawing.Point(61, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 27)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "BPM"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelTH
        '
        Me.LabelTH.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelTH.Location = New System.Drawing.Point(148, 0)
        Me.LabelTH.Name = "LabelTH"
        Me.LabelTH.Size = New System.Drawing.Size(29, 27)
        Me.LabelTH.TabIndex = 8
        Me.LabelTH.Text = "th"
        Me.LabelTH.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.22222!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24.44444!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 23.88889!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 19.44444!))
        Me.TableLayoutPanel2.Controls.Add(Me.nJackBPM, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label9, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.nJackTH, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.LabelTH, 3, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(219, 136)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(180, 27)
        Me.TableLayoutPanel2.TabIndex = 110
        '
        'OpGeneral
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(413, 547)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TBAssociateBML)
        Me.Controls.Add(Me.TBAssociateBME)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.nGridPartition)
        Me.Controls.Add(Me.cMStopPreview)
        Me.Controls.Add(Me.cAutoSave)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.NAutoSave)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CPgUpDn)
        Me.Controls.Add(Me.TBAssociateIBMSC)
        Me.Controls.Add(Me.TBAssociatePMS)
        Me.Controls.Add(Me.cMClickFocus)
        Me.Controls.Add(Me.cMEnterFocus)
        Me.Controls.Add(Me.cStop1296)
        Me.Controls.Add(Me.cBpm1296)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.cBeep)
        Me.Controls.Add(Me.TBAssociate)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CTextEncoding)
        Me.Controls.Add(Me.CWheel)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OpGeneral"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "General Settings"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NAutoSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nGridPartition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        CType(Me.nJackBPM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nJackTH, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents CWheel As System.Windows.Forms.ComboBox
    Friend WithEvents CTextEncoding As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TBAssociate As System.Windows.Forms.Button
    Friend WithEvents cBeep As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents cBpm1296 As System.Windows.Forms.CheckBox
    Friend WithEvents cStop1296 As System.Windows.Forms.CheckBox
    Friend WithEvents cMEnterFocus As System.Windows.Forms.CheckBox
    Friend WithEvents cMClickFocus As System.Windows.Forms.CheckBox
    Friend WithEvents TBAssociatePMS As System.Windows.Forms.Button
    Friend WithEvents TBAssociateIBMSC As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CPgUpDn As System.Windows.Forms.ComboBox
    Friend WithEvents NAutoSave As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cAutoSave As System.Windows.Forms.CheckBox
    Friend WithEvents cMStopPreview As System.Windows.Forms.CheckBox
    Friend WithEvents nGridPartition As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TBAssociateBME As System.Windows.Forms.Button
    Friend WithEvents TBAssociateBML As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents rMiddleAuto As System.Windows.Forms.RadioButton
    Friend WithEvents rMiddleDrag As System.Windows.Forms.RadioButton
    Friend WithEvents nJackBPM As NumericUpDown
    Friend WithEvents Label8 As Label
    Friend WithEvents nJackTH As NumericUpDown
    Friend WithEvents Label9 As Label
    Friend WithEvents LabelTH As Label
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
End Class
