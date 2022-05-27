<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.LabelTHJack = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.cAudioLine = New System.Windows.Forms.CheckBox()
        Me.cTemplateSnapToVPosition = New System.Windows.Forms.CheckBox()
        Me.NLNGap = New System.Windows.Forms.NumericUpDown()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.LabelTHLN = New System.Windows.Forms.Label()
        Me.TTemp = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.cPastePatternToVPosition = New System.Windows.Forms.CheckBox()
        Me.cPreloadBMSStruct = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NAutoSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nGridPartition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.nJackBPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nJackTH, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.NLNGap, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(229, 669)
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
        Me.CWheel.Location = New System.Drawing.Point(154, 475)
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
        Me.Label1.Location = New System.Drawing.Point(12, 477)
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
        Me.cBeep.AutoSize = True
        Me.cBeep.Checked = True
        Me.cBeep.CheckState = System.Windows.Forms.CheckState.Checked
        Me.TableLayoutPanel5.SetColumnSpan(Me.cBeep, 3)
        Me.cBeep.Dock = System.Windows.Forms.DockStyle.Left
        Me.cBeep.Location = New System.Drawing.Point(3, 28)
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
        Me.PictureBox1.Location = New System.Drawing.Point(20, 459)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(375, 1)
        Me.PictureBox1.TabIndex = 47
        Me.PictureBox1.TabStop = False
        '
        'cBpm1296
        '
        Me.cBpm1296.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cBpm1296, 3)
        Me.cBpm1296.Dock = System.Windows.Forms.DockStyle.Left
        Me.cBpm1296.Location = New System.Drawing.Point(3, 78)
        Me.cBpm1296.Name = "cBpm1296"
        Me.cBpm1296.Size = New System.Drawing.Size(254, 19)
        Me.cBpm1296.TabIndex = 13
        Me.cBpm1296.Text = "Extend number of multi-byte BPMs to 1296"
        Me.cBpm1296.UseVisualStyleBackColor = True
        '
        'cStop1296
        '
        Me.cStop1296.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cStop1296, 3)
        Me.cStop1296.Dock = System.Windows.Forms.DockStyle.Left
        Me.cStop1296.Location = New System.Drawing.Point(3, 103)
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
        Me.cMEnterFocus.Location = New System.Drawing.Point(34, 585)
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
        Me.cMClickFocus.Location = New System.Drawing.Point(34, 610)
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
        Me.Label5.Location = New System.Drawing.Point(12, 506)
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
        Me.CPgUpDn.Location = New System.Drawing.Point(154, 504)
        Me.CPgUpDn.Name = "CPgUpDn"
        Me.CPgUpDn.Size = New System.Drawing.Size(237, 23)
        Me.CPgUpDn.TabIndex = 103
        '
        'NAutoSave
        '
        Me.NAutoSave.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NAutoSave.DecimalPlaces = 1
        Me.NAutoSave.Location = New System.Drawing.Point(191, 3)
        Me.NAutoSave.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.NAutoSave.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NAutoSave.Name = "NAutoSave"
        Me.NAutoSave.Size = New System.Drawing.Size(88, 23)
        Me.NAutoSave.TabIndex = 11
        Me.NAutoSave.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label7.Location = New System.Drawing.Point(285, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(73, 25)
        Me.Label7.TabIndex = 59
        Me.Label7.Text = "minutes"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cAutoSave
        '
        Me.cAutoSave.AutoSize = True
        Me.cAutoSave.Checked = True
        Me.cAutoSave.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cAutoSave.Dock = System.Windows.Forms.DockStyle.Left
        Me.cAutoSave.Location = New System.Drawing.Point(3, 3)
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
        Me.cMStopPreview.Location = New System.Drawing.Point(34, 635)
        Me.cMStopPreview.Name = "cMStopPreview"
        Me.cMStopPreview.Size = New System.Drawing.Size(253, 19)
        Me.cMStopPreview.TabIndex = 108
        Me.cMStopPreview.Text = "Stop preview if clicked on the editing panel"
        Me.cMStopPreview.UseVisualStyleBackColor = True
        '
        'nGridPartition
        '
        Me.nGridPartition.Location = New System.Drawing.Point(197, 3)
        Me.nGridPartition.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nGridPartition.Minimum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.nGridPartition.Name = "nGridPartition"
        Me.nGridPartition.Size = New System.Drawing.Size(79, 23)
        Me.nGridPartition.TabIndex = 9
        Me.nGridPartition.Value = New Decimal(New Integer() {192, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label6.Location = New System.Drawing.Point(3, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(188, 33)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Maximum Grid Partition in BMS"
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
        Me.Label3.Location = New System.Drawing.Point(12, 534)
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
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(154, 533)
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
        Me.nJackBPM.Size = New System.Drawing.Size(50, 23)
        Me.nJackBPM.TabIndex = 9
        Me.nJackBPM.Value = New Decimal(New Integer() {255, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label8.Location = New System.Drawing.Point(3, 33)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(188, 33)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Minimum Jack Threshold"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'nJackTH
        '
        Me.nJackTH.Location = New System.Drawing.Point(101, 3)
        Me.nJackTH.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nJackTH.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nJackTH.Name = "nJackTH"
        Me.nJackTH.Size = New System.Drawing.Size(35, 23)
        Me.nJackTH.TabIndex = 9
        Me.nJackTH.Value = New Decimal(New Integer() {16, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label9.Location = New System.Drawing.Point(59, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(36, 27)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "BPM"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelTHJack
        '
        Me.LabelTHJack.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelTHJack.Location = New System.Drawing.Point(142, 0)
        Me.LabelTHJack.Name = "LabelTHJack"
        Me.LabelTHJack.Size = New System.Drawing.Size(30, 27)
        Me.LabelTHJack.TabIndex = 8
        Me.LabelTHJack.Text = "th"
        Me.LabelTHJack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.TableLayoutPanel2.Controls.Add(Me.LabelTHJack, 3, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(197, 36)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(175, 27)
        Me.TableLayoutPanel2.TabIndex = 110
        '
        'cAudioLine
        '
        Me.cAudioLine.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cAudioLine, 3)
        Me.cAudioLine.Dock = System.Windows.Forms.DockStyle.Left
        Me.cAudioLine.Location = New System.Drawing.Point(3, 128)
        Me.cAudioLine.Name = "cAudioLine"
        Me.cAudioLine.Size = New System.Drawing.Size(233, 19)
        Me.cAudioLine.TabIndex = 14
        Me.cAudioLine.Text = "Display audio lines during note preview"
        Me.cAudioLine.UseVisualStyleBackColor = True
        '
        'cTemplateSnapToVPosition
        '
        Me.cTemplateSnapToVPosition.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cTemplateSnapToVPosition, 3)
        Me.cTemplateSnapToVPosition.Dock = System.Windows.Forms.DockStyle.Left
        Me.cTemplateSnapToVPosition.Location = New System.Drawing.Point(3, 153)
        Me.cTemplateSnapToVPosition.Name = "cTemplateSnapToVPosition"
        Me.cTemplateSnapToVPosition.Size = New System.Drawing.Size(326, 19)
        Me.cTemplateSnapToVPosition.TabIndex = 111
        Me.cTemplateSnapToVPosition.Text = "Snap to Vertical Position for Moving to Template Position"
        Me.cTemplateSnapToVPosition.UseVisualStyleBackColor = True
        '
        'NLNGap
        '
        Me.NLNGap.DecimalPlaces = 4
        Me.NLNGap.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NLNGap.Location = New System.Drawing.Point(3, 3)
        Me.NLNGap.Name = "NLNGap"
        Me.NLNGap.Size = New System.Drawing.Size(74, 23)
        Me.NLNGap.TabIndex = 9
        Me.NLNGap.Value = New Decimal(New Integer() {5, 0, 0, 65536})
        '
        'Label10
        '
        Me.Label10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label10.Location = New System.Drawing.Point(3, 66)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(188, 33)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "Auto Long Note Gap"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.8617!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.1383!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label11, 0, 3)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 1, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.Label10, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel2, 1, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.Label8, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.nGridPartition, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Label6, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TTemp, 1, 3)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(15, 107)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 4
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(376, 133)
        Me.TableLayoutPanel3.TabIndex = 112
        '
        'Label11
        '
        Me.Label11.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label11.Location = New System.Drawing.Point(3, 99)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(188, 34)
        Me.Label11.TabIndex = 115
        Me.Label11.Text = "Temporary BMS Filename"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel4.Controls.Add(Me.LabelTHLN, 1, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.NLNGap, 0, 0)
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(197, 69)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(175, 27)
        Me.TableLayoutPanel4.TabIndex = 114
        '
        'LabelTHLN
        '
        Me.LabelTHLN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LabelTHLN.Location = New System.Drawing.Point(83, 0)
        Me.LabelTHLN.Name = "LabelTHLN"
        Me.LabelTHLN.Size = New System.Drawing.Size(89, 27)
        Me.LabelTHLN.TabIndex = 10
        Me.LabelTHLN.Text = "beats"
        Me.LabelTHLN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TTemp
        '
        Me.TTemp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TTemp.Location = New System.Drawing.Point(197, 102)
        Me.TTemp.Name = "TTemp"
        Me.TTemp.Size = New System.Drawing.Size(176, 23)
        Me.TTemp.TabIndex = 116
        Me.TTemp.Text = "___TempBMS.bmsc"
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel5.ColumnCount = 3
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.cAutoSave, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.NAutoSave, 1, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.cTemplateSnapToVPosition, 0, 6)
        Me.TableLayoutPanel5.Controls.Add(Me.Label7, 2, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.cBeep, 0, 1)
        Me.TableLayoutPanel5.Controls.Add(Me.cBpm1296, 0, 3)
        Me.TableLayoutPanel5.Controls.Add(Me.cStop1296, 0, 4)
        Me.TableLayoutPanel5.Controls.Add(Me.cAudioLine, 0, 5)
        Me.TableLayoutPanel5.Controls.Add(Me.cPastePatternToVPosition, 0, 7)
        Me.TableLayoutPanel5.Controls.Add(Me.cPreloadBMSStruct, 0, 2)
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(15, 246)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 8
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(376, 201)
        Me.TableLayoutPanel5.TabIndex = 113
        '
        'cPastePatternToVPosition
        '
        Me.cPastePatternToVPosition.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cPastePatternToVPosition, 3)
        Me.cPastePatternToVPosition.Dock = System.Windows.Forms.DockStyle.Left
        Me.cPastePatternToVPosition.Location = New System.Drawing.Point(3, 178)
        Me.cPastePatternToVPosition.Name = "cPastePatternToVPosition"
        Me.cPastePatternToVPosition.Size = New System.Drawing.Size(259, 20)
        Me.cPastePatternToVPosition.TabIndex = 111
        Me.cPastePatternToVPosition.Text = "Snap to Vertical Position for Pasting Patterns"
        Me.cPastePatternToVPosition.UseVisualStyleBackColor = True
        '
        'cPreloadBMSStruct
        '
        Me.cPreloadBMSStruct.AutoSize = True
        Me.TableLayoutPanel5.SetColumnSpan(Me.cPreloadBMSStruct, 3)
        Me.cPreloadBMSStruct.Location = New System.Drawing.Point(3, 53)
        Me.cPreloadBMSStruct.Name = "cPreloadBMSStruct"
        Me.cPreloadBMSStruct.Size = New System.Drawing.Size(182, 19)
        Me.cPreloadBMSStruct.TabIndex = 112
        Me.cPreloadBMSStruct.Text = "Preload all files in the files tab"
        Me.cPreloadBMSStruct.UseVisualStyleBackColor = True
        '
        'OpGeneral
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(413, 716)
        Me.Controls.Add(Me.TableLayoutPanel5)
        Me.Controls.Add(Me.TableLayoutPanel3)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TBAssociateBML)
        Me.Controls.Add(Me.TBAssociateBME)
        Me.Controls.Add(Me.cMStopPreview)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CPgUpDn)
        Me.Controls.Add(Me.TBAssociateIBMSC)
        Me.Controls.Add(Me.TBAssociatePMS)
        Me.Controls.Add(Me.cMClickFocus)
        Me.Controls.Add(Me.cMEnterFocus)
        Me.Controls.Add(Me.PictureBox1)
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
        CType(Me.NLNGap, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
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
    Friend WithEvents LabelTHJack As Label
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents cAudioLine As CheckBox
    Friend WithEvents cTemplateSnapToVPosition As CheckBox
    Friend WithEvents NLNGap As NumericUpDown
    Friend WithEvents Label10 As Label
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents LabelTHLN As Label
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents cPastePatternToVPosition As CheckBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TTemp As TextBox
    Friend WithEvents cPreloadBMSStruct As CheckBox
End Class
