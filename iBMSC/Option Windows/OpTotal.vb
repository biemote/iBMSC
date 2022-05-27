Public Class OpTotal
    Dim NoteCount As Integer
    Public TotalOption As Integer
    Dim CTotalList() As RadioButton

    Public Sub New(ByVal xNoteCount As Integer, ByVal xTotalOption As Integer, ByVal xMultiplier As Double,
                   ByVal xGlobalMultiplier As Double, ByVal xDecimal As Integer,
                   ByVal xDisplayValue As Boolean, ByVal xDisplayText As Boolean, ByVal xAutofill As Boolean)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CTotalList = {CTotalIIDX1, CTotalIIDX2, CTotalMultiplier}
        NoteCount = xNoteCount
        TotalOption = xTotalOption
        NMultiplier.Value = CDec(xMultiplier)
        NGlobalMultiplier.Value = CDec(xGlobalMultiplier)
        NDecimal.Value = xDecimal
        CBDisplayValue.Checked = xDisplayValue
        CBDisplayText.Checked = xDisplayText
        CBAutoFill.Checked = xAutofill
        CTotalList(TotalOption).Checked = True

        CalculateTotal()
    End Sub
    Private Sub OpTotal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Font = MainWindow.Font

        Text = Strings.fopTOTAL.Title

        Label1.Text = Strings.fopTOTAL.CalculationOption
        Label2.Text = Strings.fopTOTAL.Parameters
        Label3.Text = Strings.fopTOTAL.TOTAL
        Label4.Text = Strings.fopTOTAL.CalculationFormula
        CTotalIIDX1.Text = Strings.fopTOTAL.IIDXTOTALSupposition1
        CTotalIIDX2.Text = Strings.fopTOTAL.IIDXTOTALSupposition2
        CTotalMultiplier.Text = Strings.fopTOTAL.Multiplier
        Label5.Text = Strings.fopTOTAL.GlobalMultiplier
        Label11.Text = Strings.fopTOTAL.DecimalPlaces
        CBDisplayValue.Text = Strings.fopTOTAL.DisplayValue
        CBDisplayText.Text = Strings.fopTOTAL.DisplayRecommendedTOTAL.Replace("{}", Strings.StatusBar.RecommendedTotal)
        CBAutoFill.Text = Strings.fopTOTAL.AutofillTOTAL

        Dim Notes As String = Strings.fopTOTAL.Notes
        Label6.Text = "7.605 * " & Notes & " / (0.01 * " & Notes & " + 6.5)"
        Label7.Text = Notes & " < 400 ? 200 + " & Notes & " / 5 :"
        Label10.Text = Notes & " < 600 ? 280 + (" & Notes & " - 400) / 2.5 :"
        Label9.Text = Notes & " > 600 ? 360 + (" & Notes & " - 600) / 5"
        Label8.Text = Notes & " * " & Strings.fopTOTAL.Multiplier

        OK_Button.Text = Strings.OK
        Cancel_Button.Text = Strings.Cancel
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CTotalIIDX1_Click(sender As Object, e As EventArgs) Handles CTotalIIDX1.Click, CTotalIIDX2.Click, CTotalMultiplier.Click
        Dim RadioS As RadioButton = CType(sender, RadioButton)
        TotalOption = Array.IndexOf(Of RadioButton)(CTotalList, RadioS)
    End Sub

    Private Sub NGlobalMultiplier_TextChanged(sender As Object, e As EventArgs) Handles NGlobalMultiplier.TextChanged, NDecimal.TextChanged
        CalculateTotal()
    End Sub

    Private Sub CalculateTotal()
        Dim Dec = CInt(NDecimal.Value)
        LTotalIIDX1.Text = Math.Round(NoteCount * 7.605 / (0.01 * NoteCount + 6.5) * NGlobalMultiplier.Value, Dec).ToString()
        LTotalIIDX2.Text = Math.Round(CDbl(IIf(NoteCount < 400, 200 + NoteCount / 5, IIf(NoteCount < 600, 280 + (NoteCount - 400) / 2.5, 360 + (NoteCount - 600) / 5))) * CDbl(NGlobalMultiplier.Text), Dec).ToString()
        LTotalMultiplier.Text = Math.Round(NoteCount * CDbl(NMultiplier.Text) * CDbl(NGlobalMultiplier.Text), Dec).ToString()
    End Sub
End Class