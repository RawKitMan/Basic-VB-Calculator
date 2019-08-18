
Public Class Calculator
    Dim leftNumber As String = ""
    Dim rightNumber As String = ""
    Dim numOperator As String = ""
    Dim resultCounter As Integer = 0
    Dim cjCalc As New myCalculator(leftNumber, rightNumber, numOperator)

    Private Sub btnZero_Click(sender As Object, e As EventArgs) _
        Handles btnZero.Click
        textInput.Text = textInput.Text & btnZero.Text
    End Sub
    Private Sub btnOne_Click(sender As Object, e As EventArgs) _
        Handles btnOne.Click
        textInput.Text = textInput.Text & btnOne.Text
    End Sub
    Private Sub btnTwo_Click(sender As Object, e As EventArgs) _
        Handles btnTwo.Click
        textInput.Text = textInput.Text & btnTwo.Text
    End Sub
    Private Sub btnThree_Click(sender As Object, e As EventArgs) _
        Handles btnThree.Click
        textInput.Text = textInput.Text & btnThree.Text
    End Sub
    Private Sub btnFour_Click(sender As Object, e As EventArgs) _
        Handles btnFour.Click
        textInput.Text = textInput.Text & btnFour.Text
    End Sub
    Private Sub btnFive_Click(sender As Object, e As EventArgs) _
        Handles btnFive.Click
        textInput.Text = textInput.Text & btnFive.Text
    End Sub
    Private Sub btnSix_Click(sender As Object, e As EventArgs) _
        Handles btnSix.Click
        textInput.Text = textInput.Text & btnSix.Text
    End Sub
    Private Sub btnSeven_Click(sender As Object, e As EventArgs) _
        Handles btnSeven.Click
        textInput.Text = textInput.Text & btnSeven.Text
    End Sub
    Private Sub btnEight_Click(sender As Object, e As EventArgs) _
        Handles btnEight.Click
        textInput.Text = textInput.Text & btnEight.Text
    End Sub
    Private Sub btnNine_Click(sender As Object, e As EventArgs) _
        Handles btnNine.Click
        textInput.Text = textInput.Text & btnNine.Text
    End Sub


    'Be sure to account for no inputs and put up an error message box if no input is provided
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then
            'create the setLeft function and replace it here
            cjCalc.SetLeft(textInput.Text)
            cjCalc.SetOperator(btnAdd.Text)
            textInput.Text = ""
        End If

    End Sub

    Private Sub BtnSubtract_Click(sender As Object, e As EventArgs) Handles btnSubtract.Click
        If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then
            'create the setLeft function and replace it here
            cjCalc.SetLeft(textInput.Text)
            cjCalc.SetOperator(btnSubtract.Text)
            textInput.Text = ""
        End If
    End Sub

    Private Sub BtnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then
            'create the setLeft function and replace it here
            cjCalc.SetLeft(textInput.Text)
            cjCalc.SetOperator(btnMultiply.Text)
            textInput.Text = ""
        End If
    End Sub

    Private Sub BtnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then
            'create the setLeft function and replace it here
            cjCalc.SetLeft(textInput.Text)
            cjCalc.SetOperator(btnDivide.Text)
            textInput.Text = ""
        End If
    End Sub


    Public Sub BtnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        If (Not (textInput.Text Is "") And cjCalc.lNumber) Then
            cjCalc.SetRight(textInput.Text)
        End If
        If (Not (cjCalc.lNumber Is "") And Not (cjCalc.rNumber Is "") And Not (cjCalc.numOper Is "")) Then
            cjCalc.calculate()
            resultLabel.Text = cjCalc.lNumber + " " + cjCalc.numOper + " " + cjCalc.rNumber + " = " + cjCalc.getResult().ToString()
            If (cjCalc.Results.Count > 1) Then
                resultCounter += 1
                prevResultLabel.Text = cjCalc.getPreviousResult(resultCounter - 1)
            End If
            cjCalc.SetLeft("")
            cjCalc.SetRight("")

        End If

        textInput.Text = ""

    End Sub


End Class

Public Class myCalculator

    Public lNumber As String
    Public rNumber As String
    Public numOper As String
    Public Results As New List(Of Double)

    Public Sub New(ByVal leftNum As String, rightNum As String, oper As String)

        lNumber = leftNum
        rNumber = rightNum
        numOper = oper

    End Sub

    Public Sub SetLeft(ByVal leftNum As String)
        lNumber = leftNum
    End Sub
    Public Sub SetRight(ByVal rightNum As String)
        rNumber = rightNum
    End Sub
    Public Sub SetOperator(ByVal oper As String)
        numOper = oper
    End Sub

    Public Sub calculate()

        If (numOper Is "+") Then
            Results.Add(Double.Parse(lNumber) + Double.Parse(rNumber))
        ElseIf (numOper Is "-") Then
            Results.Add(Double.Parse(lNumber) - Double.Parse(rNumber))
        ElseIf (numOper Is "*") Then
            Results.Add(Double.Parse(lNumber) * Double.Parse(rNumber))
        ElseIf (numOper Is "/") Then
            Results.Add(Double.Parse(lNumber) / Double.Parse(rNumber))
        End If

    End Sub

    Public Function getResult()
        If (Results.Count > 0) Then
            Return Results(Results.Count - 1)
        Else
            Return Results(0)
        End If

    End Function

    Public Function getPreviousResult(ByVal index As Integer)
        Return Results(index)
    End Function


End Class

