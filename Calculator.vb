
Imports System.Text.RegularExpressions

Public Class Calculator
    Dim leftNumber As String = ""
    Dim rightNumber As String = ""
    Dim numOperator As String = ""
    Dim resultCounter As Integer = 0
    Dim numPattern As Regex = New Regex("^-?\d+(\.\d+)?$")
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

    Private Sub BtnDecimal_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        If (textInput.Text = "") Then
            textInput.Text = "0."
        Else
            textInput.Text = textInput.Text & "."
        End If

    End Sub

    'Our operator buttons on the Calculator GUI
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        SetItems(textInput.Text, btnAdd.Text)
    End Sub

    Private Sub BtnSubtract_Click(sender As Object, e As EventArgs) Handles btnSubtract.Click
        SetItems(textInput.Text, btnSubtract.Text)
    End Sub

    Private Sub BtnMultiply_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        SetItems(textInput.Text, btnMultiply.Text)
    End Sub

    Private Sub BtnDivide_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        SetItems(textInput.Text, btnDivide.Text)
    End Sub

    Public Sub BtnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        SetEquals()
    End Sub

    'The only symbols on a keyboard we want to use if someone is typing the expression is our math operators. So
    'lets capture those events and have them do the same thing the operator buttons in our GUI do. 

    Private Sub TextInput_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles textInput.KeyPress
        If e.KeyChar = Convert.ToChar(43) Then
            SetItems(textInput.Text, btnAdd.Text)
            e.Handled = True
        ElseIf e.KeyChar = Convert.ToChar(45) Then
            SetItems(textInput.Text, btnSubtract.Text)
            e.Handled = True
        ElseIf e.KeyChar = Convert.ToChar(42) Then
            SetItems(textInput.Text, btnMultiply.Text)
            e.Handled = True
        ElseIf e.KeyChar = Convert.ToChar(47) Then
            SetItems(textInput.Text, btnDivide.Text)
            e.Handled = True
        ElseIf e.KeyChar = Convert.ToChar(61) Then
            SetEquals()
            e.Handled = True
        ElseIf e.KeyChar = Convert.ToChar(13) Then
            SetEquals()
            e.Handled = True
        End If
    End Sub

    Public Sub SetEquals()

        If numPattern.IsMatch(textInput.Text) Then
            If (cjCalc.numOper.Equals("/") And Double.Parse(textInput.Text) = 0) Then
                MsgBox("Cannot divide by zero")
                Reset()
            ElseIf (Not (textInput.Text Is "") And cjCalc.lNumber And numPattern.IsMatch(textInput.Text)) Then
                cjCalc.SetRight(textInput.Text)
            End If

        Else
            MsgBox("Invalid Expression")
            Reset()
        End If


        If (Not (cjCalc.lNumber Is "") And Not (cjCalc.rNumber Is "") And Not (cjCalc.numOper Is "")) Then
            cjCalc.calculate()
            resultLabel.Text = cjCalc.lNumber + " " + cjCalc.numOper + " " + cjCalc.rNumber + " = " + cjCalc.getResult().ToString()
            If (cjCalc.Results.Count > 1) Then
                resultCounter += 1
                prevResultLabel.Text = cjCalc.getPreviousResult(resultCounter - 1)
            End If

            Reset()
        End If

        textInput.Text = ""
    End Sub

    Public Sub SetItems(ByVal num As String, oper As String)

        If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then
            If (numPattern.IsMatch(num)) Then
                cjCalc.SetLeft(num)
                cjCalc.SetOperator(oper)
            Else
                MsgBox("Invalid Expression")
                Reset()
            End If
        End If

        textInput.Text = ""
    End Sub

    Public Sub Reset()
        cjCalc.SetLeft("")
        cjCalc.SetRight("")
        cjCalc.SetOperator("")
    End Sub

    Private Sub BtnNegPos_Click(sender As Object, e As EventArgs) Handles btnNegPos.Click
        Console.WriteLine(textInput.Text.IndexOf("-") = 0)
        If (textInput.Text Is "") Then
            textInput.Text = "-"
        ElseIf (textInput.Text.IndexOf("-") = 0) Then
            textInput.Text = textInput.Text.Replace("-", "")
        End If
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

