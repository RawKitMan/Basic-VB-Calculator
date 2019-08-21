
Imports System.Text.RegularExpressions

Public Class CalculatorOne

    'Storage variables for our input strings that will get passed to the Calculator class
    Dim leftNumber As String = ""
    Dim rightNumber As String = ""
    Dim numOperator As String = ""

    'To access our previous result, we want to keep track of the index on the getPreviousResults()
    'method is using
    Dim resultCounter As Integer = 0

    'We only want positive and negative numbers to pass for the left and right numbers
    'This regex helps us get that
    Dim numPattern As Regex = New Regex("^-?\d+(\.\d+)?$")

    'Create an instance of our Calculator Object
    Dim cjCalc As New myCalculator(leftNumber, rightNumber, numOperator)

    'For the number and decimal buttons in our gui, we want to add their text to the text box
    'As well as set our number to be positive or negative. 
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

    'This button is the only button we use to make the number negative/positive. Will not use
    'Keyboard inputs for this
    Private Sub BtnNegPos_Click(sender As Object, e As EventArgs) Handles btnNegPos.Click


        If (textInput.Text Is "") Then
            textInput.Text = "-"
        ElseIf (Not (textInput.Text.IndexOf("-") = 0)) Then
            textInput.Text = "-" + textInput.Text
        ElseIf (textInput.Text.IndexOf("-") = 0) Then
            textInput.Text = textInput.Text.Replace("-", "")
        End If
    End Sub

    'Our operator buttons on the Calculator GUI. Each one calls the method SetItems which is explained
    'below
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


    'We click the equals sign button, we call the SetEquals method, again explained below.
    Public Sub BtnEquals_Click(sender As Object, e As EventArgs) Handles btnEquals.Click
        SetEquals()
    End Sub

    'Clear button to reset the calculator
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Reset()
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

    'This method sets the value in the text box to be the right number and then performs the calculations
    'by calling the calculate() method in our Calculator class
    Public Sub SetEquals()

        'First, we need to know if the number being put in as the Right Number is actually a number
        'If it's not, then we need to reset the calculator and try again.
        If numPattern.IsMatch(textInput.Text) Then
            'If it is a number, we want to check to see if we are dividing by 0. If so, 
            'The user is informed and the calculator resets
            If (cjCalc.numOper.Equals("/") And Double.Parse(textInput.Text) = 0) Then
                MsgBox("Cannot divide by zero")
                Reset()
                'Otherwise, we want to set the number....
            ElseIf (Not (textInput.Text Is "")) Then

                '...Assuming the number being put in falls in the range of the Double data type
                'If not, we let the user know, reset the calculator and get out of the Sub
                If ((Double.Parse(textInput.Text) < Double.MinValue) Or (Double.Parse(textInput.Text) > Double.MaxValue)) Then
                    MsgBox("Number is out of range")
                    Reset()
                    Exit Sub
                    'If all checks pass, we set the right number using the setRight() method
                Else
                    cjCalc.SetRight(textInput.Text)
                End If

            End If
        Else
            MsgBox("Invalid Expression")
            Reset()
        End If

        'If our left and right numbers are set in our cjCalc instance, let's calculate
        If (Not (cjCalc.lNumber Is "") And Not (cjCalc.rNumber Is "") And Not (cjCalc.numOper Is "")) Then
            cjCalc.calculate()

            'Grab the result and make a check that it falls within the Double data type number range
            Dim totalAnswer As Double = cjCalc.getResult()
            If ((totalAnswer < Double.MinValue) Or (totalAnswer > Double.MaxValue)) Then
                MsgBox("The answer is out of range")
                Reset()
                Exit Sub
                'If it does, show the answer
            Else
                resultLabel.Text = cjCalc.lNumber + " " + cjCalc.numOper + " " + cjCalc.rNumber + " = " + totalAnswer.ToString()
                'Since we get a result, the counter goes up.
                resultCounter += 1
            End If


            'Our array only holds up to 10 results, so if our counter goes up to 10, we need to reset
            'We also do this in the Calculator class when we input the solution into the array.
            'But the main goals is to get the element that was placed before the current result
            If (resultCounter = 10) Then
                prevResultLabel.Text = cjCalc.getPreviousResult(9)
                resultCounter = 0
            Else
                prevResultLabel.Text = cjCalc.getPreviousResult(resultCounter - 1)
            End If

            'Once everything looks fine and dandy, we reset the calculator
            Reset()
        End If

    End Sub


    'This method sets our Left number and the selected operator into the Calculator method
    'As before we check to see if the number being put in is an actual number that falls within
    'the Double data type number range
    Public Sub SetItems(ByVal num As String, oper As String)

        If numPattern.IsMatch(textInput.Text) Then


            If (Not (textInput.Text Is "") And cjCalc.lNumber Is "") Then

                If (Double.Parse(num) < Double.MinValue Or Double.Parse(num) > Double.MaxValue) Then
                    MsgBox("Number is out of range")
                    Reset()
                    Exit Sub
                End If

                cjCalc.SetLeft(num)
                cjCalc.SetOperator(oper)
                textInput.Text = ""
            End If
        Else
            MsgBox("Invalid Expression")
            Reset()
        End If
    End Sub

    'Method to reset the calculator
    Public Sub Reset()
        cjCalc.SetLeft("")
        cjCalc.SetRight("")
        cjCalc.SetOperator("")
        textInput.Text = ""
    End Sub


End Class

Public Class myCalculator

    'The variables that we store and work with for our numbers and operators
    Public lNumber As String
    Public rNumber As String
    Public numOper As String

    'The array to hold all the results we get for our calculations
    Public results(10) As Double

    'Our index to keep track of where our results are placed in the array
    Public resultIndex As Integer = 0

    'Our constructor. When constructed, they will only have empty strings, but we will set them
    'later with the methods below
    Public Sub New(ByVal leftNum As String, rightNum As String, oper As String)

        lNumber = leftNum
        rNumber = rightNum
        numOper = oper

    End Sub

    'Sets our left number
    Public Sub SetLeft(ByVal leftNum As String)
        lNumber = leftNum
    End Sub

    'Sets our right number
    Public Sub SetRight(ByVal rightNum As String)
        rNumber = rightNum
    End Sub

    'Sets our selected operator
    Public Sub SetOperator(ByVal oper As String)
        numOper = oper
    End Sub

    'Our method to calculate...obviously. We parse the number strings and then perform the
    'associated math function based on the selected operator. The result is stored in the
    'results array and we increase the index by one.
    Public Sub calculate()

        If (numOper Is "+") Then
            results(resultIndex) = Double.Parse(lNumber) + Double.Parse(rNumber)
        ElseIf (numOper Is "-") Then
            results(resultIndex) = Double.Parse(lNumber) - Double.Parse(rNumber)
        ElseIf (numOper Is "*") Then
            results(resultIndex) = Double.Parse(lNumber) * Double.Parse(rNumber)
        ElseIf (numOper Is "/") Then
            results(resultIndex) = Double.Parse(lNumber) / Double.Parse(rNumber)
        End If

        resultIndex += 1

    End Sub

    'The method to get our current result from the calculation above. Since we incremented our
    'index in the calculate method, we need to reference the element at resultIndex - 1. However
    ' If we were at the end of our array already, we show what was at the end of the array
    'and reset the index to 0. 
    Public Function getResult()

        If (resultIndex = 10) Then
            resultIndex = 0
            Return results(9)
        Else
            Return results(resultIndex - 1)
        End If

    End Function

    'Calls the element at the designated index. For our purposes, it's the one that was stored
    'prior to the current calculation being done.
    Public Function getPreviousResult(ByVal index As Integer)
        Return Results(index)
    End Function

End Class

