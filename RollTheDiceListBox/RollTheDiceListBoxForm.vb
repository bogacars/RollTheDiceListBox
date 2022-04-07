



Option Explicit On
Option Strict On
Public Class RollTheDiceListBoxForm


    Private Sub RollButton_Click(sender As Object, e As EventArgs) Handles RollButton.Click, RollToolStripMenuItem.Click
        DisplayListBox.Items.Clear()
        Dim RandomNumbers(10) As Integer
        'This makes a RandomNumbers array ^
        DisplayListBox.Items.Add("                                Random Numbers")
        For i = 1 To 1000
            RandomNumbers(GetRandomNumber(11)) += 1
            'This for loop rolls 2 dice 1000 times
        Next

        DisplayListBox.Items.Add(StrDup(77, "-")) 'These duplicate the dash 77 times

        Dim TopRow As String

        For i = 1 To 11
            TopRow &= CStr((i) + 1).PadLeft(6) & "|"
            'This for loop sets up the top row from 2 to 12

        Next
        DisplayListBox.Items.Add(TopRow)

        DisplayListBox.Items.Add(StrDup(77, "-"))

        Dim BottomRow As String

        For i = 0 To UBound(RandomNumbers)
            BottomRow &= CStr(RandomNumbers(i)).PadLeft(6) & "|"
            'This forloop is our bottom numbers that are "random"
        Next
        DisplayListBox.Items.Add(BottomRow)
        'This Puts the random numbers in the bottom row^
        DisplayListBox.Items.Add(StrDup(77, "-"))
    End Sub
    Function GetRandomNumber(MaxNumber As Integer) As Integer
        Randomize()
        Dim FirstNumber As Integer
        Dim SecondNumber As Integer
        FirstNumber = CInt(Int((6 * Rnd() * +1)))
        SecondNumber = CInt(Int((6 * Rnd() * +1)))
        Return (FirstNumber + SecondNumber)
        'This Function Simply rolls 2 dice and adds them together
        'Then when called it rolls both dice 1000 times and adds them up
    End Function
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem.Click
        Me.Close()
        'This closes the program
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem.Click
        DisplayListBox.Items.Clear()
        'This clears the program when the clear button is pressed
    End Sub
End Class
