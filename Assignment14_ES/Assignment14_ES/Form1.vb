Public Class Form1
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Close the Program

        Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear all labels and textboxes reset focus

        txtFirstName.Clear()
        txtLastName.Clear()
        txtPhoneNumber.Clear()
        txtEmail.Clear()
        txtDays.Clear()
        lblSubTotal.ResetText()
        lblTax.ResetText()
        lblTotal.ResetText()
        radStandard.Checked = True
        radAAA.Checked = False
        radAARP.Checked = False
        cboState.ResetText()
        txtFirstName.Focus()
        chkRepeat.Checked = False
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        'Declare Variables

        Dim intDays As Integer
        Dim strFirstName As String = txtFirstName.Text
        Dim strLastName As String = txtLastName.Text
        Dim strPhoneNumber As String = txtPhoneNumber.Text
        Dim strEmail As String = txtEmail.Text
        Dim dblSubTotal As Double
        Dim dblTax As Double
        Dim dblTotal As Double

        Dim blnValidated As Boolean = True

        Call Get_Validate_Input(strFirstName, strLastName, strPhoneNumber, strEmail, intDays, blnValidated)

        If blnValidated = True Then
            Call Calculations(intDays, dblSubTotal, dblTax, dblTotal)
            Call Display_Totals(dblSubTotal, dblTax, dblTotal)
        End If
    End Sub

    'Get and Validate Inputs

    Private Sub Get_Validate_Input(strFirstName As String, strLastName As String, strPhoneNumber As String, strEmail As String, ByRef intDays As Integer, ByRef blnValidated As Boolean)
        Call Get_Validate_FirstName(strFirstName, blnValidated)
        Call Get_Validate_LastName(strLastName, blnValidated)
        Call Get_Validate_PhoneNumber(strPhoneNumber, blnValidated)
        Call Get_Validate_Email(strEmail, blnValidated)
        Call Get_Validate_Days(intDays, blnValidated)
        Call Get_Validate_Season(blnValidated)
        Call Get_Validate_State(blnValidated)
    End Sub

    Private Sub Get_Validate_FirstName(strFirstName, ByRef blnValidated)
        If txtFirstName.Text = String.Empty Then
            MessageBox.Show("First Name Must be Entered")
            txtFirstName.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub


    Private Sub Get_Validate_LastName(strLastName, ByRef blnValidated)
        If txtLastName.Text = String.Empty Then
            MessageBox.Show("Last Name Must be Entered")
            txtLastName.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    Private Sub Get_Validate_PhoneNumber(strPhoneNumber, ByRef blnValidated)
        If txtPhoneNumber.Text = String.Empty Then
            MessageBox.Show("Phone Number Must be Entered")
            txtPhoneNumber.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    Private Sub Get_Validate_Email(Email, ByRef blnValidated)
        If txtEmail.Text = String.Empty Then
            MessageBox.Show("Email Must be Entered")
            txtEmail.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    Private Sub Get_Validate_Days(ByRef intDays As Integer, ByRef blnValidated As Boolean)
        If Integer.TryParse(txtDays.Text, intDays) Then
            intDays = txtDays.Text
            If (intDays > 0) And (intDays < 60) Then
            Else
                MessageBox.Show("Number of Days Must be greater than 0 and less than 60")
                txtDays.Focus()
                blnValidated = False
                Exit Sub
            End If
        Else
            MessageBox.Show("Number of Days Must be entered, and Must be Numeric")
            txtDays.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    Private Sub Get_Validate_Season(ByRef blnValidated As Boolean)
        If radOff.Checked Or radPeak.Checked Or radStandard.Checked Then
        Else
            MessageBox.Show("Season Must be Selected")
            radStandard.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    Private Sub Get_Validate_State(ByRef blnValidated As Boolean)
        If cboState.Text = String.Empty Then
            MessageBox.Show("State Must be Selected")
            cboState.Focus()
            blnValidated = False
            Exit Sub
        End If
    End Sub

    'Do Calculations

    Private Sub Calculations(ByRef intDays As Integer, ByRef dblSubTotal As Double, ByRef dblTax As Double, ByRef dblTotal As Double)
        Call Calculate_Season_Price(intDays, dblSubTotal)
        Call Calculate_Seasonal_Discount(dblSubTotal)
        Call Calculate_Determine_Discounts(intDays, dblSubTotal)
        Call Calculate_Tax(dblTax, dblSubTotal)
        Call Calculate_Total(dblTax, dblSubTotal, dblTotal)
    End Sub

    Private Sub Calculate_Determine_Discounts(ByRef intDays As Integer, ByRef dblSubtotal As Double)
        Call Calculate_Day_Discount(intDays, dblSubtotal)
        Call Calculate_Optional_Discount(dblSubtotal)
        Call Calculate_Repeating_Discount(dblSubtotal)
    End Sub

    Private Sub Calculate_Season_Price(ByRef intDays As Integer, ByRef dblSubTotal As Double)
        If radOff.Checked Then
            dblSubTotal = (50 * intDays)
        Else
            If radPeak.Checked Then
                dblSubTotal = (150 * intDays)
            Else
                dblSubTotal = (100 * intDays)
            End If
        End If
    End Sub

    Private Sub Calculate_Seasonal_Discount(ByRef dblSubTotal As Double)
        If radAAA.Checked Or radAARP.Checked Then
            dblSubTotal = dblSubTotal - dblSubTotal * 0.025
        End If
    End Sub

    Private Sub Calculate_Day_Discount(ByRef intDays As Integer, ByRef dblSubTotal As Double)
        If intDays > 30 Then
            dblSubTotal = dblSubTotal * 0.9
        Else
            If intDays > 14 Then
                dblSubTotal = dblSubTotal * 0.95
            End If
        End If
    End Sub

    Private Sub Calculate_Optional_Discount(ByRef dblSubTotal As Double)
        If radOff.Checked Then
            If dblSubTotal > 300 Then
                MessageBox.Show(FormatCurrency(dblSubTotal), "Your Discount")
                MessageBox.Show("The maximum discount for the Off Season is $300, The Maximum Discount will be applied.")
                dblSubTotal = dblSubTotal + dblSubTotal - 300
            End If
        End If

        If radStandard.Checked Then
            If dblSubTotal > 400 Then
                MessageBox.Show(FormatCurrency(dblSubTotal), "Your Discount")
                MessageBox.Show("The maximum discount for the Standard Season is $400, The Maximum Discount will be applied.")
                dblSubTotal = dblSubTotal + dblSubTotal - 400
            End If
        End If

        If radPeak.Checked Then
            If dblSubTotal > 500 Then
                MessageBox.Show(ToString(dblSubTotal), "Your Discount")
                MessageBox.Show("The maximum discount for the Peak Season is $500, The Maximum Discount will be applied.")
                dblSubTotal = dblSubTotal + dblSubTotal - 500
            End If
        End If
    End Sub

    Private Sub Calculate_Repeating_Discount(ByRef dblSubTotal)
        If chkRepeat.Checked Then
            If dblSubTotal > 300 Then
                dblSubTotal = dblSubTotal - 30
            Else
                If dblSubTotal > 200 Then
                    dblSubTotal = dblSubTotal - 20
                Else
                    If dblSubTotal > 50 Then
                        dblSubTotal = dblSubTotal - 5
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Calculate_Tax(ByRef dblTax As Double, ByRef dblSubTotal As Double)
        dblTax = dblSubTotal * 0.1
        If cboState.Text = "Florida" Then
            dblTax = 0
        Else
            dblTax = dblSubTotal * 0.1
        End If
    End Sub

    Private Sub Calculate_Total(ByRef dblSubTotal As Double, ByRef dblTax As Double, ByRef dblTotal As Double)
        dblTotal = dblSubTotal + dblTax
    End Sub

    'Display Outputs

    Private Sub Display_Totals(ByVal dblSubTotal As Double, ByVal dblTax As Double, ByVal dblTotal As Double)
        lblSubTotal.Text = dblSubTotal.ToString("c")
        lblTax.Text = dblTax.ToString("c")
        lblTotal.Text = dblTotal.ToString("c")
    End Sub

End Class
