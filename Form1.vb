Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Sex As String
        Dim press1 As Double = 0
        Dim press2 As Double = 0
        Dim press3 As Double = 0
        Dim pressStart As Double = 0
        'Dim press2Start As Double = 0
        'Dim press3Start As Double = 0
        If RadioButton1.Checked Then Sex = "Male"
        If RadioButton2.Checked Then Sex = "Female"



        Dim rng As New Random
        Dim arndint = rng.Next(2, 6)

        If TimeOfDay >= #11:59:55 PM# Then
            MsgBox("The current time is within 5 seconds of midnight" &
                   vbCrLf & "The timer returns to 0.0 at midnight")
            Return
        End If
        Dim start, finish, totalTime As Double
        Dim Msgboxstring As String

        Msgboxstring = "Ready to begin?"
        If (MsgBox(Msgboxstring, MsgBoxStyle.YesNo, "Science Project")) = MsgBoxResult.Yes Then
            'Dim rng As New Random
            arndint = rng.Next(2, 6)
            start = Microsoft.VisualBasic.DateAndTime.Timer                     'current time
            finish = start + arndint                                            'Set end time for random duration.
            Do While Microsoft.VisualBasic.DateAndTime.Timer < finish
                ' Do other processing while waiting for 5 seconds to elapse.
            Loop
            BackColor = Color.DeepSkyBlue                                       'time has passed, so change colour and start first timer
            pressStart = Microsoft.VisualBasic.DateAndTime.Timer
            'wait until the key is pressed
            MsgBox("presskey", MsgBoxStyle.Critical, "Science!")
            press1 = Microsoft.VisualBasic.DateAndTime.Timer - pressStart
            Me.Label3.Text = Math.Round(press1, 4)                                          'update screen time
            BackColor = Color.WhiteSmoke                                        'reset screen colour

        End If

        Msgboxstring = "Ready for next test?"
        If (MsgBox(Msgboxstring, MsgBoxStyle.YesNo, "Science!")) = MsgBoxResult.Yes Then
            'Dim rng As New Random
            arndint = rng.Next(2, 6)
            start = Microsoft.VisualBasic.DateAndTime.Timer                     'current time
            finish = start + arndint                                            'Set end time for random duration.
            Do While Microsoft.VisualBasic.DateAndTime.Timer < finish
                ' Do other processing while waiting for 5 seconds to elapse.
            Loop
            BackColor = Color.DeepSkyBlue                                       'time has passed, so change colour and start first timer
            pressStart = Microsoft.VisualBasic.DateAndTime.Timer
            'wait until the key is pressed
            MsgBox("presskey", MsgBoxStyle.Critical, "Science!")
            press2 = Microsoft.VisualBasic.DateAndTime.Timer - pressStart
            Me.Label4.Text = Math.Round(press2, 4)                                      'update screen time
            BackColor = Color.WhiteSmoke                                        'reset screen colour

        End If



        Msgboxstring = "Ready for final test?"
        If (MsgBox(Msgboxstring, MsgBoxStyle.YesNo, "Science!")) = MsgBoxResult.Yes Then
            'Dim rng As New Random
            arndint = rng.Next(2, 6)
            start = Microsoft.VisualBasic.DateAndTime.Timer                     'current time
            finish = start + arndint                                            'Set end time for random duration.
            Do While Microsoft.VisualBasic.DateAndTime.Timer < finish
                ' Do other processing while waiting for 5 seconds to elapse.
            Loop
            BackColor = Color.DeepSkyBlue                                       'time has passed, so change colour and start first timer
            pressStart = Microsoft.VisualBasic.DateAndTime.Timer
            'wait until the key is pressed
            MsgBox("presskey", MsgBoxStyle.Critical, "Science!")
            press3 = Microsoft.VisualBasic.DateAndTime.Timer - pressStart
            Me.Label5.Text = Math.Round(press3, 4)                                             'update screen time
            BackColor = Color.WhiteSmoke                                        'reset screen colour

        End If

        'average times of key press
        Dim pressAverage As Double = (press1 + press2 + press3) / 3
        Me.Label6.Text = Math.Round(pressAverage, 4)
        Me.Label6.Visible = True
        Me.Label7.Visible = True
        Me.Label8.Visible = True
        Me.Label9.Visible = True
        Me.Label10.Visible = True
        Me.Label11.Visible = True
        Me.ProgressBar1.Visible = True
        Me.ProgressBar2.Visible = True

        Dim myPath As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Science.csv"
        My.Computer.FileSystem.WriteAllText(myPath,
                                            vbCrLf & Me.TextBox1.Text & "," & Sex & "," & Me.CheckBox1.Checked & "," & Me.CheckBox2.Checked & "," & Me.CheckBox3.Checked & "," & press1 & "," & press2 & "," & press3 & "," & pressAverage, True)



        'put the code for average progress bars here
        'Dim myPath As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Science.csv"
        Using MyReader As New Microsoft.VisualBasic.
            FileIO.TextFieldParser(
                myPath)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            Dim totalRows As Short = 0
            Dim largestNumber As Double = 0
            Dim runningTotal As Double = 0
            Dim ageOnRecord As Double = 0
            While Not MyReader.EndOfData
                Try

                    currentRow = MyReader.ReadFields()
                    Dim currentField As String
                    Dim fieldCount As Short = 1
                    For Each currentField In currentRow
                        'processing of each of the fields goes here
                        If (fieldCount = 1) Then
                            ageOnRecord = Convert.ToInt32(currentField)
                            'MsgBox(ageOnRecord,, "ageOnRecord")
                        End If
                        If (fieldCount = 9) Then
                            If (ageOnRecord >= (Convert.ToInt32(Me.TextBox1.Text) - 3) And ageOnRecord <= (Convert.ToInt32(Me.TextBox1.Text) + 3)) Then
                                'MsgBox(Convert.ToInt32(Me.TextBox1.Text),, "textbox age")
                                'MsgBox(currentField,, "speed")
                                totalRows = totalRows + 1
                                runningTotal = runningTotal + currentField
                                If (currentField > largestNumber) Then
                                    largestNumber = currentField
                                End If
                            End If
                        End If
                        fieldCount = fieldCount + 1
                    Next

                Catch ex As Microsoft.VisualBasic.
                    FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                           "is not valid and will be skipped.")
                End Try

            End While
            'MsgBox(totalRows,, "Total Rows")
            'MsgBox(largestNumber,, "Largest Number")
            'MsgBox(runningTotal,, "Total")
            'MsgBox(runningTotal / totalRows,, "Average Value")




            Me.Label8.Visible = True
            Me.Label9.Visible = True
            Me.Label9.Text = "Average: " & Math.Round(runningTotal / totalRows, 4)
            Me.Label10.Visible = True
            Me.Label10.Text = "Max: " & Math.Round(largestNumber, 4)
            Me.Label11.Visible = True
            Me.Label13.Visible = True
            Me.Label14.Visible = True
            Me.ProgressBar1.Visible = True
            Me.ProgressBar2.Visible = True
            Me.ProgressBar1.Minimum = 0
            Me.ProgressBar1.Maximum = largestNumber * 1000
            Me.ProgressBar1.Value = (runningTotal * 1000) / totalRows
            Me.ProgressBar2.Minimum = 0
            Me.ProgressBar2.Maximum = largestNumber * 1000
            Me.ProgressBar2.Value = pressAverage * 1000

        End Using

        MsgBox("Reset Form?", MsgBoxStyle.OkOnly, "Science!")


        BackColor = Color.WhiteSmoke
        Me.Label3.Text = 0
        Me.Label4.Text = 0
        Me.Label5.Text = 0
        Me.Label9.Text = 0
        Me.Label10.Text = 0
        Me.TextBox1.Text = ""
        Me.CheckBox3.CheckState = False
        Me.CheckBox2.CheckState = False
        Me.CheckBox1.CheckState = False
        Me.RadioButton1.Checked = False
        Me.RadioButton2.Checked = False
        Me.Label6.Visible = False
        Me.Label7.Visible = False
        Me.Label8.Visible = False
        Me.Label9.Visible = False
        Me.Label10.Visible = False
        Me.Label11.Visible = False
        Me.Label13.Visible = False
        Me.Label14.Visible = False
        Me.ProgressBar1.Visible = False
        Me.ProgressBar2.Visible = False
        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = 0
        Me.ProgressBar1.Value = 0
        Me.ProgressBar2.Minimum = 0
        Me.ProgressBar2.Maximum = 0
        Me.ProgressBar2.Value = 0




    End Sub

End Class
