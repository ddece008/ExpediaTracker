Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms.DataFormats
Imports ADODB
Imports CefSharp.Structs
Imports Microsoft.Office.Interop.Access
Imports Microsoft.Office.Interop.Word
Imports NHunspell
Public Class Form1
    'Dim Path As String = $"{Application.StartupPath}"
    Dim PathF As String = Path.GetDirectoryName(Application.ExecutablePath)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ' Define the paths for the affix file and dictionary file
        'Dim affFile As String = AppDomain.CurrentDomain.BaseDirectory & "../../Dictionaries/en_us.aff"
        'Dim dicFile As String = AppDomain.CurrentDomain.BaseDirectory & "../../Dictionaries/en_us.dic"
        Dim affFile As String = PathF & "../../Dictionaries/en_us.aff"
        Dim dicFile As String = PathF & "../../Dictionaries/en_us.dic"


        ' Clear the list box that will display suggestions
        lbSuggestion.Items.Clear()

        ' Create an instance of the Hunspell class using the specified affix and dictionary files
        Using hunspell As New Hunspell(affFile, dicFile)
            ' Split the input text into individual words
            'Dim inputWords As String() = TextBox1.Text.Split(" "c)
            ' Get the sentence from textbox1
            If String.IsNullOrEmpty(TextBox1.Text) Then
                ' Display a message if TextBox2 is blank
                MessageBox.Show("TextBox1 is blank. Please enter some text.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Dim sentence As String = TextBox1.Text
                ' Split the sentence into words
                Dim words As String() = sentence.Split(" "c)

                ' Clear previous selections in lbSuggestion
                lbSuggestion.ClearSelected()

                ' Check each word for correct spelling
                For i As Integer = 0 To words.Length - 1
                    'Dim correct As Boolean = hunspell.Spell(trimmedWord)
                    Dim isSpelledCorrectly As Boolean = hunspell.Spell(words(i))

                    If isSpelledCorrectly Then
                        ' Find the index of the correct word in lbSuggestion
                        Dim index As Integer = lbSuggestion.FindStringExact(words(i))

                        ' If the word is found, select it in lbSuggestion
                        If index <> -1 Then
                            lbSuggestion.SetSelected(index, True)

                            ' Replace the word with the selected suggestion
                            words(i) = lbSuggestion.SelectedItem.ToString()
                        End If
                    Else
                        Dim index As Integer = lbSuggestion.FindStringExact(words(i))
                        ' If the word is found, select it in lbSuggestion
                        Dim suggestions As List(Of String) = hunspell.Suggest(words(i))
                        'countlabel.Text = "There are " & suggestions.Count.ToString() & " suggestions"
                        For Each suggestion As String In suggestions


                            lbSuggestion.Items.Add(suggestion)
                            lbSuggestion.SelectedIndex = 0

                        Next

                        'Form4.ListBox1.SelectedIndex = 0


                        ' Replace the word with the selected suggestion
                        words(i) = lbSuggestion.SelectedItem.ToString()
                        lbSuggestion.ClearSelected()
                        lbSuggestion.Items.Clear()

                    End If
                Next

                ' Join the corrected words back into a sentence 
                Dim correctedSentence As String = String.Join(" ", words)

                ' Update the text in textbox1
                'TextBox1.Text = correctedSentence
                TextBox2.Text = correctedSentence
            End If
        End Using
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Check if TextBox2 is blank
        If String.IsNullOrEmpty(TextBox2.Text) Then
            ' Display a message if TextBox2 is blank
            MessageBox.Show("TextBox2 is blank.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            ' Copy the text from TextBox2 to the clipboard
            Clipboard.SetText(TextBox2.Text)

            ' Optionally, you can provide a user feedback
            MessageBox.Show("Text copied to clipboard!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private remainingTime As Integer = 180 ' 3 minutes in seconds
    Private startTime As DateTime

    Private remainingTime2 As Integer = 180 ' 3 minutes in seconds
    Private startTime2 As DateTime

    Private remainingTime3 As Integer = 180 ' 3 minutes in seconds
    Private startTime3 As DateTime

    Private remainingTime4 As Integer = 180 ' 3 minutes in seconds
    Private startTime4 As DateTime

    Private remainingTime5 As Integer = 180 ' 3 minutes in seconds
    Private startTime5 As DateTime

    Private remainingTime6 As Integer = 180 ' 3 minutes in seconds
    Private startTime6 As DateTime


    Private timerRunning As Boolean = False
    Private timerPaused As Boolean = False
    'Private startTime1 As DateTime
    Private remainingTime1 As TimeSpan
    Private timerRunning2 As Boolean = False
    Private timerPaused2 As Boolean = False
    Private remainingTimex2 As TimeSpan
    Private timerRunning3 As Boolean = False
    Private timerPaused3 As Boolean = False
    Private remainingTimex3 As TimeSpan
    Private timerRunning4 As Boolean = False
    Private timerPaused4 As Boolean = False
    Private remainingTimex4 As TimeSpan
    Private timerRunning5 As Boolean = False
    Private timerPaused5 As Boolean = False
    Private remainingTimex5 As TimeSpan
    Private timerRunning6 As Boolean = False
    Private timerPaused6 As Boolean = False
    Private remainingTimex6 As TimeSpan
    Dim username As String = Environment.UserName
    Dim imagePathPause As String = $"{Application.StartupPath}\png\pause.png"
    Dim imagePathPlay As String = $"{Application.StartupPath}\png\play.png"



    Private Sub UpdateTimeLabel1()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar1.Value)
        Label12.Text = remainingTime.ToString("hh\:mm\:ss")
        'Label11.Text = remainingTime.ToString("hh\:mm\:ss")

    End Sub
    Private Sub UpdateTimeLabel2()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar2.Value)
        Label13.Text = remainingTime.ToString("hh\:mm\:ss")
    End Sub
    Private Sub UpdateTimeLabel3()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar3.Value)
        Label14.Text = remainingTime.ToString("hh\:mm\:ss")
    End Sub
    Private Sub UpdateTimeLabel4()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar4.Value)
        Label4.Text = remainingTime.ToString("hh\:mm\:ss")
    End Sub
    Private Sub UpdateTimeLabel5()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar5.Value)
        Label5.Text = remainingTime.ToString("hh\:mm\:ss")
    End Sub
    Private Sub UpdateTimeLabel6()
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(ProgressBar6.Value)
        Label6.Text = remainingTime.ToString("hh\:mm\:ss")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Calculate remaining time
        Dim elapsedTime = (DateTime.Now - startTime)
        Dim remainingTime As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime


        If elapsedTime.Minutes < 3 Then
            ProgressBar1.Maximum = 181
            Label12.Text = elapsedTime.ToString("hh\:mm\:ss")
            ProgressBar1.Value = CInt(elapsedTime.TotalSeconds)
        ElseIf elapsedTime.Minutes = 3 Then
            ProgressBar1.BackColor = Color.Red
            'Timer1.Stop()
            Label12.Text = elapsedTime.ToString("hh\:mm\:ss")
            remainingTime1 = DateTime.Now - startTime
            'timerPaused = True
            If Label12.Text = "00:03:01" Then
                'MsgBox("Chat1 - 3 minutes elapsed.")
                Label12.ForeColor = Color.Red

            End If

        ElseIf elapsedTime.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime.TotalSeconds = 180 Then
                ProgressBar1.Value = 180
            Else
                ProgressBar1.Maximum = 600 * 4
                ProgressBar1.Value = CInt(elapsedTime.TotalSeconds)
            End If
            Label12.Text = elapsedTime.ToString("hh\:mm\:ss")
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        ' Calculate remaining time
        Dim elapsedTime2 = (DateTime.Now - startTime2)
        Dim remainingTime2 As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime2

        If elapsedTime2.Minutes < 3 Then
            ProgressBar2.Maximum = 181
            Label13.Text = elapsedTime2.ToString("hh\:mm\:ss")
            ProgressBar2.Value = CInt(elapsedTime2.TotalSeconds)
        ElseIf elapsedTime2.Minutes = 3 Then
            ProgressBar2.BackColor = Color.Red
            'Timer1.Stop()
            'Timer1.Stop()
            Label13.Text = elapsedTime2.ToString("hh\:mm\:ss")
            remainingTime2 = DateTime.Now - startTime
            'timerPaused = True
            If Label13.Text = "00:03:01" Then
                'MsgBox("Chat1 - 3 minutes elapsed.")
                Label13.ForeColor = Color.Red
            End If
        ElseIf elapsedTime2.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime2.TotalSeconds = 180 Then
                ProgressBar2.Value = 180
            Else
                ProgressBar2.Maximum = 600
                ProgressBar2.Value = CInt(elapsedTime2.TotalSeconds)
            End If

            Label13.Text = elapsedTime2.ToString("hh\:mm\:ss")

        End If

    End Sub
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        ' Calculate remaining time
        Dim elapsedTime3 = (DateTime.Now - startTime3)
        Dim remainingTime3 As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime3

        If elapsedTime3.Minutes < 3 Then
            ProgressBar3.Maximum = 181
            Label14.Text = elapsedTime3.ToString("hh\:mm\:ss")
            ProgressBar3.Value = CInt(elapsedTime3.TotalSeconds)
        ElseIf elapsedTime3.Minutes = 3 Then
            ProgressBar3.BackColor = Color.Red
            'Timer1.Stop()
            'Timer1.Stop()
            Label14.Text = elapsedTime3.ToString("hh\:mm\:ss")
            remainingTime3 = DateTime.Now - startTime
            'timerPaused = True
            If Label14.Text = "00:03:01" Then
                'MsgBox("Chat1 - 3 minutes elapsed.")
                Label14.ForeColor = Color.Red
            End If

        ElseIf elapsedTime3.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime3.TotalSeconds = 180 Then
                ProgressBar3.Value = 180
            Else
                ProgressBar3.Maximum = 600
                ProgressBar3.Value = CInt(elapsedTime3.TotalSeconds)
            End If
            Label14.Text = elapsedTime3.ToString("hh\:mm\:ss")
        End If
    End Sub
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        ' Calculate remaining time
        Dim elapsedTime4 = (DateTime.Now - startTime4)
        Dim remainingTime4 As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime4

        If elapsedTime4.Minutes < 3 Then
            ProgressBar4.Maximum = 181
            Label4.Text = elapsedTime4.ToString("hh\:mm\:ss")
            ProgressBar4.Value = CInt(elapsedTime4.TotalSeconds)
        ElseIf elapsedTime4.Minutes = 3 Then
            ProgressBar4.BackColor = Color.Red
            'Timer1.Stop()
            'Timer1.Stop()
            Label4.Text = elapsedTime4.ToString("hh\:mm\:ss")
            remainingTime4 = DateTime.Now - startTime
            'timerPaused = True
            If Label4.Text = "00:03:01" Then
                'MsgBox("Chat1 - 4 minutes elapsed.")
                Label4.ForeColor = Color.Red
            End If

        ElseIf elapsedTime4.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime4.TotalSeconds = 180 Then
                ProgressBar4.Value = 180
            Else
                ProgressBar4.Maximum = 600
                ProgressBar4.Value = CInt(elapsedTime4.TotalSeconds)
            End If
            Label4.Text = elapsedTime4.ToString("hh\:mm\:ss")
        End If
    End Sub
    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        ' Calculate remaining time
        Dim elapsedTime5 = (DateTime.Now - startTime5)
        Dim remainingTime5 As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime5

        If elapsedTime5.Minutes < 3 Then
            ProgressBar5.Maximum = 181
            Label5.Text = elapsedTime5.ToString("hh\:mm\:ss")
            ProgressBar5.Value = CInt(elapsedTime5.TotalSeconds)
        ElseIf elapsedTime5.Minutes = 3 Then
            ProgressBar5.BackColor = Color.Red
            'Timer1.Stop()
            'Timer1.Stop()
            Label5.Text = elapsedTime5.ToString("hh\:mm\:ss")
            remainingTime5 = DateTime.Now - startTime
            'timerPaused = True
            If Label5.Text = "00:03:01" Then
                'MsgBox("Chat1 - 5 minutes elapsed.")
                Label5.ForeColor = Color.Red
            End If

        ElseIf elapsedTime5.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime5.TotalSeconds = 180 Then
                ProgressBar5.Value = 180
            Else
                ProgressBar5.Maximum = 600
                ProgressBar5.Value = CInt(elapsedTime5.TotalSeconds)
            End If
            Label5.Text = elapsedTime5.ToString("hh\:mm\:ss")
        End If
    End Sub


    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        ' Calculate remaining time
        Dim elapsedTime6 = (DateTime.Now - startTime6)
        Dim remainingTime6 As TimeSpan = TimeSpan.FromSeconds(180) - elapsedTime6

        If elapsedTime6.Minutes < 3 Then
            ProgressBar6.Maximum = 181
            Label6.Text = elapsedTime6.ToString("hh\:mm\:ss")
            ProgressBar6.Value = CInt(elapsedTime6.TotalSeconds)
        ElseIf elapsedTime6.Minutes = 3 Then
            ProgressBar6.BackColor = Color.Red
            'Timer1.Stop()
            'Timer1.Stop()
            Label6.Text = elapsedTime6.ToString("hh\:mm\:ss")
            remainingTime6 = DateTime.Now - startTime
            'timerPaused = True
            If Label6.Text = "00:03:01" Then
                'MsgBox("Chat1 - 6 minutes elapsed.")
                Label6.ForeColor = Color.Red
            End If

        ElseIf elapsedTime6.Minutes > 3 Then
            ' Update progress bar
            If elapsedTime6.TotalSeconds = 180 Then
                ProgressBar6.Value = 180
            Else
                ProgressBar6.Maximum = 600
                ProgressBar6.Value = CInt(elapsedTime6.TotalSeconds)
            End If
            Label6.Text = elapsedTime6.ToString("hh\:mm\:ss")
        End If
    End Sub


    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        '----- play1

        If Not timerRunning Then
            ' Start the timer
            startTime = DateTime.Now
            Timer1.Start()
            timerRunning = True
            timerPaused = False
            PictureBox8.Image = Image.FromFile(imagePathPause)

        ElseIf timerRunning And Not timerPaused Then
            ' Pause the timer and calculate remaining time
            Timer1.Stop()
            remainingTime1 = DateTime.Now - startTime
            timerPaused = True
            ' Reset the image to the original one
            PictureBox8.Image = Image.FromFile(imagePathPlay)
            'PictureBox8.Dock = DockStyle.Fill
        ElseIf timerRunning And timerPaused Then
            ' Resume the timer with remaining time
            startTime = DateTime.Now - remainingTime1
            Timer1.Start()
            'Timer1.Stop()
            timerPaused = False
            PictureBox8.Image = Image.FromFile(imagePathPause)
        End If
    End Sub

    '----- play2
    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click

        If Not timerRunning2 Then
            ' Start the timer
            startTime2 = DateTime.Now
            Timer2.Start()
            timerRunning2 = True
            timerPaused2 = False
            PictureBox10.Image = Image.FromFile(imagePathPause)
        ElseIf timerRunning2 And Not timerPaused2 Then
            ' Pause the timer and calculate remaining time
            Timer2.Stop()
            remainingTimex2 = DateTime.Now - startTime2
            timerPaused2 = True
            ' Reset the image to the original one
            PictureBox10.Image = Image.FromFile(imagePathPlay)
            'PictureBox10.Dock = DockStyle.Fill
        ElseIf timerRunning2 And timerPaused2 Then
            ' Resume the timer with remaining time
            startTime2 = DateTime.Now - remainingTimex2
            Timer2.Start()
            timerPaused2 = False
            PictureBox10.Image = Image.FromFile(imagePathPause)
        End If
    End Sub
    '----- play3
    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click
        If Not timerRunning3 Then
            ' Start the timer
            startTime3 = DateTime.Now
            Timer3.Start()
            timerRunning3 = True
            timerPaused3 = False
            PictureBox12.Image = Image.FromFile(imagePathPause)
        ElseIf timerRunning3 And Not timerPaused3 Then
            ' Pause the timer and calculate remaining time
            Timer3.Stop()
            remainingTimex3 = DateTime.Now - startTime3
            timerPaused3 = True
            ' Reset the image to the original one
            PictureBox12.Image = Image.FromFile(imagePathPlay)

            'PictureBox12.Dock = DockStyle.Fill
        ElseIf timerRunning3 And timerPaused3 Then
            ' Resume the timer with remaining time
            startTime3 = DateTime.Now - remainingTimex3
            Timer3.Start()
            timerPaused3 = False
            PictureBox12.Image = Image.FromFile(imagePathPause)
        End If
    End Sub
    '----- play4
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If Not timerRunning4 Then
            ' Start the timer
            startTime4 = DateTime.Now
            Timer4.Start()
            timerRunning4 = True
            timerPaused4 = False
            PictureBox2.Image = Image.FromFile(imagePathPause)
        ElseIf timerRunning4 And Not timerPaused4 Then
            ' Pause the timer and calculate remaining time
            Timer4.Stop()
            remainingTimex4 = DateTime.Now - startTime4
            timerPaused4 = True
            ' Reset the image to the original one
            PictureBox2.Image = Image.FromFile(imagePathPlay)
            'PictureBox2.Dock = DockStyle.Fill
        ElseIf timerRunning4 And timerPaused4 Then
            ' Resume the timer with remaining time
            startTime4 = DateTime.Now - remainingTimex4
            Timer4.Start()
            timerPaused4 = False
            PictureBox2.Image = Image.FromFile(imagePathPause)
        End If
    End Sub
    '----- play5
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If Not timerRunning5 Then
            ' Start the timer
            startTime5 = DateTime.Now
            Timer5.Start()
            timerRunning5 = True
            timerPaused5 = False
            PictureBox4.Image = Image.FromFile(imagePathPause)
        ElseIf timerRunning5 And Not timerPaused5 Then
            ' Pause the timer and calculate remaining time
            Timer5.Stop()
            remainingTimex5 = DateTime.Now - startTime5
            timerPaused5 = True
            ' Reset the image to the original one
            PictureBox4.Image = Image.FromFile(imagePathPlay)
            'PictureBox4.Dock = DockStyle.Fill
        ElseIf timerRunning5 And timerPaused5 Then
            ' Resume the timer with remaining time
            startTime5 = DateTime.Now - remainingTimex5
            Timer5.Start()
            timerPaused5 = False
            PictureBox4.Image = Image.FromFile(imagePathPause)
        End If
    End Sub
    ''----- play6
    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        If Not timerRunning6 Then
            ' Start the timer
            startTime6 = DateTime.Now
            Timer6.Start()
            timerRunning6 = True
            timerPaused6 = False
            PictureBox6.Image = Image.FromFile(imagePathPause)
        ElseIf timerRunning6 And Not timerPaused6 Then
            ' Pause the timer and calculate remaining time
            Timer6.Stop()
            remainingTimex6 = DateTime.Now - startTime6
            timerPaused6 = True
            ' Reset the image to the original one
            PictureBox6.Image = Image.FromFile(imagePathPlay)
            'PictureBox6.Dock = DockStyle.Fill
        ElseIf timerRunning6 And timerPaused6 Then
            ' Resume the timer with remaining time
            startTime6 = DateTime.Now - remainingTimex6
            Timer6.Start()
            timerPaused6 = False
            PictureBox6.Image = Image.FromFile(imagePathPause)
        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        'Timer1.Stop()
        'remainingTime1 = DateTime.Now - startTime
        'timerPaused = True
        'UpdateTimeLabel1()
        Timer1.Stop()
        remainingTime1 = DateTime.Now - startTime
        timerPaused = True
        UpdateTimeLabel1()
        'PictureBox8.Enabled = False
        PictureBox8.Image = Image.FromFile(imagePathPlay)

    End Sub

    Private Sub PictureBox11_Click_1(sender As Object, e As EventArgs) Handles PictureBox11.Click
        Timer2.Stop()
        remainingTimex2 = DateTime.Now - startTime2
        timerPaused2 = True
        UpdateTimeLabel2()
        'PictureBox10.Enabled = False
        PictureBox10.Image = Image.FromFile(imagePathPlay)
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        Timer3.Stop()
        remainingTimex3 = DateTime.Now - startTime3
        timerPaused3 = True
        UpdateTimeLabel3()
        'PictureBox12.Enabled = False
        PictureBox12.Image = Image.FromFile(imagePathPlay)
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Timer4.Stop()
        remainingTimex4 = DateTime.Now - startTime4
        timerPaused4 = True
        UpdateTimeLabel4()
        'PictureBox2.Enabled = False
        PictureBox2.Image = Image.FromFile(imagePathPlay)
    End Sub
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Timer5.Stop()
        remainingTimex5 = DateTime.Now - startTime5
        timerPaused5 = True
        UpdateTimeLabel5()
        'PictureBox4.Enabled = False
        PictureBox4.Image = Image.FromFile(imagePathPlay)
    End Sub


    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Timer6.Stop()
        remainingTimex6 = DateTime.Now - startTime6
        timerPaused6 = True
        UpdateTimeLabel6()
        'PictureBox6.Enabled = False
        PictureBox6.Image = Image.FromFile(imagePathPlay)
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim folderPath As String = Path.GetDirectoryName(Application.ExecutablePath)

        'MessageBox.Show($"{Application.StartupPath}")
        'MessageBox.Show(Path.GetDirectoryName(Application.ExecutablePath))
        rtbSpiel.Text = "Good [morning/afternoon/evening], thank you for calling Intuit.
My name is [Your Name], and I'm here to assist you today. 

May I have the pleasure of knowing your name, please? 

Great! How can I help you today? Whether it's a question, a concern, or if you just need some information, 
I'm here to provide you with the best assistance possible. 
So, go ahead and let me know what you need, and we'll work together to find a solution.

Thank you for choosing Intuit, and I appreciate the opportunity to assist you!"

        rtbFAQs.Text = "Here are some common topics covered in travel-related FAQs:

Booking Process:

How do I book a flight/hotel/car rental/package on Expedia?
Can I make changes to my reservation?
What payment methods are accepted?
Account and Profile:

How do I create an Expedia account?
How can I reset my password?
What are the benefits of creating an account?
Cancellation and Refund Policies:

What is the cancellation policy for flights/hotels/packages?
How can I request a refund?
Travel Documentation:

What travel documents do I need?
How can I access my itinerary?
Expedia Rewards Program:

How does the Expedia Rewards program work?
How can I earn and redeem Expedia points?
Customer Support:

How can I contact Expedia customer support?
What should I do if I encounter issues during my trip?
Travel Insurance:

Does Expedia offer travel insurance?
What does the travel insurance cover?
Mobile App:

How do I download and use the Expedia mobile app?
What features are available on the mobile app?
Special Requests and Preferences:

How can I make special requests for my accommodation (e.g., room preferences)?
Can I request special meals for my flight?
Promotions and Discounts:

How can I find and apply discounts or promo codes?
Are there any ongoing promotions or exclusive deals?"
    End Sub

    Private Sub PictureBox15_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click
        Timer1.Stop() ' Stop the timer
        startTime = DateTime.Now
        Label11.Text = "hh:mm:ss"
        ProgressBar1.Value = 0
        timerRunning = False
        timerPaused = True
        PictureBox8.Image = Image.FromFile(imagePathPlay)
        Label11.ForeColor = Color.Black

        Timer2.Stop() ' Stop the timer
        startTime2 = DateTime.Now
        Label12.Text = "hh:mm:ss"
        ProgressBar2.Value = 0
        timerRunning2 = False
        timerPaused2 = True
        PictureBox10.Image = Image.FromFile(imagePathPlay)
        Label12.ForeColor = Color.Black

        Timer3.Stop() ' Stop the timer
        startTime3 = DateTime.Now
        Label13.Text = "hh:mm:ss"
        ProgressBar3.Value = 0
        timerRunning3 = False
        timerPaused3 = True
        PictureBox12.Image = Image.FromFile(imagePathPlay)
        Label13.ForeColor = Color.Black

        Timer4.Stop() ' Stop the timer
        startTime4 = DateTime.Now
        Label4.Text = "hh:mm:ss"
        ProgressBar4.Value = 0
        timerRunning4 = False
        timerPaused4 = True
        PictureBox2.Image = Image.FromFile(imagePathPlay)
        Label4.ForeColor = Color.Black

        Timer5.Stop() ' Stop the timer
        startTime5 = DateTime.Now
        Label5.Text = "hh:mm:ss"
        ProgressBar5.Value = 0
        timerRunning5 = False
        timerPaused5 = True
        PictureBox4.Image = Image.FromFile(imagePathPlay)
        Label5.ForeColor = Color.Black

        Timer6.Stop() ' Stop the timer
        startTime6 = DateTime.Now
        Label6.Text = "hh:mm:ss"
        ProgressBar6.Value = 0
        timerRunning6 = False
        timerPaused6 = True
        PictureBox6.Image = Image.FromFile(imagePathPlay)
        Label6.ForeColor = Color.Black


    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://ext-expediagroup.okta.com/")
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start("https://kronos8.nac.sitel-world.net/wfc/logon")
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Process.Start("https://forms.office.com/e/6zYbPMxapG")
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Process.Start("https://outlook.com")
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Process.Start("https://learningmanager.adobe.com/ExpediaTPSP")
    End Sub

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        Process.Start("https://apac-myacademy.learning-tribes.com/account/settings")
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Process.Start("https://sitel-16.us145lxiexweb.nac.sitel-world.net")
    End Sub

    Private Sub LinkLabel8_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel8.LinkClicked
        Process.Start("https://password.foundever.com")
    End Sub

    Private Sub LinkLabel9_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel9.LinkClicked
        Process.Start("https://everconnect.foundever.com/")
    End Sub

End Class
