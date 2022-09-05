'**************************************
'*       Pyramid solitaire            *
'**************************************
'* The game can be compared to        *
'* the original version of pyramid    *
'* solitaire                          *
'**************************************
'* Iván Viemoes Cuevas   27-04-2020   *
'**************************************


Public Class Form1

    'Creates a random number, which is used to find random cards.
    Shared random As New Random()
    Dim rn As Integer

    'Creates two arrays, the first is used for all the cards, and the second for the pile.
    Dim Kort(51) As Integer
    Dim række(28) As PictureBox

    'Creates two variables, which helps "unlock" the cards when those ontop are gone.
    Dim Kort2 As Integer
    Dim NæsteRække As Integer

    'Creates an index variable, which is used to input number to the array.
    Dim i As Integer = 0

    'Creates two variable which is used to remove and find cards in the array.
    Dim o As Integer = 1
    Dim p As Integer

    'Remembers the points
    Dim point As Integer

    'Constant to make the loop run forever
    Const konstant As Boolean = True

    'Creates a dim that remembers how many cards has been turned over, one that puts a value on the card
    'and one that remembers how many times the pile has turned
    Dim k As Integer = 27
    Dim antalfyldte As Integer
    Dim bunkervendt As Integer

    'Creates a variable that remembers the sum of the cards, how many cards has been clicked, and remembers the card a picture
    Dim sum As Integer
    Dim KortKlikket As Integer
    Dim HuskKort As PictureBox

    'Bruges til at lave spillepladen, og felternes placering.
    'Is used to create the playing board and the location of the marks
    Dim ANTAL_FELTER_X As Integer
    Dim ANTAL_FELTER_Y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim PictureBoxFelt(ANTAL_FELTER_X, ANTAL_FELTER_Y) As PictureBox


    'program start
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'At program start: Create all picturebox on the playing board.
        OpretSpilleplade(ANTAL_FELTER_X, ANTAL_FELTER_Y)

        'Sets the background color, size of the image, and the frame of the picturebox for the pile.
        PicStartBunke.BackColor = Color.LightGray
        PicSlutBunke.BackColor = Color.LightGray
        PicStartBunke.SizeMode = PictureBoxSizeMode.StretchImage
        PicSlutBunke.SizeMode = PictureBoxSizeMode.StretchImage
        PicStartBunke.BorderStyle = BorderStyle.FixedSingle
        PicSlutBunke.BorderStyle = BorderStyle.FixedSingle

        'Sets the starting image on the pile as the backside of a card, to know that the game just startet.
        'It also gets the value "0".
        PicStartBunke.Image = My.Resources.ResourceManager.GetObject("back")
        PicStartBunke.Tag = 0

        'Finds the values of the pile you pull from, until the array is full.
        'Does it by finding a card between 1 & 52, and checks wether or not the value is in the array
        'if the value is not in the array, its put in at an empty space.
        Do While antalfyldte < 24
            rn = random.Next(1, 52)
            If Not Kort.Contains(rn) Then
                Kort(k) = rn
                antalfyldte = antalfyldte + 1
                k = k + 1
            End If
        Loop

    End Sub


    'Create the playing board with the pictureboxes in an array as playing fields
    Private Sub OpretSpilleplade(AntalFelterX As Integer, AntalFelterY As Integer)

        'Creates row 1
        NextRandom()
        Opretspillefelt(155, 221)
        NextRandom()
        Opretspillefelt(216, 221)
        NextRandom()
        Opretspillefelt(277, 221)
        NextRandom()
        Opretspillefelt(338, 221)
        NextRandom()
        Opretspillefelt(399, 221)
        NextRandom()
        Opretspillefelt(460, 221)
        NextRandom()
        Opretspillefelt(521, 221)
        NextRandom()

        'Creates row 2
        Opretspillefelt(186, 201)
        NextRandom()
        Opretspillefelt(247, 201)
        NextRandom()
        Opretspillefelt(308, 201)
        NextRandom()
        Opretspillefelt(369, 201)
        NextRandom()
        Opretspillefelt(430, 201)
        NextRandom()
        Opretspillefelt(491, 201)
        NextRandom()

        'Creates row 3
        Opretspillefelt(216, 181)
        NextRandom()
        Opretspillefelt(277, 181)
        NextRandom()
        Opretspillefelt(338, 181)
        NextRandom()
        Opretspillefelt(399, 181)
        NextRandom()
        Opretspillefelt(460, 181)
        NextRandom()

        'Creates row 4
        Opretspillefelt(246, 161)
        NextRandom()
        Opretspillefelt(307, 161)
        NextRandom()
        Opretspillefelt(368, 161)
        NextRandom()
        Opretspillefelt(429, 161)
        NextRandom()

        'Creates row 5
        Opretspillefelt(276, 141)
        NextRandom()
        Opretspillefelt(337, 141)
        NextRandom()
        Opretspillefelt(398, 141)
        NextRandom()

        'Creates row 6
        Opretspillefelt(306, 121)
        NextRandom()
        Opretspillefelt(367, 121)
        NextRandom()

        'Creates row 7
        Opretspillefelt(337, 101)

    End Sub

    'Click routine mutual for all the pictureboxes on the playing board
    Private Sub KlikFelt(sender As Object, e As EventArgs) Handles PicStartBunke.Click, PicSlutBunke.Click

    'resets values stored in "o" and "p"
        o = 1
        p = 1

    'Find the picturebox that was clicked.
        Dim PictureBoxKlikketPå As PictureBox = sender
        PictureBoxKlikketPå.Tag = sender.tag

        'The program enters one of the following if-statements, compared to which value the picturebox holds
        'The value is then changed to a number between 1 and 13, the same as the card itself. It also remembers how many cards was clicked
        'and marks the picturepoxes, unless it's a card without a value, as the card only showing the backside.
        If 0 < CInt(PictureBoxKlikketPå.Tag) And CInt(PictureBoxKlikketPå.Tag) <= 13 Then
            sum = sum + PictureBoxKlikketPå.Tag
            KortKlikket = KortKlikket + 1
            PictureBoxKlikketPå.BorderStyle = BorderStyle.Fixed3D
        ElseIf 13 < CInt(PictureBoxKlikketPå.Tag) And CInt(PictureBoxKlikketPå.Tag) <= 26 Then
            sum = sum + (PictureBoxKlikketPå.Tag - 13)
            KortKlikket = KortKlikket + 1
            PictureBoxKlikketPå.BorderStyle = BorderStyle.Fixed3D
        ElseIf 26 < CInt(PictureBoxKlikketPå.Tag) And CInt(PictureBoxKlikketPå.Tag) <= 39 Then
            sum = sum + (PictureBoxKlikketPå.Tag - 26)
            KortKlikket = KortKlikket + 1
            PictureBoxKlikketPå.BorderStyle = BorderStyle.Fixed3D
        ElseIf 39 < CInt(PictureBoxKlikketPå.Tag) And CInt(PictureBoxKlikketPå.Tag) <= 52 Then
            sum = sum + (PictureBoxKlikketPå.Tag - 39)
            KortKlikket = KortKlikket + 1
            PictureBoxKlikketPå.BorderStyle = BorderStyle.Fixed3D
        ElseIf CInt(PictureBoxKlikketPå.Tag) = 0 Then
            Exit Sub
        End If

        'the clicked card is remembered if only one card has been clicked
        If KortKlikket = 1 Then
            HuskKort = PictureBoxKlikketPå
        End If


        '*******************************************************************************
        '* In the following code, it figures which of the pile has been clicked        *
        '* and which one.                                                              *
        '*******************************************************************************
        '* If the card pressed is from "PicStartBunke", the value inside the array     *
        '* is set to nothing, and it shows the next possible card.                     *
        '*******************************************************************************
        '* If the card clicked is from "PicSlutBunke", the value minus 1 is set to     *
        '* nothing and shows the first possible card with a value, that was before it. *
        '* If no card was before it shows nothing.                                     *
        '*******************************************************************************
        '* If a card from "PicStartBunke" and "PicSLutBunke" is matched,               *
        '* the previous two are triggered.                                             *
        '*******************************************************************************
        '* We also put the frames back, so the arent selected, puts the sum and        *
        '* amount of times clicked to zero, and counts the points.                     *
        '*******************************************************************************
        'If the pciturebox clicked is either "PicStartBunke" or "PicSlutBunke", and the sum is 13, enter the if
        If CInt(PictureBoxKlikketPå.Tag) = CInt(PicStartBunke.Tag) And CInt(HuskKort.Tag) = CInt(PicSlutBunke.Tag) And sum = 13 Then
            Kort(k) = Nothing

            Do While Kort(k - p) = Nothing
                p = p - 1
            Loop

            Kort(k - p) = Nothing

            Do While Kort(k) = Nothing And k < 52
                k = k + 1
            Loop

            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)

            Do While Kort(k - o) = Nothing
                o = o + 1
            Loop

            If (k - o) < 27 Then
                PicSlutBunke.Image = Nothing
                PicSlutBunke.Tag = Nothing
            Else
                PicSlutBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k - o))
                PicSlutBunke.Tag = Kort(k - o)
            End If



            PicStartBunke.BorderStyle = BorderStyle.FixedSingle
            PicSlutBunke.BorderStyle = BorderStyle.FixedSingle

            sum = 0
            KortKlikket = 0
            point = point + 10

            'If the starting pile is clicked and the sum is 13, enter the if.
        ElseIf CInt(PictureBoxKlikketPå.Tag) = CInt(PicStartBunke.Tag) And sum = 13 Then
            Kort(k) = Nothing

            Do While Kort(k) = Nothing And k < 52
                k = k + 1
            Loop

            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)

            If KortKlikket = 2 Then
                HuskKort.Visible = False
            End If

            PicStartBunke.BorderStyle = BorderStyle.FixedSingle

            sum = 0
            KortKlikket = 0
            point = point + 10

            'If the second card clicked is from the start pile, and the sum is 13, enter the if.
        ElseIf CInt(HuskKort.Tag) = CInt(PicStartBunke.Tag) And sum = 13 Then
            Kort(k) = Nothing

            Do While Kort(k) = Nothing And k < 52
                k = k + 1
            Loop

            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)

            If KortKlikket = 2 Then
                PictureBoxKlikketPå.Visible = False
            End If

            PicStartBunke.BorderStyle = BorderStyle.FixedSingle

            sum = 0
            KortKlikket = 0
            point = point + 10

            'If the first card clicked is from the second pile, and the sum is 13, enter the if.
        ElseIf CInt(PictureBoxKlikketPå.Tag) = CInt(PicSlutBunke.Tag) And sum = 13 Then
            Do While Kort(k - p) = Nothing
                p = p - 1
            Loop

            Kort(k - p) = Nothing

            Do While Kort(k - o) = Nothing
                o = o + 1
            Loop

            If (k - o) < 27 Then
                PicSlutBunke.Image = Nothing
                PicSlutBunke.Tag = Nothing
            Else
                PicSlutBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k - o))
                PicSlutBunke.Tag = Kort(k - o)
            End If

            If KortKlikket = 2 Then
                HuskKort.Visible = False
            End If

            PicSlutBunke.BorderStyle = BorderStyle.FixedSingle

            sum = 0
            KortKlikket = 0
            point = point + 10

            'If the second card clicked is from the second pile, and the sum is 13, enter the if.
        ElseIf CInt(HuskKort.Tag) = CInt(PicSlutBunke.Tag) And sum = 13 Then
            Do While Kort(k - p) = Nothing
                p = p - 1
            Loop
            Kort(k - p) = Nothing

            Do While Kort(k - o) = Nothing
                o = o + 1
            Loop

            If (k - o) < 27 Then
                PicSlutBunke.Image = Nothing
                PicSlutBunke.Tag = Nothing
            Else
                PicSlutBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k - o))
                PicSlutBunke.Tag = Kort(k - o)
            End If

            If KortKlikket = 2 Then
                PictureBoxKlikketPå.Visible = False
            End If

            PicSlutBunke.BorderStyle = BorderStyle.FixedSingle

            sum = 0
            KortKlikket = 0
            point = point + 10

            'If the sum is 13, enter the if.
        ElseIf sum = 13 Then
            PictureBoxKlikketPå.Visible = False
            HuskKort.Visible = False
            sum = 0
            KortKlikket = 0
            point = point + 10

            'If two cards has been clicked, enter the if.
        ElseIf KortKlikket = 2 Then
            sum = 0
            KortKlikket = 0
            PictureBoxKlikketPå.BorderStyle = BorderStyle.None
            HuskKort.BorderStyle = BorderStyle.None
        End If

        TextBox_Point.Text = point

        'Checks if any cards has "unlocked". Also makes it possible to use them if so.
        Unlock()

    End Sub

    Sub NextRandom()
        'Runs the loop until a new value is found for the array.
        Do While konstant = True
            'Finds a random number between 1 and 52.
            rn = random.Next(1, 52)
        'Checks if the arrays contains the found number, enters the if, if not.
            If Not Kort.Contains(rn) Then
            'Inputs the found number in the array.
                Kort(i) = rn

            'Exits the loop
                Exit Do
            End If
        Loop
    End Sub

    Private Function Opretspillefelt(LocX As Single, LocY As Single) As Double
        'Creates new object
        PictureBoxFelt(x, y) = New PictureBox
        'Sets properties size, position color etc.
        PictureBoxFelt(x, y).Size = New System.Drawing.Size(62, 88)
        'Placement depends on "x" and "y".
        PictureBoxFelt(x, y).Location = New System.Drawing.Point(LocX, LocY)
        PictureBoxFelt(x, y).BackColor = Nothing
        PictureBoxFelt(x, y).SizeMode = PictureBoxSizeMode.StretchImage
        PictureBoxFelt(x, y).BorderStyle = BorderStyle.None

        'Sets the image equal to the value inside the array.
        PictureBoxFelt(x, y).Image = My.Resources.ResourceManager.GetObject("_" & Kort(i))
        'Gives the picturebox the value inside the array.
        PictureBoxFelt(x, y).Tag = Kort(i)

        'enables the first row, disables the rest of the rows
        If i > 6 Then
            PictureBoxFelt(x, y).Enabled = False
        End If

        'figures out which space the picturebox is in, form 0 to 27.
        række(i) = PictureBoxFelt(x, y)

        'Increases the i value
        i = i + 1

        'Adds the picturebox to form1
        Me.Controls.Add(Me.PictureBoxFelt(x, y))

        'Sæt klik-rutine til sub'en KlikFelt(). 
        'Dvs. KlikFelt() bliver FÆLLES klik-rutine for ALLE pictureboxene
        AddHandler PictureBoxFelt(x, y).Click, AddressOf KlikFelt
    End Function


    Private Sub ButtonNext_Click(sender As Object, e As EventArgs) Handles ButtonNext.Click
        'When the next button is pressed, the picture on "PicSlutBunke" and value is set to "PicStartBunke".
        PicSlutBunke.Image = PicStartBunke.Image
        PicSlutBunke.Tag = PicStartBunke.Tag

        'If k is less than fifty, enter the if and add to k until a usable value is in the array.
        If k < 50 Then
            Do While Kort(k + 1) = Nothing
                k = k + 1
            Loop
        End If


        'Adds 1 to k to find the next card.
        k = k + 1

        'If k becomes more than 52, set it to 27 and remember how many times the pile has been turned.
        If k >= 52 Then
            k = 27
            bunkervendt = bunkervendt + 1
        End If

        'If the card moved is the last in the pile, put the picturebox as an empty.
        If k = 51 Then
            PicStartBunke.Image = Nothing
            PicStartBunke.Tag = Nothing

            'If the game has just started put a card on the start pile and nothing in the second pile.
        ElseIf k = 27 Then
            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)
            PicSlutBunke.Image = Nothing

            'If none of the above match, run the pile and show a new card everytime you click.
        Else
            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)
        End If

        'If the pile has been turned three times, the game is over and a button is hsown to create a new game.
        If bunkervendt > 3 Then
            PictureBox_GameOver.Visible = True
            Button_NewGame.Enabled = True
            Button_NewGame.Visible = True
            Label_NewGame.Visible = True
            ButtonNext.Enabled = False
        End If
    End Sub

    '
    '
    '*****************************************************************
    '* In the following it checks wether or not some cards           *
    '* has been unlocked, and can be used.                           *
    '*****************************************************************
    '* It's done by checking all the rows, if two cards next to      *
    '* eachother are visible or not. If not, it'll unlock the card   *
    '* beneath them. It does this for every row except the last one, *
    '* if no card is on the last space, a new window shows that says *
    '* you've won.                                                   *
    '*****************************************************************
    Sub Unlock()
        For Kort As Integer = 0 To 5
            Kort2 = Kort + 1
            NæsteRække = Kort + 7
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        For Kort As Integer = 7 To 11
            Kort2 = Kort + 1
            NæsteRække = Kort + 6
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        For Kort As Integer = 13 To 16
            Kort2 = Kort + 1
            NæsteRække = Kort + 5
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        For Kort As Integer = 18 To 20
            Kort2 = Kort + 1
            NæsteRække = Kort + 4
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        For Kort As Integer = 22 To 23
            Kort2 = Kort + 1
            NæsteRække = Kort + 3
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        For Kort As Integer = 25 To 25
            Kort2 = Kort + 1
            NæsteRække = Kort + 2
            If række(Kort).Visible = False And række(Kort2).Visible = False Then
                række(NæsteRække).Enabled = True
            End If
        Next

        If række(27).Visible = False Then
            Du_vandt.Show()
            Me.Visible = False
        End If

    End Sub


    Private Sub Button_NewGame_Click(sender As Object, e As EventArgs) Handles Button_NewGame.Click
        'If the button is pressed, the game restarts.
        Application.Restart()
    End Sub
End Class
