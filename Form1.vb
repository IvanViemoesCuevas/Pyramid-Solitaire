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

    'Laver random tal, som kan bruges til at finde random kort.
    Shared random As New Random()
    Dim rn As Integer

    'Laver 2 arrays, det ene er til alle kortene, og det andet er til bunken.
    Dim Kort(51) As Integer
    Dim række(28) As PictureBox

    'Laver 2 variabler, der hjælper med at at "unlocke" kort, når andre bliver fjernet.
    Dim Kort2 As Integer
    Dim NæsteRække As Integer

    'Laver en index variabel, som bruges til at indsætte tal i arrayet.
    Dim i As Integer = 0

    'laver 2 variabler, der bruges til at fjerne og finde kort i array
    Dim o As Integer = 1
    Dim p As Integer

    'Husker pointene
    Dim point As Integer

    'Konstant der får et loop til at køre for evigt
    Const konstant As Boolean = True

    'laver en dim der husker hvor mange kort man har vendt, en der sætter en værdi på kortene,
    'og en der husker hvor mange gange hele bunken er vendt
    Dim k As Integer = 27
    Dim antalfyldte As Integer
    Dim bunkervendt As Integer

    'laver en variabel der husker summen af kortet, hvor mange kort der er blevet klikket, og husker kortet som billede
    Dim sum As Integer
    Dim KortKlikket As Integer
    Dim HuskKort As PictureBox

    'Bruges til at lave spillepladen, og felternes placering.
    Dim ANTAL_FELTER_X As Integer
    Dim ANTAL_FELTER_Y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim PictureBoxFelt(ANTAL_FELTER_X, ANTAL_FELTER_Y) As PictureBox


    'program start
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Ved program start: Opret alle picturebox felter på spillepladen
        OpretSpilleplade(ANTAL_FELTER_X, ANTAL_FELTER_Y)

        'sætter baggrundsfarven, størrelsen af billedet, og rammerne på pictureboxene for bunken
        PicStartBunke.BackColor = Color.LightGray
        PicSlutBunke.BackColor = Color.LightGray
        PicStartBunke.SizeMode = PictureBoxSizeMode.StretchImage
        PicSlutBunke.SizeMode = PictureBoxSizeMode.StretchImage
        PicStartBunke.BorderStyle = BorderStyle.FixedSingle
        PicSlutBunke.BorderStyle = BorderStyle.FixedSingle

        'sætter start billedet på PicStartBunke som bagsiden af et kort, for at man kan se at spillet lige er startet.
        'den får desuden en værdi på 0.
        PicStartBunke.Image = My.Resources.ResourceManager.GetObject("back")
        PicStartBunke.Tag = 0

        'finder værdierne til bunken man trækker fra, indtil arrayet er fyldt.
        'Dette gør den ved at finde et kort mellem 1 og 52, og tjekker om værdien er i arrayet
        'hvis værdien ikke er i arrayet, sætter den værdien ind i arrayet, på en fri plads.
        Do While antalfyldte < 24
            rn = random.Next(1, 52)
            If Not Kort.Contains(rn) Then
                Kort(k) = rn
                antalfyldte = antalfyldte + 1
                k = k + 1
            End If
        Loop

    End Sub


    'Opret spilleplade med pictureboxe i et array som spillefelter
    Private Sub OpretSpilleplade(AntalFelterX As Integer, AntalFelterY As Integer)

        'Opretter række 1
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

        'Opretter række 2
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

        'Opretter række 3
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

        'Opretter række 4
        Opretspillefelt(246, 161)
        NextRandom()
        Opretspillefelt(307, 161)
        NextRandom()
        Opretspillefelt(368, 161)
        NextRandom()
        Opretspillefelt(429, 161)
        NextRandom()

        'Opretter række 5
        Opretspillefelt(276, 141)
        NextRandom()
        Opretspillefelt(337, 141)
        NextRandom()
        Opretspillefelt(398, 141)
        NextRandom()

        'Opretter række 6
        Opretspillefelt(306, 121)
        NextRandom()
        Opretspillefelt(367, 121)
        NextRandom()

        'Opretter række 7
        Opretspillefelt(337, 101)

    End Sub

    'Klik-rutine fælles for alle pictureboxes på spillepladen
    Private Sub KlikFelt(sender As Object, e As EventArgs) Handles PicStartBunke.Click, PicSlutBunke.Click

        'resetter værdierne for henholdsvis o & p
        o = 1
        p = 1

        'Find den picturebox, der blev klikket på
        Dim PictureBoxKlikketPå As PictureBox = sender
        PictureBoxKlikketPå.Tag = sender.tag

        'programmet går ind i en af de følgende if'er i forhold til hvilken værdi pictureboxen har
        'inde i if'en laver den værdien om til et tal mellem 1 og 13, husker hvor mange kort der er blevet klikke på,
        'og markere pictureboxen, medmindre det er et kort uden værdi, som bagsiden
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

        'hvis der kun er blevet klikket på et kort husker den pictureboxen
        If KortKlikket = 1 Then
            HuskKort = PictureBoxKlikketPå
        End If


        '*******************************************************************************
        '* I følgende kode finder den ud af om der er blevet trykket på en af de to    *
        '* pictureboxe, der bruges til bunken, og hvilke der er trykket.               *
        '*******************************************************************************
        '* I tilfælde hvor et af kortene trykket er fra PicStartBunke sætter den, den  *
        '* nuværende værdi i arrayet til intet, og viser det næste mulige kort i       *
        '* arrayet.                                                                    *
        '*******************************************************************************
        '* I tilfælde hvor et af kortene trykket, er fra PicSlutBunke, sætter den,     *
        '* den nuværende værdi minus 1 i arrayet, til intet og viser det kort før,     *
        '* der har en værdi i arrayet. hvis der intet kort er før, som har en værdi i  *
        '* arrayet, viser den ikke noget kort.                                         *
        '*******************************************************************************
        '* Hvis et kort fra PicStartBunke og PicSlutBunke bliver matchet, vil der ske  *
        '* begge dele, både med kort, foran og bagved det kort der er matchet.         *
        '*******************************************************************************
        '* Udover dette, sætter vi også rammerne tilbage, så de ikke er markeret       *
        '* længere, sætter summen og kortlikkende til 0, og tæller point               *
        '*******************************************************************************
        'hvis pictureboxene der er blevet klikket på er PicStartBunke og PicSlutBunke, og summen er 13, gå ind i if'en
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

            'hvis pictureboxene der er blevet klikket på er PicStartBunke og PicSlutBunke, og summen er 13, gå ind i if'en
        ElseIf CInt(PictureBoxKlikketPå.Tag) = CInt(PicSlutBunke.Tag) And CInt(HuskKort.Tag) = CInt(PicStartBunke.Tag) And sum = 13 Then
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

            'hvis der er klikket på start bunken og summen er 13, gå ind i if'en
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

            'hvis det andet kort der er klikket på, er start bunken, og summen giver 13, gå ind i if'en
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

            'Hvis det første kort der er klikket på, er kortet i slutbunken og summen er 13, gå ind i if'en
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

            'hvis det andet kort der er klikket på er slutbunken og summen giver 13
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

            'hvis summen er 13 gå ind i if'en
        ElseIf sum = 13 Then
            PictureBoxKlikketPå.Visible = False
            HuskKort.Visible = False
            sum = 0
            KortKlikket = 0
            point = point + 10

            'hvis der er klikket på 2 kort, gå ind i if'en
        ElseIf KortKlikket = 2 Then
            sum = 0
            KortKlikket = 0
            PictureBoxKlikketPå.BorderStyle = BorderStyle.None
            HuskKort.BorderStyle = BorderStyle.None
        End If

        TextBox_Point.Text = point

        'Tjekker om der er kort der er blevet frie, og hvis der er, gør det, det muligt at bruge dem.
        Unlock()

    End Sub

    Sub NextRandom()
        'Kører loopet indtil der er fundet en ny værdi til arrayet
        Do While konstant = True
            'finder et random tal mellem 1 og 52
            rn = random.Next(1, 52)
            'Tjekker om arrayet indeholder det random tal, hvis ikke går den i if'en
            If Not Kort.Contains(rn) Then
                'sætter det random tal ind i arrayet
                Kort(i) = rn

                'går ud af loopet
                Exit Do
            End If
        Loop
    End Sub

    Private Function Opretspillefelt(LocX As Single, LocY As Single) As Double
        'Opret nyt object
        PictureBoxFelt(x, y) = New PictureBox
        'Sæt egenskaber størrelse, position, farve, osv.
        PictureBoxFelt(x, y).Size = New System.Drawing.Size(62, 88)
        'Placering afhænger af x og y:
        PictureBoxFelt(x, y).Location = New System.Drawing.Point(LocX, LocY)
        PictureBoxFelt(x, y).BackColor = Nothing
        PictureBoxFelt(x, y).SizeMode = PictureBoxSizeMode.StretchImage
        PictureBoxFelt(x, y).BorderStyle = BorderStyle.None

        'sætter billedet lig med værdien i arrayet
        PictureBoxFelt(x, y).Image = My.Resources.ResourceManager.GetObject("_" & Kort(i))
        'giver pictureboxen værdien der er i arrayet
        PictureBoxFelt(x, y).Tag = Kort(i)

        'enabler hele den første række, disabler resten af rækkerne
        If i > 6 Then
            PictureBoxFelt(x, y).Enabled = False
        End If

        'finder ud af hvilket felt pictureboxen er i fra 0..27
        række(i) = PictureBoxFelt(x, y)

        'laver en ny i-værdi
        i = i + 1

        'Tilknyt pictureboxen til Form1
        Me.Controls.Add(Me.PictureBoxFelt(x, y))

        'Sæt klik-rutine til sub'en KlikFelt(). 
        'Dvs. KlikFelt() bliver FÆLLES klik-rutine for ALLE pictureboxene
        AddHandler PictureBoxFelt(x, y).Click, AddressOf KlikFelt
    End Function


    Private Sub ButtonNext_Click(sender As Object, e As EventArgs) Handles ButtonNext.Click
        'Når der trykkes "next" sætter den PicSlutBunke's billede og værdi lig PicStartBunke
        PicSlutBunke.Image = PicStartBunke.Image
        PicSlutBunke.Tag = PicStartBunke.Tag

        'Hvis k er mindre end halvtreds går den ind i if'en og plusser k indtil der er en brugbar værdi i arrayet
        If k < 50 Then
            Do While Kort(k + 1) = Nothing
                k = k + 1
            Loop
        End If


        'plusser k med 1, for at finde det næste kort
        k = k + 1

        'hvis k bliver over 52, sætter den k lig 27 og husker hvor mange gange du har vendt bunken
        If k >= 52 Then
            k = 27
            bunkervendt = bunkervendt + 1
        End If

        'hvis det er det sidste kort der bliver flyttet viser den bunken som tom
        If k = 51 Then
            PicStartBunke.Image = Nothing
            PicStartBunke.Tag = Nothing

            'hvis man lige har startet bunken sætter den et kort på startbunken, men slutbunken er tom
        ElseIf k = 27 Then
            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)
            PicSlutBunke.Image = Nothing

            'hvis ingen af de foregående if'er gælder, kører den bunken igennem og viser et nyt kort for hver gang du trykker
        Else
            PicStartBunke.Image = My.Resources.ResourceManager.GetObject("_" & Kort(k))
            PicStartBunke.Tag = Kort(k)
        End If

        'Hvis bunken er blevet vendt 3 gange kan du ikke spille mere, og der kommer en knap frem, som bruges til at lave nyt spil
        If bunkervendt > 3 Then
            PictureBox_GameOver.Visible = True
            Button_NewGame.Enabled = True
            Button_NewGame.Visible = True
            Label_NewGame.Visible = True
            ButtonNext.Enabled = False
        End If
    End Sub


    '*****************************************************************
    '* I følgende kode tjekker den om der er nogle kort der kan      *
    '* blive unlocked og bruges i spillet                            *
    '*****************************************************************
    '* Dette gør den ved at køre den første bunke igennem som et     *
    '* loop, og tjekker for hvert "par" kort ved siden af hinanden   *
    '* om de er visible, hvis de ikke er det, skal den unlocke det   *
    '* kort under de to kort der var der før.                        *
    '* dette gør den for hver række, indtil vi når den sidste række. * 
    '* her der tjekker den om det sidste kort er der, hvis ikke,     *
    '* viser den et nyt vindue, som fortæller dig du har vundet.     *
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
        'hvis der trykket på knappen starter den spillet forfra
        Application.Restart()
    End Sub
End Class
