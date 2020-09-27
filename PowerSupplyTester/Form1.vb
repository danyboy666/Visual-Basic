'------------------------------------------------------------------------------------------------------------------------------
'Titre du projet: Power Suply Tester
'R�vision: 1.0
'Date: 06/12/2011
'Copyright: Dany Ferron et J�r�me Pelletier
'------------------------------------------------------------------------------------------------------------------------------
'Projet final cr�e par Dany Ferron et J�r�me Pelletier pour Louis Dion dans le cadre du cours de surveillance 
'de syst�mes (243-575-RK) au C�gep de Rimouski. 
'
'Ce programme permet de tester le courant entrant et sortant d'un bloc d'alimentation tout en activant des 
'relais lorsqu'il y a surcharge de courant de sortie � l'aide du module de test CB7050.
'------------------------------------------------------------------------------------------------------------------------------

Public Class Form1
    'Variable globales
    Dim WithEvents serialPort As New IO.Ports.SerialPort
    Dim style As MsgBoxStyle
    Dim response As MsgBoxResult
    Dim DataSend As String
    Dim msg As String
    Dim UserMsg As String
    Dim SeqTest As Integer
    Dim TestPass As Integer
    Dim intResponse As Integer
    'Variables utilis�es pour prise de mesures
    Dim MesureAmp As Double
    Dim MesureVr1 As Double 'Mesure de tension � la r�sistance R1
    Dim MesureVr3 As Double 'Mesure de tension � la r�sistance R3
    Dim MesureVr5 As Double 'Mesure de tension � la r�sistance R5
    Dim MesureVr6 As Double 'Mesure de tension � la r�sistance R6
    Dim MesureVr7 As Double 'Mesure de tension � la r�sistance R7

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For i As Integer = 0 To My.Computer.Ports.SerialPortNames.Count - 1
            cbbCOMPorts.Items.Add(My.Computer.Ports.SerialPortNames(i))
        Next
        btnDisconnect.Enabled = False
        'Initialisation des variables au chargement du programme
        SeqTest = 1
        response = 0
        TestPass = 0
        BtnTest.Enabled = False
    End Sub

    Private Sub DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        txtDataReceived.Invoke(New myDelegate(AddressOf updateTextBox), New Object() {})
    End Sub

    Public Delegate Sub myDelegate()

    Public Sub updateTextBox()
        With txtDataReceived
            .Font = New Font("Garamond", 12.0!, FontStyle.Bold)
            .ForeColor = Color.Red
            .AppendText(serialPort.ReadExisting)
            .ScrollToCaret()
        End With
    End Sub

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
        Try
            With serialPort
                .PortName = cbbCOMPorts.Text
                .BaudRate = 9600
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
            End With
            serialPort.Open()
            lblMessage.Text = (cbbCOMPorts.Text & " connected" & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Start test")
            btnConnect.Enabled = False
            btnDisconnect.Enabled = True
            BtnTest.Enabled = True
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnDisconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnect.Click
        Try
            serialPort.Close()
            lblMessage.Text = serialPort.PortName & " disconnected."
            btnConnect.Enabled = True
            btnDisconnect.Enabled = False
            BtnTest.Enabled = False
            SeqTest = 1
            BtnTest.Text = "Start test"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub btnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnQuit.Click
        'Boite de dialogue pour la fonction "Quit"
        intResponse = MsgBox("Do you really want to quit?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Quit?")
        'Si l'utilisateur appui sur "oui" le programme se termine
        If intResponse = vbYes Then
            End
        End If
    End Sub

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnTest.Click
        SequenceDeTest()
    End Sub

    Public Sub SequenceDeTest()
        '-----------------------------------------------------------------------------------------------------------------------
        '                           Groupe de mesures Test de tension � l'entrer � 105V
        '-----------------------------------------------------------------------------------------------------------------------

        If SeqTest = 1 Then '�tape 1 : Ajuster tension au variac � 105V.

            msg = "Ajuster la tension d'entr�e � 105 Vac..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            SeqTest = 2 'Pr�paration pour la s�quence de test suivante
            BtnTest.Text = "Continuer test" 'Texte sur le bouton de test
            Exit Sub
        End If

        If SeqTest = 2 Then 'Test de tension � l'entrer � 105V.
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#020" 'Envoyer commande de test sur le canal 0 du module pour lecture de la tension.
            MesureVr1 = 1 'Dummy value for testing purposes, test de tension appliqu� par le variac

            'MesureVr1 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7))
            'txtDataReceived.Text = Str(MesureVr1) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            'V�rification de la mesure de tension prise
            'If MesureVr1 > 7.049 And MesureVr1 < 7.791 Then 'V�rification de la tension DC � 7.42(pourcentage d'�rreur 5%) V.
            If MesureVr1 = 1 Then
                SeqTest = 3 'Pr�paration pour la s�quence de test suivante
                Exit Sub

            Else
                msg = "Veuillez rev�rifier votre montage..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Start test" 'Message � l'utilisateur
                BtnTest.Text = "Start test"
                lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
                SeqTest = 1 'Si la valeur est fause, l'utilisateur devra r�ajuster la tension au variac, retour � la s�quence de test 1
                Exit Sub
            End If
        End If

        If SeqTest = 3 Then 'Test de courant � l'entrer � 105V.
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#021" 'Envoyer commande de test sur le canal 1 du module pour lecture de la tension.
            MesureVr3 = 1 'Dummy value for testing purposes, test sur la r�sitance de 1.9Ohms � l'entr� du circuit d'alimentation

            'MesureVr3 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr3) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            MesureAmp = (MesureVr3 / 1.9) 'Calcul du courant
            MesureAmp = FormatNumber(MesureAmp, 5) 'Mise en forme de la variable pour garder 5 chiffres apr�s la virgule

            If MesureAmp > 0.449 And MesureAmp < 0.791 Then
                UserMsg = " amp�res..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
                lblMessage.Text = Str(MesureAmp) + UserMsg 'Faire afficher le message � l'utilisateur
                BtnTest.Enabled = False
                System.Threading.Thread.Sleep(1000) 'Pause de 1 seconde
                BtnTest.Enabled = True
                SeqTest = 4 'Pr�paration pour la s�quence de test suivante
                Exit Sub

            Else
                msg = "Veuillez rev�rifier votre montage..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Start test" 'Message � l'utilisateur
                BtnTest.Text = "Start test"
                lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
                SeqTest = 1 'Si la valeur est fause, l'utilisateur devra r�ajuster la tension au variac
                Exit Sub
            End If
        End If
        '-----------------------------------------------------------------------------------------------------------------------
        '                           Groupe de mesures Test de tension � l'entrer � 125V
        '-----------------------------------------------------------------------------------------------------------------------

        If SeqTest = 4 Then '�tape 4 : Ajuster tension au variac � 125V.
            msg = "Ajuster la tension d'entr�e � 125 Vac..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            SeqTest = 5 'Pr�paration pour la s�quence de test suivante
            BtnTest.Text = "Continuer test" 'Texte sur le bouton de test
            Exit Sub
        End If

        If SeqTest = 5 Then 'Test de tension � l'entrer � 125V.
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#020" 'Envoyer commande de test sur le canal 0 du module pour lecture de la tension.
            MesureVr1 = 1 'Dummy value for testing purposes, test de tension appliqu�e par le variac

            'MesureVr1 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr1) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            'V�rification de la mesure de tension prise
            'If MesureVr1 > 7.049 And MesureVr1 < 7.791 Then 'V�rification de la tension DC � 7.42(pourcentage d'�rreur 5%) V.
            If MesureVr1 = 1 Then
                SeqTest = 6 'Pr�paration pour la s�quence de test suivante
                Exit Sub
            Else
                msg = "Veuillez rev�rifier votre montage..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Start test" 'Message � l'utilisateur
                BtnTest.Text = "Start test"
                lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
                SeqTest = 4 'Si la valeur est fause, l'utilisateur devra reajuster la tension au variac
                Exit Sub
            End If
        End If

        If TestPass = 1 And SeqTest = 6 Then
            msg = "Test � la sortie du bloc d'alimentation..." & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            SeqTest = 7 'Pr�paration pour la s�quence de test suivante
            Exit Sub
        End If

        If SeqTest = 6 Then 'Test de courant � l'entrer � 125V.
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#021" 'Envoyer commande de test sur le canal 1 du module pour lecture de la tension.
            MesureVr3 = 2 'Dummy value for testing purposes, test sur la r�sitance de 1.9 Ohms � l'entr� du circuit d'alimentation

            'MesureVr3 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr3) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            MesureAmp = (MesureVr3 / 1.9) 'Calcul du courant
            MesureAmp = FormatNumber(MesureAmp, 5) 'Mise en forme de la variable pour garder 5 chiffres apr�s la virgule

            If MesureAmp > 1.0 And MesureAmp < 1.08 Then
                UserMsg = " amp�res..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
                lblMessage.Text = Str(MesureAmp) + UserMsg 'Faire afficher le message � l'utilisateur
                BtnTest.Enabled = False
                System.Threading.Thread.Sleep(1000) 'Pause de 1 seconde
                BtnTest.Enabled = True
                TestPass = 1
                Exit Sub
            Else
                msg = "Veuillez rev�rifier votre montage..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Start test" 'Message � l'utilisateur
                BtnTest.Text = "Start test"
                lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
                SeqTest = 4 'Si la valeur est fause, l'utilisateur devra reajuster la tension au variac
                Exit Sub
            End If
        End If
        '-----------------------------------------------------------------------------------------------------------------------
        '                           Groupe de mesures Test de tension � la sortie de l'alimentation
        '-----------------------------------------------------------------------------------------------------------------------

        If SeqTest = 7 Then
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#022" 'Envoyer commande de test sur le canal 2 du module pour lecture de la tension.
            MesureVr5 = 5 'Dummy value for testing purposes, test sur la r�sitance de 1.9 Ohms � l'entr� du circuit d'alimentation

            'MesureVr5 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr5) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            MesureAmp = (MesureVr5 / 3.3) 'Calcul du courant
            MesureAmp = FormatNumber(MesureAmp, 5) 'Mise en forme de la variable pour garder 5 chiffres apr�s la virgule

            If MesureAmp > 1.25 And MesureAmp < 1.75 Then
                UserMsg = " amp�res..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
                lblMessage.Text = Str(MesureAmp) + UserMsg 'Faire afficher le message � l'utilisateur
                BtnTest.Enabled = False
                System.Threading.Thread.Sleep(1000) 'Pause de 1 seconde
                BtnTest.Enabled = True
                SeqTest = 8 'Pr�paration pour la s�quence de test suivante
                Exit Sub
            Else
                UserMsg = "Courant de charge trop �lev�..." & vbCrLf & vbCrLf & "veuillez mettre hors tension l'alimentation" 'Message � l'utilisateur
                lblMessage.Text = UserMsg 'Faire afficher le message � l'utilisateur
                DataSend = "#010" 'Envoyer commande de pour activer le relais
                Exit Sub
            End If
        End If

        If SeqTest = 8 Then
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#023" 'Envoyer commande de test sur le canal 3 du module pour lecture de la tension.
            MesureVr6 = 5 'Dummy value for testing purposes, test sur la r�sitance de 1.9Ohms � l'entr� du circuit d'alimentation

            'MesureVr6 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr6) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            MesureAmp = (MesureVr6 / 10) 'Calcul du courant
            MesureAmp = FormatNumber(MesureAmp, 5) 'Mise en forme de la variable pour garder 5 chiffres apr�s la virgule

            If MesureAmp > 0.4 And MesureAmp < 0.6 Then
                UserMsg = " amp�res..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
                lblMessage.Text = Str(MesureAmp) + UserMsg 'Faire afficher le message � l'utilisateur
                BtnTest.Enabled = False
                System.Threading.Thread.Sleep(1000) 'Pause de 1 seconde
                BtnTest.Enabled = True
                SeqTest = 9 'Pr�paration pour la s�quence de test suivante
                Exit Sub
            Else
                UserMsg = "Courant de charge trop �lev�..." & vbCrLf & vbCrLf & "veuillez mettre hors tension l'alimentation" 'Message � l'utilisateur
                lblMessage.Text = UserMsg 'Faire afficher le message � l'utilisateur
                DataSend = "#011" 'Envoyer commande de pour activer le relais
                Exit Sub
            End If
        End If

        If SeqTest = 9 Then
            msg = "Prise de mesure..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
            lblMessage.Text = msg 'Faire afficher le message � l'utilisateur
            DataSend = "#024" 'Envoyer commande de test sur le canal 4 du module pour lecture de la tension.
            MesureVr7 = 5 'Dummy value for testing purposes, test sur la r�sitance de 1.9Ohms � l'entr� du circuit d'alimentation

            'MesureVr7 = Val(Microsoft.VisualBasic.Right(txtDataReceived.Text, 7)) 'Garder en variable les 7 derniers chiffres des donn�es recus par le module de test
            'txtDataReceived.Text = Str(MesureVr7) 'Convertir en chaine de charact�re les donn�es recu par le module de test

            MesureAmp = (MesureVr7 / 20) 'Calcul du courant
            MesureAmp = FormatNumber(MesureAmp, 5) 'Mise en forme de la variable pour garder 5 chiffres apr�s la virgule

            If MesureAmp > 0.15 And MesureAmp < 0.35 Then
                UserMsg = " amp�res..." & vbCrLf & vbCrLf & "Pour continuer, appuyer sur le bouton: Continuer test" 'Message � l'utilisateur
                lblMessage.Text = Str(MesureAmp) + UserMsg 'Faire afficher le message � l'utilisateur
                BtnTest.Enabled = False
                System.Threading.Thread.Sleep(1000) 'Pause de 1 seconde
                BtnTest.Enabled = True
                SeqTest = 10 'Pr�paration pour la s�quence de test suivante
                Exit Sub
            Else
                UserMsg = "Courant de charge trop �lev�..." & vbCrLf & vbCrLf & "veuillez mettre hors tension l'alimentation" 'Message � l'utilisateur
                lblMessage.Text = UserMsg 'Faire afficher le message � l'utilisateur
                DataSend = "#012" 'Envoyer commande de pour activer le relais
                Exit Sub
            End If
        End If
        If SeqTest = 10 Then
            UserMsg = "Tous les tests ont �t� effectu�s avec succ�s!" 'Message � l'utilisateur
            lblMessage.Text = UserMsg 'Faire afficher le message � l'utilisateur
            End
        End If
    End Sub

    Private Sub txtDataReceived_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataReceived.TextChanged
    End Sub

    Private Sub lblMessage_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMessage.TextChanged
        'Mise en forme de la police d'affichage pour la boite de message destin� � l'utilisateur
        With lblMessage
            .Font = New Font("Garamond", 10.0!, FontStyle.Bold)
            .ForeColor = Color.Black
        End With
    End Sub

    Private Sub lblAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblAbout.Click
        'Boite de dialogue "About"
        MsgBox("Projet final cr�e par Dany Ferron et J�r�me Pelletier pour Louis Dion dans le cadre du cours de surveillance de syst�mes (243-575-RK) au C�gep de Rimouski 2011" & vbCrLf & "Ce programme permet de tester le courant entrant et sortant d'un bloc d'alimentation tout en activant des relais � l'aide d'un module de test CB7050", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "About")
    End Sub

    Private Sub lblMessage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMessage.Click

    End Sub

    Private Sub cbbCOMPorts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbbCOMPorts.SelectedIndexChanged

    End Sub
End Class
