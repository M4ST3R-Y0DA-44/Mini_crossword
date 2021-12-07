' Nom Programme: Jeux Buggle a deux joueur
' Programme: Concu dans le cadre d'un travail pratique pour l'universite
' Programmeur: Yannick Munger 23 Mars 2021

Option Explicit On
Option Infer Off
Option Strict On



Public Class Form1

    ' Declaration de 5 variable de class qui est utilisé pour le calcul du nombre de mot total et le nombre de mot moyen ainsi que pour le chronometre. 
    Private intNombreMot1 As Integer
    Private intNombreMot2 As Integer
    Private intNombreMoyenLettre1 As Integer
    Private intNombreMoyenLettre2 As Integer
    Private intConteurSeconde As Integer

    ' Fonction qui Renvoi Vrai si aucun element de la grille est en bleu (donc qu'aucun élément a été selectionné)
    ' Cela permet de selectionner n'importe quel label dans la grille
    ' lorsque le label est le premier a etre selectionné lors de la preuve.
    Private Function Premier_click() As Boolean
        If (lblGrille1.ForeColor = Color.Black And lblGrille2.ForeColor = Color.Black And lblGrille3.ForeColor = Color.Black And
                lblGrille4.ForeColor = Color.Black And lblGrille5.ForeColor = Color.Black And lblGrille6.ForeColor = Color.Black And
        lblGrille7.ForeColor = Color.Black And lblGrille8.ForeColor = Color.Black And lblGrille9.ForeColor = Color.Black And
        lblGrille10.ForeColor = Color.Black And lblGrille11.ForeColor = Color.Black And lblGrille12.ForeColor = Color.Black And
        lblGrille13.ForeColor = Color.Black And lblGrille14.ForeColor = Color.Black And lblGrille15.ForeColor = Color.Black And
        lblGrille16.ForeColor = Color.Black) Then
            Return True
        End If
    End Function

    ' Procédure qui permet de retrouver les parametre initial enable et text des boutons des joueur ainsi que certain element du menu. 
    Private Sub Propriete_initial()
        btnJoueur1.Enabled = False
        btnJoueur2.Enabled = False
        txtMotJoueur1.Enabled = False
        txtMotJoueur2.Enabled = False
        btnTerminer.Enabled = False
        TerminerToolStripMenuItem.Enabled = False
        Joueur1ToolStripMenuItem.Enabled = False
        Joueur2ToolStripMenuItem.Enabled = False
        PreuveToolStripMenuItem.Enabled = False
        PreuveToolStripMenuItem1.Enabled = False
        btnJoueur1.Text = "P&reuve"
        btnJoueur2.Text = "Pr&euve"
    End Sub

    ' Procédure qui effectue le changement en bleu des lettre selectionné et les ajoute au label de verification pour les comparer.
    ' Si la comparaison est identique elle effectue aussi la calcul du nombre de mot et la moyenne du nombre de lettre par mots et les mets a jours. 
    ' Elle retourne les bouton a preuve et prépare pour la prochaine vérification. 
    Private Sub Verification(ByVal lblnomGrille As Label)
        If (txtMotJoueur1.Enabled = True And txtMotJoueur1.Text <> String.Empty) Or (txtMotJoueur2.Enabled = True And txtMotJoueur2.Text <> String.Empty) Then
            lblnomGrille.ForeColor = Color.Blue
            If btnJoueur1.Enabled And btnJoueur2.Enabled = False Then
                lblVerification1.Text &= lblnomGrille.Text
                If txtMotJoueur1.Text = lblVerification1.Text And lblVerification1.Text.Length >= 3 Then
                    Integer.TryParse(lblNombreMot1.Text, intNombreMot1)
                    intNombreMot1 += 1
                    lblNombreMot1.Text = intNombreMot1.ToString
                    intNombreMoyenLettre1 += txtMotJoueur1.Text.Length
                    lblNombreMoyenLettre1.Text = (intNombreMoyenLettre1 / intNombreMot1).ToString("N2")
                    ' Tout remettre en noir dans la grille pour la prochaine verification.
                    Remettre_Grille_Noir()
                    ' Vide les label et change le bouton annuler pour preuve.
                    txtMotJoueur1.Text = String.Empty
                    txtMotJoueur1.Enabled = False
                    lblVerification1.Text = String.Empty
                    btnJoueur1.Text = "P&reuve"
                    btnJoueur2.Enabled = True
                End If
            ElseIf btnJoueur2.Enabled And btnJoueur1.Enabled = False Then
                lblVerification2.Text &= lblnomGrille.Text
                If txtMotJoueur2.Text = lblVerification2.Text And lblVerification2.Text.Length >= 3 Then
                    Integer.TryParse(lblNombreMot2.Text, intNombreMot2)
                    intNombreMot2 += 1
                    lblNombreMot2.Text = intNombreMot2.ToString
                    intNombreMoyenLettre2 += txtMotJoueur2.Text.Length
                    lblNombreMoyenLettre2.Text = (intNombreMoyenLettre2 / intNombreMot2).ToString("N2")
                    ' Tout remettre en noir dans la grille pour la prochaine verification. 
                    Remettre_Grille_Noir()
                    ' Vide les label et change le bouton annuler en preuve.
                    txtMotJoueur2.Text = String.Empty
                    txtMotJoueur2.Enabled = False
                    lblVerification2.Text = String.Empty
                    btnJoueur2.Text = "Pr&euve"
                    btnJoueur1.Enabled = True
                End If
            End If
        End If
    End Sub

    ' Procédure de vérification du label et de la grille. 
    Private Sub Verification_des_mots()

    End Sub

    ' Procédure de selection des labels pour afficher des voyelle ou des consonne dans la grille de jeux en fonction d'un nombre aléatoire obetnue. 
    Private Sub Ecrire_Dans_Label_Grille(ByVal intNombreAleatoireLabel As Integer, ByVal intNombreAleatoireLettre As Integer, ByVal strTypeDeLettre As String)
        Select Case intNombreAleatoireLabel
            Case Is = 1
                lblGrille1.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 2
                lblGrille2.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 3
                lblGrille3.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 4
                lblGrille4.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 5
                lblGrille5.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 6
                lblGrille6.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 7
                lblGrille7.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 8
                lblGrille8.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 9
                lblGrille9.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 10
                lblGrille10.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 11
                lblGrille11.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 12
                lblGrille12.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 13
                lblGrille13.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 14
                lblGrille14.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 15
                lblGrille15.Text = strTypeDeLettre(intNombreAleatoireLettre)
            Case Is = 16
                lblGrille16.Text = strTypeDeLettre(intNombreAleatoireLettre)
        End Select
    End Sub

    ' Remet tout le contenue de la grille en noir.
    Private Sub Remettre_Grille_Noir()
        lblGrille1.ForeColor = Color.Black
        lblGrille2.ForeColor = Color.Black
        lblGrille3.ForeColor = Color.Black
        lblGrille4.ForeColor = Color.Black
        lblGrille5.ForeColor = Color.Black
        lblGrille6.ForeColor = Color.Black
        lblGrille7.ForeColor = Color.Black
        lblGrille8.ForeColor = Color.Black
        lblGrille9.ForeColor = Color.Black
        lblGrille10.ForeColor = Color.Black
        lblGrille11.ForeColor = Color.Black
        lblGrille12.ForeColor = Color.Black
        lblGrille13.ForeColor = Color.Black
        lblGrille14.ForeColor = Color.Black
        lblGrille15.ForeColor = Color.Black
        lblGrille16.ForeColor = Color.Black
    End Sub

    ' Vide tout le contenue des label et champs text.
    Private Sub Remettre_vide()
        lblGrille1.Text = String.Empty
        lblGrille2.Text = String.Empty
        lblGrille3.Text = String.Empty
        lblGrille4.Text = String.Empty
        lblGrille5.Text = String.Empty
        lblGrille6.Text = String.Empty
        lblGrille7.Text = String.Empty
        lblGrille8.Text = String.Empty
        lblGrille9.Text = String.Empty
        lblGrille10.Text = String.Empty
        lblGrille11.Text = String.Empty
        lblGrille12.Text = String.Empty
        lblGrille13.Text = String.Empty
        lblGrille14.Text = String.Empty
        lblGrille15.Text = String.Empty
        lblGrille16.Text = String.Empty

        txtMotJoueur1.Text = String.Empty
        txtMotJoueur2.Text = String.Empty
        lblVerification1.Text = String.Empty
        lblVerification2.Text = String.Empty
        lblNombreMot1.Text = String.Empty
        lblNombreMot2.Text = String.Empty
        lblNombreMoyenLettre1.Text = String.Empty
        lblNombreMoyenLettre2.Text = String.Empty
    End Sub

    ' Procédure a la fermeture de la form.
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Variable du contenue en string de toute les grille pour valider si la partie est debuter ou non. 
        ' Variable de retour du dialog si la partie est commencer et qu'on veut quand meme quitter.
        Dim strContenueGrille As String
        Dim dlgButton As DialogResult

        strContenueGrille = lblGrille1.Text & lblGrille2.Text & lblGrille3.Text & lblGrille4.Text &
            lblGrille5.Text & lblGrille6.Text & lblGrille7.Text & lblGrille8.Text & lblGrille9.Text &
            lblGrille10.Text & lblGrille11.Text & lblGrille12.Text & lblGrille13.Text & lblGrille14.Text &
            lblGrille15.Text & lblGrille16.Text

        If strContenueGrille <> String.Empty Then
            dlgButton = MessageBox.Show("Voulez-vous vraiment quittez la partie en cours?", "Quitter la partie", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If dlgButton = DialogResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub

    ' Bouton Quitter
    Private Sub menQuitter_Click(sender As Object, e As EventArgs) Handles menQuitter.Click
        Me.Close()
    End Sub

    ' Bouton Nouvelle Partie
    Private Sub NouvellePartieToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NouvellePartieToolStripMenuItem.Click
        ' initialisation de tout les case dans notre grille de lettres. Ces cases sont chacune exprimer a l'aide d'un label allant
        ' de lblGrille1 a lblGrille 16.

        Const strVoyelle As String = "AEIOU"
        Const strConsonne As String = "BCDFGHJKLMNPQRSTVWXYZ"

        Dim randomGridLabel As New Random
        Dim randomLetter As New Random
        Dim intRandomGridLabel As Integer
        Dim intRandomLetter As Integer
        Dim lstLabelGrille As New List(Of Integer)

        ' Vider tout les case
        Remettre_vide()

        ' Vider les valeur qui serve a calculer le total des mots et le nombre moyen de lettres.
        intNombreMot1 = 0
        intNombreMot2 = 0
        intNombreMoyenLettre1 = 0
        intNombreMoyenLettre2 = 0

        ' Tout remettre en noir
        Remettre_Grille_Noir()

        ' Remettre les boutons non accessible pour la duré du timer. 
        Propriete_initial()

        ' Création d'une list avec tout les nombre possible pour les labels de notre grille de jeux.
        For GridLabelNumber As Integer = 1 To 16
            lstLabelGrille.Add(GridLabelNumber)
        Next GridLabelNumber

        ' Assignation des voyelle et des consonne ensuite avec une structure de boucle imbriqué. 
        While lstLabelGrille.Count > 0
            While lstLabelGrille.Count > 10
                intRandomGridLabel = randomGridLabel.Next(1, 17)
                If lstLabelGrille.Contains(intRandomGridLabel) Then
                    lstLabelGrille.Remove(intRandomGridLabel)
                    intRandomLetter = randomLetter.Next(0, 5)
                    Ecrire_Dans_Label_Grille(intRandomGridLabel, intRandomLetter, strVoyelle)
                End If
            End While
            intRandomGridLabel = randomGridLabel.Next(1, 17)
            If lstLabelGrille.Contains(intRandomGridLabel) Then
                lstLabelGrille.Remove(intRandomGridLabel)
                intRandomLetter = randomLetter.Next(0, 20)
                Ecrire_Dans_Label_Grille(intRandomGridLabel, intRandomLetter, strConsonne)
            End If
        End While

        ' Initialiser le timer a 3 minutes. 
        intConteurSeconde = 180
        Timer.Start()
    End Sub

    ' empeche la selection si la lettre n'est pas collé sur la case 1. 
    Private Sub lblGrille1_Click(sender As Object, e As EventArgs) Handles lblGrille1.Click
        If Premier_click() Then
            Verification(lblGrille1)
        ElseIf (lblGrille2.ForeColor = Color.Blue Or lblGrille5.ForeColor = Color.Blue Or lblGrille6.ForeColor = Color.Blue) And
            lblGrille1.ForeColor = Color.Black Then
            Verification(lblGrille1)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 2. 
    Private Sub lblGrille2_Click(sender As Object, e As EventArgs) Handles lblGrille2.Click
        If Premier_click() Then
            Verification(lblGrille2)
        ElseIf (lblGrille1.ForeColor = Color.Blue Or lblGrille3.ForeColor = Color.Blue Or lblGrille5.ForeColor = Color.Blue Or
            lblGrille6.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue) And lblGrille2.ForeColor = Color.Black Then
            Verification(lblGrille2)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 3.
    Private Sub lblGrille3_Click(sender As Object, e As EventArgs) Handles lblGrille3.Click
        If Premier_click() Then
            Verification(lblGrille3)
        ElseIf (lblGrille2.ForeColor = Color.Blue Or lblGrille4.ForeColor = Color.Blue Or lblGrille6.ForeColor = Color.Blue Or
            lblGrille7.ForeColor = Color.Blue Or lblGrille8.ForeColor = Color.Blue) And lblGrille3.ForeColor = Color.Black Then
            Verification(lblGrille3)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 4. 
    Private Sub lblGrille4_Click(sender As Object, e As EventArgs) Handles lblGrille4.Click
        If Premier_click() Then
            Verification(lblGrille4)
        ElseIf (lblGrille3.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue Or lblGrille8.ForeColor = Color.Blue) And
            lblGrille4.ForeColor = Color.Black Then
            Verification(lblGrille4)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 5. 
    Private Sub lblGrille5_Click(sender As Object, e As EventArgs) Handles lblGrille5.Click
        If Premier_click() Then
            Verification(lblGrille5)
        ElseIf (lblGrille1.ForeColor = Color.Blue Or lblGrille2.ForeColor = Color.Blue Or lblGrille6.ForeColor = Color.Blue Or
            lblGrille9.ForeColor = Color.Blue Or lblGrille10.ForeColor = Color.Blue) And lblGrille5.ForeColor = Color.Black Then
            Verification(lblGrille5)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 6. 
    Private Sub lblGrille6_Click(sender As Object, e As EventArgs) Handles lblGrille6.Click
        If Premier_click() Then
            Verification(lblGrille6)
        ElseIf (lblGrille1.ForeColor = Color.Blue Or lblGrille2.ForeColor = Color.Blue Or lblGrille3.ForeColor = Color.Blue Or
            lblGrille5.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue Or lblGrille9.ForeColor = Color.Blue Or
            lblGrille10.ForeColor = Color.Blue Or lblGrille11.ForeColor = Color.Blue) And lblGrille6.ForeColor = Color.Black Then
            Verification(lblGrille6)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 7. 
    Private Sub lblGrille7_Click(sender As Object, e As EventArgs) Handles lblGrille7.Click
        If Premier_click() Then
            Verification(lblGrille7)
        ElseIf (lblGrille2.ForeColor = Color.Blue Or lblGrille3.ForeColor = Color.Blue Or lblGrille4.ForeColor = Color.Blue Or
            lblGrille6.ForeColor = Color.Blue Or lblGrille8.ForeColor = Color.Blue Or lblGrille10.ForeColor = Color.Blue Or
            lblGrille11.ForeColor = Color.Blue Or lblGrille12.ForeColor = Color.Blue) And lblGrille7.ForeColor = Color.Black Then
            Verification(lblGrille7)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 8. 
    Private Sub lblGrille8_Click(sender As Object, e As EventArgs) Handles lblGrille8.Click
        If Premier_click() Then
            Verification(lblGrille8)
        ElseIf (lblGrille3.ForeColor = Color.Blue Or lblGrille4.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue Or
            lblGrille11.ForeColor = Color.Blue Or lblGrille12.ForeColor = Color.Blue) And lblGrille8.ForeColor = Color.Black Then
            Verification(lblGrille8)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 9. 
    Private Sub lblGrille9_Click(sender As Object, e As EventArgs) Handles lblGrille9.Click
        If Premier_click() Then
            Verification(lblGrille9)
        ElseIf (lblGrille5.ForeColor = Color.Blue Or lblGrille6.ForeColor = Color.Blue Or lblGrille10.ForeColor = Color.Blue Or
            lblGrille13.ForeColor = Color.Blue Or lblGrille14.ForeColor = Color.Blue) And lblGrille9.ForeColor = Color.Black Then
            Verification(lblGrille9)
        End If

    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 10. 
    Private Sub lblGrille10_Click(sender As Object, e As EventArgs) Handles lblGrille10.Click
        If Premier_click() Then
            Verification(lblGrille10)
        ElseIf (lblGrille5.ForeColor = Color.Blue Or lblGrille6.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue Or
            lblGrille9.ForeColor = Color.Blue Or lblGrille11.ForeColor = Color.Blue Or lblGrille13.ForeColor = Color.Blue Or
            lblGrille14.ForeColor = Color.Blue Or lblGrille15.ForeColor = Color.Blue) And lblGrille10.ForeColor = Color.Black Then
            Verification(lblGrille10)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 11. 
    Private Sub lblGrille11_Click(sender As Object, e As EventArgs) Handles lblGrille11.Click
        If Premier_click() Then
            Verification(lblGrille11)
        ElseIf (lblGrille6.ForeColor = Color.Blue Or lblGrille7.ForeColor = Color.Blue Or lblGrille8.ForeColor = Color.Blue Or
            lblGrille10.ForeColor = Color.Blue Or lblGrille12.ForeColor = Color.Blue Or lblGrille14.ForeColor = Color.Blue Or
            lblGrille15.ForeColor = Color.Blue Or lblGrille16.ForeColor = Color.Blue) And lblGrille11.ForeColor = Color.Black Then
            Verification(lblGrille11)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 12. 
    Private Sub lblGrille12_Click(sender As Object, e As EventArgs) Handles lblGrille12.Click
        If Premier_click() Then
            Verification(lblGrille12)
        ElseIf (lblGrille7.ForeColor = Color.Blue Or lblGrille8.ForeColor = Color.Blue Or lblGrille11.ForeColor = Color.Blue Or
            lblGrille15.ForeColor = Color.Blue Or lblGrille16.ForeColor = Color.Blue) And lblGrille12.ForeColor = Color.Black Then
            Verification(lblGrille12)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 13. 
    Private Sub lblGrille13_Click(sender As Object, e As EventArgs) Handles lblGrille13.Click
        If Premier_click() Then
            Verification(lblGrille13)
        ElseIf (lblGrille9.ForeColor = Color.Blue Or lblGrille10.ForeColor = Color.Blue Or lblGrille14.ForeColor = Color.Blue) And
            lblGrille13.ForeColor = Color.Black Then
            Verification(lblGrille13)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 14. 
    Private Sub lblGrille14_Click(sender As Object, e As EventArgs) Handles lblGrille14.Click
        If Premier_click() Then
            Verification(lblGrille14)
        ElseIf (lblGrille9.ForeColor = Color.Blue Or lblGrille10.ForeColor = Color.Blue Or lblGrille11.ForeColor = Color.Blue Or
            lblGrille13.ForeColor = Color.Blue Or lblGrille15.ForeColor = Color.Blue) And lblGrille14.ForeColor = Color.Black Then
            Verification(lblGrille14)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 15. 
    Private Sub lblGrille15_Click(sender As Object, e As EventArgs) Handles lblGrille15.Click
        If Premier_click() Then
            Verification(lblGrille15)
        ElseIf (lblGrille10.ForeColor = Color.Blue Or lblGrille11.ForeColor = Color.Blue Or lblGrille12.ForeColor = Color.Blue Or
            lblGrille14.ForeColor = Color.Blue Or lblGrille16.ForeColor = Color.Blue) And lblGrille15.ForeColor = Color.Black Then
            Verification(lblGrille15)
        End If
    End Sub

    ' Empeche la selection si la lettre n'est pas collé sur la case 16. 
    Private Sub lblGrille16_Click(sender As Object, e As EventArgs) Handles lblGrille16.Click
        If Premier_click() Then
            Verification(lblGrille16)
        ElseIf (lblGrille11.ForeColor = Color.Blue Or lblGrille12.ForeColor = Color.Blue Or lblGrille15.ForeColor = Color.Blue) And
            lblGrille16.ForeColor = Color.Black Then
            Verification(lblGrille16)
        End If
    End Sub

    ' Messagebox qui apparait lors du click sur le bouton A Propos ?. 
    Private Sub ÀProposToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ÀProposToolStripMenuItem.Click
        Dim dialogReturn As DialogResult

        dialogReturn = MessageBox.Show("Ce programme a été concu par: Yannick Munger" & ControlChars.NewLine & "Numéro d'étudiant: 536 855 494", "À Propos")
    End Sub

    ' Chronometre de la partie
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Dim intSeconde As Integer
        Dim intMinute As Integer

        intConteurSeconde -= 1
        If intConteurSeconde >= 0 Then
            intSeconde = intConteurSeconde Mod 60
            intMinute = intConteurSeconde \ 60
            lblTimer.Text = (intMinute.ToString & ":" & intSeconde.ToString.PadLeft(2, "0"c))
            If intConteurSeconde = 0 Then
                btnJoueur1.Enabled = True
                btnJoueur2.Enabled = True
                btnTerminer.Enabled = True
                TerminerToolStripMenuItem.Enabled = True
                Joueur1ToolStripMenuItem.Enabled = True
                Joueur2ToolStripMenuItem.Enabled = True
                PreuveToolStripMenuItem.Enabled = True
                PreuveToolStripMenuItem1.Enabled = True
            End If
        End If
    End Sub

    ' Activation/desactivation des boutons preuve/annule et des txtbox joueur 1.
    Private Sub btnJoueur1_Click(sender As Object, e As EventArgs) Handles btnJoueur1.Click, PreuveToolStripMenuItem.Click
        If btnJoueur1.Text = "P&reuve" And btnJoueur1.Enabled = True Then
            txtMotJoueur1.Enabled = True
            txtMotJoueur1.Focus()
            btnJoueur2.Enabled = False
            btnJoueur1.Text = "&Annuler"
        ElseIf btnJoueur1.Text = "&Annuler" And btnJoueur1.Enabled Then
            txtMotJoueur1.Enabled = False
            txtMotJoueur1.Text = String.Empty
            Remettre_Grille_Noir()
            lblVerification1.Text = String.Empty
            btnJoueur2.Enabled = True
            btnJoueur1.Text = "P&reuve"
        End If
    End Sub

    ' Activation/desactivation des boutons preuve/annule et des txtbox joueur 2. 
    Private Sub btnJoueur2_Click(sender As Object, e As EventArgs) Handles btnJoueur2.Click, PreuveToolStripMenuItem1.Click
        If btnJoueur2.Text = "Pr&euve" And btnJoueur2.Enabled = True Then
            txtMotJoueur2.Enabled = True
            txtMotJoueur2.Focus()
            btnJoueur1.Enabled = False
            btnJoueur2.Text = "&Annuler"
        ElseIf btnJoueur2.Text = "&Annuler" And btnJoueur2.Enabled = True Then
            txtMotJoueur2.Enabled = False
            txtMotJoueur2.Text = String.Empty
            Remettre_Grille_Noir()
            lblVerification2.Text = String.Empty
            btnJoueur1.Enabled = True
            btnJoueur2.Text = "Pr&euve"
        End If
    End Sub

    ' Message du gagnant de la partie et retourne le jeux a sa page vide initial. 
    Private Sub Terminer_Partie(sender As Object, e As EventArgs) Handles btnTerminer.Click, TerminerToolStripMenuItem.Click
        Dim dlgButton As DialogResult
        Dim strMessage As String
        Dim strMessageGagnantJoueur1 As String
        Dim strMessageGagnantJoueur2 As String
        Dim strBrisEgaliteJoueur1Gagnant As String
        Dim strBrisEgaliteJoueur2Gagnant As String
        Dim strMessageEgalite As String
        Dim strMessageEgaliteNul As String

        ' 2 Message gagnant determiner par le nombre de mot. 
        strMessageGagnantJoueur1 = ("Félicitation Joueur 1 vous avez gagné avec un total de " & lblNombreMot1.Text &
            " mots qui contenait en moyenne " & lblNombreMoyenLettre1.Text & " lettres!")
        strMessageGagnantJoueur2 = ("Félicitation Joueur 2 vous avez gagné avec un total de " & lblNombreMot2.Text &
            " mots qui contenait en moyenne " & lblNombreMoyenLettre2.Text & " lettres!")

        ' 3 Message situation de double egalite, joueur 1 ou joueur 2 gagne avec plus de lettre par mot 
        strBrisEgaliteJoueur1Gagnant = ("Félicitation joueur 1 vous avez gagné au bris d'égalité grâce à votre moyenne de " & lblNombreMoyenLettre1.Text & " lettres par mot")
        strBrisEgaliteJoueur2Gagnant = ("Félicitation joueur 2 vous avez gagné au bris d'égalité grâce à votre moyenne de " & lblNombreMoyenLettre2.Text & " lettres par mot")
        strMessageEgalite = ("La partie termine a égalité entre le joueur 1 et le joueur 2 grace a un nombre de mot de " & lblNombreMot1.Text &
            " et un nombre moyen de " & lblNombreMoyenLettre1.Text & " lettres par mots")
        strMessageEgaliteNul = ("La partie termine a égalité entre le joueur 1 et le joueur 2 car vous avez tout les deux trouver aucun mots. Meilleur chance la prochaine fois!")

        ' 6 cas seront traité joueur 1 gagnant, joueur 2 gagnant et l'égalité nul et non nul. 
        If intNombreMot1 > intNombreMot2 Then
            strMessage = strMessageGagnantJoueur1
        ElseIf intNombreMot1 < intNombreMot2 Then
            strMessage = strMessageGagnantJoueur2
        Else
            If intNombreMoyenLettre1 > intNombreMoyenLettre2 Then
                strMessage = strBrisEgaliteJoueur1Gagnant
            ElseIf intNombreMoyenLettre1 < intNombreMoyenLettre2 Then
                strMessage = strBrisEgaliteJoueur2Gagnant
            ElseIf lblNombreMot1.Text = String.Empty And lblNombreMot2.Text = String.Empty Then
                strMessage = strMessageEgaliteNul
            Else
                strMessage = strMessageEgalite
            End If
        End If

        ' Message box qui affiche le message du gagnant a la fin de la partie. Il remet aussi la partie en mode initial. 
        dlgButton = MessageBox.Show(strMessage, "Fin de la partie", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If dlgButton = DialogResult.OK Then
            Remettre_Grille_Noir()
            Remettre_vide()
            Propriete_initial()
        End If
    End Sub
End Class
