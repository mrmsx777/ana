Public Class Formulaire
    ' Déclaration des variables globales :
    Dim numero As String
    Dim initiales As String
    Dim naissance As String
    Dim indice_sexe As String
    Dim sexe As String

    ' Chargement du formulaire : lecture de la dernière ligne enregistrée dans un fichier
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim lastLine As String = ""
        Using file_read As New System.IO.StreamReader("C:\Users\Admin\Desktop\log.txt")
            Do While Not file_read.EndOfStream
                lastLine = file_read.ReadLine()
            Loop
        End Using
        last_numero.Text = lastLine ' Affichage de la dernière ligne dans une zone de texte
    End Sub

    ' Enregistrement de la valeur du texte dans la variable "numero"
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        numero = TextBox1.Text
    End Sub

    ' Événement Click du bouton
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Vérification si le numéro suit le modèle désiré
        If numero.Length = 12 Then
            ' Vérification de chaque caractère pour correspondre au modèle souhaité
            If Char.IsLetter(numero(0)) AndAlso
               Char.IsLetter(numero(1)) AndAlso
               Char.IsLetter(numero(2)) AndAlso
               Char.IsLetter(numero(3)) AndAlso
               IsNumeric(numero(4)) AndAlso
               IsNumeric(numero(5)) Then

                ' Extraction du mois et du jour pour la vérification
                Dim mois As Integer = Integer.Parse(numero.Substring(6, 2))
                Dim jours As Integer = Integer.Parse(numero.Substring(8, 2))

                ' Vérification de la validité des mois et jours en fonction du modèle
                If (mois >= 1 AndAlso mois <= 12) OrElse
                   (mois >= 51 AndAlso mois <= 62 AndAlso
                   ((mois <= 12 AndAlso jours >= 1 AndAlso jours <= 31) OrElse
                   (mois >= 51 AndAlso mois <= 62 AndAlso jours >= 1 AndAlso jours <= 9) OrElse
                   (mois >= 51 AndAlso mois <= 62 AndAlso jours >= 10 AndAlso jours <= 29) OrElse
                   (mois >= 51 AndAlso mois <= 62 AndAlso jours >= 30 AndAlso jours <= 31))) AndAlso
                   IsNumeric(numero(6)) AndAlso
                   IsNumeric(numero(7)) AndAlso
                   IsNumeric(numero(8)) AndAlso
                   IsNumeric(numero(9)) AndAlso
                   IsNumeric(numero(10)) Then

                    ' Extraction des initiales et de la date de naissance
                    initiales = numero(0) & numero(3) & "    - "
                    naissance = numero(4) & numero(5) & "/" & numero(6) & numero(7) & "/" & numero(8) & numero(9) & "    - "
                    indice_sexe = numero(6)

                    ' Identification du sexe en fonction de l'indice du numéro
                    If indice_sexe = "5" Then
                        sexe = "Femme"
                        naissance = numero(4) & numero(5) & "/" & "0" & numero(7) & "/" & numero(8) & numero(9) & "    - "
                    ElseIf indice_sexe = "6" Then
                        sexe = "Femme"
                        naissance = numero(4) & numero(5) & "/" & "1" & numero(7) & "/" & numero(8) & numero(9) & "    - "
                    Else
                        sexe = "Homme"
                    End If

                    ' Affichage du résultat final avec les variables correctes
                    resultat.Text = "Initiales:   " & initiales & "  Naissance:    " & naissance & "  Sexe:    " & sexe
                    last_numero.Text = numero ' Affichage du numéro saisi précédemment
                Else
                    MessageBox.Show("Numero Invalide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            Else
                MessageBox.Show("Numero Invalide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Else
            MessageBox.Show("Numero Invalide", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
    End Sub

    ' Enregistrement du dernier numéro saisi dans le fichier de log avant la fermeture du programme
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim file_write As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter("C:\Users\Admin\Desktop\log.txt", True)
        file_write.WriteLine(numero)
        file_write.Close()
    End Sub
End Class
