Sub Bouton_Nettoyage()

    Dim reponse As VbMsgBoxResult

    reponse = MsgBox("Voulez-vous lancer le nettoyage des données ?", vbYesNo + vbQuestion, "Confirmation")

    If reponse = vbNo Then Exit Sub

    Call procedure_complete

    MsgBox "Nettoyage terminé avec succès !", vbInformation

End Sub
