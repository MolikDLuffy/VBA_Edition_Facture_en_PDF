# VBA_Edition_Facture_en_PDF
Je vous invite à aller voir la vidéo concernant ce code : 
https://www.youtube.com/watch?v=9fUf9X9Wm_Q&t=340s

    Sub edition_facture()
    Dim chemin As String
    Dim ligne As Long

    Application.ScreenUpdating = False

    'indiquer le chemin où vont être enregistrer les pdf
    chemin = Sheets("Paramètres").Range("A2").Value

    Feuil1.Select

    For ligne = Range("A200000").End(xlUp).Row To 2 Step -1

        If Cells(ligne, 9).Value <> "Editer" Then
    
            facture_numero = Sheets("Base").Cells(ligne, 5)
            facture_nom_el = Sheets("Base").Cells(ligne, 1)
      
        '------------------------------------------------------------
            'Aller récupérer dans la base de données les éléments
            'pour les intégrer dans la facture
            
            'la date
            Feuil3.Range("B5").Value = Cells(ligne, 4).Value
            'le numéro
            Feuil3.Range("B6").Value = Cells(ligne, 5).Value
            'le nom élève
            Feuil3.Range("D8").Value = Cells(ligne, 1).Value
            'l'adresse élève
            Feuil3.Range("D9").Value = Cells(ligne, 2).Value
            'le cp élève
            Feuil3.Range("D10").Value = Cells(ligne, 3).Value
            'la période de formation
            Feuil3.Range("A17").Value = Cells(ligne, 6).Value
            'le détail de la facture
            Feuil3.Range("C17").Value = Cells(ligne, 7).Value
            'le montant de la facture
            Feuil3.Range("E17").Value = Cells(ligne, 8).Value
        
    '----------------------------------------------------------------
        
        Feuil1.Cells(ligne, 9).Value = "Editer"

        Sheets("Facture").ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=chemin & facture_nom_el & "-" & facture_numero & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

    End If
    Next ligne

    Application.ScreenUpdating = True

    End Sub
