Sub SetLanguageDanish()
    Dim scount, j, k, fcount
    scount = ActivePresentation.Slides.Count
    For j = 1 To scount
        fcount = ActivePresentation.Slides(j).Shapes.Count
        For k = 1 To fcount 'change all shapes:
            If ActivePresentation.Slides(j).Shapes(k).HasTextFrame Then
                ActivePresentation.Slides(j).Shapes(k).TextFrame.TextRange.LanguageID = msoLanguageIDDanish
            End If
        Next k
        fcount = ActivePresentation.Slides(j).NotesPage.Shapes.Count
        For k = 1 To fcount 'change all shapes:
            If ActivePresentation.Slides(j).NotesPage.Shapes(k).HasTextFrame Then
                ActivePresentation.Slides(j).NotesPage.Shapes(k).TextFrame.TextRange.LanguageID = msoLanguageIDDanish
            End If
        Next k
    Next j
End Sub
