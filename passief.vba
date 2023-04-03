Sub PassiefTaalgebruik()

Dim range As range
Dim i As Long
Dim TargetList

TargetList = Array("word", "worden", "wordt", "werd", "werden", "zijn", "word", "Word") ' put list of terms to find here

For i = 0 To UBound(TargetList)

Set range = ActiveDocument.range

With range.Find
.Text = TargetList(i)
.Format = True
.MatchCase = True
.MatchWholeWord = False
.MatchWildcards = False
.MatchSoundsLike = False
.MatchAllWordForms = False

Do While .Execute(Forward:=True) = True
range.HighlightColorIndex = wdTurquoise

Loop

End With
Next

End Sub

