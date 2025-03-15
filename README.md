# MACRO-PROGRAM-TO-CALCULATE-VARIATION-AND-DATE-STANDARD-
Sub varandstd()Dim dataRange As RangeDim variation As DoubleDim stdDev As Double Set dataRange = ActiveSheet.Range("A2:A100") variation = WorksheetFunction.Var(dataRange)stdDev = WorksheetFunction.StDev(dataRange) MsgBox "Variation: " & variation & vbCrLf & "Standard Deviation: " & stdDevEnd Sub
