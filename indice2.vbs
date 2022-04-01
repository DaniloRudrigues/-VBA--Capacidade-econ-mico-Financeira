Sub municípios3()
'
' municípios3 Macro
'
' Atalho do teclado: Ctrl+e
'
    
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    
    i = 1
    
    While i <= 2
    
    
        k = i + 1
        j = i + 2
    
        Windows("3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm").Activate
        Cells(5, 3).Value = Worksheets("Listas").Cells(k, 14).Value
        
        
        Windows("Municípios.xlsx").Activate
        
        Range("A1").Select
        
        ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R5C3"
        Cells(j, 2).Select
                
        ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R4C7"
        Cells(j, 3).Select
                
        ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R5C7"
        Cells(j, 4).Select
               
       ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R6C7"
        Cells(j, 5).Select
        
        ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R7C7"
        Cells(j, 6).Select
        
        ActiveCell.FormulaR1C1 = _
            "='[3.3.1_EVTE v9f - GLOBAL - COMPESA - ARPE.xlsm]CONFIGURAÇÕES'!R9C7"
        Cells(j, 7).Select
        
        Range("B3:G171").Select
        Selection.Copy
        ActiveWindow.SmallScroll Down:=-174
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
        
        i = i + 1
    
    Wend
End Sub
