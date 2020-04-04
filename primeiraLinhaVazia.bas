'***************************************************************************
'Purpose: to find out the first empty row, based on a last used column
'Inputs:
'   ultimaColuna: is a parameter that defines the last column must be viewed to
'       decide if a row is empty or not
'Outputs: an Integer that indicates the first empty line
'***************************************************************************

Function primeiraLinhaVazia(ByVal ultimaColuna As Integer) As Integer
Dim ws As Worksheet, col As Integer, lin As Integer, linhaVazia As Boolean
Set ws = ActiveSheet
col = 1
lin = 1
linhaVazia = True

Do While True
    
    Do While col <= ultimaColuna
                
        If IsEmpty(ws.Cells(lin, col)) = False Then
            linhaVazia = False
            Exit Do
        End If
        col = col + 1
    Loop
    
    'Apos todas as colunas, a linha é tida como vazia
    If linhaVazia = True Then
        primeiraLinhaVazia = lin
        Exit Function
    End If
    
    lin = lin + 1
    col = 1
    linhaVazia = True
Loop

End Function

