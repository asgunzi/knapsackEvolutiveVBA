Attribute VB_Name = "M_Evolutive"
Option Explicit
Dim nseed As Long

Sub KnapEvolutive()
'Solves knapsack problem with an evolutive metaheuristic procedure

Dim evalue As Variant, eweight As Variant, maxW As Double
Dim outSolution As Variant
Dim nl As Long
Dim maxtrials As Long

Math.Randomize

Application.ScreenUpdating = False
'ReadData
Plan1.Activate
nl = Application.WorksheetFunction.CountA(Range("c8:c50000"))

If nl = 0 Then Exit Sub

evalue = Range("c8:c" & 7 + nl)
eweight = Range("d8:d" & 7 + nl)
maxW = Range("c5")
maxtrials = Range("l1")


'Run
Call knapRun(evalue, eweight, maxW, outSolution, maxtrials)

'record data

Range("h8:i5000").ClearContents
Range("h8").Resize(UBound(outSolution, 1), 2) = outSolution

End Sub

Sub knapRun(evalue As Variant, eweight As Variant, maxW As Double, outSolution As Variant, maxtrials As Long)
Dim i As Long, p As Long


Dim swapsolution As Variant
Dim maxValue As Double
Dim currValue As Double
'Configurations for evolutive method

ReDim outSolution(1 To UBound(eweight, 1), 1 To 2)

frmProcessing.Show vbModeless

For i = 1 To maxtrials
    
    ReDim swapsolution(1 To UBound(eweight, 1), 1 To 1)
    Call knaptrial(evalue, eweight, maxW, swapsolution)
    currValue = evalknap(swapsolution, evalue)
    
    If currValue > maxValue Then
        maxValue = currValue
        For p = 1 To UBound(eweight, 1)
            outSolution(p, 1) = p
            outSolution(p, 2) = swapsolution(p, 1)
        Next p
        
    End If
    
    If i Mod 1000 = 0 Then
        frmProcessing.lblProcessing = i / maxtrials * 100 & "% Concluído"
        DoEvents
    End If
Next i


frmProcessing.Hide


End Sub

Private Function evalknap(swapsolution As Variant, evalue As Variant)

Dim swap As Double
Dim i As Long

For i = 1 To UBound(swapsolution, 1)
    swap = swap + swapsolution(i, 1) * evalue(i, 1)
Next i

evalknap = swap

End Function


Private Sub knaptrial(evalue As Variant, eweight As Variant, maxW As Double, swapsolution As Variant)
'Finds one solution
Dim isEnd As Boolean
Dim currweight As Double
Dim currMax As Double, i_prob As Variant
Dim total As Double
Dim linmax As Long, idx As Long


isEnd = False
currweight = 0
While isEnd <> True
    
    currMax = maxW - currweight 'available weight
    
    'Creates the probability array
    evalProb evalue, eweight, i_prob, total, linmax, currMax, swapsolution, swapsolution
    
    
    If total > 0 Then 'there's still some item to be picked
        'Makes a single draw
        idx = draw(i_prob, total, linmax)
        
        swapsolution(idx, 1) = 1
        currweight = currweight + eweight(idx, 1)
    Else
        isEnd = True
    End If
Wend


End Sub


Private Sub evalProb(evalue As Variant, eweight As Variant, i_prob As Variant, total As Double, linmax As Long, currMax As Double, swapsolution As Variant, currSolution As Variant)
'Evaluates probabilities

Dim count As Long
Dim i As Long

total = 0

ReDim i_prob(1 To UBound(evalue, 1), 1 To 2)
count = 0
For i = 1 To UBound(evalue, 1)
    If swapsolution(i, 1) = 0 And eweight(i, 1) < currMax Then  'It is not part of the solution and there's space for it
        count = count + 1
        i_prob(count, 1) = i
        total = total + (evalue(i, 1) / (eweight(i, 1) + 0.001)) ^ 3 * (1 + swapsolution(i, 1)) ^ 0.5 'The most valuable with less weight has more chances, current solution also has a weight
        i_prob(count, 2) = total
        
    End If
        
Next i

linmax = count


End Sub


Function draw(ByVal i_prob As Variant, ByVal total As Double, ByVal linmax As Long) As Long
'Receives array with probabilities
'Returns index

Dim nsort As Double
Dim i As Long, idxout As Long
Dim aux As Double

nseed = nseed + 1
aux = Math.Rnd(-nseed)
nsort = Math.Rnd(-nseed) * total
'nsort = Math.Rnd() * total
'nsort = Rnd()

For i = 1 To linmax
    If nsort <= i_prob(i, 2) Then
        idxout = i
        Exit For
    End If
Next i

If idxout > 0 Then
    draw = i_prob(idxout, 1)
Else
    draw = 0
End If

End Function



