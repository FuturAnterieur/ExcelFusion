Attribute VB_Name = "RandomMacros"
Option Explicit

Sub testdgs()
Attribute testdgs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' testdgs Macro
'

'
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=SUBSTITUTE(R[1]C[-10], ""yotta"", ""peta"")"
    Range("M3").Select
End Sub

Sub resetrange()
    ActiveSheet.UsedRange
End Sub
