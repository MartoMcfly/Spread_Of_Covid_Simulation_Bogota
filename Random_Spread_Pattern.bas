Attribute VB_Name = "Module1"
Public lleno As Boolean
Public ableInfect As Boolean
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public matriz(1 To 54, 1 To 88) As Integer
Public periodo As Integer
'Option Explicit
Sub reiniciar()
    
    ableInfect = False
    Dim infectados As Boolean
    Dim rango As Range
    periodo = 1
    lleno = True
    Application.ScreenUpdating = False
    ActiveSheet.ChartObjects("Chart 1").Visible = False
    For i = 1 To 54
        For j = 1 To 88
            If Sheet6.Cells(i, j).Value <> "" And Sheet6.Cells(i, j).Value <> 0 Then
                Sheet4.Cells(i, j).Interior.Color = vbWhite
                Sheet4.Cells(i, j).Font.Color = vbWhite
                matriz(i, j) = 1
            End If
        Next j
    Next i
    If Not (Sheet5.Cells(2, 1) = "") Then
        Set rango = Sheet5.Range(Sheet5.Cells(2, 1), Sheet5.Cells(2, 5).End(xlDown))
        rango.Value = ""
    End If
    Application.ScreenUpdating = True
    'ActiveSheet.ChartObjects("Chart 1").Chart.Refresh
    'ActiveSheet.ChartObjects("Chart 1").Visible = True
End Sub
Sub infectar()
    ableInfect = True
    MsgBox ("Infecte las zonas que desee y simule")
End Sub
Sub simulPrueba()
    Dim random As Double
    Dim suma As Integer
    Dim fila As Integer
    lleno = False
    ableInfect = False
    Dim susep As Integer
    Dim sumaR As Integer
    Dim ciclos() As Integer
    ciclos = matriz
    
    Dim tdr As Integer
    
'Application.ScreenUpdating = False 'This line disable the on screen update for better performance, the blink you see, you could delete both lanes but it will run slower
'Dim myChart As ChartObject
'For Each myChart In ActiveSheet.ChartObjects
'    myChart.Chart.Refresh
'Next myChart
'Application.ScreenUpdating = True 'This line reenable the on screen update for better performance, the blink you see, you could delete both lanes but it will run slower

    
    Dim nueva() As Integer
    nueva = matriz
    hayinfectados = True
    While Not lleno And hayinfectados
        DoEvents
        Application.ScreenUpdating = False
        hayinfectados = False
        tdr = Sheet4.Cells(40, 112)
        For i = 1 To 54
            For j = 1 To 88
                If matriz(i, j) = 2 Then
                    hayinfectados = True
                    ciclos(i, j) = ciclos(i, j) + 1
                                      
                    If ciclos(i, j) > tdr Then
                        random = Rnd()
                        If random > Sheet7.Cells(i, j).Value * Sheet4.Cells(40, 166) Then
                            nueva(i, j) = 4
                        Else
                            nueva(i, j) = 3
                        End If
                    End If
                    
                    If lleno Then
                        GoTo salir
                    End If
                    random = Rnd()
                    If random < Sheet6.Cells(i - 1, j).Value * Sheet4.Cells(40, 139) And (nueva(i - 1, j) <> 4) And (nueva(i - 1, j) <> 3) Then
                        nueva(i - 1, j) = 2
                    End If
                    random = Rnd()
                    If random < Sheet6.Cells(i + 1, j).Value * Sheet4.Cells(40, 139) And (nueva(i + 1, j) <> 4) And (nueva(i + 1, j) <> 3) Then
                        nueva(i + 1, j) = 2
                    End If
                    random = Rnd()
                    If random < Sheet6.Cells(i, j - 1).Value * Sheet4.Cells(40, 139) And (nueva(i, j - 1) <> 4) And (nueva(i, j - 1) <> 3) Then
                        nueva(i, j - 1) = 2
                    End If
                    random = Rnd()
                    If random < Sheet6.Cells(i, j + 1).Value * Sheet4.Cells(40, 139) And (nueva(i, j + 1) <> 4) And (nueva(i, j + 1) <> 3) Then
                        nueva(i, j + 1) = 2
                    End If
                End If
            Next
        Next
        If lleno Then
            GoTo salir
        End If
        
        For i = 1 To 54
            For j = 1 To 88
                matriz(i, j) = nueva(i, j)
            Next
        Next
        suma = 0
        
        susep = 1771
        
        sumaR = 0
        
        sumaM = 0
        
        If lleno Then
            GoTo salir
        End If
        For i = 1 To 54
            For j = 1 To 88
                If matriz(i, j) = 2 Then
                    Sheet4.Cells(i, j).Interior.Color = vbRed
                    Sheet4.Cells(i, j).Font.Color = vbRed
                    suma = suma + 1
                ElseIf matriz(i, j) = 4 Then
                    Sheet4.Cells(i, j).Interior.Color = vbCyan
                    Sheet4.Cells(i, j).Font.Color = vbCyan
                    sumaR = sumaR + 1
                ElseIf matriz(i, j) = 3 Then
                    Sheet4.Cells(i, j).Interior.ColorIndex = 48
                    Sheet4.Cells(i, j).Font.ColorIndex = 48
                    sumaM = sumaM + 1
                End If
            Next
        Next
        If suma = 1771 Then
            lleno = True
        End If
        If (Sheet5.Cells(2, 1) = "") Then
            fila = 2
        Else
            fila = Sheet5.Cells(1, 1).End(xlDown).Row + 1
        End If
        
        Sheet5.Cells(fila, 3) = sumaR
        Sheet5.Cells(fila, 2) = suma + sumaR + sumaM
        Sheet5.Cells(fila, 1) = periodo
        Sheet5.Cells(fila, 4) = susep - sumaR - suma
        Sheet5.Cells(fila, 5) = sumaM
        
        periodo = periodo + 1
        Application.ScreenUpdating = True
        Sleep 100
salir:
        Application.ScreenUpdating = True
        'ActiveSheet.ChartObjects("Chart 1").Chart.Refresh
        Wend
        ActiveSheet.ChartObjects("Chart 1").Visible = True
        If Not hayinfectados Then
            ableInfect = False
        End If
End Sub
Sub detener()
    lleno = True
End Sub

