Attribute VB_Name = "modSort"
Option Compare Database
Option Explicit

Public Sub QuickSort(ByRef ArrToSort As Variant, ByVal FirstIndex As Long, ByVal LastIndex As Long)

    Dim Low As Long, High As Long
    Dim Pivot As Variant, Temp As Variant

    Low = FirstIndex
    High = LastIndex
    Pivot = ArrToSort((FirstIndex + LastIndex) \ 2)

    Do While Low <= High
        Do While ArrToSort(Low) < Pivot
            Low = Low + 1
        Loop
        Do While ArrToSort(High) > Pivot
            High = High - 1
        Loop
        If Low <= High Then
            Temp = ArrToSort(Low)
            ArrToSort(Low) = ArrToSort(High)
            ArrToSort(High) = Temp
            Low = Low + 1
            High = High - 1
        End If
    Loop

    If FirstIndex < High Then QuickSort ArrToSort, FirstIndex, High
    If Low < LastIndex Then QuickSort ArrToSort, Low, LastIndex

End Sub
