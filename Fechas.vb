Option Explicit
'Application.Text(DateValue(Fecha), "[$-es-MX]MMMDD;@")
Public Function DiaLab(FechaInicial As Variant, Dias As Integer, Festivos As Range)
    
    Do While Dias <> 0
        If Dias > 0 Then
            Do
                FechaInicial = FechaInicial + 1
            Loop While Not Festivos.Find(FechaInicial) Is Nothing Or Weekday(FechaInicial) = 1 Or Weekday(FechaInicial) = 7
            Dias = Dias - 1
        Else
            Do
                FechaInicial = FechaInicial - 1
            Loop While Not Festivos.Find(FechaInicial) Is Nothing Or Weekday(FechaInicial) = 1 Or Weekday(FechaInicial) = 7
            Dias = Dias + 1
        End If
    Loop
    
    DiaLab = FechaInicial
    
End Function

Public Function FormatoFecha(ByVal Rango As String) As Date

    FormatoFecha = DateSerial(Right(Rango, 4), Mid(Rango, 4, 2), Left(Rango, 2))

End Function