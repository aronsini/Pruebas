''----------------------------------------------------------------------------------------
'Librerias:
'Verifica si determinado archivo existe y regresa un boleano
''----------------------------------------------------------------------------------------

'Function Verificar_Archivo(ByVal Ruta_Archivo As String) As Boolean
'    On Error Resume Next
'    Ruta_Archivo = Dir(Ruta_Archivo)
'    On Error GoTo 0
'    If Ruta_Archivo <> "" Then
'        Verificar_Archivo = True
'    Else
'        Verificar_Archivo = False
'    End If
'
'End Function


'Environ("USERPROFILE") localuser folder

Private Sub ForEachFile(Ruta As String, Optional Extension As String)

    Dim Archivo As Variant
    
    If Right(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    Archivo = Dir(Ruta & "*" & Extension)
    While (Archivo <> "")
        
        'Procedimiento a cada uno de los archivos en ruta
        
        Debug.Print Archivo
        Archivo = Dir
        Wend
    End Sub