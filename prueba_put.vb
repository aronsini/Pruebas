Option Explicit
Option Base 1

'#############################################################################################
'#----------------------------------ENUMERADORES---------------------------------------------#
'#############################################################################################

Public Enum zdt_Open_Source
    zdt_From_DB = 2
    zdt_From_ZP = 3
End Enum

Public Enum psProperty
    psBenchmark = 1
    psSimulator = 3
    psConfidence1 = 5
    psConfidence2 = 6
    psConfidence3 = 7
    psModel = 9
    psTimeScalling = 11
    psCurrency = 18
End Enum

'#############################################################################################
'#----------------------------------PROPIEDADES----------------------------------------------#
'#############################################################################################

Private pZEUS As Object
Private pInputName As String
Private pNombre As String
Private pRuta_ZP As String
Private pRuta_PS As String
Private pPS_TxT() As String
Private pNumPort As Long
Private pNumIns As Long
Private pValor As Double
Private pOpenSource As zdt_Open_Source
Private pPestReporte As Worksheet
Private pHoldingAnalytics As Variant
Private pPortfolioAnalytics As Variant
Private pArrResIns As Variant
Private pArrResPort As Variant

'Esta funci�n se ingresa para evitar el error de espera de una aplicaci�n OLE  https://latecnologiaatualcance.com/arreglo-microsoft-excel-esta-esperando-otra-solicitud-para-completar-una-accion-ole/
Private Declare Function CoRegisterMessageFilter Lib "ole32" (ByVal IFilterIn As Long, ByRef PreviousFilter) As Long

Public Property Get Valor() As Double
    Valor = pValor
End Property

Public Property Get Nombre() As String
    Nombre = pNombre
End Property

Public Property Get Ruta_ZP() As String
    Ruta_ZP = pRuta_ZP
End Property

Public Property Get Ruta_PS() As String
    Ruta_PS = pRuta_PS
End Property

Public Property Get NumPort() As Long
    NumPort = pNumPort
End Property

Public Property Get NumIns() As Long
    NumIns = pNumIns
End Property

Public Property Get arrResIns() As Variant
    arrResIns = pArrResIns
End Property

Public Property Get arrResPort() As Variant
    arrResPort = pArrResPort
End Property

Public Property Get wsReporte() As Worksheet
    Set wsReporte = pPestReporte
End Property

Public Property Get OpenSource() As zdt_Open_Source
    OpenSource = pOpenSource
End Property

Public Property Let HoldingAnalytics(ByVal arrAnalytics As Variant)
    pHoldingAnalytics = arrAnalytics
End Property

Public Property Let PortfolioAnalytics(ByVal arrAnalytics As Variant)
    pPortfolioAnalytics = arrAnalytics
End Property

'#############################################################################################
'#----------------------------------COMPLEMENTARIAS------------------------------------------#
'#############################################################################################

Private Sub KillMessageFilter()

    Dim IMsgFilter As Long
    
    CoRegisterMessageFilter 0&, IMsgFilter

End Sub

Private Sub RestoreMessageFilter()

    Dim IMsgFilter As Long
    
    CoRegisterMessageFilter IMsgFilter, IMsgFilter

End Sub

'#############################################################################################
'#----------------------------------PROCEDIMIENTOS-------------------------------------------#
'#############################################################################################

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: GetPortfolioAnalytic
' Objetivo: Extrae el anal�tico de un portafolio
' Regresa: Variant
'     @ AnalyticID [String] -> Anal�tico a extraer
' ----------------------------------------------------------------
Public Function GetPortfolioAnalytic(ByVal AnalyticID As String) As Variant

    GetPortfolioAnalytic = pZEUS.GetPortfolioAnalytic(pNumPort, AnalyticID)

End Function

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: OpenPortfolio
' Objetivo: Abre un portafolio en ZEUS y extra� algunas propiedades
'     @ portfolio_name_or_path [String] -> Ruta o nombre del portafolio sobre el que se va a trabajar
' ----------------------------------------------------------------
Public Sub OpenPortfolio(ByVal portfolio_name_or_path As String)
    
    KillMessageFilter
    
    pInputName = portfolio_name_or_path
    ZDT_GetPortfolioSource
    pNumPort = pZEUS.OpenDocument(pInputName, pOpenSource)
    pValor = GetPortfolioAnalytic("Value")
    pNumIns = GetPortfolioAnalytic("NumberOfSecurities")
    
    RestoreMessageFilter
    
End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: CrearReporte
' Objetivo: Valida resultados, crea pesta�a y vacia reporte
'     @ wbDestino [Workbook] -> Libro en que se generar� reporte
' ----------------------------------------------------------------
Public Sub CrearReporte(ByVal wbDestino As Workbook)

    'verificamos si ya se extrajeron analiticos; de lo contrario y si existe la matriz de analiticos los extrae
    If Not IsArray(pArrResIns) Then
        If IsArray(pHoldingAnalytics) Then
            GetResults_Holdings
        Else
            MsgBox "No has definido los anal�ticos de instrumento a extraer"
            Exit Sub
        End If
    End If
    
    If Not IsArray(pArrResPort) Then
        If IsArray(pPortfolioAnalytics) Then
            GetResults_Portfolio
        Else
            MsgBox "No has definido los anal�ticos de portafolio a extraer"
            Exit Sub
        End If
    End If
    
    'creamos pesta�a con el nombre del portafolio
    Set pPestReporte = CrearPesta�a(wbDestino)
    
    'vaciamos matrices resultado
    With pPestReporte
        .Cells(2, 1).Resize(UBound(pArrResPort, 1), UBound(pArrResPort, 2)).Value = pArrResPort
        .Cells(5, 1).Resize(UBound(pArrResIns, 1), UBound(pArrResIns, 2)).Value = pArrResIns
    End With

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: CrearPesta�a
' Objetivo: Crea la pesta�a del portafolio, si existe la borra, crea nueva y renombra
' Regresa: Worksheet
'     @ wbLibro [Workbook] -> Libro en el que se agregar� la pesta�a
' ----------------------------------------------------------------
Private Function CrearPesta�a(ByVal wbLibro As Workbook) As Worksheet
    
    Application.DisplayAlerts = False
    
    If Worksheet_Exist(wbLibro, pNombre) Then wbLibro.Sheets(pNombre).Delete

    Set CrearPesta�a = wbLibro.Sheets.Add
    CrearPesta�a.Name = pNombre
    
    Application.DisplayAlerts = True
    
End Function

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: Worksheet_Exist
' Objetivo: Revisa si la pesta�a existe en un libro
' Regresa: Boolean
'     @ wbLibro [Workbook] -> Libro en el cual buscar
'     @ strName [String] -> Nombre de la pesta�a buscada
' ----------------------------------------------------------------
Private Function Worksheet_Exist(ByVal wbLibro As Workbook, _
                                 ByVal strName As String) As Boolean

    Dim Pest As Worksheet
    
    For Each Pest In wbLibro.Sheets
        If Pest.Name = strName Then
            Worksheet_Exist = True
            Exit Function
        End If
    Next Pest
    
    Worksheet_Exist = False
    
End Function

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: PS_UpdatePropertyValue
' Objetivo: Cambia el valor de la propiedad del archivo ps
'     @ psProperty [psProperty] -> L�nea en que se encuentra la propiedad
'     @ strNewValue [String] -> Valor que se va a colocar en la propiedad
' ----------------------------------------------------------------
Public Sub PS_UpdatePropertyValue(ByVal psProperty As psProperty, _
                                    ByVal strNewValue As String)
    Dim strCurrentValue As String

    strCurrentValue = PS_GetPropertyValue(psProperty)
    pPS_TxT(psProperty) = Replace(pPS_TxT(psProperty), strCurrentValue, strNewValue)
    PS_File ForWriting

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: PS_GetPropertyValue
' Objetivo: Extrae los valores almacenados en el archivo ps basado en la l�nea ingresada como propiedad
' Regresa: String
'     @ psProperty [psProperty] -> Propiedad de la que se quiere extraer el valor
' ----------------------------------------------------------------
Public Function PS_GetPropertyValue(ByVal psProperty As psProperty) As String

    Dim txtPropertyLine As String

    txtPropertyLine = pPS_TxT(psProperty)
    PS_GetPropertyValue = Mid(txtPropertyLine, Len(txtPropertyLine) - InStr(1, StrReverse(txtPropertyLine), vbTab) + 2)

End Function

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: PS_File
' Objetivo: Escribe o le� el archivo ps
'     @ xMode [IOMode] -> Modo de lectura o escritura
' ----------------------------------------------------------------
Private Sub PS_File(ByVal xMode As IOMode)

    Dim FSO As Object
    Dim PS_File As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set PS_File = FSO.OpenTextFile(pRuta_PS, xMode)

    Select Case xMode
        Case IOMode.ForReading
            pPS_TxT = Split(PS_File.ReadAll, vbNewLine)
        Case IOMode.ForWriting
            PS_File.Write Join(pPS_TxT, vbNewLine)
    End Select

    PS_File.Close

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: UpdateZP
' Objetivo: Cierra el portafolio y vuelve a abrirlo
' ----------------------------------------------------------------
Public Sub UpdateZP()

    ClosePortfolio
    OpenPortfolio (pInputName)

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: GetResults_Portfolio
' Objetivo: Si ya se encuentran definidos los anal�ticos del portafolio, entonces extrae los resultados como una propiedad
' ----------------------------------------------------------------
Public Sub GetResults_Portfolio()

    Dim arrRes As Variant
    Dim colAna As Long

    ReDim arrRes(1 To 2, 1 To UBound(pPortfolioAnalytics))

    For colAna = 1 To UBound(pPortfolioAnalytics)
        arrRes(1, colAna) = pPortfolioAnalytics(colAna, 1)
        arrRes(2, colAna) = Null_to_string(GetPortfolioAnalytic(pPortfolioAnalytics(colAna, 1)))
    Next colAna
    
    pArrResPort = arrRes

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: ZDT_GetPortfolioSource
' Objetivo: Se definen propiedades de clase:
'   @Si el origen del portafolio hace referencia a una ruta, abrir� desde zp
'   @De lo contrario se intentar� abrir desde base de datos
' ----------------------------------------------------------------
Private Sub ZDT_GetPortfolioSource()

    If pInputName <> "" Then
        If Right(pInputName, 3) = ".zp" _
        And Dir(pInputName) <> "" Then
            pOpenSource = zdt_From_ZP
            pRuta_ZP = pInputName
            pNombre = Replace(Dir(pInputName), ".zp", "")
            If Dir(Replace(pRuta_ZP, ".zp", ".ps")) <> "" Then
                pRuta_PS = Replace(pRuta_ZP, ".zp", ".ps")
                PS_File ForReading
            End If
        Else
            pOpenSource = zdt_From_DB
            pNombre = pInputName
        End If
    Else
        MsgBox ("No se ha ingresado un nombre de portafolio valido.")
    End If

End Sub

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: GetHoldingTicker
' Objetivo: Extrae el identificador de ZEUS del instrumento
' Regresa: Variant
'     @ NumIns [Integer] -> �ndice del instrumento dentro del portafolio
' ----------------------------------------------------------------
Public Function GetHoldingTicker(ByVal NumIns As Integer) As Variant

    GetHoldingTicker = pZEUS.GetSecurityCode(pNumPort, "T", NumIns)

End Function

' ----------------------------------------------------------------
' Fecha: 21/04/2021 // Aar�n Santana
' Nombre: GetHoldingAnalytic
' Objetivo: Extra� de ZEUS el anal�tico desde el ticker de un instrumento
' Regresa: Variant
'     @ Ticker [String] -> Identificador del instrumento
'     @ AnalyticID [String] -> Identificador del anal�tico
' ----------------------------------------------------------------
Public Function GetHoldingAnalytic(ByVal Ticker As String, ByVal AnalyticID As String) As Variant

    GetHoldingAnalytic = pZEUS.GetSecurityAnalytic(pNumPort, Ticker, "T", AnalyticID)

End Function

' ----------------------------------------------------------------
' Fecha: 15/04/2021 // Aar�n Santana
' Nombre: Null_to_string
' Objetivo: Asegura la conversi�n de un valor a cadena de texto, si es nulo devuelve cadena vacia
'     @ xValor [Variant] -> Valor a validar
' ----------------------------------------------------------------
Private Function Null_to_string(ByVal xValor As Variant) As String

    If IsNull(xValor) Then
        Null_to_string = ""
    Else
        Null_to_string = xValor
    End If

End Function

' ----------------------------------------------------------------
' Fecha: 15/04/2021 // Aar�n Santana
' Nombre: GetResults_Holdings
' Objetivo: Revisa si el array de anal�ticos tiene valores, de ser asi extrae los valores y los almacena en un array de resultados.
' ----------------------------------------------------------------
Public Sub GetResults_Holdings()

    Dim arrRes As Variant
    Dim txtTicker As Variant
    Dim colRes As Long
    Dim filaRes As Long
    
    If Not IsArray(pHoldingAnalytics) Then
        MsgBox "No se ha definido la propiedad HoldingAnalytics"
        Exit Sub
    End If
    
    ReDim arrRes(1 To pNumIns + 1, 1 To UBound(pHoldingAnalytics))

    For colRes = 1 To UBound(pHoldingAnalytics)
        arrRes(1, colRes) = pHoldingAnalytics(colRes, 1)
    Next colRes

    For filaRes = 2 To pNumIns + 1
        txtTicker = GetHoldingTicker(filaRes - 2)
        For colRes = 1 To UBound(pHoldingAnalytics)
            arrRes(filaRes, colRes) = Null_to_string(GetHoldingAnalytic(txtTicker, pHoldingAnalytics(colRes, 1)))
        Next colRes
    Next filaRes

    pArrResIns = arrRes

End Sub

' ----------------------------------------------------------------
' Fecha: 15/04/2021 // Aar�n Santana
' Nombre: Class_Initialize
' Objetivo: Inicializa la clase ZDT
' ----------------------------------------------------------------
Private Sub Class_Initialize()

    On Error GoTo Error_handler

    Set pZEUS = CreateObject("Zeus.Dev")

    Exit Sub

Error_handler:

    MsgBox "No se ha iniciado Zeus correctamente." & vbCrLf & _
            "Ejecuta en modo administrador ZDT y Zeus, si ya lo intentaste y no logr�s ejecutar ponte en contacto con L1_soporte@riskconsult.com.mx"

End Sub

' ----------------------------------------------------------------
' Fecha: 15/04/2021 // Aar�n Santana
' Nombre: ClosePortfolio
' Objetivo: En Zeus cierra el portafolio
' ----------------------------------------------------------------
Public Sub ClosePortfolio()

    pNumIns = pZEUS.CloseDocument(pNumPort)

End Sub
