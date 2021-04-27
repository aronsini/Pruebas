ESTA LÍNEA SOLO SE AGREGA COMO PRUEBA DE ACTUALIZACIÓN
Option Explicit
Option Base 1

'#############################################################################################
'#----------------------------------ENUMS----------------------------------------------------#
'#############################################################################################

Enum colFactorStrings
    COL_FACTOR_STRINGS_ID = 1
    COL_FACTOR_STRINGS_NAME = 2
    COL_FACTOR_STRINGS_DESCRIPTION = 3
    COL_FACTOR_STRINGS_TYPE = 4
    COL_FACTOR_STRINGS_COLUMN = 5
End Enum

'#############################################################################################
'#----------------------------------FUNCIONES------------------------------------------------#
'#############################################################################################

Private Function sqlGetQueryResults(ByVal CadenaConexion As String, _
                                    ByVal Consulta As String) As Variant

    Dim Conexion As New ADODB.Connection   'Connection
    Dim RecSet As New ADODB.Recordset   'Recordset
    Dim Header As Field
    Dim intCol As Integer

    Conexion.Open CadenaConexion
    RecSet.Open Consulta, Conexion
    sqlGetQueryResults = Application.WorksheetFunction.Transpose(RecSet.GetRows)

    RecSet.Close
    Conexion.Close

    Set RecSet = Nothing
    Set Conexion = Nothing

End Function

Private Function ZEUS_FactorGroup_BMV() As Variant

    ZEUS_FactorGroup_BMV = Array(300, 301, 302, 303, 304, 305, 306, 307, 308, 309, 311, 310, 1404, 1402, 1401, 1405, 1400, 1403, 44, 35)

End Function

Private Function ZEUS_FactorGroup_Currency() As Variant

    ZEUS_FactorGroup_Currency = Array(87, 107, 83, 103, 62, 79, 110, 334, 330, 44, 36, 332, 333, 329, 76)

End Function

Private Function ZEUS_FactorGroup_CurveMXN() As Variant

    ZEUS_FactorGroup_CurveMXN = Array(6, 10070, 0, 3, 10071, 4, 10, 1, 5, 2, 52, 54)

End Function

Private Function ZEUS_FactorGroup_Curves() As Variant

    ZEUS_FactorGroup_Curves = Array(87, 107, 83, 103, 62, 79, 110, 334, 330, 44, 36, 332, 333, 329, 76, 70, 64, 67, 68, 65, 69, 66, 43, 37, 40, 112, 41, _
                            113, 38, 42, 39, 6, 10070, 0, 3, 10071, 4, 10, 1, 5, 2, 46, 47, 55, 49, 51, 50, 48, 52, 54, 53, 20, 10059, 10061, 21, 10062, _
                            22, 10060, 10063, 17, 11, 14, 18, 15, 19, 12, 16, 13, 34, 28, 31, 10072, 32, 114, 29, 33, 30, 10086, 10080, 10083, 10115, 10084, _
                            10116, 10081, 10085, 10082)

End Function

Private Function ZEUS_FactorGroup_ETF() As Variant

    ZEUS_FactorGroup_ETF = Array(87, 107, 83, 103, 62, 79, 110, 334, 330, 44, 36, 332, 333, 329, 76, 70, 64, 67, 68, 65, 69, 66, 43, 37, 40, _
                          112, 41, 113, 38, 42, 39, 6, 10070, 0, 3, 10071, 4, 10, 1, 5, 2, 46, 47, 55, 49, 51, 50, 48, 52, 54, 53, 20, 10059, _
                          10061, 21, 10062, 22, 10060, 10063, 17, 11, 14, 18, 15, 19, 12, 16, 13, 34, 28, 31, 10072, 32, 114, 29, 33, 30, 10086, _
                          10080, 10083, 10115, 10084, 10116, 10081, 10085, 10082, 89, 84, 90, 75, 117, 78, 322, 118, 327, 35, 104, 321, 328, 108, _
                          105, 85, 325, 111, 81, 80, 320, 326, 324, 91, 106, 71, 323, 77, 86, 63, 88, 60, 82)

End Function

Private Function ZEUS_FactorGroup_SIC() As Variant

    ZEUS_FactorGroup_SIC = Array(87, 107, 83, 103, 62, 79, 110, 334, 330, 44, 36, 332, 333, 329, 76, 89, 84, 90, 75, 117, 78, 322, 118, _
                          327, 35, 104, 321, 328, 108, 105, 85, 325, 111, 81, 80, 320, 326, 324, 91, 106, 71, 323, 77, 86, 63, 88, 60, 82)

End Function

Public Function ZEUS_GetFactors_FullReturnsMatrix(ByVal ZeusDB_ConexionString As String, _
                                                  ByVal dteCalculationDate As Date, _
                                                  ByVal intLatestScenarios As Long) As Variant

    Dim Consulta As String
    
    Consulta = "SELECT [300], [301], [302], [303], [304], [305], [306], [307], [308], [309], [311], [310], [92], [87], [107], [83], [103], [62], [79], " & _
               "[110], [334], [330], [93], [44], [36], [332], [333], [329], [76], [70], [64], [67], [68], [65], [69], [66], [43], [37], [40], [112], " & _
               "[41], [113], [38], [42], [39], [6], [10070], [0], [3], [10071], [4], [10], [1], [5], [2], [46], [47], [55], [49], [51], [50], [48], " & _
               "[52], [54], [53], [20], [10059], [10061], [21], [10062], [22], [10060], [10063], [17], [11], [14], [18], [15], [19], [12], [16], [13], " & _
               "[34], [28], [31], [10072], [32], [114], [29], [33], [30], [10086], [10080], [10083], [10115], [10084], [10116], [10081], [10085], [10082], " & _
               "[89], [84], [90], [75], [117], [78], [322], [118], [327], [35], [104], [321], [328], [108], [105], [85], [325], [111], [81], [80], [320], " & _
               "[326], [324], [91], [106], [71], [323], [77], [86], [63], [88], [60], [82], [1404], [1402], [1401], [1405], [1400], [1403] " & _
               "FROM (SELECT TOP " & intLatestScenarios & " dteDate FROM tblDATA_BusinessDays WHERE dteDate <= '" & Format(dteCalculationDate, "YYYYMMDD") & "' ORDER BY " & _
               "dteDate DESC) D LEFT JOIN (SELECT dteDate, [300], [301], [302], [303], [304], [305], [306], [307], [308], [309], [311], [310], [92], " & _
               "[87], [107], [83], [103], [62], [79], [110], [334], [330], [93], [44], [36], [332], [333], [329], [76], [70], [64], [67], [68], [65], " & _
               "[69], [66], [43], [37], [40], [112], [41], [113], [38], [42], [39], [6], [10070], [0], [3], [10071], [4], [10], [1], [5], [2], [46], " & _
               "[47], [55], [49], [51], [50], [48], [52], [54], [53], [20], [10059], [10061], [21], [10062], [22], [10060], [10063], [17], [11], [14], " & _
               "[18], [15], [19], [12], [16], [13], [34], [28], [31], [10072], [32], [114], [29], [33], [30], [10086], [10080], [10083], [10115], [10084], " & _
               "[10116], [10081], [10085], [10082], [89], [84], [90], [75], [117], [78], [322], [118], [327], [35], [104], [321], [328], [108], [105], " & _
               "[85], [325], [111], [81], [80], [320], [326], [324], [91], [106], [71], [323], [77], [86], [63], [88], [60], [82], [1404], [1402], [1401], " & _
               "[1405], [1400], [1403] FROM tblDATA_FactorReturns F PIVOT (SUM(dblValue) FOR intID IN ([300], [301], [302], [303], [304], [305], [306], " & _
               "[307], [308], [309], [311], [310], [92], [87], [107], [83], [103], [62], [79], [110], [334], [330], [93], [44], [36], [332], [333], [329], " & _
               "[76], [70], [64], [67], [68], [65], [69], [66], [43], [37], [40], [112], [41], [113], [38], [42], [39], [6], [10070], [0], [3], [10071], " & _
               "[4], [10], [1], [5], [2], [46], [47], [55], [49], [51], [50], [48], [52], [54], [53], [20], [10059], [10061], [21], [10062], [22], [10060], " & _
               "[10063], [17], [11], [14], [18], [15], [19], [12], [16], [13], [34], [28], [31], [10072], [32], [114], [29], [33], [30], [10086], [10080], " & _
               "[10083], [10115], [10084], [10116], [10081], [10085], [10082], [89], [84], [90], [75], [117], [78], [322], [118], [327], [35], [104], [321], " & _
               "[328], [108], [105], [85], [325], [111], [81], [80], [320], [326], [324], [91], [106], [71], [323], [77], [86], [63], [88], [60], [82], " & _
               "[1404], [1402], [1401], [1405], [1400], [1403])) P) F ON D.dteDate = F.dteDate ORDER BY D.dteDate ASC"
    
    ZEUS_GetFactors_FullReturnsMatrix = sqlGetQueryResults(ZeusDB_ConexionString, Consulta)

End Function

Public Function ZEUS_GetFactors_Group(ByVal Modulo As String, _
                                      ByVal TV As String, _
                                      ByVal Moneda As String) As Variant
    
    Select Case Modulo
        Case "Zero Coupon", "Fixed Coupon", "Floater Coupon"
            Select Case TV
                Case "BI", "M", "LD", "IM", "IQ", "IS"
                    ZEUS_GetFactors_Group = ZEUS_FactorGroup_CurveMXN
                Case Else
                    ZEUS_GetFactors_Group = ZEUS_FactorGroup_Curves
            End Select
        Case "Equity"
            Select Case TV
                Case "0", "00", "1", "1B", "1R", "41", "CF", "FE", "FF", "FH", "YY", "YYSP"
                    ZEUS_GetFactors_Group = ZEUS_FactorGroup_BMV
                Case "*I", "1A", "1ASP", "1E", "1ESP", "RC"
                    ZEUS_GetFactors_Group = ZEUS_FactorGroup_SIC
                Case Else
                    ZEUS_GetFactors_Group = ZEUS_FactorGroup_ETF
            End Select
        Case "Cash"
            ZEUS_GetFactors_Group = ZEUS_FactorGroup_Currency
    End Select

End Function

'Obtiene matriz de FactorIDs desde el nombre de los factores
Public Function ZEUS_GetFactors_IDsFromDesc(ByVal arrFactorsDescriptions As Variant) As Variant
    
    Dim arrRes As Variant
    Dim filaRes As Long
    
    'Aseguramos conversion a array
    arrFactorsDescriptions = arrFactorsDescriptions
    
    'Igualamos dimensiones
    ReDim arrRes(1 To UBound(arrFactorsDescriptions), 1 To 1)
    
    'Loop que identifica cada descripción con un ID
    For filaRes = 1 To UBound(arrFactorsDescriptions)
        arrRes(filaRes, 1) = ZEUS_GetFactor_ID(arrFactorsDescriptions(filaRes, 1))
    Next filaRes
    
    ZEUS_GetFactors_IDsFromDesc = arrRes
    
End Function

'Extrae una matriz de rendimientos i x j con i = numero de filas en arrFullFactorsReturns y j = numero de IDs en arrIDS
Public Function ZEUS_GetFactors_Returns(ByVal arrFullFactorsReturns As Variant, _
                                        ByVal arrIDs As Variant) As Variant
    
    Dim arrRes As Variant
    Dim collFS As Collection
    Dim colFactor As Long
    Dim colRes As Long
    Dim filaRes As Long
    
    'Aseguramos que arrIDs se convierta en matriz
    arrIDs = arrIDs
    
    'redimensionamos matriz resultado y obtenemos colleccion de strings
    ReDim arrRes(1 To UBound(arrFullFactorsReturns), 1 To UBound(arrIDs))
    Set collFS = ZEUS_GetFactors_StringsCollection
    
    'iteramos columnas desde numero de ids
    For colRes = LBound(arrIDs) To UBound(arrIDs)
        'obtenemos la columna en la que se ubican los rendimientos del factor(colres)
        colFactor = ZEUS_GetFactor_String(arrIDs(colRes, LBound(arrIDs, 2)), COL_FACTOR_STRINGS_ID, COL_FACTOR_STRINGS_COLUMN)
        'iteramos sobre los rendimientos(filares) del factor(colres)
        For filaRes = LBound(arrFullFactorsReturns) To UBound(arrFullFactorsReturns)
            'asignamos valor que le corresponmde
            arrRes(filaRes, colRes) = arrFullFactorsReturns(filaRes, colFactor)
        Next filaRes
    Next colRes
    
    'Regresamos matriz de rendimientos para cada factor en arrIDs
    ZEUS_GetFactors_Returns = arrRes
    
End Function

Public Function ZEUS_GetFactors_StringsCollection() As Collection
    
    Dim tmpColl As New Collection
    
    With tmpColl
        .Add Array(95, "CRUDE_PALM_OIL", "Crude Palm Oil Commodity Factor (Kuala Lumpur)", "Commodity", -1)
        .Add Array(116, "GLD", "Gold Commodity Factor (ETF)", "Commodity", -1)
        .Add Array(115, "HEATING_OIL", "Heating Oil Commodity Factor (New York)", "Commodity", -1)
        .Add Array(98, "NATURAL_GAS", "Natural Gas Commodity Factor (New York)", "Commodity", -1)
        .Add Array(99, "POLYETHYLENE", "Polyethylene Commodity Factor (London)", "Commodity", -1)
        .Add Array(100, "POLYPROPYLENE", "Polypropylene Commodity Factor (London)", "Commodity", -1)
        .Add Array(94, "SOYBEAN_OIL", "Soybean Oil Commodity Factor (Chicago)", "Commodity", -1)
        .Add Array(109, "SUGAR_FSB", "Sugar Commodity Factor (New York)", "Commodity", -1)
        .Add Array(96, "WHEAT_KW", "Wheat Commodity Factor (Kansas)", "Commodity", -1)
        .Add Array(97, "WHEAT_MW", "Wheat Commodity Factor (Minneapolis)", "Commodity", -1)
        .Add Array(101, "WHEAT_VJ", "Wheat Commodity Factor (Argentina)", "Commodity", -1)
        .Add Array(102, "WHEAT_W", "Wheat Commodity Factor (Chicago)", "Commodity", -1)
        .Add Array(92, "ARS", "Argentine Peso", "Currency", 13)
        .Add Array(87, "AUD", "Australian Dollar", "Currency", 14)
        .Add Array(107, "BRL", "Brazilian Real", "Currency", 15)
        .Add Array(83, "CAD", "Canadian Dollar", "Currency", 16)
        .Add Array(103, "CNY", "Chinese Yuan", "Currency", 17)
        .Add Array(62, "EUR", "EURO", "Currency", 18)
        .Add Array(79, "GBP", "Great Britain Pound", "Currency", 19)
        .Add Array(110, "HKD", "Hong Kong Dollar", "Currency", 20)
        .Add Array(334, "INR", "India Rupee", "Currency", 21)
        .Add Array(330, "KRW", "Korea (South) Won", "Currency", 22)
        .Add Array(44, "MXN", "Mexican Peso", "Currency", 24)
        .Add Array(93, "MYR", "Malaysian Ringgit", "Currency", 23)
        .Add Array(332, "SGD", "Singapore Dollar", "Currency", 26)
        .Add Array(329, "TRY", "Turkey Lira", "Currency", 28)
        .Add Array(333, "TWD", "Taiwan New Dollar", "Currency", 27)
        .Add Array(36, "UDI", "Mexican UDI", "Currency", 25)
        .Add Array(76, "YEN", "YEN", "Currency", 29)
        .Add Array(64, "EUR1", "EUR 1m", "Curve", 31)
        .Add Array(65, "EUR2", "EUR 3m", "Curve", 34)
        .Add Array(66, "EUR3", "EUR 6m", "Curve", 36)
        .Add Array(67, "EUR4", "EUR 1y", "Curve", 32)
        .Add Array(68, "EUR5", "EUR 2y", "Curve", 33)
        .Add Array(69, "EUR6", "EUR 5y", "Curve", 35)
        .Add Array(70, "EUR7", "EUR 10y", "Curve", 30)
        .Add Array(37, "LIB1", "LIB 1m", "Curve", 38)
        .Add Array(38, "LIB2", "LIB 3m", "Curve", 43)
        .Add Array(39, "LIB3", "LIB 6m", "Curve", 45)
        .Add Array(40, "LIB4", "LIB 1y", "Curve", 39)
        .Add Array(41, "LIB5", "LIB 2y", "Curve", 41)
        .Add Array(42, "LIB6", "LIB 5y", "Curve", 44)
        .Add Array(43, "LIB7", "LIB 10y", "Curve", 37)
        .Add Array(112, "LIB8", "LIB 20y", "Curve", 40)
        .Add Array(113, "LIB9", "LIB 30y", "Curve", 42)
        .Add Array(10070, "MXN0", "MXN 1d", "Curve", 47)
        .Add Array(0, "MXN1", "MXN 1m", "Curve", 48)
        .Add Array(1, "MXN2", "MXN 3m", "Curve", 53)
        .Add Array(2, "MXN3", "MXN 6m", "Curve", 55)
        .Add Array(3, "MXN4", "MXN 1y", "Curve", 49)
        .Add Array(4, "MXN5", "MXN 2y", "Curve", 51)
        .Add Array(5, "MXN6", "MXN 5y", "Curve", 54)
        .Add Array(6, "MXN7", "MXN 10y", "Curve", 46)
        .Add Array(10071, "MXN8", "MXN 20y", "Curve", 50)
        .Add Array(10, "MXN9", "MXN 30y", "Curve", 52)
        .Add Array(10059, "TIIE1", "TIIE 1m", "Curve", 67)
        .Add Array(10060, "TIIE2", "TIIE 3m", "Curve", 72)
        .Add Array(10061, "TIIE3", "TIIE 1y", "Curve", 68)
        .Add Array(10062, "TIIE4", "TIIE 2y", "Curve", 70)
        .Add Array(10063, "TIIE5", "TIIE 5y", "Curve", 73)
        .Add Array(20, "TIIE6", "TIIE 10y", "Curve", 66)
        .Add Array(21, "TIIE7", "TIIE 20y", "Curve", 69)
        .Add Array(22, "TIIE8", "TIIE 30y", "Curve", 71)
        .Add Array(11, "TRS1", "TRS 1m", "Curve", 75)
        .Add Array(12, "TRS2", "TRS 3m", "Curve", 80)
        .Add Array(13, "TRS3", "TRS 6m", "Curve", 82)
        .Add Array(14, "TRS4", "TRS 1y", "Curve", 76)
        .Add Array(15, "TRS5", "TRS 2y", "Curve", 78)
        .Add Array(16, "TRS6", "TRS 5y", "Curve", 81)
        .Add Array(17, "TRS7", "TRS 10y", "Curve", 74)
        .Add Array(18, "TRS8", "TRS 20y", "Curve", 77)
        .Add Array(19, "TRS9", "TRS 30y", "Curve", 79)
        .Add Array(28, "UDI1", "UDI 1m", "Curve", 84)
        .Add Array(29, "UDI2", "UDI 3m", "Curve", 89)
        .Add Array(30, "UDI3", "UDI 6m", "Curve", 91)
        .Add Array(31, "UDI4", "UDI 1y", "Curve", 85)
        .Add Array(32, "UDI5", "UDI 2y", "Curve", 87)
        .Add Array(33, "UDI6", "UDI 5y", "Curve", 90)
        .Add Array(34, "UDI7", "UDI 10y", "Curve", 83)
        .Add Array(10072, "UDI8", "UDI 20y", "Curve", 86)
        .Add Array(114, "UDI9", "UDI 30y", "Curve", 88)
        .Add Array(10080, "UMS1", "UMS 1m", "Curve", 93)
        .Add Array(10081, "UMS2", "UMS 3m", "Curve", 98)
        .Add Array(10082, "UMS3", "UMS 6m", "Curve", 100)
        .Add Array(10083, "UMS4", "UMS 1y", "Curve", 94)
        .Add Array(10084, "UMS5", "UMS 2y", "Curve", 96)
        .Add Array(10085, "UMS6", "UMS 5y", "Curve", 99)
        .Add Array(10086, "UMS7", "UMS 10y", "Curve", 92)
        .Add Array(10115, "UMS8", "UMS 20y", "Curve", 95)
        .Add Array(10116, "UMS9", "UMS 30y", "Curve", 97)
        .Add Array(10094, "df_EUR1", "DiscFactor EUR 1m", "DiscountFactor", -1)
        .Add Array(10095, "df_EUR2", "DiscFactor EUR 3m", "DiscountFactor", -1)
        .Add Array(10096, "df_EUR3", "DiscFactor EUR 6m", "DiscountFactor", -1)
        .Add Array(10097, "df_EUR4", "DiscFactor EUR 1y", "DiscountFactor", -1)
        .Add Array(10098, "df_EUR5", "DiscFactor EUR 2y", "DiscountFactor", -1)
        .Add Array(10099, "df_EUR6", "DiscFactor EUR 5y", "DiscountFactor", -1)
        .Add Array(10100, "df_EUR7", "DiscFactor EUR 10y", "DiscountFactor", -1)
        .Add Array(10037, "df_LIB1", "DiscFactor LIB 1m", "DiscountFactor", -1)
        .Add Array(10038, "df_LIB2", "DiscFactor LIB 3m", "DiscountFactor", -1)
        .Add Array(10039, "df_LIB3", "DiscFactor LIB 6m", "DiscountFactor", -1)
        .Add Array(10040, "df_LIB4", "DiscFactor LIB 1y", "DiscountFactor", -1)
        .Add Array(10041, "df_LIB5", "DiscFactor LIB 2y", "DiscountFactor", -1)
        .Add Array(10042, "df_LIB6", "DiscFactor LIB 5y", "DiscountFactor", -1)
        .Add Array(10043, "df_LIB7", "DiscFactor LIB 10y", "DiscountFactor", -1)
        .Add Array(10112, "df_LIB8", "DiscFactor LIB 20y", "DiscountFactor", -1)
        .Add Array(10113, "df_LIB9", "DiscFactor LIB 30y", "DiscountFactor", -1)
        .Add Array(10073, "df_MXN0", "DiscFactor MXN 1d", "DiscountFactor", -1)
        .Add Array(10000, "df_MXN1", "DiscFactor MXN 1m", "DiscountFactor", -1)
        .Add Array(10001, "df_MXN2", "DiscFactor MXN 3m", "DiscountFactor", -1)
        .Add Array(10002, "df_MXN3", "DiscFactor MXN 6m", "DiscountFactor", -1)
        .Add Array(10003, "df_MXN4", "DiscFactor MXN 1y", "DiscountFactor", -1)
        .Add Array(10004, "df_MXN5", "DiscFactor MXN 2y", "DiscountFactor", -1)
        .Add Array(10005, "df_MXN6", "DiscFactor MXN 5y", "DiscountFactor", -1)
        .Add Array(10006, "df_MXN7", "DiscFactor MXN 10y", "DiscountFactor", -1)
        .Add Array(10074, "df_MXN8", "DiscFactor MXN 20y", "DiscountFactor", -1)
        .Add Array(10010, "df_MXN9", "DiscFactor MXN 30y", "DiscountFactor", -1)
        .Add Array(10047, "df_SPD_BDELT", "DiscFactor Spread BDE LT", "DiscountFactor", -1)
        .Add Array(10046, "df_SPD_BDENW", "DiscFactor Spread BDE LP", "DiscountFactor", -1)
        .Add Array(10052, "df_SPD_BPABP1", "DiscFactor Spread BPABP1", "DiscountFactor", -1)
        .Add Array(10048, "df_SPD_BPSBP1", "DiscFactor Spread BPSBP1", "DiscountFactor", -1)
        .Add Array(10053, "df_SPD_BPTBP1", "DiscFactor Spread BPTBP1", "DiscountFactor", -1)
        .Add Array(10055, "df_SPD_CETCTI", "DiscFactor Spread CetesCTI", "DiscountFactor", -1)
        .Add Array(10049, "df_SPD_PLV3A", "DiscFactor Spread PLV 3A", "DiscountFactor", -1)
        .Add Array(10051, "df_SPD_PLVP0", "DiscFactor Spread PLV P0", "DiscountFactor", -1)
        .Add Array(10050, "df_SPD_PLVP8", "DiscFactor Spread PLV P8", "DiscountFactor", -1)
        .Add Array(10054, "df_SPD_XAXA0", "DiscFactor Spread XAXA0", "DiscountFactor", -1)
        .Add Array(10064, "df_TIIE1", "DiscFactor TIIE 1m", "DiscountFactor", -1)
        .Add Array(10065, "df_TIIE2", "DiscFactor TIIE 3m", "DiscountFactor", -1)
        .Add Array(10066, "df_TIIE3", "DiscFactor TIIE 1Y", "DiscountFactor", -1)
        .Add Array(10067, "df_TIIE4", "DiscFactor TIIE 2Y", "DiscountFactor", -1)
        .Add Array(10068, "df_TIIE5", "DiscFactor TIIE 5Y", "DiscountFactor", -1)
        .Add Array(10020, "df_TIIE6", "DiscFactor TIIE 10y", "DiscountFactor", -1)
        .Add Array(10021, "df_TIIE7", "DiscFactor TIIE 20y", "DiscountFactor", -1)
        .Add Array(10022, "df_TIIE8", "DiscFactor TIIE 30y", "DiscountFactor", -1)
        .Add Array(10011, "df_TRS1", "DiscFactor TRS 1m", "DiscountFactor", -1)
        .Add Array(10012, "df_TRS2", "DiscFactor TRS 3m", "DiscountFactor", -1)
        .Add Array(10013, "df_TRS3", "DiscFactor TRS 6m", "DiscountFactor", -1)
        .Add Array(10014, "df_TRS4", "DiscFactor TRS 1y", "DiscountFactor", -1)
        .Add Array(10015, "df_TRS5", "DiscFactor TRS 2y", "DiscountFactor", -1)
        .Add Array(10016, "df_TRS6", "DiscFactor TRS 5y", "DiscountFactor", -1)
        .Add Array(10017, "df_TRS7", "DiscFactor TRS 10y", "DiscountFactor", -1)
        .Add Array(10018, "df_TRS8", "DiscFactor TRS 20y", "DiscountFactor", -1)
        .Add Array(10019, "df_TRS9", "DiscFactor TRS 30y", "DiscountFactor", -1)
        .Add Array(10028, "df_UDI1", "DiscFactor UDI 1m", "DiscountFactor", -1)
        .Add Array(10029, "df_UDI2", "DiscFactor UDI 3m", "DiscountFactor", -1)
        .Add Array(10030, "df_UDI3", "DiscFactor UDI 6m", "DiscountFactor", -1)
        .Add Array(10031, "df_UDI4", "DiscFactor UDI 1y", "DiscountFactor", -1)
        .Add Array(10032, "df_UDI5", "DiscFactor UDI 2y", "DiscountFactor", -1)
        .Add Array(10033, "df_UDI6", "DiscFactor UDI 5y", "DiscountFactor", -1)
        .Add Array(10034, "df_UDI7", "DiscFactor UDI 10y", "DiscountFactor", -1)
        .Add Array(10075, "df_UDI8", "DiscFactor UDI 20y", "DiscountFactor", -1)
        .Add Array(10114, "df_UDI9", "DiscFactor UDI 30y", "DiscountFactor", -1)
        .Add Array(10087, "df_UMS1", "DiscFactor UMS 1m", "DiscountFactor", -1)
        .Add Array(10088, "df_UMS2", "DiscFactor UMS 3m", "DiscountFactor", -1)
        .Add Array(10089, "df_UMS3", "DiscFactor UMS 6m", "DiscountFactor", -1)
        .Add Array(10090, "df_UMS4", "DiscFactor UMS 1y", "DiscountFactor", -1)
        .Add Array(10091, "df_UMS5", "DiscFactor UMS 2y", "DiscountFactor", -1)
        .Add Array(10092, "df_UMS6", "DiscFactor UMS 5y", "DiscountFactor", -1)
        .Add Array(10093, "df_UMS7", "DiscFactor UMS 10y", "DiscountFactor", -1)
        .Add Array(10117, "df_UMS8", "DiscFactor UMS 20y", "DiscountFactor", -1)
        .Add Array(10118, "df_UMS9", "DiscFactor UMS 30y", "DiscountFactor", -1)
        .Add Array(88, "ASX50", "S&P ASX 50 IND", "Index", 131)
        .Add Array(89, "CAC", "CAC 40 IND", "Index", 101)
        .Add Array(90, "DAX", "DAX IND", "Index", 103)
        .Add Array(75, "DJIA", "Dow Jones IA", "Index", 104)
        .Add Array(108, "EWZ", "MSCI Brazil Index", "Index", 114)
        .Add Array(78, "FTSE", "FTSE 100", "Index", 106)
        .Add Array(322, "HSCEI", "HK Stock Exc. Hang Seng China Ent. Index", "Index", 107)
        .Add Array(118, "IBEX", "IBEX 35 Index Factor", "Index", 108)
        .Add Array(35, "IPC", "MEXBOL IND", "Index", 110)
        .Add Array(327, "KOSPI", "Korea Stock Exchange KOSPI Index", "Index", 109)
        .Add Array(81, "MSDUJN", "MSCI Japan", "Index", 119)
        .Add Array(91, "MSDUUK", "MSCI United Kingdom", "Index", 124)
        .Add Array(328, "MXAPJ", "MSCI Asia Pacific ex Japan", "Index", 113)
        .Add Array(321, "MXASJ", "MSCI Asia ex Japan", "Index", 112)
        .Add Array(105, "MXEF", "MSCI Emerging MKT", "Index", 115)
        .Add Array(85, "MXEM", "MSCI EMU", "Index", 116)
        .Add Array(325, "MXEUG", "MSCI Europe ex UK", "Index", 117)
        .Add Array(111, "MXHK", "MSCI Hong Kong", "Index", 118)
        .Add Array(80, "MXPCJ", "MSCI Pacific x Japan", "Index", 120)
        .Add Array(320, "MXSG", "MSCI Singapore", "Index", 121)
        .Add Array(324, "MXTR", "MSCI Turkey", "Index", 123)
        .Add Array(104, "MXWD", "MSCI AC World", "Index", 111)
        .Add Array(106, "MXWO", "MSCI World", "Index", 125)
        .Add Array(71, "NASDAQ", "NASDAQ", "Index", 126)
        .Add Array(323, "NIFTY", "National Stock Exchange CNX Nifty Index", "Index", 127)
        .Add Array(77, "NIKKEI", "NIKKEI", "Index", 128)
        .Add Array(86, "RAY", "Russel 3000 IND", "Index", 129)
        .Add Array(63, "S&P", "S&P", "Index", 130)
        .Add Array(117, "SX5E", "EURO STOXX 50 Index Factor", "Index", 105)
        .Add Array(326, "TAMSCI", "MSCI Taiwan", "Index", 122)
        .Add Array(60, "TFB_IDX", "TFB", "Index", 132)
        .Add Array(82, "TOPIX", "TOPIX", "Index", 133)
        .Add Array(84, "TSX60", "Canadian Equity IND", "Index", 102)
        .Add Array(300, "AlimentoBebidaTabaco", "ALIMENTO BEBIDA TABACO", "Industry", 1)
        .Add Array(301, "Autoservicios", "AUTOSERVICIOS", "Industry", 2)
        .Add Array(302, "Cemento", "CEMENTO", "Industry", 3)
        .Add Array(303, "Comercio", "COMERCIO", "Industry", 4)
        .Add Array(304, "Comunicaciones", "COMUNICACIONES", "Industry", 5)
        .Add Array(305, "Construccion", "CONSTRUCCION", "Industry", 6)
        .Add Array(306, "GruposFinanceiros", "GRUPOS FINANCIEROS", "Industry", 7)
        .Add Array(307, "Holdings", "HOLDINGS", "Industry", 8)
        .Add Array(308, "IndTransformacion", "INDUSTRIA DE LA TRANSFORMACION", "Industry", 9)
        .Add Array(309, "MediosYEntretenim", "MEDIOS Y ENTRETENIMIENTO", "Industry", 10)
        .Add Array(311, "Transporte", "TRANSPORTE", "Industry", 11)
        .Add Array(310, "Varios", "VARIOS", "Industry", 12)
        .Add Array(47, "SPD_BDELT", "SPD BDE LT", "Spread", 57)
        .Add Array(46, "SPD_BDENW", "SPD BDE LP", "Spread", 56)
        .Add Array(52, "SPD_BPABP1", "SPD Rev. 28d IP 3y", "Spread", 63)
        .Add Array(48, "SPD_BPSBP1", "SPD Rev. 182d 3y", "Spread", 62)
        .Add Array(53, "SPD_BPTBP1", "SPD Rev. 90d 3y", "Spread", 65)
        .Add Array(55, "SPD_CETCTI", "SPD CetesCTI", "Spread", 58)
        .Add Array(49, "SPD_PLV3A", "SPD PLV 3A", "Spread", 59)
        .Add Array(51, "SPD_PLVP0", "SPD PLV P0", "Spread", 60)
        .Add Array(50, "SPD_PLVP8", "SPD PLV P8", "Spread", 61)
        .Add Array(54, "SPD_XAXA0", "SPD Rev. 28d LD_XA 1y", "Spread", 64)
        .Add Array(1404, "Dolar", "Dolar", "Style", 134)
        .Add Array(1402, "Momentum", "Momentum", "Style", 135)
        .Add Array(1401, "Size", "Size", "Style", 136)
        .Add Array(1405, "TradingActivity", "TradingActivity", "Style", 137)
        .Add Array(1400, "Value", "Value", "Style", 138)
        .Add Array(1403, "Volatility", "Volatility", "Style", 139)
        .Add Array(10103, "Vol HO", "HO Volatility Factor", "Volatility", -1)
        .Add Array(10101, "Vol KW", "KW Volatility Factor", "Volatility", -1)
        .Add Array(10102, "Vol MW", "MW Volatility Factor", "Volatility", -1)
    End With

    Set ZEUS_GetFactors_StringsCollection = tmpColl

End Function

Public Function ZEUS_GetFactor_ID(ByVal strFactor As String) As Long

    Dim collFact As Collection
    Dim arrFactor As Variant
    
    Set collFact = ZEUS_GetFactors_StringsCollection
    
    For Each arrFactor In collFact
        If arrFactor(COL_FACTOR_STRINGS_DESCRIPTION) = strFactor _
        Or arrFactor(COL_FACTOR_STRINGS_NAME) = strFactor Then
            ZEUS_GetFactor_ID = arrFactor(COL_FACTOR_STRINGS_ID)
            Exit Function
        End If
    Next arrFactor
    
    ZEUS_GetFactor_ID = -999
    
End Function

'Obtiene los últimos intLatestScenarios rendimientos desde dteCalculationDate para el factor intFactorID
Public Function ZEUS_GetFactor_LastReturns(ByVal ZeusDB_ConexionString As String, _
                                           ByVal intFactorID As Long, _
                                           ByVal intLatestScenarios As Long, _
                                           ByVal dteCalculationDate As Date, _
                                           Optional ByVal DisplayDate As Boolean = False) As Variant
    
    Dim arrRes As Variant
    Dim FactMult As Long
    
    'Si el FactorID corresponde al de una curva, el vector de rendimientos es multiplicado por -1
    'If ZEUS_GetFactorType(intFactorID) = "Curve" Then FactMult = -1 Else
    FactMult = 1
    
    'Consulta SQL
    arrRes = sqlGetQueryResults(ZeusDB_ConexionString, _
                                "SELECT D.dteDate, F.dblValue * " & FactMult & _
                                "FROM (SELECT TOP " & intLatestScenarios & " dteDate " & _
                                        "FROM tblDATA_BusinessDays " & _
                                        "WHERE dteDate <= '" & Format(dteCalculationDate, "YYYYMMDD") & "' " & _
                                        "ORDER BY dteDate DESC) D " & _
                                "LEFT JOIN (SELECT dteDate, dblValue " & _
                                            "FROM tblDATA_FactorReturns " & _
                                            "WHERE intID = " & intFactorID & ") F " & _
                                "ON D.dteDate = F.dteDate " & _
                                "ORDER BY D.dteDate ASC")
                                    
    'Si es solicitado puede desplegarse columna de fechas
    If DisplayDate Then
        ZEUS_GetFactor_LastReturns = arrRes
    Else
        ZEUS_GetFactor_LastReturns = Application.WorksheetFunction.index(arrRes, 0, UBound(arrRes, 2))
    End If

End Function

Public Function ZEUS_GetFactor_String(ByVal Valor_Input As Variant, _
                                      ByVal Col_Input_Value As colFactorStrings, _
                                      ByVal Col_Output_Value As colFactorStrings) As Variant
    Dim collFS As Collection
    Dim arrFactor As Variant
    
    Set collFS = ZEUS_GetFactors_StringsCollection
    
    For Each arrFactor In collFS
        If arrFactor(Col_Input_Value) = Valor_Input Then
            ZEUS_GetFactor_String = arrFactor(Col_Output_Value)
            Exit Function
        End If
    Next arrFactor
    
    ZEUS_GetFactor_String = -999

End Function

Public Function ZEUS_GetFactor_HistoricalValues(ByVal conexion_string_zeus As String, _
                                                ByVal intID As Long, _
                                                Optional ByVal OrderValues As String = "ASC", _
                                                Optional ByVal DisplayDates As Boolean = False) As Variant

    Dim Consulta As String
    Dim arrRes As Variant
    
    Consulta = "SELECT dteDate, dblValue FROM tblDATA_Factors WHERE intID = " & intID & " ORDER BY dteDate " & OrderValues
    arrRes = sqlGetQueryResults(conexion_string_zeus, Consulta)
    
    If DisplayDates = False Then
        arrRes = WorksheetFunction.index(arrRes, 0, UBound(arrRes, 2))
    End If
    
    ZEUS_GetFactor_HistoricalValues = arrRes
    
End Function

Public Function ZEUS_GetFactor_Value(ByVal conexion_string_zeus As String, _
                                     ByVal intID As Long, _
                                     ByVal dteDate As Date) As Variant

    Dim Consulta As String
    Dim arrRes As Variant
    
    Consulta = "SELECT dblValue FROM tblDATA_Factors WHERE intID = " & intID & " AND dteDate = '" & Format(dteDate, "YYYYMMDD") & "'"
    arrRes = sqlGetQueryResults(conexion_string_zeus, Consulta)
    
    ZEUS_GetFactor_Value = arrRes
    
End Function
