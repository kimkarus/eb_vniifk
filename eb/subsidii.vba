Sub StartCreateXMLFile()
    Dim wb As Workbook
    Dim nameSheet As String
    Set wb = ActiveWorkbook
    Dim ws_xml As Worksheet
    Set ws_xml = wb.Sheets("XML")
    ws_xml.Cells.Clear
    Call SetSheetHeadsXML(ws_xml)
    For i = 1 To wb.Sheets.Count
        nameSheet = wb.Sheets(i).name
        If nameSheet = "2020" Then
            Call FillXMLTableForYear1(nameSheet, wb, ws_xml, 2, 4, 500)
        End If
        If nameSheet = "2021" Then
            'Call FillXMLTableForYear2(nameSheet, wb, ws_xml)
        End If
        If nameSheet = "2022" Then
            'Call FillXMLTableForYear3(nameSheet, wb, ws_xml)
        End If
    Next i

    Call WriteXMLFile(ws_xml, wb)
End Sub
Sub SetSheetHeadsXML(ws_xml As Worksheet)
    'Plan
    'ws_xml.Range("AR2:BC2").Value = ws.Range("G7:R7").Value
    'ws_xml.Range("BD2:BO2").Value = ws.Range("G7:R7").Value
    'ws_xml.Range("BP2:CA2").Value = ws.Range("G7:R7").Value
    'Current
    'ws_xml.Range("H2:S2").Value = ws.Range("G8:R8").Value
    'ws_xml.Range("T2:AE2").Value = ws.Range("G8:R8").Value
    'ws_xml.Range("AF2:AQ2").Value = ws.Range("G8:R8").Value
    'Values
    'ws_xml.Range("E2:G2").Value = ws.Range("D8:F8").Value
    'KBK 1
    ws_xml.Range("A1").Value = "Kbk_code"
    'B - Наименование услуги
    ws_xml.Range("B1").Value = "Наименование_услуги"
    'C - Наименование организации
    ws_xml.Range("C1").Value = "Наименование организации"
    'D - ИНН
    ws_xml.Range("D1").Value = "Inn"
    'E - год1
    ws_xml.Range("E1").Value = "Value1"
    'E - год2
    ws_xml.Range("F1").Value = "Value2"
    'E - год3
    ws_xml.Range("G1").Value = "Value3"
    'CB - КПП
    ws_xml.Range("CB1").Value = "Kpp"
    'CC - Регистровый номер услуги
    ws_xml.Range("CC1").Value = "RegNumber"
    'CD - Наименование индикатора
    ws_xml.Range("CD1").Value = "Наименование_индикатора"
    'CE - Единица измерения
    ws_xml.Range("CE1").Value = "Единица_измерения"
    'CF - Код услуги
    ws_xml.Range("CF1").Value = "Код_услуги"
    'CG - Код ед. изм
    ws_xml.Range("CG1").Value = "Код ед. изм"
    'CH - Код индикатора
    ws_xml.Range("CH1").Value = "Код_индикатора"
    'CI - Дата начала услуги
    ws_xml.Range("CI1").Value = "Начало"
    'CJ - Дата окончания услуги
    ws_xml.Range("CJ1").Value = "Конец"
    '1
    'G - ОТ1 - H = 7
    ws_xml.Range("H1").Value = "Insrns_Pmnt_val_1"
    'H - МЗ - I = 8
    ws_xml.Range("I1").Value = "Mz_val_1"
    'I - ФР1 - J = 9
    ws_xml.Range("J1").Value = "Fr_val_1"
    'J - ИНЗ - K = 10
    ws_xml.Range("K1").Value = "Inz_val_1"
    'K - КУ - L = 11
    ws_xml.Range("L1").Value = "Ku_val_1"
    'L - СНИ - M = 12
    ws_xml.Range("M1").Value = "Sni_val_1"
    'M - СОЦДИ - N = 13
    ws_xml.Range("N1").Value = "Socdi_val_1"
    'N - ФР2 - O = 14
    ws_xml.Range("O1").Value = "Fr2_val_1"
    'O - УС - P = 15
    ws_xml.Range("P1").Value = "Us_val_1"
    'P - ТУ - Q = 16
    ws_xml.Range("Q1").Value = "Tu_val_1"
    'Q - ОТ2 - R = 17
    ws_xml.Range("R1").Value = "Othr_Pmnt_val_1"
    'R - ПНЗ - S = 18
    ws_xml.Range("S1").Value = "Pnz_val_1"
    '2
    'G - ОТ1 - T = 19
    ws_xml.Range("T1").Value = "Insrns_Pmnt_val_2"
    'H - МЗ - U = 20
    ws_xml.Range("U1").Value = "Mz_val_2"
    'I - ФР1 - V = 21
     ws_xml.Range("V1").Value = "Fr_val_2"
    'J - ИНЗ - W = 22
    ws_xml.Range("W1").Value = "Inz_val_2"
    'K - КУ - X = 23
    ws_xml.Range("X1").Value = "Ku_val_2"
    'L - СНИ - Y = 24
    ws_xml.Range("Y1").Value = "Sni_val_2"
    'M - СОЦДИ - Z = 25
    ws_xml.Range("Z1").Value = "Socdi_val_2"
    'N - ФР2 - AA = 26
    ws_xml.Range("AA1").Value = "Fr2_val_2"
    'O - УС - AB = 27
    ws_xml.Range("AB1").Value = "Us_val_2"
    'P - ТУ - AC = 28
    ws_xml.Range("AC1").Value = "Tu_val_2"
    'Q - ОТ2 - AD = 29
    ws_xml.Range("AD1").Value = "Othr_Pmnt_val_2"
    'R - ПНЗ - AE = 30
    ws_xml.Range("AE1").Value = "Pnz_val_2"
    '3
    'G - ОТ1 - AF = 31
    ws_xml.Range("AF1").Value = "Insrns_Pmnt_val_3"
    'H - МЗ - AG = 32
    ws_xml.Range("AG1").Value = "Mz_val_3"
    'I - ФР1 - AH = 33
    ws_xml.Range("AH1").Value = "Fr_val_3"
    'J - ИНЗ - AI = 34
    ws_xml.Range("AI1").Value = "Inz_val_3"
    'K - КУ - AJ = 35
    ws_xml.Range("AJ1").Value = "Ku_val_3"
    'L - СНИ - AK = 36
    ws_xml.Range("AK1").Value = "Sni_val_3"
    'M - СОЦДИ - AL = 37
    ws_xml.Range("AL1").Value = "Socdi_val_3"
    'N - ФР2 - AM = 38
    ws_xml.Range("AM1").Value = "Fr2_val_3"
    'O - УС - AN = 39
    ws_xml.Range("AN1").Value = "Us_val_3"
    'P - ТУ - AO = 40
    ws_xml.Range("AO1").Value = "Tu_val_3"
    'Q - ОТ2 - AP = 41
    ws_xml.Range("AP1").Value = "Othr_Pmnt_val_3"
    'R - ПНЗ - AQ = 42
    ws_xml.Range("AQ1").Value = "Pnz_val_3"
    'NAME
    'ws_xml.Range("B2").Value = ws.Range("A7").Value
    
    ws_xml.Range("CG1:CG65000").NumberFormat = "000"
    ws_xml.Range("CH1:CH65000").NumberFormat = "000"
    ws_xml.Range("CI1:CI65000").NumberFormat = "0000000000000000000"
End Sub
Sub FillXMLTableFromServiceTable(wb As Workbook, ws_xml, ws As Worksheet, rowStart, rowEnd As Integer)

    Dim ws_regNumbers As Worksheet
    Set ws_regNumbers = wb.Sheets("RegNumbers")
    For i = rowStart To rowEnd
        ws_xml.Cells(i, 3).Value = ws.Range("D2").Value
        ws_xml.Cells(i, 4).Value = ws.Range("N2").Value
        ws_xml.Cells(i, 80).Value = ws.Range("O2").Value
        ws_xml.Cells(i, 87).Value = ws.Range("P2").Value
        Dim RowCount As Integer
        RowCount = 0
        Dim RegNumber, IndicatorName, OkeiName, BaseCode, IndCode, OkeiCode, RegName As String
        RegNumber = ""
        IndicatorName = ""
        OkeiName = ""
        BaseCode = ""
        IndCode = ""
        OkeiCode = ""
        RegName = ""
        For Each rw In ws_regNumbers.Rows
            If Trim(CStr(ws_xml.Cells(i, 81).Value)) = Trim(CStr(rw.Cells(1, 2).Value)) Then
                'RegNumber = CStr(rw.Cells(1, 2).Value)
                RegName = CStr(rw.Cells(1, 1).Value)
                IndicatorName = CStr(rw.Cells(1, 3).Value)
                OkeiName = CStr(rw.Cells(1, 4).Value)
                OkeiCode = CStr(rw.Cells(1, 8).Value)
                BaseCode = CStr(rw.Cells(1, 5).Value)
                IndCode = CStr(rw.Cells(1, 9).Value)
                DateFrom = CStr(rw.Cells(1, 6).Value)
                DateBefore = CStr(rw.Cells(1, 7).Value)
                Exit For
            End If
            If CStr(rw.Cells(1, 1).Value) = "" Then
                Exit For
            End If
        Next rw
       ' ws_xml.Cells(i, 81).Value = RegNumber
        ws_xml.Cells(i, 2).Value = RegName
        ws_xml.Cells(i, 82).Value = IndicatorName
        ws_xml.Cells(i, 83).Value = OkeiName
        ws_xml.Cells(i, 84).Value = BaseCode
        ws_xml.Cells(i, 85).Value = OkeiCode
        ws_xml.Cells(i, 86).Value = IndCode
        ws_xml.Cells(i, 87).Value = DateFrom
        ws_xml.Cells(i, 88).Value = DateBefore
    Next i
End Sub
Sub FillXMLTableForCSP(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Обеспечение участия сборных команд Российской федерации в международных спортивных соревнованиях, Олимпийских играх.
    'На территории Российской Федерации
    '1
    'Plan
    ws_xml.Range("AR2:BC2").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD2:BO2").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP2:CA2").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H2:S2").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T2:AE2").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF2:AQ2").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC2").Value = ws.Range("V7").Value
    'Values
    ws_xml.Range("E2:G2").Value = ws.Range("D8:F8").Value
    'KBK
    ws_xml.Range("A2").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B2").Value = ws.Range("A7").Value
    
    '2
    'Plan
    ws_xml.Range("AR3:BC3").Value = ws.Range("G9:R9").Value
    ws_xml.Range("BD3:BO3").Value = ws.Range("G9:R9").Value
    ws_xml.Range("BP3:CA3").Value = ws.Range("G9:R9").Value
    'Current
    ws_xml.Range("H3:S3").Value = ws.Range("G10:R10").Value
    ws_xml.Range("T3:AE3").Value = ws.Range("G10:R10").Value
    ws_xml.Range("AF3:AQ3").Value = ws.Range("G10:R10").Value
    'RegNumber
    ws_xml.Range("CC3").Value = ws.Range("V10").Value
    'Values
    ws_xml.Range("E3:G3").Value = ws.Range("D10:F10").Value
    'KBK
    ws_xml.Range("A3").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B3").Value = ws.Range("A7").Value
    'RegNumber
    ws_xml.Range("CC2").Value = ws.Range("V7").Value
    'Обеспечение участия сборных команд Российской федерации в международных спортивных соревнованиях, Олимпийских играх.
    'За пределами территории Российской Федерации
    '1
    'Plan
    ws_xml.Range("AR4:BC4").Value = ws.Range("G14:R14").Value
    ws_xml.Range("BD4:BO4").Value = ws.Range("G14:R14").Value
    ws_xml.Range("BP4:CA4").Value = ws.Range("G14:R14").Value
    'Current
    ws_xml.Range("H4:S4").Value = ws.Range("G15:R15").Value
    ws_xml.Range("T4:AE4").Value = ws.Range("G15:R15").Value
    ws_xml.Range("AF4:AQ4").Value = ws.Range("G15:R15").Value
    'RegNumber
    ws_xml.Range("CC4").Value = ws.Range("V15").Value
    'Values
    ws_xml.Range("E4:G4").Value = ws.Range("D15:F15").Value
    'KBK
    ws_xml.Range("A4").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B4").Value = ws.Range("A7").Value
    '2
    'Plan
    ws_xml.Range("AR5:BC5").Value = ws.Range("G16:R16").Value
    ws_xml.Range("BD5:BO5").Value = ws.Range("G16:R16").Value
    ws_xml.Range("BP5:CA5").Value = ws.Range("G16:R16").Value
    'Current
    ws_xml.Range("H5:S5").Value = ws.Range("G17:R17").Value
    ws_xml.Range("T5:AE5").Value = ws.Range("G18:R18").Value
    ws_xml.Range("AF5:AQ5").Value = ws.Range("G19:R19").Value
    'RegNumber
    ws_xml.Range("CC5").Value = ws.Range("V19").Value
    'Values
    ws_xml.Range("E5").Value = ws.Range("D17").Value
    ws_xml.Range("F5").Value = ws.Range("E18").Value
    ws_xml.Range("G5").Value = ws.Range("F19").Value
    'KBK
    ws_xml.Range("A5").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B5").Value = ws.Range("A7").Value
    'Организация и проведение официальных спортивных мероприятий.
    'Международные, на территории Российской Федерации
    '1
    'Plan
    ws_xml.Range("AR6:BC6").Value = ws.Range("G23:R23").Value
    ws_xml.Range("BD6:BO6").Value = ws.Range("G23:R23").Value
    ws_xml.Range("BP6:CA6").Value = ws.Range("G23:R23").Value
    'Current
    ws_xml.Range("H6:S6").Value = ws.Range("G24:R24").Value
    ws_xml.Range("T6:AE6").Value = ws.Range("G24:R24").Value
    ws_xml.Range("AF6:AQ6").Value = ws.Range("G24:R24").Value
    'RegNumber
    ws_xml.Range("CC6").Value = ws.Range("V24").Value
    'Values
    ws_xml.Range("E6:G6").Value = ws.Range("D24:F24").Value
    'KBK
    ws_xml.Range("A6").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B6").Value = ws.Range("A23").Value
    'Всероссийские, на территории Российской Федерации
    '2
    'Plan
    ws_xml.Range("AR7:BC7").Value = ws.Range("G25:R25").Value
    ws_xml.Range("BD7:BO7").Value = ws.Range("G25:R25").Value
    ws_xml.Range("BP7:CA7").Value = ws.Range("G25:R25").Value
    'Current
    ws_xml.Range("H7:S7").Value = ws.Range("G26:R26").Value
    ws_xml.Range("T7:AE7").Value = ws.Range("G26:R26").Value
    ws_xml.Range("AF7:AQ7").Value = ws.Range("G26:R26").Value
    'RegNumber
    ws_xml.Range("CC7").Value = ws.Range("V26").Value
    'Values
    ws_xml.Range("E7:G7").Value = ws.Range("D26:F26").Value
    'KBK
    ws_xml.Range("A7").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B7").Value = ws.Range("A23").Value
    'Организация мероприятий по подготовке спортивных сборных команд.
    '
    '1
    'Plan
    ws_xml.Range("AR8:BC8").Value = ws.Range("G28:R28").Value
    ws_xml.Range("BD8:BO8").Value = ws.Range("G28:R28").Value
    ws_xml.Range("BP8:CA8").Value = ws.Range("G28:R28").Value
    'Current
    ws_xml.Range("H8:S8").Value = ws.Range("G29:R29").Value
    ws_xml.Range("T8:AE8").Value = ws.Range("G29:R29").Value
    ws_xml.Range("AF8:AQ8").Value = ws.Range("G29:R29").Value
    'RegNumber
    ws_xml.Range("CC8").Value = ws.Range("V29").Value
    'Values
    ws_xml.Range("E8:G8").Value = ws.Range("D29:F29").Value
    'KBK
    ws_xml.Range("A8").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B8").Value = ws.Range("A28").Value
    '2
    'Plan
    ws_xml.Range("AR9:BC9").Value = ws.Range("G30:R30").Value
    ws_xml.Range("BD9:BO9").Value = ws.Range("G30:R30").Value
    ws_xml.Range("BP9:CA9").Value = ws.Range("G30:R30").Value
    'Current
    ws_xml.Range("H9:S9").Value = ws.Range("G31:R31").Value
    ws_xml.Range("T9:AE9").Value = ws.Range("G31:R31").Value
    ws_xml.Range("AF9:AQ9").Value = ws.Range("G31:R31").Value
    'RegNumber
    ws_xml.Range("CC9").Value = ws.Range("V31").Value
    'Values
    ws_xml.Range("E9:G9").Value = ws.Range("D31:F31").Value
    'KBK
    ws_xml.Range("A9").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B9").Value = ws.Range("A28").Value
    '3
    'Plan
    ws_xml.Range("AR10:BC10").Value = ws.Range("G32:R32").Value
    ws_xml.Range("BD10:BO10").Value = ws.Range("G32:R32").Value
    ws_xml.Range("BP10:CA10").Value = ws.Range("G32:R32").Value
    'Current
    ws_xml.Range("H10:S10").Value = ws.Range("G33:R33").Value
    ws_xml.Range("T10:AE10").Value = ws.Range("G33:R33").Value
    ws_xml.Range("AF10:AQ10").Value = ws.Range("G33:R33").Value
    'RegNumber
    ws_xml.Range("CC10").Value = ws.Range("V33").Value
    'Values
    ws_xml.Range("E10").Value = ws.Range("D33").Value
    ws_xml.Range("F10").Value = ws.Range("E33").Value
    ws_xml.Range("G10").Value = ws.Range("F33").Value
    'KBK
    ws_xml.Range("A10").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B10").Value = ws.Range("A28").Value
    '4
    'Plan
    ws_xml.Range("AR11:BC11").Value = ws.Range("G36:R36").Value
    ws_xml.Range("BD11:BO11").Value = ws.Range("G36:R36").Value
    ws_xml.Range("BP11:CA11").Value = ws.Range("G36:R36").Value
    'Current
    ws_xml.Range("H11:S11").Value = ws.Range("G37:R37").Value
    ws_xml.Range("T11:AE11").Value = ws.Range("G37:R37").Value
    ws_xml.Range("AF11:AQ11").Value = ws.Range("G37:R37").Value
    'RegNumber
    ws_xml.Range("CC11").Value = ws.Range("V37").Value
    'Values
    ws_xml.Range("E11:G11").Value = ws.Range("D37:F37").Value
    'KBK
    ws_xml.Range("A11").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B11").Value = ws.Range("A28").Value
    '5
    'Plan
    ws_xml.Range("AR12:BC12").Value = ws.Range("G38:R38").Value
    ws_xml.Range("BD12:BO12").Value = ws.Range("G38:R38").Value
    ws_xml.Range("BP12:CA12").Value = ws.Range("G38:R38").Value
    'Current
    ws_xml.Range("H12:S12").Value = ws.Range("G39:R39").Value
    ws_xml.Range("T12:AE12").Value = ws.Range("G39:R39").Value
    ws_xml.Range("AF12:AQ12").Value = ws.Range("G39:R39").Value
    'RegNumber
    ws_xml.Range("CC12").Value = ws.Range("V39").Value
    'Values
    ws_xml.Range("E12:G12").Value = ws.Range("D39:F39").Value
    'KBK
    ws_xml.Range("A12").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B12").Value = ws.Range("A28").Value
    '6
    'Plan
    ws_xml.Range("AR13:BC13").Value = ws.Range("G44:R44").Value
    ws_xml.Range("BD13:BO13").Value = ws.Range("G44:R44").Value
    ws_xml.Range("BP13:CA13").Value = ws.Range("G44:R44").Value
    'Current
    ws_xml.Range("H13:S13").Value = ws.Range("G45:R45").Value
    ws_xml.Range("T13:AE13").Value = ws.Range("G45:R45").Value
    ws_xml.Range("AF13:AQ13").Value = ws.Range("G45:R45").Value
    'RegNumber
    ws_xml.Range("CC13").Value = ws.Range("V45").Value
    'Values
    ws_xml.Range("E13:G13").Value = ws.Range("D45:F45").Value
    'KBK
    ws_xml.Range("A13").Value = ws.Range("A49").Value
    'NAME
    ws_xml.Range("B13").Value = ws.Range("A28").Value
    'Организация мероприятий по научно-методическому обеспечению спортивных сборных команд.
    '0000000001100077708  30042100100000000004100103
    '1
    'Plan
    ws_xml.Range("AR14:BC14").Value = ws.Range("G59:R59").Value
    ws_xml.Range("BD14:BO14").Value = ws.Range("G59:R59").Value
    ws_xml.Range("BP14:CA14").Value = ws.Range("G59:R59").Value
    'Current
    ws_xml.Range("H14:S14").Value = ws.Range("G60:R60").Value
    ws_xml.Range("T14:AE14").Value = ws.Range("G60:R60").Value
    ws_xml.Range("AF14:AQ14").Value = ws.Range("G60:R60").Value
    'RegNumber
    ws_xml.Range("CC14").Value = ws.Range("V60").Value
    'Values
    ws_xml.Range("E14:G14").Value = ws.Range("D60:F60").Value
    'KBK
    ws_xml.Range("A14").Value = ws.Range("A61").Value
    'NAME
    ws_xml.Range("B14").Value = ws.Range("A59").Value
    'заполняем реквезиты 12 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 2, 14)
    
End Sub
Sub FillXMLTableForUSM(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Организация и проведение официальных физкультурных (физкультурно-оздоровительных) мероприятий.
    'Международные на территории Российской Федерации
    '1
    'Plan
    ws_xml.Range("AR15:BC15").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD15:BO15").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP15:CA15").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H15:S15").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T15:AE15").Value = ws.Range("G9:R9").Value
    ws_xml.Range("AF15:AQ15").Value = ws.Range("G10:R10").Value
    'RegNumber
    ws_xml.Range("CC15").Value = ws.Range("X7").Value
    'Values
    'ws_xml.Range("E15:G15").Value = ws.Range("D8:F8").Value
    ws_xml.Range("E15").Value = ws.Range("D8").Value
    ws_xml.Range("F15").Value = ws.Range("E9").Value
    ws_xml.Range("G15").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A15").Value = ws.Range("A23").Value
    'NAME
    ws_xml.Range("B15").Value = ws.Range("A7").Value
    
    '2
    'Plan
    ws_xml.Range("AR16:BC16").Value = ws.Range("G11:R11").Value
    ws_xml.Range("BD16:BO16").Value = ws.Range("G11:R11").Value
    ws_xml.Range("BP16:CA16").Value = ws.Range("G1:R11").Value
    'Current
    ws_xml.Range("H16:S16").Value = ws.Range("G12:R12").Value
    ws_xml.Range("T16:AE16").Value = ws.Range("G12:R12").Value
    ws_xml.Range("AF16:AQ16").Value = ws.Range("G12:R12").Value
    'RegNumber
    ws_xml.Range("CC16").Value = ws.Range("X12").Value
    'Values
    ws_xml.Range("E16:G16").Value = ws.Range("D12:F12").Value
    'ws_xml.Range("E16").Value = ws.Range("D8").Value
    'ws_xml.Range("F15").Value = ws.Range("E9").Value
    'ws_xml.Range("G15").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A16").Value = ws.Range("A23").Value
    'NAME
    ws_xml.Range("B16").Value = ws.Range("A7").Value
    
    '3
    'Plan
    ws_xml.Range("AR17:BC17").Value = ws.Range("G13:R13").Value
    ws_xml.Range("BD17:BO17").Value = ws.Range("G13:R13").Value
    ws_xml.Range("BP17:CA17").Value = ws.Range("G13:R13").Value
    'Current
    ws_xml.Range("H17:S17").Value = ws.Range("G14:R14").Value
    ws_xml.Range("T17:AE17").Value = ws.Range("G15:R15").Value
    ws_xml.Range("AF17:AQ17").Value = ws.Range("G16:R16").Value
    'RegNumber
    ws_xml.Range("CC17").Value = ws.Range("X13").Value
    'Values
    'ws_xml.Range("E17:G17").Value = ws.Range("D12:F12").Value
    ws_xml.Range("E17").Value = ws.Range("D14").Value
    ws_xml.Range("F17").Value = ws.Range("E15").Value
    ws_xml.Range("G17").Value = ws.Range("F16").Value
    'KBK
    ws_xml.Range("A17").Value = ws.Range("A23").Value
    'NAME
    ws_xml.Range("B17").Value = ws.Range("A7").Value

    '4
    'Plan
    ws_xml.Range("AR18:BC18").Value = ws.Range("G17:R17").Value
    ws_xml.Range("BD18:BO18").Value = ws.Range("G17:R17").Value
    ws_xml.Range("BP18:CA18").Value = ws.Range("G17:R17").Value
    'Current
    ws_xml.Range("H18:S18").Value = ws.Range("G18:R18").Value
    ws_xml.Range("T18:AE18").Value = ws.Range("G18:R18").Value
    ws_xml.Range("AF18:AQ18").Value = ws.Range("G18:R18").Value
    'RegNumber
    ws_xml.Range("CC18").Value = ws.Range("X18").Value
    'Values
    ws_xml.Range("E18:G18").Value = ws.Range("D18:F18").Value
    'ws_xml.Range("E18").Value = ws.Range("D14").Value
    'ws_xml.Range("F18").Value = ws.Range("E15").Value
    'ws_xml.Range("G18").Value = ws.Range("F16").Value
    'KBK
    ws_xml.Range("A18").Value = ws.Range("A23").Value
    'NAME
    ws_xml.Range("B18").Value = ws.Range("A7").Value
    
    'Работа 2. Организация и проведение официальных спортивных мероприятий.
    '
    '1
    'Plan
    ws_xml.Range("AR19:BC19").Value = ws.Range("G32:R32").Value
    ws_xml.Range("BD19:BO19").Value = ws.Range("G32:R32").Value
    ws_xml.Range("BP19:CA19").Value = ws.Range("G32:R32").Value
    'Current
    ws_xml.Range("H19:S19").Value = ws.Range("G33:R33").Value
    ws_xml.Range("T19:AE19").Value = ws.Range("G33:R33").Value
    ws_xml.Range("AF19:AQ19").Value = ws.Range("G33:R33").Value
    'RegNumber
    ws_xml.Range("CC19").Value = ws.Range("X32").Value
    'Values
    ws_xml.Range("E19:G19").Value = ws.Range("D33:F33").Value
    'ws_xml.Range("E19").Value = ws.Range("D8").Value
    'ws_xml.Range("F19").Value = ws.Range("E9").Value
    'ws_xml.Range("G19").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A19").Value = ws.Range("A35").Value
    'NAME
    ws_xml.Range("B19").Value = ws.Range("A32").Value
    
    'Работа 2. Организация и проведение официальных спортивных мероприятий.
    '
    '1
    'Plan
    ws_xml.Range("AR20:BC20").Value = ws.Range("G42:R42").Value
    ws_xml.Range("BD20:BO20").Value = ws.Range("G42:R42").Value
    ws_xml.Range("BP20:CA20").Value = ws.Range("G42:R42").Value
    'Current
    ws_xml.Range("H20:S20").Value = ws.Range("G43:R43").Value
    ws_xml.Range("T20:AE20").Value = ws.Range("G43:R43").Value
    ws_xml.Range("AF20:AQ20").Value = ws.Range("G43:R43").Value
    'RegNumber
    ws_xml.Range("CC20").Value = ws.Range("X42").Value
    'Values
    ws_xml.Range("E20:G20").Value = ws.Range("D43:F43").Value
    'ws_xml.Range("E19").Value = ws.Range("D8").Value
    'ws_xml.Range("F19").Value = ws.Range("E9").Value
    'ws_xml.Range("G19").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A20").Value = ws.Range("A47").Value
    'NAME
    ws_xml.Range("B20").Value = ws.Range("A42").Value
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 15, 20)
    
End Sub
Sub FillXMLTableForFCPSR(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR21:BC21").Value = ws.Range("G6:R6").Value
    ws_xml.Range("BD21:BO21").Value = ws.Range("G6:R6").Value
    ws_xml.Range("BP21:CA21").Value = ws.Range("G6:R6").Value
    'Current
    ws_xml.Range("H21:S21").Value = ws.Range("G7:R7").Value
    ws_xml.Range("T21:AE21").Value = ws.Range("G7:R7").Value
    ws_xml.Range("AF21:AQ21").Value = ws.Range("G7:R7").Value
    'RegNumber
    ws_xml.Range("CC21").Value = ws.Range("V6").Value
    'Values
    ws_xml.Range("E21:G21").Value = ws.Range("D7:F7").Value
    'ws_xml.Range("E21").Value = ws.Range("D8").Value
    'ws_xml.Range("F21").Value = ws.Range("E9").Value
    'ws_xml.Range("G21").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A21").Value = ws.Range("A13").Value
    'NAME
    ws_xml.Range("B21").Value = ws.Range("A6").Value
    
    'Организация и проведение официальных спортивных мероприятий.
    '
    '1
    'Plan
    ws_xml.Range("AR22:BC22").Value = ws.Range("G8:R8").Value
    ws_xml.Range("BD22:BO22").Value = ws.Range("G8:R8").Value
    ws_xml.Range("BP22:CA22").Value = ws.Range("G8:R8").Value
    'Current
    ws_xml.Range("H22:S22").Value = ws.Range("G9:R9").Value
    ws_xml.Range("T22:AE22").Value = ws.Range("G9:R9").Value
    ws_xml.Range("AF22:AQ22").Value = ws.Range("G9:R9").Value
    'RegNumber
    ws_xml.Range("CC22").Value = ws.Range("V8").Value
    'Values
    ws_xml.Range("E22:G22").Value = ws.Range("D7:F7").Value
    'ws_xml.Range("E22").Value = ws.Range("D8").Value
    'ws_xml.Range("F22").Value = ws.Range("E9").Value
    'ws_xml.Range("G22").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A22").Value = ws.Range("A13").Value
    'NAME
    ws_xml.Range("B22").Value = ws.Range("A6").Value
    
    'Организация и обеспечение экспериментальной и инновационной деятельности в области физкультуры и спорта.
    '
    '1
    'Plan
    ws_xml.Range("AR23:BC23").Value = ws.Range("G10:R10").Value
    ws_xml.Range("BD23:BO23").Value = ws.Range("G10:R10").Value
    ws_xml.Range("BP23:CA23").Value = ws.Range("G10:R10").Value
    'Current
    ws_xml.Range("H23:S23").Value = ws.Range("G11:R11").Value
    ws_xml.Range("T23:AE23").Value = ws.Range("G11:R11").Value
    ws_xml.Range("AF23:AQ23").Value = ws.Range("G11:R11").Value
    'RegNumber
    ws_xml.Range("CC23").Value = ws.Range("V10").Value
    'Values
    ws_xml.Range("E23:G23").Value = ws.Range("D11:F11").Value
    'ws_xml.Range("E23").Value = ws.Range("D8").Value
    'ws_xml.Range("F23").Value = ws.Range("E9").Value
    'ws_xml.Range("G23").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A23").Value = ws.Range("A13").Value
    'NAME
    ws_xml.Range("B23").Value = ws.Range("A6").Value
    
    'Работа 4. Организация и проведение официальных физкультурных (физкультурно-оздоровительных) мероприятий.
    '
    '1
    'Plan
    ws_xml.Range("AR24:BC24").Value = ws.Range("G20:R20").Value
    ws_xml.Range("BD24:BO24").Value = ws.Range("G20:R20").Value
    ws_xml.Range("BP24:CA24").Value = ws.Range("G20:R20").Value
    'Current
    ws_xml.Range("H24:S24").Value = ws.Range("G21:R21").Value
    ws_xml.Range("T24:AE24").Value = ws.Range("G21:R21").Value
    ws_xml.Range("AF24:AQ24").Value = ws.Range("G21:R21").Value
    'RegNumber
    ws_xml.Range("CC24").Value = ws.Range("V20").Value
    'Values
    ws_xml.Range("E24:G24").Value = ws.Range("D21:F21").Value
    'ws_xml.Range("E24").Value = ws.Range("D8").Value
    'ws_xml.Range("F24").Value = ws.Range("E9").Value
    'ws_xml.Range("G24").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A24").Value = ws.Range("A23").Value
    'NAME
    ws_xml.Range("B24").Value = ws.Range("A20").Value
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 21, 24)
    
End Sub
Sub FillXMLTableForUGSPORT(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR25:BC25").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD25:BO25").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP25:CA25").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H25:S25").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T25:AE25").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF25:AQ25").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC25").Value = ws.Range("V7").Value
    'Values
    'ws_xml.Range("E25:G25").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E25").Value = ws.Range("D8").Value
    ws_xml.Range("F25").Value = ws.Range("E9").Value
    ws_xml.Range("G25").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A25").Value = ws.Range("A12").Value
    'NAME
    ws_xml.Range("B25").Value = ws.Range("A7").Value
    
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 25, 25)
    
End Sub
Sub FillXMLTableForOZEROKRUG(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR26:BC26").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD26:BO26").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP26:CA26").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H26:S26").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T26:AE26").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF26:AQ26").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC26").Value = ws.Range("V7").Value
    'Values
    'ws_xml.Range("E26:G26").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E26").Value = ws.Range("D8").Value
    ws_xml.Range("F26").Value = ws.Range("E9").Value
    ws_xml.Range("G26").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A26").Value = ws.Range("A12").Value
    'NAME
    ws_xml.Range("B26").Value = ws.Range("A7").Value
    
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 26, 26)
    
End Sub
Sub FillXMLTableForNOVOGORSK(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR27:BC27").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD27:BO27").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP27:CA27").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H27:S27").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T27:AE27").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF27:AQ27").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC27").Value = ws.Range("V7").Value
    'Values
    'ws_xml.Range("E27:G27").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E27").Value = ws.Range("D8").Value
    ws_xml.Range("F27").Value = ws.Range("E9").Value
    ws_xml.Range("G27").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A27").Value = ws.Range("A12").Value
    'NAME
    ws_xml.Range("B27").Value = ws.Range("A7").Value
    
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 27, 27)
    
End Sub

Sub FillXMLTableForOKA(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR28:BC28").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD28:BO28").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP28:CA28").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H28:S28").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T28:AE28").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF28:AQ28").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC28").Value = ws.Range("V7").Value
    'Values
    'ws_xml.Range("E27:G27").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E28").Value = ws.Range("D8").Value
    ws_xml.Range("F28").Value = ws.Range("E9").Value
    ws_xml.Range("G28").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A28").Value = ws.Range("A12").Value
    'NAME
    ws_xml.Range("B28").Value = ws.Range("A7").Value
    
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 28, 28)
    
End Sub
Sub FillXMLTableForKRIM(nameSheet As String, wb As Workbook, ws_xml As Worksheet)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    '   A = kbk
    '   E:G = Values - D-F
    '   H:S = Current Values1 - G:R
    '   T:AE = Current Values2 - G:R
    '   AF:AQ = Current Values3 - G:R
    '   AR:BC = Plan Values1 - G:R
    '   BD:BO = Plan Values2 - G:R
    '   BP:CA = Plan Values3 - G:R
    'Работа 1. Организация и обеспечение координации деятельности физкультурно-спортивных организаций по подготовке спортивного резерва.
    '
    '1
    'Plan
    ws_xml.Range("AR29:BC29").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BD29:BO29").Value = ws.Range("G7:R7").Value
    ws_xml.Range("BP29:CA29").Value = ws.Range("G7:R7").Value
    'Current
    ws_xml.Range("H29:S29").Value = ws.Range("G8:R8").Value
    ws_xml.Range("T29:AE29").Value = ws.Range("G8:R8").Value
    ws_xml.Range("AF29:AQ29").Value = ws.Range("G8:R8").Value
    'RegNumber
    ws_xml.Range("CC29").Value = ws.Range("V7").Value
    'Values
    'ws_xml.Range("E29:G29").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E29").Value = ws.Range("D8").Value
    ws_xml.Range("F29").Value = ws.Range("E9").Value
    ws_xml.Range("G29").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A29").Value = ws.Range("A12").Value
    'NAME
    ws_xml.Range("B29").Value = ws.Range("A7").Value
    
    
    'заполняем реквезиты 6 строк
    Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 29, 29)
    
End Sub

Sub WriteXMLFile(ws_xml As Worksheet, wb As Workbook)
    
    Call WriteXMLFileForNormativ(ws_xml, wb)
    Call WriteXMLFileForOFO(ws_xml, wb)
End Sub
Sub WriteXMLFileForNormativ(ws_xml As Worksheet, wb As Workbook)
    
    Dim SortRange As String
    SortRange = ws_xml.Cells(Cells.Rows.Count, "B").End(xlUp).Row
    With ws_xml.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B2:B" & SortRange), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("CD2:CD" & SortRange), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        '
        .SetRange Range("A2:CJ" & SortRange)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim RowCount, CountNormativ As Integer
    Dim rw As Range
    RowCount = 0
    CountNormativ = 0
    Dim base_name_prev, base_name_cur, okei_code_prev, okei_code_cur, base_code_prev, base_code_cur, ind_name_prev, ind_name_cur As String
    Dim isTheSame As Boolean
    isTheSame = False
    base_name_prev = ""
    base_code_prev = ""
    okei_code_prev = ""
    ind_name_prev = ""
    For Each rw In ws_xml.Rows

        If RowCount > 0 Then
            base_name_cur = CStr(rw.Cells(1, 2).Value)
            okei_code_cur = CStr(Format(rw.Cells(1, 85), "000"))
            base_code_cur = CStr(rw.Cells(1, 84))
            ind_name_cur = CStr(rw.Cells(1, 82))
                If base_name_cur <> base_name_prev Or okei_code_cur <> okei_code_prev Or ind_name_prev <> ind_name_cur Then
                    If RowCount > 1 Then
                        CountNormativ = CountNormativ + 1
                        Dim path_xml As String
                        'path_xml = CStr(GetFilePath(wb.path, wb.Name, CStr(CountNormativ) & " Нормативы", CStr(base_name_prev)))
                        path_xml = CStr(GetFilePath(wb.path, wb.name, CStr(CountNormativ) & "_Нормативы", CStr(base_code_prev) & "_" & CStr(ind_name_prev)))
                        XDoc.Save path_xml
                    End If
                    If base_name_cur = "" Then
                        Exit For
                    End If
                    Set XDoc = CreateObject("MSXML2.DOMDocument")
                    XDoc.appendChild XDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'")
                    isTheSame = False
                    If ws_xml.Cells(rw.Row, 1).Value <> "" Then
                        base_name_prev = base_name_cur
                        okei_code_prev = okei_code_cur
                        base_code_prev = base_code_cur
                        ind_name_prev = ind_name_cur
                    End If
                Else
                    isTheSame = True
                End If
                If Not isTheSame Then
                    Set root = XDoc.createElement("BNZ_INF_GRBS")
                    XDoc.appendChild root
            
                    Set Base_code = XDoc.createElement("Base_code")
                    root.appendChild Base_code
                    Base_code.Text = CStr(rw.Cells(1, 84))
                    
                    Set Base_name = XDoc.createElement("Base_name")
                    root.appendChild Base_name
                    Base_name.Text = base_name_cur

                    Set Volume_indicator_code = XDoc.createElement("Volume_indicator_code")
                    root.appendChild Volume_indicator_code
                    Volume_indicator_code.Text = CStr(Format(rw.Cells(1, 86), "000"))
                    
                    Set Volume_indicator_name = XDoc.createElement("Volume_indicator_name")
                    root.appendChild Volume_indicator_name
                    Volume_indicator_name.Text = CStr(rw.Cells(1, 82))
                    
                    Set Volume_indicator_okei_code = XDoc.createElement("Volume_indicator_okei_code")
                    root.appendChild Volume_indicator_okei_code
                    Volume_indicator_okei_code.Text = CStr(Format(rw.Cells(1, 85), "000"))
                    
                    Set Volume_indicator_okei_name = XDoc.createElement("Volume_indicator_okei_name")
                    root.appendChild Volume_indicator_okei_name
                    Volume_indicator_okei_name.Text = CStr(rw.Cells(1, 83))
                    
                    Set EffectiveFrom = XDoc.createElement("EffectiveFrom")
                    root.appendChild EffectiveFrom
                    If CStr(rw.Cells(1, 84)) = "" Then
                            EffectiveFrom.Text = "22.07.2015"
                        Else
                            EffectiveFrom.Text = CStr(rw.Cells(1, 87))
                    End If
                    
                    Set EffectiveBefore = XDoc.createElement("EffectiveBefore")
                    root.appendChild EffectiveBefore
                    
                    If CStr(rw.Cells(1, 84)) = "" Then
                            EffectiveBefore.Text = "31.12.2099"
                        Else
                            EffectiveBefore.Text = CStr(rw.Cells(1, 88))
                    End If

                    Set Inst_code_oiv = XDoc.createElement("Inst_code_grbs")
                    root.appendChild Inst_code_oiv
                    Inst_code_oiv.Text = CStr(Format(rw.Cells(1, 87), "0000000000000000000"))
                    
                    Set Inst_name_oiv = XDoc.createElement("Inst_name_grbs")
                    root.appendChild Inst_name_oiv
                    Inst_name_oiv.Text = "МИНИСТЕРСТВО СПОРТА РОССИЙСКОЙ ФЕДЕРАЦИИ"
                    
                    Set Inst_inn = XDoc.createElement("Inst_inn")
                    root.appendChild Inst_inn
                    Inst_inn.Text = "7703771271"
                    
                    Set Inst_kpp = XDoc.createElement("Inst_kpp")
                    root.appendChild Inst_kpp
                    Inst_kpp.Text = "770901001"
                    
                    Set Registry_records = XDoc.createElement("Registry_records")
                    root.appendChild Registry_records
                    
                    'Set Dprtm_values = XDoc.createElement("Dprtm_values")
                    'root.appendChild Dprtm_values
                End If

                Set Registry_record = XDoc.createElement("Registry_record")
                Registry_records.appendChild Registry_record
                
                
                Set RegNumber = XDoc.createElement("RegNumber")
                Registry_record.appendChild RegNumber
  
                Dim strRegNumber, newStrRegNumber As String
                strRegNumber = RemoveWhiteSpace(CStr(rw.Cells(1, 81)))
                newStrRegNumber = getStringBeforeSpace(strRegNumber, "_")
                RegNumber.Text = newStrRegNumber
                
                'Учреждения
                Set Dprtm_records = XDoc.createElement("Dprtm_records")
                Registry_record.appendChild Dprtm_records
                
                'Запись учреждения
                Set Dprtm_record = XDoc.createElement("Dprtm_record")
                Dprtm_records.appendChild Dprtm_record
                
                Set Dprtm_code = XDoc.createElement("Dprtm_code")
                Dprtm_record.appendChild Dprtm_code
                Dprtm_code.Text = ""
                
                Set Dprtm_name = XDoc.createElement("Dprtm_name")
                Dprtm_record.appendChild Dprtm_name
                Dprtm_name.Text = CStr(rw.Cells(1, 3))
                
                Set Dprtm_inn = XDoc.createElement("Dprtm_inn")
                Dprtm_record.appendChild Dprtm_inn
                Dprtm_inn.Text = CStr(rw.Cells(1, 4))
                
                Set Dprtm_kpp = XDoc.createElement("Dprtm_kpp")
                Dprtm_record.appendChild Dprtm_kpp
                Dprtm_kpp.Text = CStr(rw.Cells(1, 80))
                
                'Значения учреждения
                Set Dprtm_values = XDoc.createElement("Dprtm_values")
                Dprtm_record.appendChild Dprtm_values
                
                'Set Bnz_avrg_pmnt = XDoc.createElement("Bnz_avrg_pmnt")
                'Dprtm_record.appendChild Bnz_avrg_pmnt
                
                Dim rel As Object
                'Insrns_Pmnt_val
                Set Insrns_Pmnt = XDoc.createElement("Insrns_Pmnt_Ot1_dprtm")
                Dprtm_values.appendChild Insrns_Pmnt
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 8))
                Insrns_Pmnt.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 20))
                Insrns_Pmnt.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 32))
                Insrns_Pmnt.setAttributeNode rel
                
                'Mz_val
                Set Mz = XDoc.createElement("Mz_dprtm")
                Dprtm_values.appendChild Mz
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 9))
                Mz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 21))
                Mz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 33))
                Mz.setAttributeNode rel
                
                'Fr_val
                Set Fr = XDoc.createElement("Fr_dprtm")
                Dprtm_values.appendChild Fr
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 10))
                Fr.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 22))
                Fr.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 34))
                Fr.setAttributeNode rel
                
                'Inz_val
                Set Inz = XDoc.createElement("Inz_dprtm")
                Dprtm_values.appendChild Inz
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 11))
                Inz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 23))
                Inz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 35))
                Inz.setAttributeNode rel
                
                'Ku_val
                Set Ku = XDoc.createElement("Ku_dprtm")
                Dprtm_values.appendChild Ku
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 12))
                Ku.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 24))
                Ku.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 36))
                Ku.setAttributeNode rel
                
                'Sni_val
                Set Sni = XDoc.createElement("Sni_dprtm")
                Dprtm_values.appendChild Sni
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 13))
                Sni.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 25))
                Sni.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 37))
                Sni.setAttributeNode rel
                
                'Socdi_val
                Set Socdi = XDoc.createElement("Socdi_dprtm")
                Dprtm_values.appendChild Socdi
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 14))
                Socdi.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 26))
                Socdi.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 38))
                Socdi.setAttributeNode rel
                
                'Fr2_val
        '        Set rel = XDoc.createAttribute("val_1")
        '        rel.NodeValue = CStr(rw.Cells(1, 11))
        '        Insrns_Pmnt.setAttributeNode rel
        '        Set rel = XDoc.createAttribute("val_2")
        '        rel.NodeValue = CStr(rw.Cells(1, 23))
        '        Insrns_Pmnt.setAttributeNode rel
        '        Set rel = XDoc.createAttribute("val_3")
        '        rel.NodeValue = CStr(rw.Cells(1, 35))
        '        Insrns_Pmnt.setAttributeNode rel
                
                'Us_val
                Set Us = XDoc.createElement("Us_dprtm")
                Dprtm_values.appendChild Us
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 16))
                Us.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 28))
                Us.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 40))
                Us.setAttributeNode rel
                
                'Tu_val
                Set Tu = XDoc.createElement("Tu_dprtm")
                Dprtm_values.appendChild Tu
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 17))
                Tu.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 29))
                Tu.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 41))
                Tu.setAttributeNode rel
                
                'Othr_Pmnt_val
                Set Othr_Pmnt = XDoc.createElement("Othr_Pmnt_Ot2_dprtm")
                Dprtm_values.appendChild Othr_Pmnt
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 18))
                Othr_Pmnt.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 30))
                Othr_Pmnt.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 42))
                Othr_Pmnt.setAttributeNode rel
                
                'Pnz_val
                Set Pnz = XDoc.createElement("Pnz_dprtm")
                Dprtm_values.appendChild Pnz
                Set rel = XDoc.createAttribute("val_1")
                rel.NodeValue = CStr(rw.Cells(1, 19))
                Pnz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_2")
                rel.NodeValue = CStr(rw.Cells(1, 31))
                Pnz.setAttributeNode rel
                Set rel = XDoc.createAttribute("val_3")
                rel.NodeValue = CStr(rw.Cells(1, 43))
                Pnz.setAttributeNode rel
                
                Set Kbk_codes = XDoc.createElement("Kbk_codes")
                Dprtm_values.appendChild Kbk_codes
                
                Set Kbk_code = XDoc.createElement("Kbk_code")
                Kbk_codes.appendChild Kbk_code
                
                Kbk_code.Text = RemoveWhiteSpace(CStr(rw.Cells(1, 1)))
        End If
        RowCount = RowCount + 1
    Next rw
    
    
'    Dim rel As Object
'    Set rel = XDoc.createAttribute("Attrib")
'    rel.NodeValue = "Attrib value"
'    elem.setAttributeNode rel
        
    
End Sub

Sub WriteXMLFileForOFO(ws_xml As Worksheet, wb As Workbook)
    
        
    Dim SortRange As String

    SortRange = ws_xml.Cells(Cells.Rows.Count, "A").End(xlUp).Row

    With ws_xml.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("C2:C" & SortRange), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("CD2:CD" & SortRange), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        '
        .SetRange Range("A2:CJ" & SortRange)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Dim RowCount, CountNormativ As Integer
    Dim rw As Range
    RowCount = 0
    Dim dep_name_prev, dep_name_cur, okei_code_prev, okei_code_cur, base_code_prev, base_code_cur, ind_name_prev, ind_name_cur As String
    
    Dim strRegNumber, newStrRegNumber As String
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.appendChild XDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'")
    
    Set root = XDoc.createElement("OFO_INF")
    XDoc.appendChild root
                    
    Set Inst_name_oiv = XDoc.createElement("Inst_name")
    root.appendChild Inst_name_oiv
    Inst_name_oiv.Text = "МИНИСТЕРСТВО СПОРТА РОССИЙСКОЙ ФЕДЕРАЦИИ"
                    
    Set Inst_inn = XDoc.createElement("Inst_inn")
    root.appendChild Inst_inn
    Inst_inn.Text = "7703771271"
                    
    Set Inst_kpp = XDoc.createElement("Inst_kpp")
    root.appendChild Inst_kpp
    Inst_kpp.Text = "770901001"
    
    Set Dprtm_records = XDoc.createElement("Dprtm_records")
    root.appendChild Dprtm_records
            
    
                    
    For Each rw In ws_xml.Rows
        If RowCount > 0 Then

            If CStr(ws_xml.Cells(rw.Row, 1).Value) = "" Then
                If RowCount > 1 Then
                    Dim path_xml As String
                    'path_xml = CStr(GetFilePath(wb.path, wb.Name, CStr(CountNormativ) & " Нормативы", CStr(base_name_prev)))
                    path_xml = CStr(GetFilePath(wb.path, wb.name, "0" & "_ОФО", "" & "" & ""))
                    XDoc.Save path_xml
                End If
                If CStr(ws_xml.Cells(rw.Row, 1).Value) = "" Then
                    Exit For
                End If

            End If
            
            dep_name_cur = CStr(rw.Cells(1, 3).Value)
            okei_code_cur = CStr(Format(rw.Cells(1, 85), "000"))
            base_code_cur = CStr(rw.Cells(1, 84))
            ind_name_cur = CStr(rw.Cells(1, 82))
            If dep_name_cur <> dep_name_prev Then
                isTheSame = False
                If ws_xml.Cells(rw.Row, 1).Value <> "" Then
                    dep_name_prev = dep_name_cur
                    okei_code_prev = okei_code_cur
                    base_code_prev = base_code_cur
                    ind_name_prev = ind_name_cur
                End If
            Else
                isTheSame = True
            End If
            If Not isTheSame Then
                Set Dprtm_record = XDoc.createElement("Dprtm_record")
                Dprtm_records.appendChild Dprtm_record
                
                Set Dprtm_code = XDoc.createElement("Dprtm_code")
                Dprtm_record.appendChild Dprtm_code
                Dprtm_code.Text = ""
                
                Set Dprtm_name = XDoc.createElement("Dprtm_name")
                Dprtm_record.appendChild Dprtm_name
                Dprtm_name.Text = CStr(rw.Cells(1, 3))
                
                Set Dprtm_inn = XDoc.createElement("Dprtm_inn")
                Dprtm_record.appendChild Dprtm_inn
                Dprtm_inn.Text = CStr(rw.Cells(1, 4))
                
                Set Dprtm_kpp = XDoc.createElement("Dprtm_kpp")
                Dprtm_record.appendChild Dprtm_kpp
                Dprtm_kpp.Text = CStr(rw.Cells(1, 80))
                
                Set Srvc_records = XDoc.createElement("Srvc_records")
                Dprtm_record.appendChild Srvc_records
                                
            End If
            
            Set Srvc_record = XDoc.createElement("Srvc_record")
            Srvc_records.appendChild Srvc_record
                
            'RegNumber
            Set RegNumber = XDoc.createElement("RegNumber")
            Srvc_record.appendChild RegNumber
                
            strRegNumber = RemoveWhiteSpace(CStr(rw.Cells(1, 81)))
            newStrRegNumber = getStringBeforeSpace(strRegNumber, "_")
            RegNumber.Text = newStrRegNumber
                
            'Vlm_indctr_records
            Set Vlm_indctr_records = XDoc.createElement("Vlm_indctr_records")
            Srvc_record.appendChild Vlm_indctr_records
            'Vlm_indctr_record
            Set Vlm_indctr_record = XDoc.createElement("Vlm_indctr_record")
            Vlm_indctr_records.appendChild Vlm_indctr_record
            'Vlm_indctr_code
            Set Vlm_indctr_code = XDoc.createElement("Vlm_indctr_code")
            Vlm_indctr_record.appendChild Vlm_indctr_code
            Vlm_indctr_code.Text = CStr(Format(rw.Cells(1, 86), "000"))
            'Vlm_indctr_name
            Set Vlm_indctr_name = XDoc.createElement("Vlm_indctr_name")
            Vlm_indctr_record.appendChild Vlm_indctr_name
            Vlm_indctr_name.Text = CStr(rw.Cells(1, 82))
            'Vlm_indctr_name_1
            Set Vlm_indctr_name_1 = XDoc.createElement("Vlm_indctr_name_1")
            Vlm_indctr_record.appendChild Vlm_indctr_name_1
            Vlm_indctr_name_1.Text = FindNameOKEI(CStr(rw.Cells(1, 85)), wb)
            'Value_1
            Set Value_1 = XDoc.createElement("Value_1")
            Vlm_indctr_record.appendChild Value_1
            Value_1.Text = CStr(rw.Cells(1, 5))
            'Value_2
            Set Value_2 = XDoc.createElement("Value_2")
            Vlm_indctr_record.appendChild Value_2
            Value_2.Text = CStr(rw.Cells(1, 6))
            'Value_3
            Set Value_3 = XDoc.createElement("Value_3")
            Vlm_indctr_record.appendChild Value_3
            Value_3.Text = CStr(rw.Cells(1, 7))
            'Kbk_code
            Set Kbk_code = XDoc.createElement("Kbk_code")
            Vlm_indctr_record.appendChild Kbk_code
            Kbk_code.Text = RemoveWhiteSpace(CStr(rw.Cells(1, 1)))
                 
        End If
        RowCount = RowCount + 1
    Next rw
    
    
'    Dim rel As Object
'    Set rel = XDoc.createAttribute("Attrib")
'    rel.NodeValue = "Attrib value"
'    elem.setAttributeNode rel
End Sub
Public Function FindNameOKEI(code As String, wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = wb.Sheets("OKEI")
    Dim RowCount As Integer
    RowCount = 0
    Dim name As String
    name = ""
    For Each rw In ws.Rows
        If Trim(CStr(rw.Cells(1, 2).Value)) = Trim(CStr(code)) Then
            name = CStr(rw.Cells(1, 1).Value)
            Exit For
        End If
        If CStr(rw.Cells(1, 1).Value) = "" Then
            Exit For
        End If
    Next rw

    FindNameOKEI = name
End Function
Public Function RemoveWhiteSpace(target As String) As String
    RemoveWhiteSpace = Replace(target, " ", "")
'    With New RegExp
'        .Pattern = "\s"
'        .MultiLine = True
'        .Global = True
'        RemoveWhiteSpace = .Replace(target, vbNullString)
'    End With
End Function
Public Function GetFilePath(folder_path, Filename, TagStart, TagEnd As String) As String

    If InStr(Filename, ".") > 0 Then
        Filename = Left(Filename, InStr(Filename, ".") - 1)
    End If
    GetFilePath = folder_path & "\" & TagStart & "_" & Filename & "_" & TagEnd & ".xml"
    
End Function
Function getStringBeforeSpace(inputString, inputChar As String)
    Dim Index As Integer
    Index = InStr(1, inputString, inputChar) - 1
    If Index > 0 Then
        getStringBeforeSpace = Mid(inputString, 1, Index)
    Else
        getStringBeforeSpace = inputString
    End If
End Function
Sub FillXMLTableForYear1(nameSheet As String, wb As Workbook, ws_xml As Worksheet, StartRowCountD, StartRowCountS, EndRowCount As Integer)
    Dim ws As Worksheet
    Set ws = wb.Sheets(nameSheet)
    Dim RowCount As Integer
    RowCount = 0
    For i = StartRowCountS To EndRowCount
        RowCount = RowCount + 1
        If RowCount >= StartRowCountS Then
            If CStr(ws.Cells(i, 3).Value) = "" Then
                Exit For
            End If
            'ФР = AVA = 1249 - 1 / val1 = J = 10
            ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ИНЗ = AVB = 1250 - 2 / val1 = K = 11
            ws_xml.Cells(StartRowCountD, 11).Value = ws.Cells(i, 1250).Value
            'уч литература = AVC = 1251 - 3 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'практика = AVD = 1252 - 4 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'повыш квалиф = AVE = 1253 - 5 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'медосмотры = AVF = 1254 - 6 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'КУ = AVG = 1255 - 7 / val1 = L = 12
            ws_xml.Cells(StartRowCountD, 12).Value = ws.Cells(i, 1255).Value
            'СНИ = AVH = 1256 - 8 / val1 = M = 13
            ws_xml.Cells(StartRowCountD, 13).Value = ws.Cells(i, 1256).Value
            'СОЦДИ = AVI = 1257 - 9 / val1 = N = 14
            ws_xml.Cells(StartRowCountD, 14).Value = ws.Cells(i, 1257).Value
            'ФР2 = AVJ = 1258 - 10 / val1 = O = 15
            ws_xml.Cells(StartRowCountD, 15).Value = ws.Cells(i, 1258).Value
            'УС = AVK = 1259 - 11 / val1 = P = 16
            ws_xml.Cells(StartRowCountD, 16).Value = ws.Cells(i, 1259).Value
            'ТУ = AVL = 1260 - 12 / val1 = Q = 17
            ws_xml.Cells(StartRowCountD, 17).Value = ws.Cells(i, 1260).Value
            'ОТ 2 = AVM = 1261 - 13 / val1 = R = 18
            ws_xml.Cells(StartRowCountD, 18).Value = ws.Cells(i, 1261).Value
            'ОТ2 зп = AVN = 1262 - 14 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'начисления = AVO = 1263 - 15 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ПНЗ = AVP = 1264 - 16 / val1 = S = 19
            ws_xml.Cells(StartRowCountD, 19).Value = ws.Cells(i, 1264).Value
            'Values
            'Value1 = P = 16 / val1 = E = 5
            ws_xml.Cells(StartRowCountD, 5).Value = ws.Cells(i, 16).Value
            'Value2 = Q = 17 / val1 = F = 6
            'Value2 = R = 18 / val1 = G = 7
            
            'Org
            Dim org As Variant
            org = FindOrg(CStr(ws.Cells(i, 3).Value), wb, 2, 50)
            ws_xml.Cells(StartRowCountD, 3).Value = org(0)
            ws_xml.Cells(StartRowCountD, 4).Value = org(1)
            ws_xml.Cells(StartRowCountD, 80).Value = org(2)
            
            'RegNumber
            Dim reg As Variant
            reg = FindRegNumber(CStr(ws.Cells(i, 13).Value), wb, 2, 1000)
            'Наименование услуги
            ws_xml.Cells(StartRowCountD, 2).Value = reg(0)
            'Номер
            ws_xml.Cells(StartRowCountD, 81).Value = CStr(ws.Cells(i, 13).Value)
            'Наименование объема
            ws_xml.Cells(StartRowCountD, 82).Value = reg(2)
            'Еденица измерения объема
            ws_xml.Cells(StartRowCountD, 83).Value = reg(3)
            'Базовый код
            ws_xml.Cells(StartRowCountD, 84).Value = reg(4)
            'Начало работ
            ws_xml.Cells(StartRowCountD, 87).Value = reg(5)
            'Конец работ
            ws_xml.Cells(StartRowCountD, 88).Value = reg(6)
            'Код ОКЕИ
            ws_xml.Cells(StartRowCountD, 85).Value = reg(7)
            'Код Индикатора
            ws_xml.Cells(StartRowCountD, 86).Value = reg(8)
            
            'Кбк
            If ws_xml.Cells(StartRowCountD, 1).Value = "" Then
                ws_xml.Cells(StartRowCountD, 1).Value = "0"
            End If
            If ws_xml.Cells(StartRowCountD, 82).Value = "" Then
                ws_xml.Cells(StartRowCountD, 82).Value = "0"
            End If
            If ws_xml.Cells(StartRowCountD, 2).Value = "" Then
                ws_xml.Cells(StartRowCountD, 2).Value = "0"
            End If
            
            StartRowCountD = StartRowCountD + 1
        End If
        

        
    Next i
    'заполняем реквезиты 12 строк
    'Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 2, 14)
    
End Sub
Public Function FindOrg(shortName As String, wb As Workbook, StartRowCountS, EndRowCount As Integer) As String()
    Dim ws As Worksheet
    Set ws = wb.Sheets("Учр")
    Dim returnVal(3) As String
    Index = 0
    For i = StartRowCountS To EndRowCount
        If UCase(Trim(CStr(ws.Cells(i, 7).Value))) = UCase(Trim(CStr(shortName))) Then
            Index = i
            Exit For
        End If
        If CStr(ws.Cells(i, 3).Value) = "" Then
            Exit For
        End If
    Next i
    If Index > 0 Then
        returnVal(0) = CStr(ws.Cells(Index, 3).Value)
        returnVal(1) = CStr(ws.Cells(Index, 5).Value)
        returnVal(2) = CStr(ws.Cells(Index, 6).Value)
    Else
        returnVal(0) = ""
        returnVal(1) = ""
        returnVal(2) = ""
    End If
    FindOrg = returnVal
End Function
Public Function FindRegNumber(name As String, wb As Workbook, StartRowCountS, EndRowCount As Integer) As String()
    Dim ws As Worksheet
    Set ws = wb.Sheets("RegNumbers")
    Dim returnVal(9) As String
    Index = 0
    For i = StartRowCountS To EndRowCount
        If UCase(Trim(CStr(ws.Cells(i, 2).Value))) = UCase(Trim(CStr(name))) Then
            Index = i
            Exit For
        End If
        If CStr(ws.Cells(i, 2).Value) = "" Then
            Exit For
        End If
    Next i
    If Index > 0 Then
        returnVal(0) = CStr(ws.Cells(Index, 1).Value)
        returnVal(1) = CStr(ws.Cells(Index, 2).Value)
        returnVal(2) = CStr(ws.Cells(Index, 3).Value)
        returnVal(3) = CStr(ws.Cells(Index, 4).Value)
        returnVal(4) = CStr(ws.Cells(Index, 5).Value)
        returnVal(5) = CStr(ws.Cells(Index, 6).Value)
        returnVal(6) = CStr(ws.Cells(Index, 7).Value)
        returnVal(7) = CStr(ws.Cells(Index, 8).Value)
        returnVal(8) = CStr(ws.Cells(Index, 9).Value)
    Else
        returnVal(0) = ""
        returnVal(1) = ""
        returnVal(2) = ""
        returnVal(3) = ""
        returnVal(4) = ""
        returnVal(5) = ""
        returnVal(6) = ""
        returnVal(7) = ""
        returnVal(8) = ""
    End If
    FindRegNumber = returnVal
End Function
