'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
        If nameSheet = "2020" & WorksheetIsExist(nameSheet) Then
            Call FillXMLTableForYear1(nameSheet, wb, ws_xml, 2, 4, 500)
        End If
        If nameSheet = "2021" & WorksheetIsExist(nameSheet) Then
            Call FillXMLTableForYear1(nameSheet, wb, ws_xml, 501, 4, 1000)
        End If
        If nameSheet = "2022" & WorksheetIsExist(nameSheet) Then
            Call FillXMLTableForYear1(nameSheet, wb, ws_xml, 1001, 4, 1500)
        End If
    Next i

    Call WriteXMLFile(ws_xml, wb)
End Sub
Private Function WorksheetIsExist(iName$) As Boolean
    On Error Resume Next
    WorksheetIsExist = (TypeOf Worksheets(iName$) Is Worksheet)
End Function
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
    ws_xml.Range("A1").value = "Kbk_code"
    'B - Íàèìåíîâàíèå óñëóãè
    ws_xml.Range("B1").value = "Íàèìåíîâàíèå_óñëóãè"
    'C - Íàèìåíîâàíèå îðãàíèçàöèè
    ws_xml.Range("C1").value = "Íàèìåíîâàíèå îðãàíèçàöèè"
    'D - ÈÍÍ
    ws_xml.Range("D1").value = "Inn"
    'CK - Êîä ó÷ðåæäåíèÿ
    ws_xml.Range("CK1").value = "Êîä ó÷ðåæäåíèÿ"
    'E - ãîä1
    ws_xml.Range("E1").value = "Value1"
    'E - ãîä2
    ws_xml.Range("F1").value = "Value2"
    'E - ãîä3
    ws_xml.Range("G1").value = "Value3"
    'CB - ÊÏÏ
    ws_xml.Range("CB1").value = "Kpp"
    'CC - Ðåãèñòðîâûé íîìåð óñëóãè
    ws_xml.Range("CC1").value = "RegNumber"
    'CD - Íàèìåíîâàíèå èíäèêàòîðà
    ws_xml.Range("CD1").value = "Íàèìåíîâàíèå_èíäèêàòîðà"
    'CE - Åäèíèöà èçìåðåíèÿ
    ws_xml.Range("CE1").value = "Åäèíèöà_èçìåðåíèÿ"
    'CF - Êîä óñëóãè
    ws_xml.Range("CF1").value = "Êîä_óñëóãè"
    'CG - Êîä åä. èçì
    ws_xml.Range("CG1").value = "Êîä åä. èçì"
    'CH - Êîä èíäèêàòîðà
    ws_xml.Range("CH1").value = "Êîä_èíäèêàòîðà"
    'CI - Äàòà íà÷àëà óñëóãè
    ws_xml.Range("CI1").value = "Íà÷àëî"
    'CJ - Äàòà îêîí÷àíèÿ óñëóãè
    ws_xml.Range("CJ1").value = "Êîíåö"
    '1
    'G - ÎÒ1 - H = 7
    ws_xml.Range("H1").value = "Insrns_Pmnt_val_1"
    'H - ÌÇ - I = 8
    ws_xml.Range("I1").value = "Mz_val_1"
    'I - ÔÐ1 - J = 9
    ws_xml.Range("J1").value = "Fr_val_1"
    'J - ÈÍÇ - K = 10
    ws_xml.Range("K1").value = "Inz_val_1"
    'K - ÊÓ - L = 11
    ws_xml.Range("L1").value = "Ku_val_1"
    'L - ÑÍÈ - M = 12
    ws_xml.Range("M1").value = "Sni_val_1"
    'M - ÑÎÖÄÈ - N = 13
    ws_xml.Range("N1").value = "Socdi_val_1"
    'N - ÔÐ2 - O = 14
    ws_xml.Range("O1").value = "Fr2_val_1"
    'O - ÓÑ - P = 15
    ws_xml.Range("P1").value = "Us_val_1"
    'P - ÒÓ - Q = 16
    ws_xml.Range("Q1").value = "Tu_val_1"
    'Q - ÎÒ2 - R = 17
    ws_xml.Range("R1").value = "Othr_Pmnt_val_1"
    'R - ÏÍÇ - S = 18
    ws_xml.Range("S1").value = "Pnz_val_1"
    '2
    'G - ÎÒ1 - T = 19
    ws_xml.Range("T1").value = "Insrns_Pmnt_val_2"
    'H - ÌÇ - U = 20
    ws_xml.Range("U1").value = "Mz_val_2"
    'I - ÔÐ1 - V = 21
     ws_xml.Range("V1").value = "Fr_val_2"
    'J - ÈÍÇ - W = 22
    ws_xml.Range("W1").value = "Inz_val_2"
    'K - ÊÓ - X = 23
    ws_xml.Range("X1").value = "Ku_val_2"
    'L - ÑÍÈ - Y = 24
    ws_xml.Range("Y1").value = "Sni_val_2"
    'M - ÑÎÖÄÈ - Z = 25
    ws_xml.Range("Z1").value = "Socdi_val_2"
    'N - ÔÐ2 - AA = 26
    ws_xml.Range("AA1").value = "Fr2_val_2"
    'O - ÓÑ - AB = 27
    ws_xml.Range("AB1").value = "Us_val_2"
    'P - ÒÓ - AC = 28
    ws_xml.Range("AC1").value = "Tu_val_2"
    'Q - ÎÒ2 - AD = 29
    ws_xml.Range("AD1").value = "Othr_Pmnt_val_2"
    'R - ÏÍÇ - AE = 30
    ws_xml.Range("AE1").value = "Pnz_val_2"
    '3
    'G - ÎÒ1 - AF = 31
    ws_xml.Range("AF1").value = "Insrns_Pmnt_val_3"
    'H - ÌÇ - AG = 32
    ws_xml.Range("AG1").value = "Mz_val_3"
    'I - ÔÐ1 - AH = 33
    ws_xml.Range("AH1").value = "Fr_val_3"
    'J - ÈÍÇ - AI = 34
    ws_xml.Range("AI1").value = "Inz_val_3"
    'K - ÊÓ - AJ = 35
    ws_xml.Range("AJ1").value = "Ku_val_3"
    'L - ÑÍÈ - AK = 36
    ws_xml.Range("AK1").value = "Sni_val_3"
    'M - ÑÎÖÄÈ - AL = 37
    ws_xml.Range("AL1").value = "Socdi_val_3"
    'N - ÔÐ2 - AM = 38
    ws_xml.Range("AM1").value = "Fr2_val_3"
    'O - ÓÑ - AN = 39
    ws_xml.Range("AN1").value = "Us_val_3"
    'P - ÒÓ - AO = 40
    ws_xml.Range("AO1").value = "Tu_val_3"
    'Q - ÎÒ2 - AP = 41
    ws_xml.Range("AP1").value = "Othr_Pmnt_val_3"
    'R - ÏÍÇ - AQ = 42
    ws_xml.Range("AQ1").value = "Pnz_val_3"
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
        ws_xml.Cells(i, 3).value = ws.Range("D2").value
        ws_xml.Cells(i, 4).value = ws.Range("N2").value
        ws_xml.Cells(i, 80).value = ws.Range("O2").value
        ws_xml.Cells(i, 87).value = ws.Range("P2").value
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
            If Trim(CStr(ws_xml.Cells(i, 81).value)) = Trim(CStr(rw.Cells(1, 2).value)) Then
                'RegNumber = CStr(rw.Cells(1, 2).Value)
                RegName = CStr(rw.Cells(1, 1).value)
                IndicatorName = CStr(rw.Cells(1, 3).value)
                OkeiName = CStr(rw.Cells(1, 4).value)
                OkeiCode = CStr(rw.Cells(1, 8).value)
                BaseCode = CStr(rw.Cells(1, 5).value)
                IndCode = CStr(rw.Cells(1, 9).value)
                DateFrom = CStr(rw.Cells(1, 6).value)
                DateBefore = CStr(rw.Cells(1, 7).value)
                Exit For
            End If
            If CStr(rw.Cells(1, 1).value) = "" Then
                Exit For
            End If
        Next rw
       ' ws_xml.Cells(i, 81).Value = RegNumber
        ws_xml.Cells(i, 2).value = RegName
        ws_xml.Cells(i, 82).value = IndicatorName
        ws_xml.Cells(i, 83).value = OkeiName
        ws_xml.Cells(i, 84).value = BaseCode
        ws_xml.Cells(i, 85).value = OkeiCode
        ws_xml.Cells(i, 86).value = IndCode
        ws_xml.Cells(i, 87).value = DateFrom
        ws_xml.Cells(i, 88).value = DateBefore
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
    'Îáåñïå÷åíèå ó÷àñòèÿ ñáîðíûõ êîìàíä Ðîññèéñêîé ôåäåðàöèè â ìåæäóíàðîäíûõ ñïîðòèâíûõ ñîðåâíîâàíèÿõ, Îëèìïèéñêèõ èãðàõ.
    'Íà òåððèòîðèè Ðîññèéñêîé Ôåäåðàöèè
    '1
    'Plan
    ws_xml.Range("AR2:BC2").value = ws.Range("G7:R7").value
    ws_xml.Range("BD2:BO2").value = ws.Range("G7:R7").value
    ws_xml.Range("BP2:CA2").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H2:S2").value = ws.Range("G8:R8").value
    ws_xml.Range("T2:AE2").value = ws.Range("G8:R8").value
    ws_xml.Range("AF2:AQ2").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC2").value = ws.Range("V7").value
    'Values
    ws_xml.Range("E2:G2").value = ws.Range("D8:F8").value
    'KBK
    ws_xml.Range("A2").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B2").value = ws.Range("A7").value
    
    '2
    'Plan
    ws_xml.Range("AR3:BC3").value = ws.Range("G9:R9").value
    ws_xml.Range("BD3:BO3").value = ws.Range("G9:R9").value
    ws_xml.Range("BP3:CA3").value = ws.Range("G9:R9").value
    'Current
    ws_xml.Range("H3:S3").value = ws.Range("G10:R10").value
    ws_xml.Range("T3:AE3").value = ws.Range("G10:R10").value
    ws_xml.Range("AF3:AQ3").value = ws.Range("G10:R10").value
    'RegNumber
    ws_xml.Range("CC3").value = ws.Range("V10").value
    'Values
    ws_xml.Range("E3:G3").value = ws.Range("D10:F10").value
    'KBK
    ws_xml.Range("A3").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B3").value = ws.Range("A7").value
    'RegNumber
    ws_xml.Range("CC2").value = ws.Range("V7").value
    'Îáåñïå÷åíèå ó÷àñòèÿ ñáîðíûõ êîìàíä Ðîññèéñêîé ôåäåðàöèè â ìåæäóíàðîäíûõ ñïîðòèâíûõ ñîðåâíîâàíèÿõ, Îëèìïèéñêèõ èãðàõ.
    'Çà ïðåäåëàìè òåððèòîðèè Ðîññèéñêîé Ôåäåðàöèè
    '1
    'Plan
    ws_xml.Range("AR4:BC4").value = ws.Range("G14:R14").value
    ws_xml.Range("BD4:BO4").value = ws.Range("G14:R14").value
    ws_xml.Range("BP4:CA4").value = ws.Range("G14:R14").value
    'Current
    ws_xml.Range("H4:S4").value = ws.Range("G15:R15").value
    ws_xml.Range("T4:AE4").value = ws.Range("G15:R15").value
    ws_xml.Range("AF4:AQ4").value = ws.Range("G15:R15").value
    'RegNumber
    ws_xml.Range("CC4").value = ws.Range("V15").value
    'Values
    ws_xml.Range("E4:G4").value = ws.Range("D15:F15").value
    'KBK
    ws_xml.Range("A4").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B4").value = ws.Range("A7").value
    '2
    'Plan
    ws_xml.Range("AR5:BC5").value = ws.Range("G16:R16").value
    ws_xml.Range("BD5:BO5").value = ws.Range("G16:R16").value
    ws_xml.Range("BP5:CA5").value = ws.Range("G16:R16").value
    'Current
    ws_xml.Range("H5:S5").value = ws.Range("G17:R17").value
    ws_xml.Range("T5:AE5").value = ws.Range("G18:R18").value
    ws_xml.Range("AF5:AQ5").value = ws.Range("G19:R19").value
    'RegNumber
    ws_xml.Range("CC5").value = ws.Range("V19").value
    'Values
    ws_xml.Range("E5").value = ws.Range("D17").value
    ws_xml.Range("F5").value = ws.Range("E18").value
    ws_xml.Range("G5").value = ws.Range("F19").value
    'KBK
    ws_xml.Range("A5").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B5").value = ws.Range("A7").value
    'Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ñïîðòèâíûõ ìåðîïðèÿòèé.
    'Ìåæäóíàðîäíûå, íà òåððèòîðèè Ðîññèéñêîé Ôåäåðàöèè
    '1
    'Plan
    ws_xml.Range("AR6:BC6").value = ws.Range("G23:R23").value
    ws_xml.Range("BD6:BO6").value = ws.Range("G23:R23").value
    ws_xml.Range("BP6:CA6").value = ws.Range("G23:R23").value
    'Current
    ws_xml.Range("H6:S6").value = ws.Range("G24:R24").value
    ws_xml.Range("T6:AE6").value = ws.Range("G24:R24").value
    ws_xml.Range("AF6:AQ6").value = ws.Range("G24:R24").value
    'RegNumber
    ws_xml.Range("CC6").value = ws.Range("V24").value
    'Values
    ws_xml.Range("E6:G6").value = ws.Range("D24:F24").value
    'KBK
    ws_xml.Range("A6").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B6").value = ws.Range("A23").value
    'Âñåðîññèéñêèå, íà òåððèòîðèè Ðîññèéñêîé Ôåäåðàöèè
    '2
    'Plan
    ws_xml.Range("AR7:BC7").value = ws.Range("G25:R25").value
    ws_xml.Range("BD7:BO7").value = ws.Range("G25:R25").value
    ws_xml.Range("BP7:CA7").value = ws.Range("G25:R25").value
    'Current
    ws_xml.Range("H7:S7").value = ws.Range("G26:R26").value
    ws_xml.Range("T7:AE7").value = ws.Range("G26:R26").value
    ws_xml.Range("AF7:AQ7").value = ws.Range("G26:R26").value
    'RegNumber
    ws_xml.Range("CC7").value = ws.Range("V26").value
    'Values
    ws_xml.Range("E7:G7").value = ws.Range("D26:F26").value
    'KBK
    ws_xml.Range("A7").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B7").value = ws.Range("A23").value
    'Îðãàíèçàöèÿ ìåðîïðèÿòèé ïî ïîäãîòîâêå ñïîðòèâíûõ ñáîðíûõ êîìàíä.
    '
    '1
    'Plan
    ws_xml.Range("AR8:BC8").value = ws.Range("G28:R28").value
    ws_xml.Range("BD8:BO8").value = ws.Range("G28:R28").value
    ws_xml.Range("BP8:CA8").value = ws.Range("G28:R28").value
    'Current
    ws_xml.Range("H8:S8").value = ws.Range("G29:R29").value
    ws_xml.Range("T8:AE8").value = ws.Range("G29:R29").value
    ws_xml.Range("AF8:AQ8").value = ws.Range("G29:R29").value
    'RegNumber
    ws_xml.Range("CC8").value = ws.Range("V29").value
    'Values
    ws_xml.Range("E8:G8").value = ws.Range("D29:F29").value
    'KBK
    ws_xml.Range("A8").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B8").value = ws.Range("A28").value
    '2
    'Plan
    ws_xml.Range("AR9:BC9").value = ws.Range("G30:R30").value
    ws_xml.Range("BD9:BO9").value = ws.Range("G30:R30").value
    ws_xml.Range("BP9:CA9").value = ws.Range("G30:R30").value
    'Current
    ws_xml.Range("H9:S9").value = ws.Range("G31:R31").value
    ws_xml.Range("T9:AE9").value = ws.Range("G31:R31").value
    ws_xml.Range("AF9:AQ9").value = ws.Range("G31:R31").value
    'RegNumber
    ws_xml.Range("CC9").value = ws.Range("V31").value
    'Values
    ws_xml.Range("E9:G9").value = ws.Range("D31:F31").value
    'KBK
    ws_xml.Range("A9").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B9").value = ws.Range("A28").value
    '3
    'Plan
    ws_xml.Range("AR10:BC10").value = ws.Range("G32:R32").value
    ws_xml.Range("BD10:BO10").value = ws.Range("G32:R32").value
    ws_xml.Range("BP10:CA10").value = ws.Range("G32:R32").value
    'Current
    ws_xml.Range("H10:S10").value = ws.Range("G33:R33").value
    ws_xml.Range("T10:AE10").value = ws.Range("G33:R33").value
    ws_xml.Range("AF10:AQ10").value = ws.Range("G33:R33").value
    'RegNumber
    ws_xml.Range("CC10").value = ws.Range("V33").value
    'Values
    ws_xml.Range("E10").value = ws.Range("D33").value
    ws_xml.Range("F10").value = ws.Range("E33").value
    ws_xml.Range("G10").value = ws.Range("F33").value
    'KBK
    ws_xml.Range("A10").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B10").value = ws.Range("A28").value
    '4
    'Plan
    ws_xml.Range("AR11:BC11").value = ws.Range("G36:R36").value
    ws_xml.Range("BD11:BO11").value = ws.Range("G36:R36").value
    ws_xml.Range("BP11:CA11").value = ws.Range("G36:R36").value
    'Current
    ws_xml.Range("H11:S11").value = ws.Range("G37:R37").value
    ws_xml.Range("T11:AE11").value = ws.Range("G37:R37").value
    ws_xml.Range("AF11:AQ11").value = ws.Range("G37:R37").value
    'RegNumber
    ws_xml.Range("CC11").value = ws.Range("V37").value
    'Values
    ws_xml.Range("E11:G11").value = ws.Range("D37:F37").value
    'KBK
    ws_xml.Range("A11").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B11").value = ws.Range("A28").value
    '5
    'Plan
    ws_xml.Range("AR12:BC12").value = ws.Range("G38:R38").value
    ws_xml.Range("BD12:BO12").value = ws.Range("G38:R38").value
    ws_xml.Range("BP12:CA12").value = ws.Range("G38:R38").value
    'Current
    ws_xml.Range("H12:S12").value = ws.Range("G39:R39").value
    ws_xml.Range("T12:AE12").value = ws.Range("G39:R39").value
    ws_xml.Range("AF12:AQ12").value = ws.Range("G39:R39").value
    'RegNumber
    ws_xml.Range("CC12").value = ws.Range("V39").value
    'Values
    ws_xml.Range("E12:G12").value = ws.Range("D39:F39").value
    'KBK
    ws_xml.Range("A12").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B12").value = ws.Range("A28").value
    '6
    'Plan
    ws_xml.Range("AR13:BC13").value = ws.Range("G44:R44").value
    ws_xml.Range("BD13:BO13").value = ws.Range("G44:R44").value
    ws_xml.Range("BP13:CA13").value = ws.Range("G44:R44").value
    'Current
    ws_xml.Range("H13:S13").value = ws.Range("G45:R45").value
    ws_xml.Range("T13:AE13").value = ws.Range("G45:R45").value
    ws_xml.Range("AF13:AQ13").value = ws.Range("G45:R45").value
    'RegNumber
    ws_xml.Range("CC13").value = ws.Range("V45").value
    'Values
    ws_xml.Range("E13:G13").value = ws.Range("D45:F45").value
    'KBK
    ws_xml.Range("A13").value = ws.Range("A49").value
    'NAME
    ws_xml.Range("B13").value = ws.Range("A28").value
    'Îðãàíèçàöèÿ ìåðîïðèÿòèé ïî íàó÷íî-ìåòîäè÷åñêîìó îáåñïå÷åíèþ ñïîðòèâíûõ ñáîðíûõ êîìàíä.
    '0000000001100077708  30042100100000000004100103
    '1
    'Plan
    ws_xml.Range("AR14:BC14").value = ws.Range("G59:R59").value
    ws_xml.Range("BD14:BO14").value = ws.Range("G59:R59").value
    ws_xml.Range("BP14:CA14").value = ws.Range("G59:R59").value
    'Current
    ws_xml.Range("H14:S14").value = ws.Range("G60:R60").value
    ws_xml.Range("T14:AE14").value = ws.Range("G60:R60").value
    ws_xml.Range("AF14:AQ14").value = ws.Range("G60:R60").value
    'RegNumber
    ws_xml.Range("CC14").value = ws.Range("V60").value
    'Values
    ws_xml.Range("E14:G14").value = ws.Range("D60:F60").value
    'KBK
    ws_xml.Range("A14").value = ws.Range("A61").value
    'NAME
    ws_xml.Range("B14").value = ws.Range("A59").value
    'çàïîëíÿåì ðåêâåçèòû 12 ñòðîê
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
    'Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ôèçêóëüòóðíûõ (ôèçêóëüòóðíî-îçäîðîâèòåëüíûõ) ìåðîïðèÿòèé.
    'Ìåæäóíàðîäíûå íà òåððèòîðèè Ðîññèéñêîé Ôåäåðàöèè
    '1
    'Plan
    ws_xml.Range("AR15:BC15").value = ws.Range("G7:R7").value
    ws_xml.Range("BD15:BO15").value = ws.Range("G7:R7").value
    ws_xml.Range("BP15:CA15").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H15:S15").value = ws.Range("G8:R8").value
    ws_xml.Range("T15:AE15").value = ws.Range("G9:R9").value
    ws_xml.Range("AF15:AQ15").value = ws.Range("G10:R10").value
    'RegNumber
    ws_xml.Range("CC15").value = ws.Range("X7").value
    'Values
    'ws_xml.Range("E15:G15").Value = ws.Range("D8:F8").Value
    ws_xml.Range("E15").value = ws.Range("D8").value
    ws_xml.Range("F15").value = ws.Range("E9").value
    ws_xml.Range("G15").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A15").value = ws.Range("A23").value
    'NAME
    ws_xml.Range("B15").value = ws.Range("A7").value
    
    '2
    'Plan
    ws_xml.Range("AR16:BC16").value = ws.Range("G11:R11").value
    ws_xml.Range("BD16:BO16").value = ws.Range("G11:R11").value
    ws_xml.Range("BP16:CA16").value = ws.Range("G1:R11").value
    'Current
    ws_xml.Range("H16:S16").value = ws.Range("G12:R12").value
    ws_xml.Range("T16:AE16").value = ws.Range("G12:R12").value
    ws_xml.Range("AF16:AQ16").value = ws.Range("G12:R12").value
    'RegNumber
    ws_xml.Range("CC16").value = ws.Range("X12").value
    'Values
    ws_xml.Range("E16:G16").value = ws.Range("D12:F12").value
    'ws_xml.Range("E16").Value = ws.Range("D8").Value
    'ws_xml.Range("F15").Value = ws.Range("E9").Value
    'ws_xml.Range("G15").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A16").value = ws.Range("A23").value
    'NAME
    ws_xml.Range("B16").value = ws.Range("A7").value
    
    '3
    'Plan
    ws_xml.Range("AR17:BC17").value = ws.Range("G13:R13").value
    ws_xml.Range("BD17:BO17").value = ws.Range("G13:R13").value
    ws_xml.Range("BP17:CA17").value = ws.Range("G13:R13").value
    'Current
    ws_xml.Range("H17:S17").value = ws.Range("G14:R14").value
    ws_xml.Range("T17:AE17").value = ws.Range("G15:R15").value
    ws_xml.Range("AF17:AQ17").value = ws.Range("G16:R16").value
    'RegNumber
    ws_xml.Range("CC17").value = ws.Range("X13").value
    'Values
    'ws_xml.Range("E17:G17").Value = ws.Range("D12:F12").Value
    ws_xml.Range("E17").value = ws.Range("D14").value
    ws_xml.Range("F17").value = ws.Range("E15").value
    ws_xml.Range("G17").value = ws.Range("F16").value
    'KBK
    ws_xml.Range("A17").value = ws.Range("A23").value
    'NAME
    ws_xml.Range("B17").value = ws.Range("A7").value

    '4
    'Plan
    ws_xml.Range("AR18:BC18").value = ws.Range("G17:R17").value
    ws_xml.Range("BD18:BO18").value = ws.Range("G17:R17").value
    ws_xml.Range("BP18:CA18").value = ws.Range("G17:R17").value
    'Current
    ws_xml.Range("H18:S18").value = ws.Range("G18:R18").value
    ws_xml.Range("T18:AE18").value = ws.Range("G18:R18").value
    ws_xml.Range("AF18:AQ18").value = ws.Range("G18:R18").value
    'RegNumber
    ws_xml.Range("CC18").value = ws.Range("X18").value
    'Values
    ws_xml.Range("E18:G18").value = ws.Range("D18:F18").value
    'ws_xml.Range("E18").Value = ws.Range("D14").Value
    'ws_xml.Range("F18").Value = ws.Range("E15").Value
    'ws_xml.Range("G18").Value = ws.Range("F16").Value
    'KBK
    ws_xml.Range("A18").value = ws.Range("A23").value
    'NAME
    ws_xml.Range("B18").value = ws.Range("A7").value
    
    'Ðàáîòà 2. Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ñïîðòèâíûõ ìåðîïðèÿòèé.
    '
    '1
    'Plan
    ws_xml.Range("AR19:BC19").value = ws.Range("G32:R32").value
    ws_xml.Range("BD19:BO19").value = ws.Range("G32:R32").value
    ws_xml.Range("BP19:CA19").value = ws.Range("G32:R32").value
    'Current
    ws_xml.Range("H19:S19").value = ws.Range("G33:R33").value
    ws_xml.Range("T19:AE19").value = ws.Range("G33:R33").value
    ws_xml.Range("AF19:AQ19").value = ws.Range("G33:R33").value
    'RegNumber
    ws_xml.Range("CC19").value = ws.Range("X32").value
    'Values
    ws_xml.Range("E19:G19").value = ws.Range("D33:F33").value
    'ws_xml.Range("E19").Value = ws.Range("D8").Value
    'ws_xml.Range("F19").Value = ws.Range("E9").Value
    'ws_xml.Range("G19").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A19").value = ws.Range("A35").value
    'NAME
    ws_xml.Range("B19").value = ws.Range("A32").value
    
    'Ðàáîòà 2. Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ñïîðòèâíûõ ìåðîïðèÿòèé.
    '
    '1
    'Plan
    ws_xml.Range("AR20:BC20").value = ws.Range("G42:R42").value
    ws_xml.Range("BD20:BO20").value = ws.Range("G42:R42").value
    ws_xml.Range("BP20:CA20").value = ws.Range("G42:R42").value
    'Current
    ws_xml.Range("H20:S20").value = ws.Range("G43:R43").value
    ws_xml.Range("T20:AE20").value = ws.Range("G43:R43").value
    ws_xml.Range("AF20:AQ20").value = ws.Range("G43:R43").value
    'RegNumber
    ws_xml.Range("CC20").value = ws.Range("X42").value
    'Values
    ws_xml.Range("E20:G20").value = ws.Range("D43:F43").value
    'ws_xml.Range("E19").Value = ws.Range("D8").Value
    'ws_xml.Range("F19").Value = ws.Range("E9").Value
    'ws_xml.Range("G19").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A20").value = ws.Range("A47").value
    'NAME
    ws_xml.Range("B20").value = ws.Range("A42").value
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR21:BC21").value = ws.Range("G6:R6").value
    ws_xml.Range("BD21:BO21").value = ws.Range("G6:R6").value
    ws_xml.Range("BP21:CA21").value = ws.Range("G6:R6").value
    'Current
    ws_xml.Range("H21:S21").value = ws.Range("G7:R7").value
    ws_xml.Range("T21:AE21").value = ws.Range("G7:R7").value
    ws_xml.Range("AF21:AQ21").value = ws.Range("G7:R7").value
    'RegNumber
    ws_xml.Range("CC21").value = ws.Range("V6").value
    'Values
    ws_xml.Range("E21:G21").value = ws.Range("D7:F7").value
    'ws_xml.Range("E21").Value = ws.Range("D8").Value
    'ws_xml.Range("F21").Value = ws.Range("E9").Value
    'ws_xml.Range("G21").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A21").value = ws.Range("A13").value
    'NAME
    ws_xml.Range("B21").value = ws.Range("A6").value
    
    'Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ñïîðòèâíûõ ìåðîïðèÿòèé.
    '
    '1
    'Plan
    ws_xml.Range("AR22:BC22").value = ws.Range("G8:R8").value
    ws_xml.Range("BD22:BO22").value = ws.Range("G8:R8").value
    ws_xml.Range("BP22:CA22").value = ws.Range("G8:R8").value
    'Current
    ws_xml.Range("H22:S22").value = ws.Range("G9:R9").value
    ws_xml.Range("T22:AE22").value = ws.Range("G9:R9").value
    ws_xml.Range("AF22:AQ22").value = ws.Range("G9:R9").value
    'RegNumber
    ws_xml.Range("CC22").value = ws.Range("V8").value
    'Values
    ws_xml.Range("E22:G22").value = ws.Range("D7:F7").value
    'ws_xml.Range("E22").Value = ws.Range("D8").Value
    'ws_xml.Range("F22").Value = ws.Range("E9").Value
    'ws_xml.Range("G22").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A22").value = ws.Range("A13").value
    'NAME
    ws_xml.Range("B22").value = ws.Range("A6").value
    
    'Îðãàíèçàöèÿ è îáåñïå÷åíèå ýêñïåðèìåíòàëüíîé è èííîâàöèîííîé äåÿòåëüíîñòè â îáëàñòè ôèçêóëüòóðû è ñïîðòà.
    '
    '1
    'Plan
    ws_xml.Range("AR23:BC23").value = ws.Range("G10:R10").value
    ws_xml.Range("BD23:BO23").value = ws.Range("G10:R10").value
    ws_xml.Range("BP23:CA23").value = ws.Range("G10:R10").value
    'Current
    ws_xml.Range("H23:S23").value = ws.Range("G11:R11").value
    ws_xml.Range("T23:AE23").value = ws.Range("G11:R11").value
    ws_xml.Range("AF23:AQ23").value = ws.Range("G11:R11").value
    'RegNumber
    ws_xml.Range("CC23").value = ws.Range("V10").value
    'Values
    ws_xml.Range("E23:G23").value = ws.Range("D11:F11").value
    'ws_xml.Range("E23").Value = ws.Range("D8").Value
    'ws_xml.Range("F23").Value = ws.Range("E9").Value
    'ws_xml.Range("G23").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A23").value = ws.Range("A13").value
    'NAME
    ws_xml.Range("B23").value = ws.Range("A6").value
    
    'Ðàáîòà 4. Îðãàíèçàöèÿ è ïðîâåäåíèå îôèöèàëüíûõ ôèçêóëüòóðíûõ (ôèçêóëüòóðíî-îçäîðîâèòåëüíûõ) ìåðîïðèÿòèé.
    '
    '1
    'Plan
    ws_xml.Range("AR24:BC24").value = ws.Range("G20:R20").value
    ws_xml.Range("BD24:BO24").value = ws.Range("G20:R20").value
    ws_xml.Range("BP24:CA24").value = ws.Range("G20:R20").value
    'Current
    ws_xml.Range("H24:S24").value = ws.Range("G21:R21").value
    ws_xml.Range("T24:AE24").value = ws.Range("G21:R21").value
    ws_xml.Range("AF24:AQ24").value = ws.Range("G21:R21").value
    'RegNumber
    ws_xml.Range("CC24").value = ws.Range("V20").value
    'Values
    ws_xml.Range("E24:G24").value = ws.Range("D21:F21").value
    'ws_xml.Range("E24").Value = ws.Range("D8").Value
    'ws_xml.Range("F24").Value = ws.Range("E9").Value
    'ws_xml.Range("G24").Value = ws.Range("F10").Value
    'KBK
    ws_xml.Range("A24").value = ws.Range("A23").value
    'NAME
    ws_xml.Range("B24").value = ws.Range("A20").value
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR25:BC25").value = ws.Range("G7:R7").value
    ws_xml.Range("BD25:BO25").value = ws.Range("G7:R7").value
    ws_xml.Range("BP25:CA25").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H25:S25").value = ws.Range("G8:R8").value
    ws_xml.Range("T25:AE25").value = ws.Range("G8:R8").value
    ws_xml.Range("AF25:AQ25").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC25").value = ws.Range("V7").value
    'Values
    'ws_xml.Range("E25:G25").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E25").value = ws.Range("D8").value
    ws_xml.Range("F25").value = ws.Range("E9").value
    ws_xml.Range("G25").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A25").value = ws.Range("A12").value
    'NAME
    ws_xml.Range("B25").value = ws.Range("A7").value
    
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR26:BC26").value = ws.Range("G7:R7").value
    ws_xml.Range("BD26:BO26").value = ws.Range("G7:R7").value
    ws_xml.Range("BP26:CA26").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H26:S26").value = ws.Range("G8:R8").value
    ws_xml.Range("T26:AE26").value = ws.Range("G8:R8").value
    ws_xml.Range("AF26:AQ26").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC26").value = ws.Range("V7").value
    'Values
    'ws_xml.Range("E26:G26").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E26").value = ws.Range("D8").value
    ws_xml.Range("F26").value = ws.Range("E9").value
    ws_xml.Range("G26").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A26").value = ws.Range("A12").value
    'NAME
    ws_xml.Range("B26").value = ws.Range("A7").value
    
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR27:BC27").value = ws.Range("G7:R7").value
    ws_xml.Range("BD27:BO27").value = ws.Range("G7:R7").value
    ws_xml.Range("BP27:CA27").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H27:S27").value = ws.Range("G8:R8").value
    ws_xml.Range("T27:AE27").value = ws.Range("G8:R8").value
    ws_xml.Range("AF27:AQ27").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC27").value = ws.Range("V7").value
    'Values
    'ws_xml.Range("E27:G27").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E27").value = ws.Range("D8").value
    ws_xml.Range("F27").value = ws.Range("E9").value
    ws_xml.Range("G27").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A27").value = ws.Range("A12").value
    'NAME
    ws_xml.Range("B27").value = ws.Range("A7").value
    
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR28:BC28").value = ws.Range("G7:R7").value
    ws_xml.Range("BD28:BO28").value = ws.Range("G7:R7").value
    ws_xml.Range("BP28:CA28").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H28:S28").value = ws.Range("G8:R8").value
    ws_xml.Range("T28:AE28").value = ws.Range("G8:R8").value
    ws_xml.Range("AF28:AQ28").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC28").value = ws.Range("V7").value
    'Values
    'ws_xml.Range("E27:G27").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E28").value = ws.Range("D8").value
    ws_xml.Range("F28").value = ws.Range("E9").value
    ws_xml.Range("G28").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A28").value = ws.Range("A12").value
    'NAME
    ws_xml.Range("B28").value = ws.Range("A7").value
    
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    'Ðàáîòà 1. Îðãàíèçàöèÿ è îáåñïå÷åíèå êîîðäèíàöèè äåÿòåëüíîñòè ôèçêóëüòóðíî-ñïîðòèâíûõ îðãàíèçàöèé ïî ïîäãîòîâêå ñïîðòèâíîãî ðåçåðâà.
    '
    '1
    'Plan
    ws_xml.Range("AR29:BC29").value = ws.Range("G7:R7").value
    ws_xml.Range("BD29:BO29").value = ws.Range("G7:R7").value
    ws_xml.Range("BP29:CA29").value = ws.Range("G7:R7").value
    'Current
    ws_xml.Range("H29:S29").value = ws.Range("G8:R8").value
    ws_xml.Range("T29:AE29").value = ws.Range("G8:R8").value
    ws_xml.Range("AF29:AQ29").value = ws.Range("G8:R8").value
    'RegNumber
    ws_xml.Range("CC29").value = ws.Range("V7").value
    'Values
    'ws_xml.Range("E29:G29").Value = ws.Range("D7:F7").Value
    ws_xml.Range("E29").value = ws.Range("D8").value
    ws_xml.Range("F29").value = ws.Range("E9").value
    ws_xml.Range("G29").value = ws.Range("F10").value
    'KBK
    ws_xml.Range("A29").value = ws.Range("A12").value
    'NAME
    ws_xml.Range("B29").value = ws.Range("A7").value
    
    
    'çàïîëíÿåì ðåêâåçèòû 6 ñòðîê
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
    Dim base_name_prev, base_name_cur, okei_code_prev, okei_code_cur, base_code_prev, base_code_cur, ind_name_prev, ind_name_cur, kbk_prev, kbk_cur As String
    Dim isTheSame As Boolean
    isTheSame = False
    base_name_prev = ""
    base_code_prev = ""
    okei_code_prev = ""
    ind_name_prev = ""
    kbk_prev = ""
    
    Dim normativi_folder_path As String
    normativi_folder_path = wb.path + "\Íîðìàòèâû"
    
    If Dir(normativi_folder_path, vbDirectory) <> "" Then 'ïðîâåðÿåì åñòü ëè ïàïêà "èìÿ ïàïêè"
    
    Else
        MkDir (normativi_folder_path) 'ñîçäà¸ì ïàïêó "èìÿ ïàïêè"
    End If
    
    
    For Each rw In ws_xml.Rows

        If RowCount > 0 Then
            base_name_cur = CStr(rw.Cells(1, 2).value)
            okei_code_cur = CStr(Format(rw.Cells(1, 85), "000"))
            base_code_cur = CStr(rw.Cells(1, 84))
            ind_name_cur = CStr(rw.Cells(1, 82))
            kbk_cur = CStr(rw.Cells(1, 1))
                If kbk_cur <> "" And kbk_cur <> "0" Then
                    If base_name_cur <> base_name_prev Or okei_code_cur <> okei_code_prev Or ind_name_prev <> ind_name_cur Then
                        If RowCount > 1 Then
                            CountNormativ = CountNormativ + 1
                            Dim path_xml As String
                            'path_xml = CStr(GetFilePath(wb.path, wb.Name, CStr(CountNormativ) & " Íîðìàòèâû", CStr(base_name_prev)))
                            'path_xml = CStr(GetFilePath(wb.path, wb.name, CStr(CountNormativ) & "_Íîðìàòèâû", CStr(base_code_prev) & "_" & CStr(ind_name_prev)))
                            path_xml = CStr(GetFilePath(normativi_folder_path, wb.name, CStr(CountNormativ) & "_Íîðìàòèâû", CStr(base_code_prev) & "_" & CStr(ind_name_prev)))
                            XDoc.Save path_xml
                        End If
                        If base_name_cur = "" Then
                            Exit For
                        End If
                        Set XDoc = CreateObject("MSXML2.DOMDocument")
                        XDoc.appendChild XDoc.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'")
                        isTheSame = False
                        If ws_xml.Cells(rw.Row, 1).value Then
                            base_name_prev = base_name_cur
                            okei_code_prev = okei_code_cur
                            base_code_prev = base_code_cur
                            ind_name_prev = ind_name_cur
                            kbk_prev = kbk_cur
                        End If
                    Else
                        isTheSame = True
                    End If
                End If
                Dim strRegNumber, newStrRegNumber As String
                strRegNumber = RemoveWhiteSpace(CStr(rw.Cells(1, 81)))
                
                isGood = True
                If strRegNumber = "" Or kbk_cur = "0" Then
                    isGood = False
                End If
                If isGood Then
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
                        Inst_name_oiv.Text = "ÌÈÍÈÑÒÅÐÑÒÂÎ ÑÏÎÐÒÀ ÐÎÑÑÈÉÑÊÎÉ ÔÅÄÅÐÀÖÈÈ"
                        
                        Set Inst_inn = XDoc.createElement("Inst_inn")
                        root.appendChild Inst_inn
                        Inst_inn.Text = "7703771271"
                        
                        Set Inst_kpp = XDoc.createElement("Inst_kpp")
                        root.appendChild Inst_kpp
                        Inst_kpp.Text = "770901001"
                        
                        Set Inst_code = XDoc.createElement("Inst_code")
                        root.appendChild Inst_code
                        Inst_code.Text = Inst_inn.Text + Inst_kpp.Text
                        
                        Set Registry_records = XDoc.createElement("Registry_records")
                        root.appendChild Registry_records
                        
                        'Set Dprtm_values = XDoc.createElement("Dprtm_values")
                        'root.appendChild Dprtm_values
                    End If
    
                    Set Registry_record = XDoc.createElement("Registry_record")
                    Registry_records.appendChild Registry_record
                    
                    
                    Set RegNumber = XDoc.createElement("RegNumber")
                    Registry_record.appendChild RegNumber
      
                    newStrRegNumber = getStringBeforeSpace(strRegNumber, "_")
                    RegNumber.Text = newStrRegNumber
                    
                    'Ó÷ðåæäåíèÿ
                    Set Dprtm_records = XDoc.createElement("Dprtm_records")
                    Registry_record.appendChild Dprtm_records
                    
                    'Çàïèñü ó÷ðåæäåíèÿ
                    Set Dprtm_record = XDoc.createElement("Dprtm_record")
                    Dprtm_records.appendChild Dprtm_record
                    
                    Set Dprtm_code = XDoc.createElement("Dprtm_code")
                    Dprtm_record.appendChild Dprtm_code
                    Dprtm_code.Text = CStr(rw.Cells(1, 89))
                    
                    Set Dprtm_name = XDoc.createElement("Dprtm_name")
                    Dprtm_record.appendChild Dprtm_name
                    Dprtm_name.Text = CStr(rw.Cells(1, 3))
                    
                    Set Dprtm_inn = XDoc.createElement("Dprtm_inn")
                    Dprtm_record.appendChild Dprtm_inn
                    Dprtm_inn.Text = CStr(rw.Cells(1, 4))
                    
                    Set Dprtm_kpp = XDoc.createElement("Dprtm_kpp")
                    Dprtm_record.appendChild Dprtm_kpp
                    Dprtm_kpp.Text = CStr(rw.Cells(1, 80))
                    
                    'Çíà÷åíèÿ ó÷ðåæäåíèÿ
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
        End If
        RowCount = RowCount + 1
    Next rw
    
    'Call PackArchive(normativi_folder_path, wb.path, "Íîðìàòèâû")
        
    
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
    Inst_name_oiv.Text = "ÌÈÍÈÑÒÅÐÑÒÂÎ ÑÏÎÐÒÀ ÐÎÑÑÈÉÑÊÎÉ ÔÅÄÅÐÀÖÈÈ"
                    
    Set Inst_inn = XDoc.createElement("Inst_inn")
    root.appendChild Inst_inn
    Inst_inn.Text = "7703771271"
                    
    Set Inst_kpp = XDoc.createElement("Inst_kpp")
    root.appendChild Inst_kpp
    Inst_kpp.Text = "770901001"
    
    Set Inst_code = XDoc.createElement("Inst_code")
    root.appendChild Inst_code
    Inst_code.Text = "00100777"
    'Inst_code.Text = Inst_inn.Text + Inst_kpp.Text
    
    Set Dprtm_records = XDoc.createElement("Dprtm_records")
    root.appendChild Dprtm_records
            
                    
    For Each rw In ws_xml.Rows
        If RowCount > 0 Then

            If CStr(ws_xml.Cells(rw.Row, 1).value) = "" Then
                If RowCount > 1 Then
                    Dim path_xml As String
                    'path_xml = CStr(GetFilePath(wb.path, wb.Name, CStr(CountNormativ) & " Íîðìàòèâû", CStr(base_name_prev)))
                    path_xml = CStr(GetFilePath(wb.path, wb.name, "0" & "_ÎÔÎ", "" & "" & ""))
                    XDoc.Save path_xml
                End If
                If CStr(ws_xml.Cells(rw.Row, 1).value) = "" Then
                    Exit For
                End If

            End If
            
            dep_name_cur = CStr(rw.Cells(1, 3).value)
            okei_code_cur = CStr(Format(rw.Cells(1, 85), "000"))
            base_code_cur = CStr(rw.Cells(1, 84))
            ind_name_cur = CStr(rw.Cells(1, 82))
            If dep_name_cur <> dep_name_prev Then
                isTheSame = False
                If ws_xml.Cells(rw.Row, 1).value <> "" Then
                    dep_name_prev = dep_name_cur
                    okei_code_prev = okei_code_cur
                    base_code_prev = base_code_cur
                    ind_name_prev = ind_name_cur
                End If
            Else
                isTheSame = True
            End If
            
            strRegNumber = RemoveWhiteSpace(CStr(rw.Cells(1, 81)))
            isGood = True
            If strRegNumber = "" Then
                isGood = False
            End If
            If isGood Then
                If Not isTheSame Then
                    Set Dprtm_record = XDoc.createElement("Dprtm_record")
                    Dprtm_records.appendChild Dprtm_record
                    
                    Set Dprtm_code = XDoc.createElement("Dprtm_code")
                    Dprtm_record.appendChild Dprtm_code
                    Dprtm_code.Text = CStr(rw.Cells(1, 89))
                    
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
                Value_1.Text = GetValueWithZero(rw.Cells(1, 5))
                'Value_2
                Set Value_2 = XDoc.createElement("Value_2")
                Vlm_indctr_record.appendChild Value_2
                Value_2.Text = GetValueWithZero(rw.Cells(1, 6))
                'Value_3
                Set Value_3 = XDoc.createElement("Value_3")
                Vlm_indctr_record.appendChild Value_3
                Value_3.Text = GetValueWithZero(rw.Cells(1, 7))
                'Kbk_code
                Set Kbk_code = XDoc.createElement("Kbk_code")
                Vlm_indctr_record.appendChild Kbk_code
                Kbk_code.Text = RemoveWhiteSpace(CStr(rw.Cells(1, 1)))
            End If
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
        If Trim(CStr(rw.Cells(1, 2).value)) = Trim(CStr(code)) Then
            name = CStr(rw.Cells(1, 1).value)
            Exit For
        End If
        If CStr(rw.Cells(1, 1).value) = "" Then
            Exit For
        End If
    Next rw

    FindNameOKEI = name
End Function
Public Function RemoveWhiteSpace(target As String) As String
    RemoveWhiteSpace = Replace(target, " ", "")
End Function
Public Function GetValueWithZero(value As String) As String
    value = CStr(value)
    value = RemoveWhiteSpace(value)
    If value = "" Then
        GetValueWithZero = 0
        Exit Function
    End If
    GetValueWithZero = value

End Function
Public Function GetFilePath(folder_path, FileName, TagStart, TagEnd As String) As String

    If InStr(FileName, ".") > 0 Then
        FileName = Left(FileName, InStr(FileName, ".") - 1)
    End If
    GetFilePath = folder_path & "\" & TagStart & "_" & FileName & "_" & TagEnd & ".xml"
    
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
            If CStr(ws.Cells(i, 3).value) = "" Then
                Exit For
            End If
            'ÔÐ = AVA = 1249 - 1 / val1 = J = 10
            ws_xml.Cells(StartRowCountD, 10).value = ws.Cells(i, 1249).value
            'ÈÍÇ = AVB = 1250 - 2 / val1 = K = 11
            ws_xml.Cells(StartRowCountD, 11).value = ws.Cells(i, 1250).value
            'ó÷ ëèòåðàòóðà = AVC = 1251 - 3 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ïðàêòèêà = AVD = 1252 - 4 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ïîâûø êâàëèô = AVE = 1253 - 5 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ìåäîñìîòðû = AVF = 1254 - 6 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ÊÓ = AVG = 1255 - 7 / val1 = L = 12
            ws_xml.Cells(StartRowCountD, 12).value = ws.Cells(i, 1255).value
            'ÑÍÈ = AVH = 1256 - 8 / val1 = M = 13
            ws_xml.Cells(StartRowCountD, 13).value = ws.Cells(i, 1256).value
            'ÑÎÖÄÈ = AVI = 1257 - 9 / val1 = N = 14
            ws_xml.Cells(StartRowCountD, 14).value = ws.Cells(i, 1257).value
            'ÔÐ2 = AVJ = 1258 - 10 / val1 = O = 15
            ws_xml.Cells(StartRowCountD, 15).value = ws.Cells(i, 1258).value
            'ÓÑ = AVK = 1259 - 11 / val1 = P = 16
            ws_xml.Cells(StartRowCountD, 16).value = ws.Cells(i, 1259).value
            'ÒÓ = AVL = 1260 - 12 / val1 = Q = 17
            ws_xml.Cells(StartRowCountD, 17).value = ws.Cells(i, 1260).value
            'ÎÒ 2 = AVM = 1261 - 13 / val1 = R = 18
            ws_xml.Cells(StartRowCountD, 18).value = ws.Cells(i, 1261).value
            'ÎÒ2 çï = AVN = 1262 - 14 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'íà÷èñëåíèÿ = AVO = 1263 - 15 / val1 = J
            'ws_xml.Cells(StartRowCountD, 10).Value = ws.Cells(i, 1249).Value
            'ÏÍÇ = AVP = 1264 - 16 / val1 = S = 19
            ws_xml.Cells(StartRowCountD, 19).value = ws.Cells(i, 1264).value
            'Values
            'Value1 = P = 16 / val1 = E = 5
            ws_xml.Cells(StartRowCountD, 5).value = ws.Cells(i, 16).value
            'Value2 = Q = 17 / val1 = F = 6
            'Value2 = R = 18 / val1 = G = 7
            
            'Org
            Dim org As Variant
            org = FindOrg(CStr(ws.Cells(i, 3).value), wb, 2, 50)
            ws_xml.Cells(StartRowCountD, 3).value = org(0)
            ws_xml.Cells(StartRowCountD, 4).value = org(1)
            ws_xml.Cells(StartRowCountD, 80).value = org(2)
            'Êîä ó÷ðåæäåíèÿ
            ws_xml.Cells(StartRowCountD, 89).value = org(3)
            
            'RegNumber
            Dim reg As Variant
            reg = FindRegNumber(CStr(ws.Cells(i, 13).value), wb, 2, 1000)
            'Íàèìåíîâàíèå óñëóãè
            ws_xml.Cells(StartRowCountD, 2).value = reg(0)
            'KBK
            ws_xml.Cells(StartRowCountD, 1).value = reg(9)
            'Íîìåð
            ws_xml.Cells(StartRowCountD, 81).value = CStr(ws.Cells(i, 13).value)
            'Íàèìåíîâàíèå îáúåìà
            ws_xml.Cells(StartRowCountD, 82).value = reg(2)
            'Åäåíèöà èçìåðåíèÿ îáúåìà
            ws_xml.Cells(StartRowCountD, 83).value = reg(3)
            'Áàçîâûé êîä
            ws_xml.Cells(StartRowCountD, 84).value = reg(4)
            'Íà÷àëî ðàáîò
            ws_xml.Cells(StartRowCountD, 87).value = reg(5)
            'Êîíåö ðàáîò
            ws_xml.Cells(StartRowCountD, 88).value = reg(6)
            'Êîä ÎÊÅÈ
            ws_xml.Cells(StartRowCountD, 85).value = reg(7)
            'Êîä Èíäèêàòîðà
            ws_xml.Cells(StartRowCountD, 86).value = reg(8)
            
            'Êáê
            If ws_xml.Cells(StartRowCountD, 1).value = "" Then
                ws_xml.Cells(StartRowCountD, 1).value = "0"
            End If
            If ws_xml.Cells(StartRowCountD, 82).value = "" Then
                ws_xml.Cells(StartRowCountD, 82).value = "0"
            End If
            If ws_xml.Cells(StartRowCountD, 2).value = "" Then
                ws_xml.Cells(StartRowCountD, 2).value = "0"
            End If
            
            StartRowCountD = StartRowCountD + 1
        End If
        

        
    Next i
    'çàïîëíÿåì ðåêâåçèòû 12 ñòðîê
    'Call FillXMLTableFromServiceTable(wb, ws_xml, ws, 2, 14)
    
End Sub
Public Function FindOrg(shortName As String, wb As Workbook, StartRowCountS, EndRowCount As Integer) As String()
    Dim ws As Worksheet
    Set ws = wb.Sheets("Ó÷ð")
    Dim returnVal(3) As String
    Index = 0
    For i = StartRowCountS To EndRowCount
        If UCase(Trim(CStr(ws.Cells(i, 7).value))) = UCase(Trim(CStr(shortName))) Then
            Index = i
            Exit For
        End If
        If CStr(ws.Cells(i, 3).value) = "" Then
            Exit For
        End If
    Next i
    If Index > 0 Then
        returnVal(0) = CStr(ws.Cells(Index, 3).value)
        returnVal(1) = CStr(ws.Cells(Index, 5).value)
        returnVal(2) = CStr(ws.Cells(Index, 6).value)
        returnVal(3) = CStr(ws.Cells(Index, 2).value)
    Else
        returnVal(0) = ""
        returnVal(1) = ""
        returnVal(2) = ""
        returnVal(3) = ""
    End If
    FindOrg = returnVal
End Function
Public Function FindRegNumber(name As String, wb As Workbook, StartRowCountS, EndRowCount As Integer) As String()
    Dim ws As Worksheet
    Set ws = wb.Sheets("RegNumbers")
    Dim returnVal(9) As String
    Index = 0
    For i = StartRowCountS To EndRowCount
        If UCase(Trim(CStr(ws.Cells(i, 2).value))) = UCase(Trim(CStr(name))) Then
            Index = i
            Exit For
        End If
        If CStr(ws.Cells(i, 2).value) = "" Then
            Exit For
        End If
    Next i
    If Index > 0 Then
        returnVal(0) = CStr(ws.Cells(Index, 1).value)
        returnVal(1) = CStr(ws.Cells(Index, 2).value)
        returnVal(2) = CStr(ws.Cells(Index, 3).value)
        returnVal(3) = CStr(ws.Cells(Index, 4).value)
        returnVal(4) = CStr(ws.Cells(Index, 5).value)
        returnVal(5) = CStr(ws.Cells(Index, 6).value)
        returnVal(6) = CStr(ws.Cells(Index, 7).value)
        returnVal(7) = CStr(ws.Cells(Index, 8).value)
        returnVal(8) = CStr(ws.Cells(Index, 9).value)
        returnVal(9) = CStr(ws.Cells(Index, 10).value)
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
        returnVal(9) = ""
    End If
    FindRegNumber = returnVal
End Function

Private Sub PackArchive(path As String, newPath As String, name As String)
    'Set ZIP = New Class1.ZipClass
    'If (ZIP.CreateArchive(path)) Then  ' ñòàðûé àðõèâ çàòèðàåòñÿ
    '    ZIP.CopyFolderToArchive newPath
    'End If
    Dim path_file_zip As String
    path_file_zip = newPath + "\" + name + ".zip"
    CreateArchive (path_file_zip)
    
    Dim FSO As Object
    Dim SourceFolder As Object
    Dim SubFolder As Object
    Dim FileItem As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.getfolder(path)
    
    For Each FileItem In SourceFolder.Files
        Call CopyFileToArchiv(path_file_zip, FileItem.path)
    Next FileItem
    
End Sub

Function CreateArchive(ZipArchivePath) As Boolean
    Dim Shell As Object
    Dim FileSystemObject As Object
    Dim ArchiveFolder As Object

     Set Shell = CreateObject("Shell.Application")
     Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

     ' Ïðîâåðêà íàëè÷èÿ ðàñøèðåíèÿ zip â ïîëíîì ïóòè-èìåíè ôàéëà
     If UCase(FileSystemObject.GetExtensionName(ZipArchivePath)) <> "ZIP" Then
          Exit Function
          
     End If
     
     ' Ñîçäàíèå ïóñòîãî zip àðõèâà
    Dim ZipFileHeader As String
    ZipFileHeader = "PK" & Chr(5) & Chr(6) & String(18, 0)
    FileSystemObject.OpenTextFile(ZipArchivePath, 2, True).Write ZipFileHeader
    Set ArchiveFolder = Shell.Namespace((ZipArchivePath))
    ' ïðîâåðêà ñîçäàíèÿ àðõèâà
    If Not (ArchiveFolder Is Nothing) Then CreateArchive = True
       
End Function
Sub CopyFileToArchiv(ZipName As String, FileName As String)
    ' ZipName - ïîëíûé ïóòü ê àðõèâó
    ' FileName - ïîëíûé ïóòü ê àðõèâèðóåìîìó ôàéëó
    Dim ShellApp As Object
    Dim DestFolder As Object

     Set ShellApp = CreateObject("Shell.Application")
     Set DestFolder = ShellApp.Namespace((ZipName))
     ' êîïèðóåìûé âûáðàííûé ôàéë â zip ïàïêó
     DestFolder.CopyHere (FileName)
     ' îæèäàåì îêîí÷àíèå ñæàòèÿ ôàéëà
     Do Until DestFolder.Items.Count = 1
          Sleep 100
     Loop

     Set ShellApp = Nothing

End Sub

