Attribute VB_Name = "UpdateProjectReview"
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object

Function UpdateCover(wb As Workbook, wsCJI3 As Worksheet)
    UpdateComentarios wb, wsCJI3
End Function

Sub AtualizarTudo(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler
    
ErrorSection = "Initialization"

Dim temp As Double
temp = Timer

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim DR As String
    Dim IsProjectReview As Boolean
    Dim NoProjectReviewFound As Boolean
    Dim response As VbMsgBoxResult
    Dim PEPList() As String
    
    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True

ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit  ' Exit the function or sub
        End If
    Loop

ErrorSection = "WorkbookSearch"
    
    NoProjectReviewFound = True
    
    ReDim PEPList(0)
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsProjectReview = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
                    
                    ' Check if E2 value is "PEP"
                    If UCase(ws.Range("E2").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E3").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("G2").Value) = "DR" Then
                        DR = ws.Range("G3").Value
                    End If
                
                ElseIf InStr(1, LCase(ws.Name), "ata", vbTextCompare) > 0 Then
                
                    ' Check if E4 value is "PEP"
                    If UCase(ws.Range("E4").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E5").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("F4").Value) = "DR" Then
                        DR = ws.Range("F5").Value
                    End If
                    
                End If
                
                If PEP <> "" Then
                    Exit For ' Stop loop after finding the first matching sheet
                End If
                
            Next ws
        End If
    
        If Not IsProjectReview Then
            GoTo NextWorkbook
        Else
            NoProjectReviewFound = False
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
    
        If Not UpdateCJI3(wb, PEP) Then
            MsgBox "Não foi possível atualizar CJI3 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        If Not UpdateCJI5(wb, PEP) Then
            'MsgBox "Não foi possível atualizar CJI5 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        If Not UpdateZTPP092(wb, DR) Then
            MsgBox "Não foi possível atualizar ZTPP092 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        If Not UpdateMapa(wb) Then
            MsgBox "Não foi possível atualizar Mapa de Suprimentos de " & vbCrLf & wb.Name, vbInformation
        End If
        
        Application.StatusBar = False
        
        ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If NoProjectReviewFound Then
        MsgBox "Nenhum Project Review foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Project Reviews atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
    Debug.Print "Atualizar Tudo - Total execution time: "; Timer - temp
    
CleanExit:

    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarTudo"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Sub AtualizarMapa(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer

    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim DR As String
    Dim IsProjectReview As Boolean
    Dim NoProjectReviewFound As Boolean
    Dim response As VbMsgBoxResult
    
ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit  ' Exit the function or sub
        End If
    Loop
    
ErrorSection = "WorkbookSearch"

    NoProjectReviewFound = True
    
    ReDim PEPList(0)
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsProjectReview = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
                
                    ' Check if E2 value is "PEP"
                    If UCase(ws.Range("E2").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E3").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("G2").Value) = "DR" Then
                        IsProjectReview = True
                        DR = ws.Range("G3").Value
                    End If
                    
                ElseIf InStr(1, LCase(ws.Name), "ata", vbTextCompare) > 0 Then
                
                    ' Check if E4 value is "PEP"
                    If UCase(ws.Range("E4").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E5").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("F4").Value) = "DR" Then
                        DR = ws.Range("F5").Value
                    End If
        
                End If
                
                If PEP <> "" Then
                    Exit For ' Stop loop after finding the first matching sheet
                End If
                
            Next ws
        End If
    
        If Not IsProjectReview Then
            GoTo NextWorkbook
        Else
            NoProjectReviewFound = False
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
        
        If Not UpdateMapa(wb) Then
            MsgBox "Não foi possível atualizar Mapa de Suprimentos de " & vbCrLf & wb.Name, vbInformation
        End If
        
        ' UpdateCover wb, wsCJI3
        
        Application.StatusBar = False

        ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If NoProjectReviewFound Then
        MsgBox "Nenhum Project Review foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Project Reviews atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
Debug.Print "Atualizar Mapa - Total execution time: "; Timer - temp
    
CleanExit:

    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarMapa"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Sub AtualizarZTPP092(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer

    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim DR As String
    Dim IsProjectReview As Boolean
    Dim NoProjectReviewFound As Boolean
    Dim response As VbMsgBoxResult
    
ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit  ' Exit the function or sub
        End If
    Loop
    
ErrorSection = "WorkbookSearch"
    
    NoProjectReviewFound = True
    
    ReDim PEPList(0)
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsProjectReview = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
            
                    ' Check if E2 value is "PEP"
                    If UCase(ws.Range("E2").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E3").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("G2").Value) = "DR" Then
                        IsProjectReview = True
                        DR = ws.Range("G3").Value
                    End If
        
                ElseIf InStr(1, LCase(ws.Name), "ata", vbTextCompare) > 0 Then
                
                    ' Check if E4 value is "PEP"
                    If UCase(ws.Range("E4").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E5").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("F4").Value) = "DR" Then
                        DR = ws.Range("F5").Value
                    End If
        
                End If
                
                If PEP <> "" Then
                    Exit For ' Stop loop after finding the first matching sheet
                End If
                
            Next ws
        End If
    
        If Not IsProjectReview Then
            GoTo NextWorkbook
        Else
            NoProjectReviewFound = False
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
    
        If Not UpdateZTPP092(wb, DR) Then
            MsgBox "Não foi possível atualizar ZTPP092 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        ' UpdateCover wb, wsCJI3
        
        Application.StatusBar = False

        ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If NoProjectReviewFound Then
        MsgBox "Nenhum Project Review foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Project Reviews atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
    
Debug.Print "Atualizar ZTPP02 - Total execution time: "; Timer - temp
    
CleanExit:
    
    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarZTPP092"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Sub AtualizarCJI5(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer

    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim IsProjectReview As Boolean
    Dim NoProjectReviewFound As Boolean
    Dim response As VbMsgBoxResult

ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit  ' Exit the function or sub
        End If
    Loop
    
ErrorSection = "WorkbookSearch"
    
    NoProjectReviewFound = True
    
    ReDim PEPList(0)

    ' Loop through all open workbooks
    For Each wb In Workbooks
ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsProjectReview = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
            
                    ' Check if E2 value is "PEP"
                    If UCase(ws.Range("E2").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E3").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                ElseIf InStr(1, LCase(ws.Name), "ata", vbTextCompare) > 0 Then
                
                    ' Check if E4 value is "PEP"
                    If UCase(ws.Range("E4").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E5").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("F4").Value) = "DR" Then
                        DR = ws.Range("F5").Value
                    End If
        
                    Exit For ' Stop loop after finding the first matching sheet
                    
                End If
                
                If PEP <> "" Then
                    Exit For ' Stop loop after finding the first matching sheet
                End If
                
            Next ws
        End If
    
        If Not IsProjectReview Then
            GoTo NextWorkbook
        Else
            NoProjectReviewFound = False
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
    
        If Not UpdateCJI5(wb, PEP) Then
            'MsgBox "Não foi possível atualizar CJI5 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        ' UpdateCover wb, wsCJI3
        
        Application.StatusBar = False

        ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If NoProjectReviewFound Then
        MsgBox "Nenhum Project Review foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Project Reviews atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
Debug.Print "Atualizar CJI5 - Total execution time: "; Timer - temp
    
CleanExit:

    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarCJI5"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Sub AtualizarCJI3(Optional ShowOnMacroList As Boolean = False)

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer

    ' Otimiza o tempo de execução do código
    OptimizeCodeExecution True
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim PEP As String
    Dim IsProjectReview As Boolean
    Dim NoProjectReviewFound As Boolean
    Dim response As VbMsgBoxResult
    
ErrorSection = "SAPSetup"

    ' Setup SAP and check if it is running
    Do While Not SetupSAPScripting
        ' Ask the user to initiate SAP or cancel
        response = MsgBox("SAP não está acessível. Inicie o SAP e clique em OK para tentar novamente, ou Cancelar para sair.", vbOKCancel + vbExclamation, "Aguardando SAP")
    
        If response = vbCancel Then
            MsgBox "Execução terminada pelo usuário.", vbInformation
            GoTo SuccessefulExit  ' Exit the function or sub
        End If
    Loop
    
    NoProjectReviewFound = True
    
    ReDim PEPList(0)
    
ErrorSection = "WorkbookSearch"

    ' Loop through all open workbooks
    For Each wb In Workbooks
ErrorSection = "WorkbookSearchFor-" & wb.Name
        IsProjectReview = False
        
        ' Avoid checking the workbook where this code is running (optional)
        If wb.Name <> ThisWorkbook.Name Then
            ' Loop through all sheets in the workbook
            For Each ws In wb.Sheets
                If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
                
                    ' Check if E2 value is "PEP"
                    If UCase(ws.Range("E2").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E3").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                ElseIf InStr(1, LCase(ws.Name), "ata", vbTextCompare) > 0 Then
                
                    ' Check if E4 value is "PEP"
                    If UCase(ws.Range("E4").Value) = "PEP" Then
                        IsProjectReview = True
                        PEP = ws.Range("E5").Value
                        
                        ReDim Preserve PEPList(UBound(PEPList) + 1)  ' Resize array dynamically
                        PEPList(UBound(PEPList)) = PEP
                    End If
                    
                    ' Check if G2 value is "DR"
                    If UCase(ws.Range("F4").Value) = "DR" Then
                        DR = ws.Range("F5").Value
                    End If
        
                End If
                
                If PEP <> "" Then
                    Exit For ' Stop loop after finding the first matching sheet
                End If
                
            Next ws
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
    
        If Not IsProjectReview Then
            GoTo NextWorkbook
        Else
            NoProjectReviewFound = False
        End If
    
        Application.StatusBar = "Trabalhando em " & wb.Name
        
        If Not UpdateCJI3(wb, PEP) Then
            MsgBox "Não foi possível atualizar CJI3 de " & vbCrLf & wb.Name, vbInformation
        End If
        
        ws.Activate
        
NextWorkbook:
    Next wb
    
SuccessefulExit:
    EndSAPScripting
    
    If NoProjectReviewFound Then
        MsgBox "Nenhum Project Review foi encontrado.", vbInformation
    Else
        ' Join the PEP array into a string for display
        MsgBox "Project Reviews atualizados com sucesso:" & Join(PEPList, vbCrLf), vbInformation
    End If
    
Debug.Print "Atualizar CJI3 - Total execution time: "; Timer - temp
    
CleanExit:

    Application.StatusBar = False
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    
    Exit Sub

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in AtualizarCJI3"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Sub

Function UpdateMapa(wb As Workbook) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateMapa Start"

    Dim exportWb As Workbook
    Dim wbIter As Workbook ' Iterator for workbooks
    Dim Workbook As Workbook
    Dim wsZTPP092 As Worksheet
    Dim wsMapa As Worksheet
    Dim exportWs As Worksheet
    Dim Row As Range
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim StartDate As String
    Dim EndDate As String
    Dim ordem As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    Dim wsMapaLR As Long
    Dim exportWsLR As Long
    Dim currentRows As Long
    Dim requiredRows As Long
    Dim foundCell As Range
    Dim gerador As String

    On Error Resume Next
    Set wsMapa = wb.Sheets("Mapa de Suprimentos")
    On Error GoTo ErrorHandler

    If wsMapa Is Nothing Then
        UpdateMapa = False
        Exit Function
    End If
    
    gerador = wsMapa.Range("A2").Value

    If gerador = "" Then
        UpdateMapa = False
        Exit Function
    End If
    
    ' Find wsMapa last row and save to wsMapaLR
    If UCase(wsMapa.Cells(wsMapa.Cells(wsMapa.Rows.Count, "B").End(xlUp).Row, 2).Value) = UCase("SOBRA") Then
        wsMapaLR = wsMapa.Cells(wsMapa.Rows.Count, "B").End(xlUp).Row - 1
    Else
        wsMapaLR = wsMapa.Cells(wsMapa.Rows.Count, "B").End(xlUp).Row
    End If
    
    ' Name of the workbook to find
    exportWbName = "CS11-" & gerador

    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count
    
Debug.Print "Setup time: " & Timer - temp
temp = Timer

ErrorSection = "SAPNavigation"

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncs11"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").Text = gerador
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").Text = "1341"
    session.findById("wnd[0]/usr/txtRC29L-STLAL").Text = "1"
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").Text = "BEST"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler

    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text

    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = exportWbPath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[0]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop

    ' If exported file not found, open it
    If Not found Then
        Set exportWb = Workbooks.Open(exportWbPath & "\" & exportWbName)
    End If

ErrorSection = "ExportFormating"

    Set exportWs = exportWb.Sheets(1)
    
    ' Find exportWs last row and save to exportWsLR (using column C as reference)
    exportWsLR = exportWs.Cells(exportWs.Rows.Count, "C").End(xlUp).Row
    
    ' Delete dummy itens
    For Each Row In exportWs.Range("A2:A" & exportWsLR)
        If Row.Cells(1, 8).Value <> "" Then
            'Row.Delete
            Row.Cells(1, 5).Value = 0
        End If
    Next Row
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer

ErrorSection = "PasteData"

    ' Clear, copy and paste values and formats without breaking formulas and headers
    If wsMapa.AutoFilterMode Then wsMapa.AutoFilter.ShowAllData ' Clear any applied filters
    
    ' Compare groups line by line.
    ' Assumption: wsMapa data starts at row 5 and exportWs data starts at row 2.
    Dim wsMapaCurrentRow As Long, exportWsCurrentRow As Long
    Dim groupExportStart As Long, groupExportEnd As Long, groupExportCount As Long
    Dim groupMapaStart As Long, groupMapaEnd As Long, groupMapaCount As Long
    Dim lastGroupMapaStart As Long, lastGroupMapaEnd As Long
    Dim wsMapaGroupFound As Boolean
    
    wsMapaCurrentRow = 5
    exportWsCurrentRow = 2
    
    lastGroupMapaStart = wsMapaCurrentRow - 1
    lastGroupMapaEnd = wsMapaCurrentRow - 1
    
    ' Strikethrough cells from wsMapa
    wsMapa.Range("A" & wsMapaCurrentRow & ":C" & wsMapaLR).Font.Strikethrough = True
    
    Do While exportWsCurrentRow <= exportWsLR
ErrorSection = "PasteDataWhile-" & exportWsCurrentRow
        ' Identify group start in exportWs: a row with 0 in column C
        If Trim(exportWs.Cells(exportWsCurrentRow, "E").Value) = "0" Or exportWs.Cells(exportWsCurrentRow, "E").Value = 0 Then
ErrorSection = "ExportLimits-" & exportWsCurrentRow
            groupExportStart = exportWsCurrentRow
            groupExportEnd = groupExportStart
            ' Determine the end of this exportWs group:
            Do While groupExportEnd + 1 <= exportWsLR And Not (Trim(exportWs.Cells(groupExportEnd + 1, "E").Value) = "0" Or exportWs.Cells(groupExportEnd + 1, "E").Value = 0)
                groupExportEnd = groupExportEnd + 1
            Loop
            groupExportCount = groupExportEnd - groupExportStart + 1
            
ErrorSection = "MatchGroup-" & exportWsCurrentRow
            ' Look for a matching group in wsMapa
            wsMapaGroupFound = False
            For wsMapaCurrentRow = lastGroupMapaEnd + 1 To wsMapaLR
                If wsMapaCurrentRow <= wsMapaLR And (Trim(wsMapa.Cells(wsMapaCurrentRow, "C").Value) = "0" Or wsMapa.Cells(wsMapaCurrentRow, "C").Value = 0) And Trim(wsMapa.Cells(wsMapaCurrentRow, "A").Value) = Trim(exportWs.Cells(groupExportStart, "C").Value) Then
                    groupMapaStart = wsMapaCurrentRow
                    groupMapaEnd = groupMapaStart
                    ' Determine the end of this exportWs group:
                    Do While groupMapaEnd + 1 <= wsMapaLR And Not (Trim(wsMapa.Cells(groupMapaEnd + 1, "C").Value) = "0" Or wsMapa.Cells(groupMapaEnd + 1, "C").Value = 0)
                        groupMapaEnd = groupMapaEnd + 1
                    Loop
                    groupMapaCount = groupMapaEnd - groupMapaStart + 1
                    wsMapaGroupFound = True
                    Exit For
                End If
            Next wsMapaCurrentRow
            
            If Not wsMapaGroupFound Then
ErrorSection = "CreateGroup-" & exportWsCurrentRow
                ' Define the start
                wsMapaCurrentRow = lastGroupMapaEnd
                exportWsCurrentRow = groupExportStart
                groupMapaStart = lastGroupMapaEnd + 1
            
                Do While exportWsCurrentRow <= groupExportEnd
ErrorSection = "CreateGroupWhile-" & exportWsCurrentRow
                    ' They are different: insert a row below copying the existing row and fill the new row green
                    wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                    ' Option 1: Copy the original row as base
                    wsMapa.Rows(wsMapaCurrentRow).Copy
                    wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                    Application.CutCopyMode = False
                    ' Replace compared columns with exportWs values
                    wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                    wsMapa.Cells(wsMapaCurrentRow + 1, "B").Value = exportWs.Cells(exportWsCurrentRow, "D").Value
                    wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                    ' Fill the new row (green) for columns A to C
                    wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow + 1, "A"), wsMapa.Cells(wsMapaCurrentRow + 1, "C")).Font.Strikethrough = False
                    groupMapaEnd = wsMapaCurrentRow + 1 ' Adjust the last group row marker after inserting a row
                    wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
            ElseIf groupExportCount = 1 And groupExportCount < groupMapaCount Then
ErrorSection = "IfExportGroupSmaller-" & exportWsCurrentRow
                ' Reorder the groups to keep the same order as exportWs
                If groupMapaStart <> lastGroupMapaEnd + 1 Then
                    wsMapa.Rows(groupMapaStart & ":" & groupMapaEnd).Cut
                    wsMapa.Rows(lastGroupMapaEnd + 1).Insert Shift:=xlDown
                    Application.CutCopyMode = False
                    ' Update groupMapaStart and groupMapaEnd based on the new location.
                    groupMapaStart = lastGroupMapaEnd + 1
                    groupMapaEnd = groupMapaStart + groupMapaCount - 1
                End If
            
                ' Define the start
                wsMapaCurrentRow = groupMapaStart
                exportWsCurrentRow = groupExportStart
                ' This means a lone header wasn't found on wsMapa, so a header must be found without messing with the groupMapa found
                ' They are different: insert a row below copying the existing row and fill the new row green
                wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                ' Option 1: Copy the original row as base
                wsMapa.Rows(wsMapaCurrentRow).Copy
                wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                Application.CutCopyMode = False
                ' Replace compared columns with exportWs values
                wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                wsMapa.Cells(wsMapaCurrentRow + 1, "B").Value = exportWs.Cells(exportWsCurrentRow, "D").Value
                wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                ' Fill the new row (green) for columns A to C
                wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow, "A"), wsMapa.Cells(wsMapaCurrentRow, "C")).Font.Strikethrough = False
                groupMapaEnd = wsMapaCurrentRow ' Adjust the last group row marker after inserting a row
                wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                wsMapaCurrentRow = wsMapaCurrentRow + 1
                exportWsCurrentRow = exportWsCurrentRow + 1
            Else
ErrorSection = "IfExportGroupBigger-" & exportWsCurrentRow
                ' Reorder the groups to keep the same order as exportWs
                If groupMapaStart <> lastGroupMapaEnd + 1 Then
                    wsMapa.Rows(groupMapaStart & ":" & groupMapaEnd).Cut
                    wsMapa.Rows(lastGroupMapaEnd + 1).Insert Shift:=xlDown
                    Application.CutCopyMode = False
                    ' Update groupMapaStart and groupMapaEnd based on the new location.
                    groupMapaStart = lastGroupMapaEnd + 1
                    groupMapaEnd = groupMapaStart + groupMapaCount - 1
                End If
            
                ' Define the start
                wsMapaCurrentRow = groupMapaStart
                exportWsCurrentRow = groupExportStart
                
                ' Compare line by line the values from wsMapa column A and C and exportWs column C and E.
                ' Assumption: wsMapa data starts at row 5 and exportWs data starts at row 2.
                Do While wsMapaCurrentRow <= groupMapaEnd
ErrorSection = "IfExportGroupBiggerWhile-" & exportWsCurrentRow
                    If exportWsCurrentRow <= groupExportEnd Then
                        ' Compare wsMapa col A with exportWs col C and wsMapa col C with exportWs col E
                        If Trim(wsMapa.Cells(wsMapaCurrentRow, "A").Value) <> Trim(exportWs.Cells(exportWsCurrentRow, "C").Value) Or _
                           Trim(wsMapa.Cells(wsMapaCurrentRow, "C").Value) <> Trim(exportWs.Cells(exportWsCurrentRow, "E").Value) Then
                            ' They are different: insert a row below copying the existing row and fill the new row green
                            wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                            ' Option 1: Copy the original row as base
                            wsMapa.Rows(wsMapaCurrentRow).Copy
                            wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                            Application.CutCopyMode = False
                            ' Replace compared columns with exportWs values
                            wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                            wsMapa.Cells(wsMapaCurrentRow + 1, "B").Value = exportWs.Cells(exportWsCurrentRow, "D").Value
                            wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                            ' Fill the new row (green) for columns A to C
                            wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow + 1, "A"), wsMapa.Cells(wsMapaCurrentRow + 1, "C")).Font.Strikethrough = False
                            groupMapaEnd = groupMapaEnd + 1 ' Adjust the last group row marker after inserting a row
                            wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                            wsMapaCurrentRow = wsMapaCurrentRow + 1 ' Skip the newly inserted row
                        Else
                            ' Remove strikethrough cells from wsMapa
                            wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow, "A"), wsMapa.Cells(wsMapaCurrentRow, "C")).Font.Strikethrough = False
                        End If
                    End If
                    
                    ' Move pointer to next group element in both sheets
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
                
                ' Continue adding inexistent values if groupMapaCount < groupExportCount
                Do While exportWsCurrentRow <= groupExportEnd
ErrorSection = "IfExportSheetBiggerWhile-" & exportWsCurrentRow
                    wsMapaCurrentRow = groupMapaEnd
                    If exportWsCurrentRow <= groupExportEnd Then
                        ' They are different: insert a row below copying the existing row and fill the new row green
                        wsMapa.Rows(wsMapaCurrentRow + 1).Insert Shift:=xlDown
                        ' Option 1: Copy the original row as base
                        wsMapa.Rows(wsMapaCurrentRow).Copy
                        wsMapa.Rows(wsMapaCurrentRow + 1).PasteSpecial Paste:=xlPasteAll
                        Application.CutCopyMode = False
                        ' Replace compared columns with exportWs values
                        wsMapa.Cells(wsMapaCurrentRow + 1, "A").Value = exportWs.Cells(exportWsCurrentRow, "C").Value
                        wsMapa.Cells(wsMapaCurrentRow + 1, "B").Value = exportWs.Cells(exportWsCurrentRow, "D").Value
                        wsMapa.Cells(wsMapaCurrentRow + 1, "C").Value = exportWs.Cells(exportWsCurrentRow, "E").Value
                        ' Fill the new row (green) for columns A to C
                        wsMapa.Range(wsMapa.Cells(wsMapaCurrentRow + 1, "A"), wsMapa.Cells(wsMapaCurrentRow + 1, "C")).Font.Strikethrough = False
                        groupMapaEnd = groupMapaEnd + 1 ' Adjust the last group row marker after inserting a row
                        wsMapaLR = wsMapaLR + 1  ' Adjust the last row marker after inserting a row
                        wsMapaCurrentRow = wsMapaCurrentRow + 1 ' Skip the newly inserted row
                    End If
                    
                    ' Move pointer to next group element in both sheets
                    wsMapaCurrentRow = wsMapaCurrentRow + 1
                    exportWsCurrentRow = exportWsCurrentRow + 1
                Loop
            End If
            
            ' Move exportRow pointer past this group.
            exportWsCurrentRow = groupExportEnd + 1
        Else
            exportWsCurrentRow = exportWsCurrentRow + 1
        End If
        
        lastGroupMapaStart = groupMapaStart
        lastGroupMapaEnd = groupMapaEnd
    Loop

ErrorSection = "Ending"

    UpdateMapa = True

    ' Close the exported workbook without saving changes
    exportWb.Close SaveChanges:=False

    ' Delete the exported workbook file
    On Error Resume Next ' In case the file is not found or cannot be deleted
    Kill exportWbPath & "\" & exportWbName
    On Error GoTo ErrorHandler

    ' Cleanup
    Application.CutCopyMode = False
    
Debug.Print "Project Review Mapa de Suprimentos Sheet update: " & Timer - temp
temp = Timer

CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateMapa"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function UpdateZTPP092(wb As Workbook, DR As String) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateZTPP092 Start"

    Dim exportWb As Workbook
    Dim Workbook As Workbook
    Dim wsZTPP092 As Worksheet
    Dim exportWs As Worksheet
    Dim Row As Range
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim StartDate As String
    Dim EndDate As String
    Dim ordem As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    
    ' Check DR number
    If DR = "" Then
        MsgBox "O DR não foi encontrado. Verifique se G2 = DR e G3 contem o número do DR.", vbInformation
        UpdateZTPP092 = False
        Exit Function
    End If
    
    On Error Resume Next
    Set wsZTPP092 = wb.Sheets("ZTPP092")
    On Error GoTo ErrorHandler
    
    If wsZTPP092 Is Nothing Then
        ' If the sheet does not exist, create it as the last sheet
        Set wsZTPP092 = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsZTPP092.Name = "ZTPP092"
    End If
    
    ' Name of the workbook to find
    exportWbName = "ZTPP092-" & DR
    
    StartDate = Format(DateSerial(2000, 1, 1), "dd.mm.yyyy")
    EndDate = Format(DateSerial(2100, 12, 31), "dd.mm.yyyy")
    
    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count

Debug.Print "Setup time: " & Timer - temp
temp = Timer

ErrorSection = "SAPNavigation"

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZTPP092"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "HENCKE"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    
    Dim grid As Object
    Dim iRow As Long
    Dim searchValue As String
    Dim colName As String
    
    ' Set your search value and the column name (as defined in the grid)
    searchValue = "FALTANTES_PLA"
    colName = "VARIANT"   ' Replace with the actual column name
    
    ' Get the grid control
    Set grid = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    
    ' Loop through all rows
    For iRow = 0 To grid.RowCount - 1
        If grid.GetCellValue(iRow, colName) = searchValue Then
            ' When found, set the current cell to the matching row
            grid.CurrentCellRow = iRow
            ' Depending on your setup, SelectedRows may require a string
            grid.SelectedRows = CStr(iRow)
            ' Double-click the cell to perform the action
            grid.doubleClickCurrentCell
            Exit For   ' Exit the loop once the desired row is found
        End If
    Next iRow

    session.findById("wnd[0]/usr/ctxtS_NETWK-LOW").Text = DR
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = StartDate
    session.findById("wnd[0]/usr/ctxtS_ECKST-HIGH").Text = EndDate
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[46]").press
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler
    
    On Error GoTo EmptyZTPP092
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    On Error GoTo ErrorHandler
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

    If False Then
EmptyZTPP092:
        On Error GoTo ErrorHandler
ErrorSection = "EmptyZTPP092"
        UpdateZTPP092 = False
        GoTo CleanExit
    End If

ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop
    
    If Not found Then
        Set exportWb = Workbooks.Open(exportWbPath & "\" & exportWbName)
    End If
    
ErrorSection = "ExportFormating"

    Set exportWs = exportWb.Sheets(1)
    
    For Each Row In exportWs.Rows
        With Row
            If .Cells(1, 1) = "" And .Cells(1, 3) = "" Then
                Exit For
            ElseIf .Cells(1, 3) = "" Then
                ordem = .Cells(1, 1).Value
            ElseIf .Cells(1, 1) = "" Then
                .Cells(1, 1) = ordem
            End If
        End With
    Next Row
    
    For Each Row In exportWs.Rows
        With Row
            If .Cells(1, 1) = "" And .Cells(1, 2) = "" Then
                Exit For
            ElseIf .Cells(1, 3) = "" Then
                .Delete
            End If
        End With
    Next Row
    
    exportWs.Rows(1).Delete
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer
    
ErrorSection = "PasteData"

    ' Clear, copy and paste values and formats without breaking formulas and headers
    If wsZTPP092.AutoFilterMode Then wsZTPP092.AutoFilter.ShowAllData ' Clear any applied filters
    wsZTPP092.UsedRange.ClearContents
    exportWs.UsedRange.Copy
    wsZTPP092.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    ' Clear the clipboard
    Application.CutCopyMode = False
     
    ' Close the workbook without saving changes
    exportWb.Close SaveChanges:=False
    
    'Application.Wait (Now + TimeValue("00:00:03"))
    
    ' Delete the workbook file
    On Error Resume Next ' In case the file is not found or cannot be deleted
    Kill exportWbPath & "\" & exportWbName
    On Error GoTo ErrorHandler
    
    ' Cleanup
    Application.CutCopyMode = False
    
    UpdateZTPP092 = True
    
Debug.Print "Project Review ZTPP02 Sheet update: " & Timer - temp
temp = Timer
    
CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateZTPP092"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function UpdateCJI5(wb As Workbook, PEP As String) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateCJI5 Start"

    Dim exportWb As Workbook
    Dim Workbook As Workbook
    Dim wsCJI5 As Worksheet
    Dim exportWs As Worksheet
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim StartDate As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    Dim CJI5IsEmpty As Boolean
    
    CJI5IsEmpty = False
    
    On Error Resume Next
    Set wsCJI5 = wb.Sheets("CJI5 (Compromisso)")
    On Error GoTo ErrorHandler
    
    ' Check if the "CJI3" sheet exists
    If wsCJI5 Is Nothing Then
        UpdateCJI5 = False
        Exit Function
    End If
    
    ' Name of the workbook to find
    exportWbName = "CJI5-" & PEP
    
    ' Set end date
    StartDate = "01.01.2000"
    ' StartDate = Format(DateSerial(Year(Date), Month(Date), 1), "dd.mm.yyyy")
    
    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count

Debug.Print "Setup time: " & Timer - temp
temp = Timer
    
ErrorSection = "SAPNavigation"

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncji5"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
    session.findById("wnd[1]").sendVKey 0
    On Error GoTo ErrorHandler
    
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = PEP
    session.findById("wnd[0]/usr/ctxtR_OBDAT-LOW").Text = StartDate
    session.findById("wnd[0]/usr/ctxtR_OBDAT-HIGH").Text = "31.12.2100"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/WEG_SOLAR"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    On Error GoTo ErrorHandler
    
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo EmptyCJI5
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    On Error GoTo ErrorHandler

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

    If False Then
EmptyCJI5:
        On Error GoTo ErrorHandler
ErrorSection = "EmptyCJI5"
        UpdateCJI5 = False
        GoTo CleanExit
    End If

ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
    
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop
    
    ' Validate if the workbook was opened successfully
    If exportWb Is Nothing Then
        wsCJI5.UsedRange.ClearContents
        UpdateCJI5 = False
        Exit Function
    End If
    
    Set exportWs = exportWb.Sheets(1)

Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer

ErrorSection = "PasteData"

    ' Copy data from exportWs to wsCJI5
    If wsCJI5.AutoFilterMode Then wsCJI5.AutoFilter.ShowAllData ' Clear any applied filters
    wsCJI5.UsedRange.ClearContents
    exportWs.UsedRange.Copy
    wsCJI5.UsedRange.PasteSpecial
    
    ' Cleanup
    Application.CutCopyMode = False
    exportWb.Close False  ' Close the exported workbook without saving
    
    UpdateCJI5 = True

Debug.Print "Project Review CJI5 Sheet update: " & Timer - temp
temp = Timer
    
CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateCJI5"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function UpdateCJI3(wb As Workbook, PEP As String) As Boolean

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

Dim temp As Double
temp = Timer
Debug.Print "UpdateCJI3 Start"

    Dim exportWb As Workbook
    Dim Workbook As Workbook
    Dim wsCJI3 As Worksheet
    Dim exportWs As Worksheet
    Dim exportWbName As String
    Dim exportWbPath As String
    Dim EndDate As String
    Dim attempt As Long
    Dim found As Boolean
    Dim wbCount As Long
    
    On Error Resume Next
    Set wsCJI3 = wb.Sheets("CJI3")
    On Error GoTo ErrorHandler
    
    ' Check if the "CJI3" sheet exists
    If wsCJI3 Is Nothing Then
        UpdateCJI3 = False
        Exit Function
    End If
    
    ' Name of the workbook to find
    exportWbName = "CJI3-" & PEP
    
    ' Set end date
    EndDate = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "dd.mm.yyyy")
    
    ' Capture initial workbook count
    wbCount = Application.Workbooks.Count

ErrorSection = "SAPNavigation"

Debug.Print "Setup time: " & Timer - temp
temp = Timer

    ' SAP Navigation and Export
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncji3"
    session.findById("wnd[0]").sendVKey 0
    
    On Error Resume Next
    session.findById("wnd[1]/usr/ctxtTCNT-PROF_DB").Text = "000000000001"
    session.findById("wnd[1]").sendVKey 0
    On Error GoTo ErrorHandler
    
    ' Clear other fields
    session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PROJN-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_NETNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_ACTVT-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCN_MATNR-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtR_KSTAR-HIGH").Text = ""
    
    ' Search PEP
    session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").Text = PEP
    session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").Text = "01.11.2000"
    session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").Text = EndDate
    session.findById("wnd[0]/usr/ctxtP_DISVAR").Text = "/CUSTO_CIDIO"
    session.findById("wnd[0]/usr/ctxtP_DISVAR").SetFocus
    session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo EmptyCJI3
    session.findById("wnd[0]/tbar[1]/btn[43]").press
    On Error GoTo ErrorHandler
    
    ' Close the file extension pop-up
    On Error Resume Next
    If session.findById("wnd[1]/usr/ctxtDY_FILENAME") Is Nothing Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    On Error GoTo ErrorHandler
    
    exportWbName = Replace(session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text, "export", exportWbName)
    exportWbPath = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
    
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = exportWbName
    session.findById("wnd[1]/tbar[0]/btn[11]").press

Debug.Print "SAP nav: " & Timer - temp
temp = Timer

    If False Then
EmptyCJI3:
        On Error GoTo ErrorHandler
ErrorSection = "EmptyCJI3"
        UpdateCJI3 = False
        GoTo CleanExit
    End If
    
ErrorSection = "ExportWorkbook"

    ' Wait for a new workbook to appear
    Do
        If Application.Workbooks.Count > wbCount Then
            ' Name of the workbook to find
            found = False
            
            ' Loop through all open workbooks
            For Each Workbook In Application.Workbooks
                If UCase(Workbook.Name) = UCase(exportWbName) Then
                    Set exportWb = Workbook
                    found = True
                    Exit For
                End If
            Next Workbook
            
            Exit Do
        End If
        
        DoEvents
    Loop
    
    ' Validate if the workbook was opened successfully
    If exportWb Is Nothing Then
        wsCJI3.UsedRange.ClearContents
        UpdateCJI3 = False
        Exit Function
    End If
    
    Set exportWs = exportWb.Sheets(1)
    
Debug.Print "Export sheet treatment: " & Timer - temp
temp = Timer
    
ErrorSection = "PasteData"

    ' Clear, copy and paste data from exportWs to wsCJI3
    If wsCJI3.AutoFilterMode Then wsCJI3.AutoFilter.ShowAllData ' Clear any applied filters
    wsCJI3.UsedRange.ClearContents
    exportWs.UsedRange.Copy
    wsCJI3.UsedRange.PasteSpecial
    
    ' Ensure columns A and B are converted to numbers
    With wsCJI3
        .Columns("A:B").NumberFormat = "0"  ' Set format to number
        .Columns("A:B").Value = .Columns("A:B").Value  ' Convert text to numbers
    End With
    
    ' Cleanup
    Application.CutCopyMode = False
    exportWb.Close False  ' Close the exported workbook without saving

Debug.Print "Project Review CJI3 Sheet update: " & Timer - temp
temp = Timer

ErrorSection = "UpdateComentarios"

    UpdateComentarios wb, wsCJI3
    
Debug.Print "Project Review comments update: " & Timer - temp
temp = Timer
    
    UpdateCJI3 = True
    
CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateCJI3"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function
Function UpdateComentarios(wb As Workbook, wsCJI3 As Worksheet)
    
    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

    Dim lastCol As Long, col As Long
    Dim lastRow As Long
    Dim commentRange As Range
    Dim headerRow As Long
    Dim searchCol As Long
    Dim cell As Range
    Dim HeaderRowNotFound As Boolean
    Dim resposta As VbMsgBoxResult
    
    ' Set references to relevant worksheets
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        If InStr(1, LCase(ws.Name), "project review", vbTextCompare) > 0 Then
            Exit For
        End If
    Next ws
    
    ' Define rows and columns
    headerRow = 20
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row ' Adjusted based on column C
    
    HeaderRowNotFound = True
    
    resposta = Empty
    
    Application.CalculateFull
    
ErrorSection = "WriteComments"

    ' Loop through all columns from L to the last column
    For col = 13 To lastCol  ' Column M = 13
ErrorSection = "WriteCommentsCol-" & col
        If InStr(1, ws.Cells(headerRow, col).Value, "Comentário", vbTextCompare) > 0 Then
            HeaderRowNotFound = False
            
            ' Loop through each row below the header to apply the formula logic
            For Each cell In ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(lastRow, col))
ErrorSection = "WriteCommentsColFor-" & cell.Row
                
                Dim result As String
                
                If cell.Offset(0, -1).Value <> 0 And cell.Value = "" Then
                
                ' Apply the equivalent logic of the complex formula here:
                result = BuildCommentText(ws, wsCJI3, cell)
                
                ' Insert value and remove formula
                cell.Value = result
                
                ElseIf cell.Offset(0, -1).Value <> 0 And cell.Value <> "" And cell.Column < 15 Then
                
                    If IsEmpty(resposta) Then
                        resposta = MsgBox("Comentário do mês de" & Replace(ws.Cells(headerRow, col).Value, "Comentário", "") & _
                                " atividade " & ws.Cells(cell.Row, 6).Value & " já existe. Deseja atualizar todos os comentários?", vbYesNo)
                    End If
                    
                    If resposta = vbYes Then
                        
                        ' Apply the equivalent logic of the complex formula here:
                        result = BuildCommentText(ws, wsCJI3, cell)
                        
                        ' Insert value and remove formula
                        cell.Value = result
                    End If
                End If
            Next cell
        End If
    Next col
    
    If HeaderRowNotFound Then
        MsgBox "Os cabeçalhos do Project Review não estão na linha 20. Verifique a posição e tente novamente.", vbInformation
    End If
    
CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in UpdateComentarios"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

' Custom function to replicate the logic from the formula
Private Function BuildCommentText(ws As Worksheet, wsCJI3 As Worksheet, dataCell As Range) As String

    ' Enable error handling
    Dim ErrorSection As String
    On Error GoTo ErrorHandler

ErrorSection = "Initialization"

    Dim output As String
    Dim searchRange As Range, matchCell As Range
    
    Set searchRange = wsCJI3.Range("B:B")

ErrorSection = "WriteComment"

    If dataCell.Row < 42 And ws.Cells(dataCell.Row, 3) <> "" Then
    
        ' Loop through the range to build the combined text
        For Each matchCell In searchRange
ErrorSection = "WriteCommentIfFor-" & matchCell.Row
            If matchCell.Offset(1, 0).Value = "" And matchCell.Value = "" Then
                Exit For
            End If
            
            If matchCell.Value = ws.Cells(18, dataCell.Column).Value And _
               matchCell.Offset(0, -1).Value = ws.Cells(19, dataCell.Column).Value Then
                
                If matchCell.Offset(0, 5).Value Like "* 0" & ws.Cells(dataCell.Row, 3).Value & "*" Or matchCell.Offset(0, 5).Value Like "* " & ws.Cells(dataCell.Row, 3).Value & "*" Then
                
                    ' Constructing text based on conditions
                    output = output & matchCell.Offset(0, 21).Value & _
                             " - R$ " & matchCell.Offset(0, 13).Value & vbNewLine
                
                End If
            End If
        Next matchCell
    
    ElseIf ws.Cells(dataCell.Row, 3) <> "" Then
    
        ' Loop through the range to build the combined text
        For Each matchCell In searchRange
ErrorSection = "WriteCommentElseFor-" & matchCell.Row
            If matchCell.Offset(1, 0).Value = "" And matchCell.Value = "" Then
                Exit For
            End If
            
            If matchCell.Value = ws.Cells(18, dataCell.Column).Value And _
               matchCell.Offset(0, -1).Value = ws.Cells(19, dataCell.Column).Value Then
                
                If matchCell.Offset(0, 5).Value Like "* 0" & ws.Cells(dataCell.Row, 3).Value & "*" Or matchCell.Offset(0, 5).Value Like "* " & ws.Cells(dataCell.Row, 3).Value & "*" Then
                
                    ' Constructing text based on conditions
                    output = output & matchCell.Offset(0, 6).Value & " - " & matchCell.Offset(0, 7).Value & " - " & matchCell.Offset(0, 8).Value & " Un. - R$" & matchCell.Offset(0, 13).Value & vbNewLine
                
                End If
            End If
        Next matchCell
    
    End If
    
    BuildCommentText = output

CleanExit:
    
    Exit Function

ErrorHandler:
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Error " & Err.Number & " at section " & ErrorSection & ": " & Err.Description, vbCritical, "Error in BuildCommentText"
    
    ' Resume cleanup to ensure that settings are restored
    Resume CleanExit
End Function

Function SetupSAPScripting() As Boolean
    
    ' Create the SAP GUI scripting engine object
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    On Error GoTo ErrorHandler
    
    If Not IsObject(SapGuiAuto) Or SapGuiAuto Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    On Error Resume Next
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    On Error GoTo ErrorHandler
    
    If Not IsObject(SAPApplication) Or SAPApplication Is Nothing Then
        SetupSAPScripting = False
        Exit Function
    End If
    
    ' Get the first connection and session
    On Error GoTo ErrorHandler
    Set Connection = SAPApplication.Children(0)
    Set session = Connection.Children(0)
    On Error GoTo ErrorHandler
    
    SetupSAPScripting = True
    
    If False Then
ErrorHandler:
    SetupSAPScripting = False
    End If
    
End Function

Function EndSAPScripting()
    ' Clean up
    Set session = Nothing
    Set Connection = Nothing
    Set SAPApplication = Nothing
    Set SapGuiAuto = Nothing
End Function

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function


