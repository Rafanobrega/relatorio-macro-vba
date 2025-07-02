Sub ImportarEFormatarRelatorio_Completo()

    Dim arquivoSelecionado As String
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsBase As Worksheet
    Dim ultimaLinha As Long, ultimaLinhaBase As Long
    Dim i As Long, colData As Long, colHora As Long, colDataReal As Long
    Dim dictBase As Object
    Dim chave As String

    Application.ScreenUpdating = False

    '=== Seleção de arquivo de dados de entrada ===
    ' Informação original substituída por confidencialidade
    arquivoSelecionado = Application.GetOpenFilename("Arquivos Excel (*.xlsx), *.xlsx", , "Selecione o arquivo de dados de entrada")
    If arquivoSelecionado = "Falso" Then Exit Sub

    Set wbOrigem = Workbooks.Open(arquivoSelecionado)
    Set wsOrigem = wbOrigem.Sheets(1)

    '=== Cria aba de destino com nome genérico ===
    Set wsDestino = ThisWorkbook.Sheets.Add
    wsDestino.Name = "RelatorioFormatado_" & Format(Now, "dd-MM-yyyy")

    '=== Copia dados ===
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, 3).End(xlUp).Row
    wsOrigem.Range("C1:G" & ultimaLinha).Copy Destination:=wsDestino.Cells(1, 1)
    wbOrigem.Close SaveChanges:=False

    '=== Abre base de dados externa ===
    ' Caminho real removido por confidencialidade
    Dim wbBase As Workbook
    Set wbBase = Workbooks.Open(ThisWorkbook.Path & "\BasePessoas.xlsx", ReadOnly:=True)
    Set wsBase = wbBase.Sheets(1)

    '=== Remove linha vazia se existir ===
    If Application.WorksheetFunction.CountA(wsDestino.Rows(1)) = 0 Then
        wsDestino.Rows(1).Delete
    End If

    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row

    '=== Define colunas ===
    colData = 2
    colHora = 3
    colDataReal = 6

    '=== Inserção de colunas adicionais ===
    Dim colRestaurante As Long
    colRestaurante = colDataReal

    wsDestino.Columns(colRestaurante).Insert Shift:=xlToRight
    wsDestino.Cells(1, colRestaurante).Value = "RESTAURANTE"

    colDataReal = colRestaurante + 1
    wsDestino.Columns(colDataReal).Insert Shift:=xlToRight
    wsDestino.Cells(1, colDataReal).Value = "DataReal"

    '=== Conversão de datas/horas e mapeamento de restaurantes ===
    For i = 2 To ultimaLinha
        Dim dataTexto As String, horaTexto As String
        Dim partes() As String
        Dim dataConvertida As Date, horaConvertida As Date
        Dim idtValor As Variant
        Dim nomeRestaurante As String

        dataTexto = Trim(wsDestino.Cells(i, colData).Text)
        horaTexto = Trim(wsDestino.Cells(i, colHora).Text)

        If InStr(dataTexto, ".") > 0 Then
            partes = Split(dataTexto, ".")
            If UBound(partes) = 2 Then
                dataConvertida = DateSerial(CInt(partes(2)), CInt(partes(1)), CInt(partes(0)))
            End If
        ElseIf IsDate(dataTexto) Then
            dataConvertida = CDate(dataTexto)
        End If

        If IsDate(horaTexto) Then
            horaConvertida = TimeValue(horaTexto)
            If Hour(horaConvertida) < 3 Then
                wsDestino.Cells(i, colDataReal).Value = dataConvertida - 1
            Else
                wsDestino.Cells(i, colDataReal).Value = dataConvertida
            End If
        Else
            wsDestino.Cells(i, colDataReal).Value = dataConvertida
        End If

        ' Mapeamento fictício de restaurante
        idtValor = wsDestino.Cells(i, 5).Value
        Select Case idtValor
            Case 1001: nomeRestaurante = "Restaurante A"
            Case 1002: nomeRestaurante = "Restaurante B"
            Case 1003: nomeRestaurante = "Restaurante C"
            Case Else: nomeRestaurante = "IDT não mapeado"
        End Select
        wsDestino.Cells(i, colRestaurante).Value = nomeRestaurante
    Next i

    '=== Cruzamento com base externa ===
    Dim colEmpresa As Long, colCNPJEmp As Long, colSubcon As Long, colCNPJSubcon As Long
    Dim colNomeDestino As Long, colCPFDestino As Long

    colEmpresa = colDataReal + 1
    colCNPJEmp = colEmpresa + 1
    colSubcon = colCNPJEmp + 1
    colCNPJSubcon = colSubcon + 1
    colNomeDestino = colCNPJSubcon + 1
    colCPFDestino = colNomeDestino + 1

    wsDestino.Cells(1, colEmpresa).Value = "EMPRESA"
    wsDestino.Cells(1, colCNPJEmp).Value = "CNPJ_EMPRESA"
    wsDestino.Cells(1, colSubcon).Value = "SUBCONTRATADA"
    wsDestino.Cells(1, colCNPJSubcon).Value = "CNPJ_SUBCONTRATADA"
    wsDestino.Cells(1, colNomeDestino).Value = "NOME"
    wsDestino.Cells(1, colCPFDestino).Value = "CPF"

    '=== Colunas da base (mapeamento fictício) ===
    Dim colLegajo As Long, colEmp As Long, colCNPJEmpBase As Long
    Dim colSubc As Long, colCNPJSubc As Long, colNomeBase As Long, colCPFBase As Long

    colLegajo = 1
    colEmp = 2
    colCNPJEmpBase = 3
    colSubc = 4
    colCNPJSubc = 5
    colNomeBase = 6
    colCPFBase = 7

    Set dictBase = CreateObject("Scripting.Dictionary")
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, colLegajo).End(xlUp).Row

    For i = 2 To ultimaLinhaBase
        Dim valorCelula As Variant
        valorCelula = wsBase.Cells(i, colLegajo).Value

        If Not IsEmpty(valorCelula) And Not IsError(valorCelula) And IsNumeric(valorCelula) Then
            chave = CStr(Fix(CDbl(valorCelula)))
            If Not dictBase.exists(chave) Then
                dictBase(chave) = Array( _
                    wsBase.Cells(i, colEmp).Value, _
                    wsBase.Cells(i, colCNPJEmpBase).Value, _
                    wsBase.Cells(i, colSubc).Value, _
                    wsBase.Cells(i, colCNPJSubc).Value, _
                    wsBase.Cells(i, colNomeBase).Value, _
                    wsBase.Cells(i, colCPFBase).Value _
                )
            End If
        End If
    Next i

    wbBase.Close SaveChanges:=False

    For i = 2 To ultimaLinha
        Dim valorRelatorio As Variant
        valorRelatorio = wsDestino.Cells(i, 1).Value

        If Not IsEmpty(valorRelatorio) And Not IsError(valorRelatorio) And IsNumeric(valorRelatorio) Then
            chave = CStr(Fix(CDbl(valorRelatorio)))
            If dictBase.exists(chave) Then
                wsDestino.Cells(i, colEmpresa).Value = dictBase(chave)(0)
                wsDestino.Cells(i, colCNPJEmp).Value = dictBase(chave)(1)
                wsDestino.Cells(i, colSubcon).Value = dictBase(chave)(2)
                wsDestino.Cells(i, colCNPJSubcon).Value = dictBase(chave)(3)
                wsDestino.Cells(i, colNomeDestino).Value = dictBase(chave)(4)
                wsDestino.Cells(i, colCPFDestino).Value = dictBase(chave)(5)
            End If
        End If
    Next i

    '=== Identificação do tipo de refeição ===
    Dim colRefeicao As Long
    colRefeicao = colCPFDestino + 1
    wsDestino.Cells(1, colRefeicao).Value = "Tipo de Refeição"

    For i = 2 To ultimaLinha
        Dim horaTextoRefeicao As String
        Dim horaConvertidaRefeicao As Date
        Dim tipoRefeicao As String

        horaTextoRefeicao = Trim(wsDestino.Cells(i, colHora).Text)

        If IsDate(horaTextoRefeicao) Then
            horaConvertidaRefeicao = TimeValue(horaTextoRefeicao)
            Select Case True
                Case horaConvertidaRefeicao >= TimeValue("06:00:00") And horaConvertidaRefeicao <= TimeValue("09:00:00")
                    tipoRefeicao = "Desjejum"
                Case horaConvertidaRefeicao >= TimeValue("11:00:00") And horaConvertidaRefeicao <= TimeValue("15:00:00")
                    tipoRefeicao = "Almoço"
                Case horaConvertidaRefeicao >= TimeValue("19:00:00") And horaConvertidaRefeicao <= TimeValue("21:00:00")
                    tipoRefeicao = "Jantar"
                Case (horaConvertidaRefeicao >= TimeValue("23:00:00") And horaConvertidaRefeicao <= TimeValue("23:59:59")) _
                   Or (horaConvertidaRefeicao >= TimeValue("00:00:00") And horaConvertidaRefeicao <= TimeValue("02:30:00"))
                    tipoRefeicao = "Ceia"
                Case Else
                    tipoRefeicao = "Fora do horário"
            End Select
            wsDestino.Cells(i, colRefeicao).Value = tipoRefeicao
        Else
            wsDestino.Cells(i, colRefeicao).Value = "Hora Inválida"
        End If
    Next i

    '=== Formatação visual geral ===
    With wsDestino
        .Activate
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        .Columns.AutoFit
        .Range(.Cells(1, 1), .Cells(1, .Cells(1, .Columns.Count).End(xlToLeft).Column)).Font.Bold = True
        .AutoFilterMode = True
    End With

    '=== Limpeza de valores nulos ===
    wsDestino.UsedRange.Replace What:="NULL", Replacement:="", LookAt:=xlWhole, MatchCase:=False

    Application.ScreenUpdating = True
    MsgBox "Relatório gerado e pronto para visualização!", vbInformation

End Sub
