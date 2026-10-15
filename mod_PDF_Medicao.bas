Attribute VB_Name = "mod_PDF_Medicao"
Option Explicit

'==========================================================================
'  GERADOR DE PDF DAS ABAS DA MEDIÇÃO (BM)
'
'  Macros para rodar (Alt+F8):
'    BM_ConfigurarEGerarPDFs  -> ajusta a impressão de cada aba e gera os PDFs
'    BM_ApenasConfigurar      -> só ajusta a impressão (para conferir antes)
'    BM_ApenasGerarPDFs       -> gera os PDFs com a configuração atual
'
'  Quais abas:
'    - Se você selecionar várias abas (Ctrl+clique nas guias) -> só essas
'    - Se só uma aba estiver selecionada -> todas as abas VISÍVEIS
'    (abas ocultas nunca entram)
'
'  Lógica de escala:
'    Toda aba sai em PÁGINA ÚNICA (melhor orientação), exceto as "longas"
'    (as que precisariam de mais de ~2 páginas de altura), que saem com
'    1 página de largura e várias de altura. ABAS_PAGINA_UNICA força a
'    página única pelo nome. Papel: A4; passa para A3 só se no A4 a fonte
'    efetiva (fonte x escala) ficar abaixo de FONTE_MIN_PT.
'==========================================================================

'---------------------------- CONFIGURAÇÕES -------------------------------
Private Const FONTE_MIN_PT As Double = 7#            ' fonte "efetiva" desejada no PDF (fonte x escala)
Private Const FATOR_PAGINA_UNICA As Double = 0.5     ' só quebra em várias páginas se a página única
                                                     ' exigir menos da metade da escala "ajustada à largura"
                                                     ' (na prática: abas com mais de ~2 páginas de altura)
' Listas separadas por ";" (aceita * e ?). Nomes comparados sem espaços nas pontas.
Private Const ABAS_PAGINA_UNICA As String = "RESUMO_MED;MC_8*;MC_20*"
Private Const ABAS_VARIAS_PAGINAS As String = ""
Private Const RODAPE_PAG As String = "&8Página &P de &N"
Private Const PERMITIR_A3 As Boolean = True          ' False = tudo em A4
Private Const PAPEL_FIXO As String = ""              ' "" = automático | "A4" | "A3"
Private Const MARGEM_LR_CM As Double = 1#            ' margem esquerda/direita
Private Const MARGEM_TB_CM As Double = 1.5           ' margem superior/inferior
Private Const MARGEM_CAB_ROD_CM As Double = 0.6      ' margem do cabeçalho/rodapé
Private Const USAR_AREA_IMPRESSAO_EXISTENTE As Boolean = True
Private Const LIMPAR_QUEBRAS_MANUAIS As Boolean = True   ' apaga quebras de página inseridas à mão
Private Const NUMERAR_PAGINAS As Boolean = True      ' "Página X de Y" (só se o rodapé estiver vazio)
Private Const GERAR_PDF_POR_ABA As Boolean = True
Private Const GERAR_PDF_UNICO As Boolean = True      ' um PDF com todas as abas, na ordem das guias
Private Const ABRIR_PASTA_NO_FIM As Boolean = True
Private Const MAX_CELULAS_AMOSTRA As Long = 3000     ' células lidas para achar a fonte predominante
'--------------------------------------------------------------------------

Private Const PT_A4_L As Double = 595.28
Private Const PT_A4_A As Double = 841.89
Private Const PT_A3_L As Double = 841.89
Private Const PT_A3_A As Double = 1190.55

'=========================== MACROS PÚBLICAS ===============================
Public Sub BM_ConfigurarEGerarPDFs()
    Executar True, True
End Sub

Public Sub BM_ApenasConfigurar()
    Executar True, False
End Sub

Public Sub BM_ApenasGerarPDFs()
    Executar False, True
End Sub

'============================== PRINCIPAL =================================
Private Sub Executar(ByVal configurar As Boolean, ByVal exportar As Boolean)
    Dim wb As Workbook
    Dim abas As Collection
    Dim ws As Worksheet
    Dim i As Long
    Dim modo As String
    Dim pastaSaida As String
    Dim relatorio As String
    Dim linha As String
    Dim erros As Long
    Dim arq As String
    Dim nomes As Variant

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Set abas = AbasAlvo(wb, modo)
    If abas.Count = 0 Then
        MsgBox "Nenhuma aba visível encontrada.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Abas a processar: " & abas.Count & " (" & modo & ")" & vbCrLf & vbCrLf & _
              IIf(configurar, "- Ajustar papel/orientação/escala" & vbCrLf, "") & _
              IIf(exportar, "- Gerar PDFs" & vbCrLf, "") & vbCrLf & _
              "Continuar?", vbQuestion + vbYesNo, "PDF da Medição") <> vbYes Then Exit Sub

    If exportar Then
        pastaSaida = EscolherPasta(wb)
        If pastaSaida = "" Then Exit Sub
        pastaSaida = pastaSaida & "PDF " & LimparNome(NomeSemExtensao(wb.Name)) & "\"
        If Dir(pastaSaida, vbDirectory) = "" Then MkDir pastaSaida
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    wb.Activate
    abas(1).Select            ' desagrupa as abas antes de mexer no PageSetup

    relatorio = "RELATÓRIO - " & wb.Name & " - " & Format(Now, "dd/mm/yyyy hh:nn") & vbCrLf
    relatorio = relatorio & "Fonte mínima efetiva: " & FONTE_MIN_PT & " pt" & vbCrLf & String(90, "-") & vbCrLf

    For i = 1 To abas.Count
        Set ws = abas(i)
        Application.StatusBar = "Processando " & i & "/" & abas.Count & ": " & ws.Name
        If configurar Then
            linha = ConfigurarAba(ws)
        Else
            linha = Trim$(ws.Name) & " | configuração atual mantida"
        End If

        If exportar And GERAR_PDF_POR_ABA Then
            arq = pastaSaida & Format(i, "00") & " - " & LimparNome(ws.Name) & ".pdf"
            On Error Resume Next
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=arq, Quality:=xlQualityStandard, _
                IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            If Err.Number <> 0 Then
                linha = linha & " | ERRO ao gerar PDF: " & Err.Description
                erros = erros + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
        relatorio = relatorio & linha & vbCrLf
    Next i

    If exportar And GERAR_PDF_UNICO And abas.Count > 1 Then
        Application.StatusBar = "Gerando PDF único..."
        ReDim nomes(0 To abas.Count - 1)
        For i = 1 To abas.Count
            nomes(i - 1) = abas(i).Name
        Next i
        arq = pastaSaida & "00 - " & LimparNome(NomeSemExtensao(wb.Name)) & " (COMPLETO).pdf"
        On Error Resume Next
        wb.Worksheets(nomes).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=arq, Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number <> 0 Then
            relatorio = relatorio & "ERRO no PDF único: " & Err.Description & vbCrLf
            erros = erros + 1
            Err.Clear
        End If
        abas(1).Select
        On Error GoTo 0
    End If

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Debug.Print relatorio
    If exportar Then
        GravarTexto pastaSaida & "_relatorio.txt", relatorio
        If ABRIR_PASTA_NO_FIM Then Shell "explorer.exe """ & pastaSaida & """", vbNormalFocus
        MsgBox "Concluído. " & abas.Count & " aba(s), " & erros & " erro(s)." & vbCrLf & vbCrLf & _
               "Pasta: " & pastaSaida & vbCrLf & _
               "Detalhes (papel, escala, avisos) em _relatorio.txt", _
               IIf(erros = 0, vbInformation, vbExclamation), "PDF da Medição"
    Else
        MsgBox "Configuração aplicada em " & abas.Count & " aba(s)." & vbCrLf & _
               "Confira em Exibir > Visualização da Quebra de Página." & vbCrLf & _
               "Relatório na Janela Imediata do VBA (Ctrl+G).", vbInformation, "PDF da Medição"
    End If
End Sub

'======================== CONFIGURAÇÃO DE UMA ABA =========================
Private Function ConfigurarAba(ws As Worksheet) As String
    Dim rng As Range, ar As Range
    Dim larg As Double, alt As Double, fonte As Double, escMin As Double
    Dim mLR As Double, mTB As Double
    Dim papeis As Variant, p As Long
    Dim pw As Double, ph As Double, codPapel As Long, nomePapel As String
    Dim sRet As Double, sPai As Double, wRet As Double, wPai As Double
    Dim achou As Boolean, umaPag As Boolean
    Dim orient As Long, esc As Double
    Dim aviso As String, nPag As Long

    Set rng = AreaImpressao(ws)
    If rng Is Nothing Then
        ConfigurarAba = Trim$(ws.Name) & " | aba vazia - ignorada"
        Exit Function
    End If

    ' Dimensões reais (pt). Linhas/colunas ocultas já contam como zero.
    For Each ar In rng.Areas
        If ar.Width > larg Then larg = ar.Width
        alt = alt + ar.Height
    Next ar
    If larg <= 0 Or alt <= 0 Then
        ConfigurarAba = Trim$(ws.Name) & " | área de impressão sem dimensão - ignorada"
        Exit Function
    End If

    fonte = FontePredominante(rng)
    escMin = FONTE_MIN_PT / fonte
    If escMin > 1 Then escMin = 1

    mLR = Application.CentimetersToPoints(MARGEM_LR_CM)
    mTB = Application.CentimetersToPoints(MARGEM_TB_CM)

    Select Case UCase$(PAPEL_FIXO)
        Case "A4": papeis = Array("A4")
        Case "A3": papeis = Array("A3")
        Case Else
            If PERMITIR_A3 Then papeis = Array("A4", "A3") Else papeis = Array("A4")
    End Select

    Dim sPag As Double, oPag As Long, sLar As Double, oLar As Long
    Dim ultimo As Boolean, forcarUnica As Boolean, forcarVarias As Boolean, longa As Boolean

    forcarUnica = NaLista(ws.Name, ABAS_PAGINA_UNICA)
    forcarVarias = NaLista(ws.Name, ABAS_VARIAS_PAGINAS)

    For p = LBound(papeis) To UBound(papeis)
        DadosPapel CStr(papeis(p)), pw, ph, codPapel
        nomePapel = CStr(papeis(p))
        ultimo = (p = UBound(papeis))

        ' Escala para caber TUDO em uma página (melhor orientação)
        sRet = Menor3((pw - 2 * mLR) / larg, (ph - 2 * mTB) / alt, 1)
        sPai = Menor3((ph - 2 * mLR) / larg, (pw - 2 * mTB) / alt, 1)
        If sRet >= sPai Then sPag = sRet: oPag = xlPortrait Else sPag = sPai: oPag = xlLandscape

        ' Escala para caber só na LARGURA (várias páginas de altura)
        wRet = Menor3((pw - 2 * mLR) / larg, 1, 1)
        wPai = Menor3((ph - 2 * mLR) / larg, 1, 1)
        If wRet >= escMin Then
            sLar = wRet: oLar = xlPortrait      ' lista longa fica em pé
        Else
            sLar = wPai: oLar = xlLandscape
        End If

        ' Aba "longa": caber em uma página exigiria reduzir demais em relação a ajustar só a largura
        If forcarUnica Then
            longa = False
        ElseIf forcarVarias Then
            longa = True
        Else
            longa = (sPag < FATOR_PAGINA_UNICA * sLar)
        End If

        If Not longa Then
            ' PÁGINA ÚNICA. Fica no A4 se a fonte ficar legível; senão tenta o próximo papel.
            If sPag >= escMin Or ultimo Then
                umaPag = True: orient = oPag: esc = sPag: achou = True
                Exit For
            End If
        Else
            ' ABA LONGA: 1 página de largura, várias de altura
            If sLar >= escMin Or ultimo Then
                umaPag = False: orient = oLar: esc = sLar: achou = True
                Exit For
            End If
        End If
    Next p

    If esc * fonte < FONTE_MIN_PT - 0.05 Then
        aviso = " | AVISO: fonte efetiva ~" & Format(esc * fonte, "0.0") & " pt (abaixo de " & FONTE_MIN_PT & ")"
    End If

    ' ---- remove quebras de página manuais antigas (causavam páginas quase vazias)
    If LIMPAR_QUEBRAS_MANUAIS Then
        On Error Resume Next
        ws.ResetAllPageBreaks
        On Error GoTo 0
    End If

    ' ---- aplica (PrintCommunication acelera muito; existe a partir do Excel 2010)
    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo 0

    With ws.PageSetup
        If USAR_AREA_IMPRESSAO_EXISTENTE = False Or Len(.PrintArea) = 0 Then
            .PrintArea = rng.Address
        End If
        .Orientation = orient
        .LeftMargin = mLR
        .RightMargin = mLR
        .TopMargin = mTB
        .BottomMargin = mTB
        .HeaderMargin = Application.CentimetersToPoints(MARGEM_CAB_ROD_CM)
        .FooterMargin = Application.CentimetersToPoints(MARGEM_CAB_ROD_CM)
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesWide = 1
        If umaPag Then
            .FitToPagesTall = 1
        Else
            .FitToPagesTall = False
        End If
        ' Numeração só quando tiver mais de uma página; remove a que a versão anterior pôs
        If .RightFooter = RODAPE_PAG Then .RightFooter = ""
        If NUMERAR_PAGINAS And Not umaPag Then
            If Len(.RightFooter) = 0 And Len(.CenterFooter) = 0 Then .RightFooter = RODAPE_PAG
        End If
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    Err.Clear
    ws.PageSetup.PaperSize = codPapel
    If Err.Number <> 0 Or ws.PageSetup.PaperSize <> codPapel Then
        aviso = aviso & " | AVISO: a impressora padrão não aceitou " & nomePapel & _
                " (defina 'Microsoft Print to PDF' como padrão e rode de novo)"
    End If
    Err.Clear
    nPag = ws.PageSetup.Pages.Count
    If Err.Number <> 0 Then nPag = 0
    On Error GoTo 0

    ConfigurarAba = Trim$(ws.Name) & " | " & nomePapel & " " & IIf(orient = xlPortrait, "retrato", "paisagem") & _
        " | " & IIf(umaPag, "página única", "1 pág. de largura") & _
        IIf(nPag > 0, " | " & nPag & " pág.", "") & _
        " | escala ~" & Format(esc * 100, "0") & "%" & _
        " | fonte " & Format(fonte, "0.#") & " pt -> ~" & Format(fonte * esc, "0.0") & " pt" & _
        " | área " & rng.Address(False, False) & aviso
End Function

'=============================== AUXILIARES ===============================
Private Function AbasAlvo(wb As Workbook, ByRef modo As String) As Collection
    Dim col As New Collection
    Dim sh As Object

    If ActiveWorkbook Is wb And ActiveWindow.SelectedSheets.Count > 1 Then
        modo = "abas selecionadas"
        For Each sh In ActiveWindow.SelectedSheets
            If TypeOf sh Is Worksheet Then
                If sh.Visible = xlSheetVisible Then col.Add sh
            End If
        Next sh
    Else
        modo = "todas as abas visíveis"
        For Each sh In wb.Worksheets
            If sh.Visible = xlSheetVisible Then col.Add sh
        Next sh
    End If
    Set AbasAlvo = col
End Function

Private Function AreaImpressao(ws As Worksheet) As Range
    Dim pa As String
    Dim rUlt As Range, cUlt As Range, rPri As Range, cPri As Range

    If USAR_AREA_IMPRESSAO_EXISTENTE Then
        pa = ws.PageSetup.PrintArea
        If Len(pa) > 0 Then
            On Error Resume Next
            Set AreaImpressao = ws.Range(pa)
            If AreaImpressao Is Nothing Then Set AreaImpressao = ws.Range(Replace(pa, ";", ","))
            On Error GoTo 0
            If Not AreaImpressao Is Nothing Then Exit Function
        End If
    End If

    ' Sem área definida: do primeiro ao último conteúdo (valores visíveis)
    With ws.Cells
        Set rUlt = .Find(What:="*", After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If rUlt Is Nothing Then Exit Function
        Set cUlt = .Find(What:="*", After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
                         SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        Set rPri = .Find(What:="*", After:=.Cells(.Rows.Count, .Columns.Count), LookIn:=xlValues, _
                         LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
        Set cPri = .Find(What:="*", After:=.Cells(.Rows.Count, .Columns.Count), LookIn:=xlValues, _
                         LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    End With
    Set AreaImpressao = ws.Range(ws.Cells(rPri.Row, cPri.Column), ws.Cells(rUlt.Row, cUlt.Column))
End Function

' Fonte mais frequente entre as células preenchidas (amostra limitada).
Private Function FontePredominante(rng As Range) As Double
    Dim d As Object
    Dim cel As Range
    Dim v As Variant, sz As Variant, k As Variant
    Dim n As Long, melhorN As Long, melhor As Double

    Set d = CreateObject("Scripting.Dictionary")
    For Each cel In rng.Cells
        v = cel.Value2
        If Not IsEmpty(v) Then
            If IsError(v) Then
                sz = cel.Font.Size
            ElseIf Len(CStr(v)) > 0 Then
                sz = cel.Font.Size
            Else
                sz = Null
            End If
            If Not IsNull(sz) Then
                If sz > 0 Then
                    d(CStr(sz)) = d(CStr(sz)) + 1
                    n = n + 1
                    If n >= MAX_CELULAS_AMOSTRA Then Exit For
                End If
            End If
        End If
    Next cel

    For Each k In d.Keys
        If d(k) > melhorN Then melhorN = d(k): melhor = CDbl(k)
    Next k
    If melhor <= 0 Then melhor = 10
    FontePredominante = melhor
End Function

Private Sub DadosPapel(ByVal nome As String, ByRef largPt As Double, ByRef altPt As Double, ByRef codigo As Long)
    If nome = "A3" Then
        largPt = PT_A3_L: altPt = PT_A3_A: codigo = xlPaperA3
    Else
        largPt = PT_A4_L: altPt = PT_A4_A: codigo = xlPaperA4
    End If
End Sub

Private Function NaLista(ByVal nome As String, ByVal lista As String) As Boolean
    Dim itens As Variant, i As Long, it As String
    If Len(Trim$(lista)) = 0 Then Exit Function
    nome = UCase$(Trim$(nome))
    itens = Split(lista, ";")
    For i = LBound(itens) To UBound(itens)
        it = UCase$(Trim$(itens(i)))
        If Len(it) > 0 Then
            If nome Like it Then NaLista = True: Exit Function
        End If
    Next i
End Function

Private Function Menor3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Menor3 = a
    If b < Menor3 Then Menor3 = b
    If c < Menor3 Then Menor3 = c
End Function

Private Function EscolherPasta(wb As Workbook) As String
    Dim base As String
    base = wb.Path
    If base = "" Or LCase$(Left$(base, 4)) = "http" Then base = Environ$("USERPROFILE") & "\Desktop"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Escolha a pasta onde salvar os PDFs"
        .InitialFileName = base & "\"
        If .Show <> -1 Then Exit Function
        EscolherPasta = .SelectedItems(1)
    End With
    If Right$(EscolherPasta, 1) <> "\" Then EscolherPasta = EscolherPasta & "\"
End Function

Private Function LimparNome(ByVal s As String) As String
    Dim proibidos As Variant, i As Long
    proibidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(proibidos) To UBound(proibidos)
        s = Replace(s, proibidos(i), "_")
    Next i
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    LimparNome = s
End Function

Private Function NomeSemExtensao(ByVal s As String) As String
    If InStrRev(s, ".") > 0 Then s = Left$(s, InStrRev(s, ".") - 1)
    NomeSemExtensao = s
End Function

Private Sub GravarTexto(ByVal caminho As String, ByVal texto As String)
    Dim f As Integer
    f = FreeFile
    Open caminho For Output As #f
    Print #f, texto
    Close #f
End Sub
