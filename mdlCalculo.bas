Attribute VB_Name = "mdlCalculo"
Option Explicit

'-------------------------------------------------------------------
' Comentários para o RIC - Registro de Identificação Civil
'-------------------------------------------------------------------
'
'  * O melhor que eu achei:
'
'
'http://ghiorzi.org/cgcancpf.htm#z
'
'  Texto no site: "... Eu presumo que seguirá a regra abaixo ..."
'
'-------------------------------------------------------------------
'
'  * Outros interessantes que eu achei:
'
'
'http://goncin.wordpress.com/2010/10/20/o-novo-registro-de-identidade-civil-ric-e-as-implicacoes-para-quem-e-desenvolvedor/
'
'http://www.cjdinfo.com.br/publicacao-calculo-digito-verificador
'
'-------------------------------------------------------------------
'
'  * O que é RIC:
'
'
'http://www.brasil.gov.br/para/servicos/documentacao/conheca-o-novo-registro-de-identidade-civil-ric
'
'http://pt.wikipedia.org/wiki/Registro_de_Identidade_Civil
'
'http://portalcapacitar.com.br/noticias/emissao-gratuita-do-registro-de-identidade-civil-ric-e-aprovada/
'
'http://www.arpensp.org.br/principal/index.cfm?tipo_layout=SISTEMA&url=noticia_mostrar.cfm&id=12816
'
'-------------------------------------------------------------------
'
'  * RIC é tema da última palestra do 2º Congresso
'
'
'http://www.iti.gov.br/noticias/indice-de-noticias/3569-ric-e-tema-da-ultima-palestra-do-2-congresso
'
'-------------------------------------------------------------------
'
'  * Problemas Técnicos com o RIC...
'
'
'http://g1.globo.com/bom-dia-brasil/noticia/2013/03/projeto-que-torna-novo-documento-de-identidade-gratuito-e-aprovado.html
'
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'
'***** IMAGENS NA INTERNET:
'
'-------------------------------------------------------------------
'http://brasil.vamoscurtir.com.br/2013/03/falha-no-sistema-de-seguranca-emperra.html
'
'                - "0000000001-9" -> CERTO! Combina com a validação "MOD11"
'
'-------------------------------------------------------------------
'http://portalcapacitar.com.br/noticias/emissao-gratuita-do-registro-de-identidade-civil-ric-e-aprovada/
'
'                - "0000000002-7" ->  CERTO! Combina com a validação "MOD11"
'
'-------------------------------------------------------------------
'http://portalintegracao.com/portal/2012/05/31/ric-saiba-como-vai-funcionar-o-novo-registro-de-identidade-civil/
'
'                - "0009404129-6" ->  CERTO! Combina com a validação "MOD11"
'
'-------------------------------------------------------------------
'http://www.minhainternetinteligente.com/app/cpf-rg-e-outros-documentos/pt-br?camp_id=5228&gclid=CLKElKLbgbgCFZPm7AodB08AgA
'
'                - "0000000005-9" ->  ERRO! O dígito verificador seria "1" se fosse "MOD11".
'
'-------------------------------------------------------------------
'http://www.linkatual.com/nova-carteira-identidade-registro-identidade-civil-ric.html
'
'                - "1234567890-2" -> ERRO! O dígito verificador seria "1" se fosse "MOD11".
'
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'          ***** REGRA PARA O CÁLCULO DO MOD11 *****
'-------------------------------------------------------------------
'
'http://www.banknote.com.br/module.htm
'
'http://pt.wikipedia.org/wiki/D%C3%ADgito_verificador'
'Observação: para o código de barras, sempre que o resto for 0, 1 ou 10, deverá ser utilizado o dígito 1.
'
'http://www.sefaz.ba.gov.br/contribuinte/informacoes_fiscais/doc_fiscal/calculodv.htm
'
'http://www.cjdinfo.com.br/publicacao-calculo-digito-verificador
'Cálculos variantes poderiam ocorrer, tal como substituir por uma letra, quando o cálculo do dígito final der 10 ou outro número escolhido.
'
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'   A REGRA QUE EU RECEBI TRATA DIFERENTE O RESTO!!!
'-------------------------------------------------------------------
'
'Extraído do Capítulo “SITDF.007.0001.02-Calculo-Dígitos-Verificadores.doc”
'
'1.3 - Cálculo genérico do módulo 11 (Pesos de 2 a 9):
'
'------------------------------------------------------
'Configuração: NNNNNNNNNNN -D
'onde,   NNNNNNNNNNN = Número Básico
'    D = Dígito Verificador
'
'O número básico pode variar conforme o tamanho do campo utilizado.
'------------------------------------------------------
'Módulo 11 - (a partir da direita) pelos pesos: 2 a 9
'------------------------------------------------------
'Exemplo: 23456789012-8 (Número básico com tamanho 11)
'------------------------------------------------------
'Cálculo do Dígito Verificador
'Número Básico: 2 3 4 5 6 7 8 9 0 1 2
'Pesos : 4 3 2 9 8 7 6 5 4 3 2
'(4x2 + 3x3 + 2x4 + 9x5 + 8x6 + 7x7 + 6x8 + 5x9 + 4x0 + 3x1 + 2x2) = 267
'267 : 11 = Resto 3
'11 - 3 = 8, então:
'Dígito Verificador = 8
'------------------------------------------------------
'Obs: Se resto 0 (zero) ou 1 (um), então DV = 0
'------------------------------------------------------
'
'------------------------------------------------------
' DÚVIDA: **** E SE FOR RESTO "10"?!?!?!?!?!?!?!?!?!?!?
'------------------------------------------------------
'REGISTRO DE IDENTIDADE CIVIL - RIC
'Nota importante: O cálculo do DV do RIC - Registro de Identidade Civil (nova carteira de identidade dos brasileiros), ainda não está claramente explicado. Eu presumo que seguirá a regra abaixo, mas ainda dependemos da confirmação, que virá a partir da emissão dos primeiros cartões, prevista para dezembro de 2010. Se o prezado leitor já tem RIC, peço o obséquio de conferir o seu DV na rotina abaixo e dizer-me se meu algoritmo funcionou.
'Saiba como se calcula o DV (Dígito Verificador) do Registro de Identidade Civil - RIC e veja o DV de qualquer número, utilizando a rotina abaixo. O DV corresponde ao resto da divisão por 11 do somatório da multiplicação de cada algarismo da base respectivamente por 9, 8, 7, 6, 5, 4, 3, 2, 9 e 8, a partir da unidade. Siga o exemplo abaixo:
'1  3  3  9  7  0  5  1  2  7
'x  x  x  x  x  x  x  x  x  x
'8  9  2  3  4  5  6  7  8  9
'----------------------------
'8+27+ 6+27+28+ 0+30+ 7+16+63 = 212÷11=19, com resto 3 (este é o DV).
'
'Nota:
'Se o resto for 10, o DV será "0".
'Introduza o número do RIC (exemplo: 1339705127)
'------------------------------------------------------


Public Function Mod_dig11(ByVal cVariavel As String) As String
Dim lRetorno As String
Dim nSoma As Integer
Dim nMult As Integer
Dim nIndice As Integer

    lRetorno = "0"
    nSoma = 0
    nMult = 2
    
    For nIndice = Len(cVariavel) To 1 Step -1
        nSoma = nSoma + (Asc(Mid(cVariavel, nIndice, 1)) - 48) * nMult

        If nMult = 9 Then
            nMult = 2
        Else
            nMult = nMult + 1
        End If

    Next
    
    nSoma = nSoma * 10
    nSoma = nSoma Mod 11
    
    If nSoma = 10 Then
        lRetorno = "0"
    Else
        lRetorno = Chr(nSoma + 48)
    End If
    
    Mod_dig11 = lRetorno

End Function

'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
' Comentários para o RG - Registro Geral
'-------------------------------------------------------------------
'Exemplos RG - Registro Geral :
'-------------------------------------------------------------------
'RG: 42.943.412-1
'RG: 42.943.425-X
'-------------------------------------------------------------------

Function CalculoDV11(strNumero As String) As String
'declara as variáveis
Dim intcontador, intnumero, intTotalNumero, intMultiplicador, intResto As Integer

    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
        CalculoDV11 = ""
        Exit Function
    End If
    
    'inicia o multiplicador
    intMultiplicador = 9
    
    'pega cada caracter do numero a partir da esquerda
    For intcontador = 1 To Len(strNumero)
    
        'extrai o caracter e multiplica pelo multiplicador
        intnumero = Val(Mid(strNumero, intcontador, 1)) * intMultiplicador
    
        'soma o resultado para totalização
        intTotalNumero = intTotalNumero + intnumero
    
        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)
    
     Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exceções (intResto = 10 então DV = "X")

    If intResto < 10 Then
        CalculoDV11 = Trim(Str(intResto))
    Else
        CalculoDV11 = "X"
    End If

End Function
