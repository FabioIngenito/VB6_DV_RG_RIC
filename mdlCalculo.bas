Attribute VB_Name = "mdlCalculo"
Option Explicit

'-------------------------------------------------------------------
' Coment�rios para o RIC - Registro de Identifica��o Civil
'-------------------------------------------------------------------
'
'  * O melhor que eu achei:
'
'
'http://ghiorzi.org/cgcancpf.htm#z
'
'  Texto no site: "... Eu presumo que seguir� a regra abaixo ..."
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
'  * O que � RIC:
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
'  * RIC � tema da �ltima palestra do 2� Congresso
'
'
'http://www.iti.gov.br/noticias/indice-de-noticias/3569-ric-e-tema-da-ultima-palestra-do-2-congresso
'
'-------------------------------------------------------------------
'
'  * Problemas T�cnicos com o RIC...
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
'                - "0000000001-9" -> CERTO! Combina com a valida��o "MOD11"
'
'-------------------------------------------------------------------
'http://portalcapacitar.com.br/noticias/emissao-gratuita-do-registro-de-identidade-civil-ric-e-aprovada/
'
'                - "0000000002-7" ->  CERTO! Combina com a valida��o "MOD11"
'
'-------------------------------------------------------------------
'http://portalintegracao.com/portal/2012/05/31/ric-saiba-como-vai-funcionar-o-novo-registro-de-identidade-civil/
'
'                - "0009404129-6" ->  CERTO! Combina com a valida��o "MOD11"
'
'-------------------------------------------------------------------
'http://www.minhainternetinteligente.com/app/cpf-rg-e-outros-documentos/pt-br?camp_id=5228&gclid=CLKElKLbgbgCFZPm7AodB08AgA
'
'                - "0000000005-9" ->  ERRO! O d�gito verificador seria "1" se fosse "MOD11".
'
'-------------------------------------------------------------------
'http://www.linkatual.com/nova-carteira-identidade-registro-identidade-civil-ric.html
'
'                - "1234567890-2" -> ERRO! O d�gito verificador seria "1" se fosse "MOD11".
'
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'          ***** REGRA PARA O C�LCULO DO MOD11 *****
'-------------------------------------------------------------------
'
'http://www.banknote.com.br/module.htm
'
'http://pt.wikipedia.org/wiki/D%C3%ADgito_verificador'
'Observa��o: para o c�digo de barras, sempre que o resto for 0, 1 ou 10, dever� ser utilizado o d�gito 1.
'
'http://www.sefaz.ba.gov.br/contribuinte/informacoes_fiscais/doc_fiscal/calculodv.htm
'
'http://www.cjdinfo.com.br/publicacao-calculo-digito-verificador
'C�lculos variantes poderiam ocorrer, tal como substituir por uma letra, quando o c�lculo do d�gito final der 10 ou outro n�mero escolhido.
'
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'   A REGRA QUE EU RECEBI TRATA DIFERENTE O RESTO!!!
'-------------------------------------------------------------------
'
'Extra�do do Cap�tulo �SITDF.007.0001.02-Calculo-D�gitos-Verificadores.doc�
'
'1.3 - C�lculo gen�rico do m�dulo 11 (Pesos de 2 a 9):
'
'------------------------------------------------------
'Configura��o: NNNNNNNNNNN -D
'onde,   NNNNNNNNNNN = N�mero B�sico
'    D = D�gito Verificador
'
'O n�mero b�sico pode variar conforme o tamanho do campo utilizado.
'------------------------------------------------------
'M�dulo 11 - (a partir da direita) pelos pesos: 2 a 9
'------------------------------------------------------
'Exemplo: 23456789012-8 (N�mero b�sico com tamanho 11)
'------------------------------------------------------
'C�lculo do D�gito Verificador
'N�mero B�sico: 2 3 4 5 6 7 8 9 0 1 2
'Pesos : 4 3 2 9 8 7 6 5 4 3 2
'(4x2 + 3x3 + 2x4 + 9x5 + 8x6 + 7x7 + 6x8 + 5x9 + 4x0 + 3x1 + 2x2) = 267
'267 : 11 = Resto 3
'11 - 3 = 8, ent�o:
'D�gito Verificador = 8
'------------------------------------------------------
'Obs: Se resto 0 (zero) ou 1 (um), ent�o DV = 0
'------------------------------------------------------
'
'------------------------------------------------------
' D�VIDA: **** E SE FOR RESTO "10"?!?!?!?!?!?!?!?!?!?!?
'------------------------------------------------------
'REGISTRO DE IDENTIDADE CIVIL - RIC
'Nota importante: O c�lculo do DV do RIC - Registro de Identidade Civil (nova carteira de identidade dos brasileiros), ainda n�o est� claramente explicado. Eu presumo que seguir� a regra abaixo, mas ainda dependemos da confirma��o, que vir� a partir da emiss�o dos primeiros cart�es, prevista para dezembro de 2010. Se o prezado leitor j� tem RIC, pe�o o obs�quio de conferir o seu DV na rotina abaixo e dizer-me se meu algoritmo funcionou.
'Saiba como se calcula o DV (D�gito Verificador) do Registro de Identidade Civil - RIC e veja o DV de qualquer n�mero, utilizando a rotina abaixo. O DV corresponde ao resto da divis�o por 11 do somat�rio da multiplica��o de cada algarismo da base respectivamente por 9, 8, 7, 6, 5, 4, 3, 2, 9 e 8, a partir da unidade. Siga o exemplo abaixo:
'1  3  3  9  7  0  5  1  2  7
'x  x  x  x  x  x  x  x  x  x
'8  9  2  3  4  5  6  7  8  9
'----------------------------
'8+27+ 6+27+28+ 0+30+ 7+16+63 = 212�11=19, com resto 3 (este � o DV).
'
'Nota:
'Se o resto for 10, o DV ser� "0".
'Introduza o n�mero do RIC (exemplo: 1339705127)
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
' Coment�rios para o RG - Registro Geral
'-------------------------------------------------------------------
'Exemplos RG - Registro Geral :
'-------------------------------------------------------------------
'RG: 42.943.412-1
'RG: 42.943.425-X
'-------------------------------------------------------------------

Function CalculoDV11(strNumero As String) As String
'declara as vari�veis
Dim intcontador, intnumero, intTotalNumero, intMultiplicador, intResto As Integer

    ' se nao for um valor numerico sai da fun��o
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
    
        'soma o resultado para totaliza��o
        intTotalNumero = intTotalNumero + intnumero
    
        'se o multiplicador for maior que 2 decrementa-o caso contrario atribuir valor padrao original
        intMultiplicador = IIf(intMultiplicador > 2, intMultiplicador - 1, 9)
    
     Next

    'calcula o resto da divisao do total por 11
    intResto = intTotalNumero Mod 11

    'verifica as exce��es (intResto = 10 ent�o DV = "X")

    If intResto < 10 Then
        CalculoDV11 = Trim(Str(intResto))
    Else
        CalculoDV11 = "X"
    End If

End Function
