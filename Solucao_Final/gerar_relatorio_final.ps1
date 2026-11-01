$ErrorActionPreference = 'Stop'
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$modelo  = "C:\Users\amand\OneDrive\Documentos\Faculdade\PI IV\Modelo_Relatorio_Final.docx"
$outDocx = "C:\Users\amand\OneDrive\Documentos\Faculdade\PI IV\Relatorio_Final_PI_Chamados_N2.docx"
$outPdf  = "C:\Users\amand\OneDrive\Documentos\Faculdade\PI IV\Relatorio_Final_PI_Chamados_N2.pdf"
$figDash = "C:\Users\amand\OneDrive\Documentos\Faculdade\PI IV\Solucao_Final\dashboard_solucao_final.png"

Copy-Item $modelo $outDocx -Force
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($outDocx)
$doc.Content.Delete() | Out-Null
$sel = $word.Selection

$doc.PageSetup.TopMargin = 85.05
$doc.PageSetup.BottomMargin = 56.7
$doc.PageSetup.LeftMargin = 85.05
$doc.PageSetup.RightMargin = 56.7

function S([string]$n) { try { $script:sel.Style = $n; $true } catch { $false } }
function PB() { $sel.InsertBreak(7) }
function Blank([int]$n=1) { for ($i=0;$i -lt $n;$i++){ $sel.TypeParagraph() } }
function CapaLine([string]$t, $bold=$false) {
  S "d) capa" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=14; $sel.Font.Bold=$bold
  $sel.ParagraphFormat.Alignment=1
  $sel.ParagraphFormat.SpaceAfter=0; $sel.ParagraphFormat.SpaceBefore=0
  if ($t) { $sel.TypeText($t) }
  $sel.TypeParagraph()
}
function Body([string]$t) {
  S "a) texto-base" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false
  $sel.ParagraphFormat.Alignment=3
  $sel.ParagraphFormat.LineSpacingRule=4
  $sel.ParagraphFormat.LineSpacing=18
  $sel.ParagraphFormat.SpaceAfter=12
  $sel.ParagraphFormat.FirstLineIndent=0
  $sel.ParagraphFormat.LeftIndent=0
  $sel.TypeText($t)
  $sel.TypeParagraph()
}
function H1([string]$t) {
  PB
  S "1) título nivel 1" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$true
  $sel.ParagraphFormat.Alignment=0
  $sel.ParagraphFormat.SpaceAfter=18
  $sel.TypeText($t.ToUpper())
  $sel.TypeParagraph()
}
function H2([string]$t) {
  S "2) título nivel 2" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$true
  $sel.ParagraphFormat.Alignment=0
  $sel.ParagraphFormat.SpaceAfter=12
  $sel.TypeText($t)
  $sel.TypeParagraph()
}
function Cap([string]$t) {
  S "c) citação" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=10; $sel.Font.Bold=$false
  $sel.ParagraphFormat.Alignment=3
  $sel.ParagraphFormat.LeftIndent=113.4
  $sel.ParagraphFormat.LineSpacingRule=0
  $sel.ParagraphFormat.SpaceAfter=12
  $sel.TypeText($t)
  $sel.TypeParagraph()
  $sel.ParagraphFormat.LeftIndent=0
}
function Bullet([string]$t) {
  S "b) texto com bullets" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false
  $sel.ParagraphFormat.Alignment=3
  $sel.ParagraphFormat.LineSpacingRule=4
  $sel.ParagraphFormat.LineSpacing=18
  $sel.ParagraphFormat.SpaceAfter=6
  $sel.TypeText($t)
  $sel.TypeParagraph()
}

$titulo = "Predição e classificação operacional de chamados N2"
$nome="Amanda Gonçalves Carneiro"; $curso="Bacharelado em Ciência de Dados e Engenharia de Computação"
$tutor="Daniel Rossado Oliveira"; $polo="Campinas"; $cidade="Campinas"; $ano="2026"

# CAPA
CapaLine "UNIVERSIDADE VIRTUAL DO ESTADO DE SÃO PAULO" $true
Blank 5
CapaLine $nome
Blank 5
CapaLine $titulo $true
Blank 2
CapaLine "Vídeo de apresentação do Projeto Integrador"
CapaLine "[inserir o link do vídeo após a publicação]"
Blank 6
CapaLine "$cidade - SP"
CapaLine $ano

# FOLHA DE ROSTO
PB
CapaLine "UNIVERSIDADE VIRTUAL DO ESTADO DE SÃO PAULO" $true
Blank 6
CapaLine $titulo $true
Blank 3
S "e) capa descrição" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false
$sel.ParagraphFormat.Alignment=3
$sel.ParagraphFormat.LeftIndent=198
$sel.TypeText("Relatório Técnico-Científico apresentado na disciplina de Projeto Integrador em Computação IV para o curso de $curso da Universidade Virtual do Estado de São Paulo (UNIVESP).")
$sel.TypeParagraph()
$sel.ParagraphFormat.LeftIndent=198
$sel.TypeText("Orientador: $tutor")
$sel.TypeParagraph()
$sel.ParagraphFormat.LeftIndent=0
Blank 7
CapaLine "$cidade - SP"
CapaLine $ano

# FICHA / RESUMO
PB
Blank 10
S "a) texto-base" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=10; $sel.Font.Bold=$false
$sel.ParagraphFormat.Alignment=3
$sel.ParagraphFormat.LineSpacingRule=0
$sel.TypeText("CARNEIRO, Amanda Gonçalves. $titulo. Relatório Técnico-Científico. $curso - Universidade Virtual do Estado de São Paulo. Tutor: $tutor. Polo $polo, $ano.")
$sel.TypeParagraph()
Blank 1
S "a) texto-base" | Out-Null
$sel.Font.Size=12; $sel.Font.Bold=$true; $sel.ParagraphFormat.Alignment=1
$sel.TypeText("RESUMO"); $sel.TypeParagraph()

$resumo = "A sustentação N2 do Dynamics 365 na Unimed Porto Alegre concentra chamados de escopos distintos (N1, N2 e N3), o que dificulta a triagem, a priorização e o cumprimento de prazos. Este Projeto Integrador IV desenvolveu uma solução analítico-preditiva sobre 614 chamados do Qualitor, com 370 registros previamente classificados. O objetivo geral foi classificar o nível operacional dos chamados e estimar o risco de ciclo longo, publicando os resultados em interface de visualização. A metodologia seguiu o Design Thinking (ouvir, criar e implementar), com análise exploratória, Naive Bayes para o nível sugerido e regressão logística para o risco de ciclo superior a dez dias. Na validação holdout, o modelo de nível obteve acurácia de 78,4% e o modelo de risco, F1 de 0,81 na classe alta. A solução final reduz a fila a uma lista acionável: devolver ao N1, escalar ao N3 ou priorizar no N2. Conclui-se que dados existentes em escala, aprendizagem de máquina e dashboard atendem ao tema norteador e apoiam a operação, com limitações de desbalanceamento da classe N1 e de campos de resolução incompletos."
S "f) Resumo/Referências" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false
$sel.ParagraphFormat.Alignment=3
$sel.ParagraphFormat.LineSpacingRule=0
$sel.TypeText($resumo)
$sel.TypeParagraph()
Blank 1
S "a) texto-base" | Out-Null
$sel.Font.Size=12; $sel.Font.Bold=$false; $sel.ParagraphFormat.Alignment=3
$sel.TypeText("PALAVRAS-CHAVE: Aprendizagem de máquina; Análise exploratória de dados; Power BI; Chamados N2; Dynamics 365.")
$sel.TypeParagraph()

# LISTAS
PB
S "a) texto-base" | Out-Null
$sel.Font.Bold=$true; $sel.ParagraphFormat.Alignment=0
$sel.TypeText("LISTA DE ILUSTRAÇÕES"); $sel.TypeParagraph()
Blank 1
S "Sumário 1" | Out-Null; $sel.Font.Bold=$false
$sel.TypeText("Figura 1  - Interface da solução final (dashboard preditivo)"); $sel.TypeParagraph()

PB
S "a) texto-base" | Out-Null
$sel.Font.Bold=$true; $sel.ParagraphFormat.Alignment=0
$sel.TypeText("LISTA DE TABELAS"); $sel.TypeParagraph()
Blank 1
S "Sumário 1" | Out-Null; $sel.Font.Bold=$false
$sel.TypeText("Tabela 1  - Distribuição da base por família operacional"); $sel.TypeParagraph()
S "Sumário 1" | Out-Null
$sel.TypeText("Tabela 2  - Métricas dos modelos preditivos (validação holdout)"); $sel.TypeParagraph()

# SUMÁRIO
PB
S "Normal" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$true
$sel.ParagraphFormat.Alignment=1
$sel.TypeText("SUMÁRIO"); $sel.TypeParagraph()
Blank 1
function Sum1([string]$t){ S "Sumário 1"|Out-Null; $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$true; $sel.ParagraphFormat.Alignment=0; $sel.TypeText($t); $sel.TypeParagraph() }
function Sum2([string]$t){ S "Sumário 2"|Out-Null; $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false; $sel.ParagraphFormat.Alignment=0; $sel.TypeText($t); $sel.TypeParagraph() }
Sum1 "1 INTRODUÇÃO"
Sum1 "2 DESENVOLVIMENTO"
Sum2 "2.1 Objetivos"
Sum2 "2.2 Justificativa e delimitação do problema"
Sum2 "2.3 Fundamentação teórica"
Sum2 "2.4 Aplicação das disciplinas estudadas no projeto integrador"
Sum2 "2.5 Metodologia"
Sum1 "3 RESULTADOS: SOLUÇÃO FINAL"
Sum1 "4 CONSIDERAÇÕES FINAIS"
Sum1 "REFERÊNCIAS"
Sum1 "APÊNDICES"

# INTRODUÇÃO  - section break
$sel.InsertBreak(2) | Out-Null
S "1) título nivel 1" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$true
$sel.ParagraphFormat.Alignment=0
$sel.TypeText("1 INTRODUÇÃO"); $sel.TypeParagraph()

Body "A operação de sustentação de plataformas empresariais, em especial o Microsoft Dynamics 365 integrado a ferramentas de service desk, depende de decisões rápidas de triagem, priorização e escalonamento. Quando a mesma fila reúne dúvidas e direitos de acesso típicos de primeiro nível, melhorias e incidentes de segundo nível e problemas de integração ou desenvolvimento de terceiro nível, a ausência de instrumentos analíticos e preditivos compromete prazos e a comunicação entre equipes."
Body "O tema norteador do Projeto Integrador em Computação IV da Universidade Virtual do Estado de São Paulo (Univesp) solicita análise de dados em escala, com conjuntos existentes ou provenientes de Internet das Coisas, associada a aprendizagem de máquina e a uma interface para visualização dos resultados. Este trabalho materializa esse tema em um estudo de caso na sustentação N2 do Dynamics 365 da Unimed Porto Alegre, com dados do Qualitor e contexto operacional da equipe responsável pelo atendimento."
Body "A escolha do cenário decorreu da existência de base histórica classificada, de um problema recorrente de governança da fila e da aderência às competências em análise de dados, Power Platform e Dynamics. O objeto de estudo são os chamados da operação N2 e a possibilidade de apoiar, por modelos supervisionados e por dashboard, a classificação operacional e a estimativa de risco de ciclo longo."
Body "Este relatório final retoma o problema, os objetivos, a fundamentação e a metodologia, apresenta a solução final implementada e discute resultados, contribuições e limitações, incorporando o feedback do relatório parcial."

H1 "2 Desenvolvimento"

H2 "2.1 OBJETIVOS"
Body "Objetivo geral: desenvolver um modelo preditivo e uma interface de visualização que classifiquem o nível operacional de chamados N2 e estimem o risco de ciclo longo, apoiando a priorização da fila."
Body "Objetivos específicos:"
Bullet "Identificar, junto à comunidade externa, os problemas da fila N2 relacionados ao tema norteador."
Bullet "Caracterizar a base histórica por status, família operacional, cluster, aging e ciclo."
Bullet "Determinar variáveis-alvo e atributos para a modelagem supervisionada."
Bullet "Treinar e avaliar modelos de classificação de nível e de risco de ciclo longo."
Bullet "Publicar os resultados em dashboard e em base consumível pelo Power BI."
Bullet "Verificar, com a operação, a utilidade da lista priorizada para a rotina semanal."

H2 "2.2 JUSTIFICATIVA E DELIMITAÇÃO DO PROBLEMA"
Body "A pergunta de pesquisa é: como utilizar análise de dados em escala e aprendizagem de máquina, com interface de visualização, para apoiar a classificação operacional e a previsão de risco de ciclo longo em chamados N2 de sustentação Dynamics 365, no contexto da Unimed Porto Alegre?"
Body "A justificativa é prática e acadêmica. Na operação, a base de 26 de outubro de 2026 reúne 614 chamados, dos quais 27 permaneciam abertos. A família operacional mostra 253 itens de N2 direto, 224 dependentes de N3 ou integração, 84 com perfil de N1 ou Inteligência e 53 mistos. A mediana de ciclo do N2 direto ficou em 18,3 dias, incompatível com matrizes de SLA em 8 a 16 horas quando há homologação, dossiê ou espera de outra área. Do ponto de vista acadêmico, o projeto emprega dados existentes em escala, técnicas de exploração e de aprendizado supervisionado e uma interface para comunicar o resultado, conforme o tema norteador."
Body "Delimita-se o escopo à análise, à modelagem e à visualização. Ficam fora do PI a reescrita da integração Qualitor -Dynamics, a zeragem forçada da fila e a resolução do backlog de N3."

H2 "2.3 FUNDAMENTAÇÃO TEÓRICA"
Body "A análise de dados exige compreender tipos e escalas das variáveis e explorar relações univariadas e multivariadas antes de qualquer modelagem (FACELI et al., 2011). Sem essa etapa, o modelo preditivo herda ruído, categorias mal definidas e interpretações equivocadas de desempenho."
Body "A análise exploratória de dados descreve distribuições, assimetrias e associações. Ferreira et al. destacam estatística descritiva, histogramas, boxplots e análises multivariadas como base para formular hipóteses (FERREIRA et al., [s.d.]). Neste projeto, a exploração cobriu volume mensal, famílias, razões de status e tempo de ciclo."
Body "A aprendizagem de máquina supervisionada induz funções a partir de exemplos rotulados (FACELI et al., 2011). Foram adotados dois modelos interpretáveis e adequados ao tamanho da base: Naive Bayes, alinhado aos métodos probabilísticos do texto-base, para o nível sugerido; e regressão logística para o risco de ciclo longo. A avaliação reporta acurácia e F1 por classe, em razão do desbalanceamento."
Cap "A qualidade do pré-processamento  - tratamento de ausências, normalização de categorias e engenharia de atributos  - condiciona o desempenho preditivo e a interpretabilidade dos resultados. (FACELI et al., 2011)"
Body "A visualização fecha o ciclo analítico. Milani et al. tratam visualização como processo de contar a história dos dados, e não apenas de desenhar gráficos (MILANI et al., [s.d.]). O dashboard e o Power BI materializam essa etapa: indicadores, fila priorizada e ação sugerida."
Body "A literatura de Internet das Coisas, também prevista no tema, reforça a ideia de dados em escala e de plataformas de coleta (SANTOS et al., [s.d.]; MORAIS et al., [s.d.]). Optou-se por dados existentes de service desk, com acesso viável e aderência à comunidade, em vez de sensores IoT."

H2 "2.4 APLICAÇÃO DAS DISCIPLINAS ESTUDADAS NO PROJETO INTEGRADOR"
Body "O trabalho articula conteúdos do próprio Projeto Integrador e de disciplinas correlatas do curso de Ciência de Dados e Engenharia de Computação."
Bullet "Análise de dados e estatística descritiva: caracterização da base, tipos de variáveis, criados versus encerrados, medianas e percentuais por família."
Bullet "Aprendizado de máquina: pré-processamento, modelos preditivos probabilísticos e avaliação holdout, conforme Faceli et al."
Bullet "Visualização de dados: escolha de indicadores, tabelas de priorização e dashboard para data storytelling."
Bullet "Metodologia do PI e Design Thinking: etapas ouvir, criar e implementar, com interação com a comunidade externa."
Bullet "Sistemas corporativos e integrações: o contexto Dynamics 365, Qualitor e Power Platform informa as regras N1, N2 e N3, sem tornar a integração o produto do PI."

H2 "2.5 METODOLOGIA"
Body "A pesquisa é aplicada, descritiva e exploratória, organizada nas três fases do Design Thinking adotadas pela Univesp."
Body "Ouvir e interpretar o contexto. A comunidade externa é a sustentação N2 Dynamics da Unimed Porto Alegre. As informações vieram de análise documental e de discussão operacional (situação da fila, relatório de sustentação, matrizes de SLA e exportações Qualitor/Dynamics). Os sujeitos são analistas de N2 e a coordenação da operação. Observaram-se fila mista, itens de N3 no nome do N2, homologações paradas e conflito entre SLA em horas e esforço de melhoria."
Body "Criar e prototipar. A base final tem 614 registros e colunas de identificador, categoria, descrição, prioridade, status, razão, título, responsável, datas, classificação e duração. Foram criados cluster, família, nível, ciclo em dias e atributos textuais. O protótipo inicial do relatório parcial era o desenho do pipeline e do dashboard. A solução final treinou Naive Bayes no título, na categoria e na prioridade, e regressão logística com prioridade, nível, palavras-chave e razão de status. O alvo de risco é ciclo maior que dez dias."
Body "Implementar e testar. Os scores foram gravados em planilha e publicados em dashboard HTML, compatível com o Power BI já existente. A validação usou holdout 70/30. A devolutiva da operação, desde o alinhamento de agosto até a base de outubro, deslocou o foco de “mostrar volume” para “sugerir ação”. A fila aberta caiu de 61 para 27 chamados no período, e a interface passou a listar risco alto, perfil N1 e perfil N3. Limitações de data de resolução vazia e de LGPD foram tratadas com proxy de ciclo, quando possível, e sem publicação indevida da base bruta."

H1 "3 Resultados: solução final"
Body "No passo ouvir, o diagnóstico se confirmou com números. A Tabela 1 resume a família operacional da base de 26 de outubro de 2026. Plataforma comercial (124), integração (100), fila (89) e árvore (74) concentram o backlog histórico. Entre os 27 abertos, 12 estavam novos, 6 aguardando cliente e 5 aguardando fornecedor."

S "a) texto-base" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=10; $sel.Font.Bold=$true
$sel.ParagraphFormat.Alignment=1
$sel.ParagraphFormat.SpaceAfter=6
$sel.TypeText("Tabela 1  - Distribuição da base por família operacional")
$sel.TypeParagraph()

$t1 = $doc.Tables.Add($sel.Range, 5, 3)
$t1.Borders.Enable = $true
$hdrs = @("Família operacional","Quantidade","Percentual")
$vals = @(
  @("N2 direto","253","41,2%"),
  @("Dependente de N3 / integração","224","36,5%"),
  @("Perfil N1 / Inteligência","84","13,7%"),
  @("Misto (comercial / dados)","53","8,6%")
)
for ($c=0; $c -lt 3; $c++) { $t1.Cell(1,$c+1).Range.Text = $hdrs[$c]; $t1.Cell(1,$c+1).Range.Font.Bold = $true; $t1.Cell(1,$c+1).Range.Font.Size = 10 }
for ($r=0; $r -lt 4; $r++) {
  for ($c=0; $c -lt 3; $c++) { $t1.Cell($r+2,$c+1).Range.Text = $vals[$r][$c]; $t1.Cell($r+2,$c+1).Range.Font.Size = 10; $t1.Cell($r+2,$c+1).Range.Font.Bold = $false }
}
$sel.EndKey(6) | Out-Null
$sel.TypeParagraph()

Body "No passo criar, a solução final ficou em três camadas. A primeira é o dataset tratado, com rótulos e atributos. A segunda são os dois modelos. A terceira é a interface: o dashboard da Figura 1 e a planilha Chamados_N2_com_predicao.xlsx, que alimenta o Power BI da operação. Cada chamado recebe nível predito, confiança, risco de ciclo longo, score e ação sugerida."

S "a) texto-base" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=10; $sel.Font.Bold=$true
$sel.ParagraphFormat.Alignment=1
$sel.ParagraphFormat.SpaceAfter=6
$sel.TypeText("Figura 1  - Interface da solução final (dashboard preditivo)")
$sel.TypeParagraph()
if (Test-Path $figDash) {
  $pic = $sel.InlineShapes.AddPicture($figDash)
  $pic.LockAspectRatio = $true
  $pic.Width = 450
}
$sel.TypeParagraph()
S "a) texto-base" | Out-Null
$sel.Font.Size=10; $sel.ParagraphFormat.Alignment=1
$sel.TypeText("Fonte: elaborada pela autora (2026), a partir da base Qualitor de 26/10/2026.")
$sel.TypeParagraph()

Body "No passo implementar, a Tabela 2 apresenta as métricas de teste. O Naive Bayes obteve acurácia de 78,4%. O F1 de N2 foi 0,83 e o de N3, 0,88. O F1 de N1 ficou em 0,33, por desbalanceamento e menor recall, o que impede o uso automático isolado dessa classe. A regressão logística de risco, com corte de dez dias, teve acurácia de 72,9% e F1 de 0,81 na classe alta."

S "a) texto-base" | Out-Null
$sel.Font.Name="Times New Roman"; $sel.Font.Size=10; $sel.Font.Bold=$true
$sel.ParagraphFormat.Alignment=1
$sel.TypeText("Tabela 2  - Métricas dos modelos preditivos (validação holdout)")
$sel.TypeParagraph()
$t2 = $doc.Tables.Add($sel.Range, 4, 4)
$t2.Borders.Enable = $true
$h2 = @("Modelo","Alvo","Acurácia","F1 (classe crítica)")
$v2 = @(
  @("Naive Bayes","Nível N1/N2/N3/Misto","78,4%","N3 = 0,88; N2 = 0,83"),
  @("Regressão logística","Risco de ciclo > 10 dias","72,9%","Alto = 0,81"),
  @("Uso combinado","Ação sugerida na fila aberta"," -","23 riscos altos; 2 N1; 9 N3")
)
for ($c=0; $c -lt 4; $c++) { $t2.Cell(1,$c+1).Range.Text = $h2[$c]; $t2.Cell(1,$c+1).Range.Font.Bold = $true; $t2.Cell(1,$c+1).Range.Font.Size = 10 }
for ($r=0; $r -lt 3; $r++) {
  for ($c=0; $c -lt 4; $c++) { $t2.Cell($r+2,$c+1).Range.Text = $v2[$r][$c]; $t2.Cell($r+2,$c+1).Range.Font.Size = 10; $t2.Cell($r+2,$c+1).Range.Font.Bold = $false }
}
$sel.EndKey(6) | Out-Null
$sel.TypeParagraph()

Body "Em relação à solução inicial do relatório parcial, houve três melhorias pedidas pela comunidade e pelo orientador: (1) passar do desenho para modelos treinados e avaliados; (2) usar a exportação atualizada, na qual a fila aberta já havia caído de 61 para 27; (3) separar com clareza objetivos, justificativa e metodologia e publicar uma lista acionável, não apenas gráficos de volume. A duração em horas do Dynamics mostrou-se, na maior parte dos casos, tempo de ação, e não ciclo completo; por isso o risco passou a usar dias entre criação e resolução."

H1 "4 Considerações finais"
Body "O objetivo geral foi atingido: existe um modelo preditivo avaliado e uma interface que classifica o chamado e estima risco de ciclo longo. Os objetivos específicos de caracterizar a base, treinar os modelos e publicar o dashboard foram cumpridos. A verificação com a operação permanece como rito contínuo, já iniciado nos alinhamentos de sustentação."
Body "À luz de Faceli et al. (2011) e da literatura de visualização, a contribuição não está em um algoritmo inédito, e sim na aplicação disciplinada de exploração, modelagem interpretável e comunicação visual a um problema real de service desk. Para a comunidade, o impacto é uma fila menor de itens indevidos no N2 e uma ordem de trabalho semanal baseada em risco e em nível sugerido."
Body "As limitações são o desbalanceamento da classe N1, campos de resolução incompletos, a necessidade de retreino a cada nova exportação e o fato de a integração Qualitor -Dynamics permanecer fora do escopo. Como continuidade, recomenda-se enriquecer o rótulo de N1, automatizar o refresh da planilha e acoplar alertas de risco alto ao fluxo já usado pela equipe."

H1 "Referências"
$refs = @(
"ABNT - ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. NBR 14724: informação e documentação  - trabalhos acadêmicos  - apresentação. Rio de Janeiro: ABNT, 2011.",
"BROWN, Tim. Design thinking. Harvard Business Review, p. 84-92, 2008.",
"FACELI, Katti; LORENA, Ana Carolina; GAMA, João; ALMEIDA, Tiago Agostinho de; CARVALHO, André Carlos Ponce de Leon Ferreira de. Inteligência artificial: uma abordagem de aprendizado de máquina. Rio de Janeiro: LTC, 2011.",
"FERREIRA, Rafael Gastão Coimbra et al. Preparação e análise exploratória de dados. [S.l.]: [s.n.], [s.d.]. Capítulos 1 a 6.",
"MILANI, Alessandra Maciel Paz et al. Visualização de dados. [S.l.]: [s.n.], [s.d.]. Capítulos 1 a 10.",
"MORAIS, Izabelly Soares de et al. Introdução a Big Data e Internet das Coisas (IoT). [S.l.]: [s.n.], [s.d.]. Subtítulo Internet das Coisas.",
"PIRES, Paulo F. et al. Plataformas para a Internet das Coisas. In: ______. Capítulo 3. [S.l.]: [s.n.], [s.d.]. p. 115-121.",
"SANTOS, Bruno P. et al. Internet das Coisas: da teoria à prática. In: ______. Capítulo 1. [S.l.]: [s.n.], [s.d.]. p. 1-5, 13-19, 35-45.",
"SEVERINO, Antônio Joaquim. Metodologia do trabalho científico. 22. ed. rev. e ampl. São Paulo: Cortez, 2002.",
"UNIVERSIDADE VIRTUAL DO ESTADO DE SÃO PAULO. Modelo de Relatório Final  - Projeto Integrador. Ambiente Virtual de Aprendizagem, 2026."
)
foreach ($r in $refs) {
  S "f) Resumo/Referências" | Out-Null
  $sel.Font.Name="Times New Roman"; $sel.Font.Size=12; $sel.Font.Bold=$false
  $sel.ParagraphFormat.Alignment=0
  $sel.ParagraphFormat.LineSpacingRule=0
  $sel.ParagraphFormat.FirstLineIndent=-36
  $sel.ParagraphFormat.LeftIndent=36
  $sel.ParagraphFormat.SpaceAfter=8
  $sel.TypeText($r); $sel.TypeParagraph()
}
$sel.ParagraphFormat.FirstLineIndent=0; $sel.ParagraphFormat.LeftIndent=0

H1 "Apêndices"
Body "Apêndice A  - Artefatos da solução final, na pasta Solucao_Final: dashboard_solucao_final.html; Chamados_N2_com_predicao.xlsx; gerar_solucao.js; metricas.json; LEIA-ME_SOLUCAO.md. O relatório Power BI da operação permanece em Relatorio de Chamados.pbix."
Body "Apêndice B  - Roteiro para gravação do vídeo: Roteiro_Video_PI_IV.txt, na pasta do PI IV."

# page numbers
Write-Output "Sections=$($doc.Sections.Count)"
$sec1=$doc.Sections.Item(1)
try { while($sec1.Headers.Item(1).PageNumbers.Count -gt 0){ $sec1.Headers.Item(1).PageNumbers.Item(1).Delete() } } catch {}
$sec1.Headers.Item(1).Range.Text=""
if ($doc.Sections.Count -ge 2) {
  $sec2=$doc.Sections.Item(2)
  $sec2.Headers.Item(1).LinkToPrevious=$false
  try { while($sec2.Headers.Item(1).PageNumbers.Count -gt 0){ $sec2.Headers.Item(1).PageNumbers.Item(1).Delete() } } catch {}
  $sec2.Headers.Item(1).Range.Text=""
  $pn=$sec2.Headers.Item(1).PageNumbers
  $pn.RestartNumberingAtSection=$true
  $pn.StartingNumber=1
  $pn.Add(2) | Out-Null
}

Write-Output "==== ORDER ===="
for ($i=1; $i -le $doc.Paragraphs.Count; $i++) {
  $p=$doc.Paragraphs.Item($i)
  $t=($p.Range.Text -replace "[\r\a\u0007]","").Trim()
  if ($t.Length -eq 0) { continue }
  if ($t -match '^(UNIVERSIDADE|Amanda|Predição|Vídeo|Relatório Técnico|Orientador:|Campinas|CARNEIRO|RESUMO|PALAVRAS|LISTA|SUMÁRIO|1 INTRODU|2 DESENVOLV|2\.[1-5]|3 RESULT|4 CONSID|REFER|APÊNDICE|Tabela |Figura )') {
    $pg=$p.Range.Information(1)
    $short= if($t.Length-gt 70){$t.Substring(0,70)}else{$t}
    Write-Output "pg$pg|$short"
  }
}
for ($s=1; $s -le $doc.Sections.Count; $s++) {
  $sec=$doc.Sections.Item($s)
  $t=""
  for ($j=1; $j -le $sec.Range.Paragraphs.Count; $j++) {
    $t=($sec.Range.Paragraphs.Item($j).Range.Text -replace "[\r\a\u0007]","").Trim()
    if ($t.Length -gt 0) { break }
  }
  Write-Output ("SEC$s => " + $(if($t.Length-gt 50){$t.Substring(0,50)}else{$t}))
}

$doc.Save()
if (Test-Path $outPdf) { Remove-Item $outPdf -Force }
$doc.ExportAsFixedFormat($outPdf, 17)
$doc.Close($false)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)|Out-Null
Write-Output "OK docx=$((Get-Item $outDocx).Length) pdf=$((Get-Item $outPdf).Length)"
