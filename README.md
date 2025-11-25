# Sistema de Sorteio Público Auditável - COLUNI-UFF

Sobre o Projeto:

Este software foi desenvolvido com o objetivo de garantir transparência, segurança e auditoria no processo de sorteio público para ingresso de novos alunos no Colégio Universitário Geraldo Reis (COLUNI-UFF).
O sistema automatiza o sorteio aleatório de candidatos, respeitando rigorosamente as cotas, a ordem de prioridade e o remanejamento automático de vagas não preenchidas para a Ampla Concorrência, conforme as regras de editais de concursos públicos.
Créditos e Licença
* Desenvolvedor: Alessandro Santos
* GitHub: @alessandrosanntos
* Versão Atual: 4.4
* Versão de lançamento: 1.0
* Licença: GNU GPLv3 (Código Aberto).

Funcionalidades Principais (Versão 4.4)
1. Auditoria e Transparência:
   * Semente de Aleatoriedade (Seed): O operador insere um dado público (ex: data/hora ou número da Loteria Federal) que torna o sorteio matemático e reproduzível.
   * Hash SHA-256: O sistema calcula a assinatura digital do arquivo Excel. Qualquer alteração no arquivo original muda este código, garantindo a integridade da base de dados.
   * Logs Detalhados: Gera arquivos de texto individuais por turma e um resumo consolidado ao final, salvos em uma pasta exclusiva com data e hora.
2. Gestão de Vagas e Cotas:
   * Configuração manual de vagas Imediata e Cadastro de Reserva por cota.
   * Remanejamento Automático: Vagas não preenchidas nas cotas prioritárias são somadas matematicamente às vagas da Ampla Concorrência (AC), com exibição detalhada (Originais + Sobras).
   * Replicação: Botão para copiar o quadro de vagas da turma anterior, agilizando a configuração de turmas similares.
   * Pular Turma: Botão dedicado para registrar turmas sem vagas ou sem sorteio, mantendo o registro no log de auditoria ("Turma Pulada").
3. Interface Visual e Suspense:
   * Resolução otimizada (1366x768).
   * Controle Passo-a-Passo: Botão "SORTEAR PRÓXIMO" permite que o operador controle o ritmo do evento.
   * Animação de Revelação:
      1. Animação "Sorteando..." (5s).
      2. Exibição do Número de Inscrição.
      3. Pausa de suspense (2s) + Mensagem "Buscando informações..." (3s).
      4. Revelação do Nome do Candidato.
   * Visualização em planilha (tabela) ao final de cada turma.

Estrutura do Arquivo de Dados (base.xlsx)
O arquivo Excel deve conter apenas 4 colunas na seguinte ordem exata (A a D):
Coluna
	Cabeçalho Sugerido
	Conteúdo
	A
	Inscrição
	Número de inscrição do candidato.
	B
	Nome
	Nome completo do candidato.
	C
	Turma
	Turma pretendida (ex: "TURMA 1", "1º ANO").
	D
	Cota
	Sigla da Cota (PPI1, PPI2, RF, PCD, AC).
	Importante:
* O sistema remove automaticamente candidatos cujo nome seja "NÚMERO CANCELADO".
* O sistema normaliza automaticamente nomes de turmas (ex: converte "1º ANO DO ENSINO FUNDAMENTAL" para "1º ANO EF").
________________
Requisitos do Sistema
* Sistema Operacional: Windows 10/11.
* Arquivos Obrigatórios na Pasta:
   1. O executável (.exe) ou script (.py).
   2. base.xlsx (A planilha de dados).
   3. README.md (Este arquivo).
   4. TUTORIAL_USUARIO.txt (Guia operacional).
   5. LICENSE (Termos da GPLv3).
Como Contribuir
Sendo um projeto de código aberto, contribuições são bem-vindas. Sinta-se à vontade para fazer um fork deste projeto no GitHub, implementar melhorias e enviar um pull request.
Desenvolvido com ❤️ para a educação pública.
