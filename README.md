# TÖNCHTORY V2.4

Correções desta versão:

- Visual geral padronizado no sistema inteiro.
- Splash centralizado com animação de opacidade e limite máximo de 7,7 segundos.
- Removidas referências textuais externas no splash e na interface.
- Funções anteriores preservadas:
  - Importar Planilha
  - Nova Contagem sem Planilha
  - Continuar Sessão
  - Desfazer Última
  - Recontar Selecionado
  - Zerar Selecionado
  - Salvar Inventário em XLSX ou CSV
- No modo Nova Contagem sem Planilha, cada bip/scan/teclado atualiza simultaneamente:
  - Quantidade
  - Disponível
  - Contagem Inventário
- O Status fica como NOVO para itens criados na contagem do zero.
- Quando uma planilha é importada, o fluxo permanece normal:
  - Quantidade e Disponível vêm da planilha.
  - Contagem Inventário recebe os bips.
  - Status calcula OK / SOBRA / FALTA / ZERADO.

## Como usar

1. Execute `start.bat`.
2. Escolha uma opção:
   - Importar Planilha
   - Nova Contagem sem Planilha
3. Bipe ou digite o código do item.
4. Cadastre o item quando solicitado.
5. Clique em Salvar Inventário para exportar XLSX ou CSV.
