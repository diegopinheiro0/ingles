# Guia de pronúncia de inglês (Diego)

Objetivo: pronunciar corretamente em **inglês americano**. O guia cresce com conteúdo que o Diego
encontra na internet. As mesmas regras estão na skill `guia-pronuncia` (repositório `skills`).

## Regras

- `Guia_Ingles_Pronuncia.xlsx` é a fonte única. O PDF é sempre gerado dela; nunca editar o PDF.
- Antes de escrever qualquer pronúncia, ler `NOTACAO.md`: notação, o que é proibido e os erros
  que já aconteceram. Ex.: work = WÂRK (nunca "UÓR"), sem vogal nasal, i/ii e u/uu.
- Toda palavra nova leva IPA (Cambridge US, sem barras), grupo, fonte e Status = treinar.
- Mostrar a tabela das inclusões ao Diego e gravar só depois do "ok". Pedido pequeno e explícito
  (marcar como dominada, corrigir uma pronúncia) dispensa a aprovação.
- Não apagar palavras nem grupos sem pedido. Não mudar a notação por conta própria.

## Rodada de atualização

```bash
pip install -r requirements.txt
python adicionar.py Guia_Ingles_Pronuncia.xlsx novos.json   # formato no topo do script
python validar.py Guia_Ingles_Pronuncia.xlsx                # zero erros antes de seguir
python gerar_pdf.py Guia_Ingles_Pronuncia.xlsx
```

Commit da planilha e do PDF juntos, com mensagem dizendo o que entrou e de qual fonte.
O Diego autorizou gravar direto no `main`.
