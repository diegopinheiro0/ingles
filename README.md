# ingles

Repositório para expandir o estudo de inglês (foco em pronúncia americana).

## Arquivos

| Arquivo | O que é |
|---|---|
| `Guia_Ingles_Pronuncia.xlsx` | Planilha mestre. **É aqui que se edita.** |
| `Guia_Ingles_Pronuncia.pdf` | Guia para ler e imprimir, gerado a partir da planilha. |
| `gerar_pdf.py` | Script que gera o PDF a partir da planilha. |
| `validar.py` | Confere a planilha (notação, IPA, duplicadas) antes de gerar o PDF. |
| `adicionar.py` | Acrescenta linhas aprovadas à planilha a partir de um JSON. |
| `NOTACAO.md` | Regras da notação de pronúncia. |
| `CLAUDE.md` | Instruções para o Claude manter o guia. |

## Como adicionar conteúdo

1. Abra a planilha e adicione linhas na aba **Palavras** (ou um novo grupo na aba **Grupos**). A aba **Como usar** explica cada coluna.
2. Gere o PDF de novo:
   ```
   pip install -r requirements.txt
   python gerar_pdf.py Guia_Ingles_Pronuncia.xlsx
   ```
3. Suba a planilha e o PDF atualizados para o repositório.

## Onde conferir a pronúncia

- Cambridge Dictionary (inglês–português, pronúncia US): https://dictionary.cambridge.org/pt/dicionario/ingles-portugues/
- Youglish (a palavra falada em vídeos reais, opção US): https://youglish.com
