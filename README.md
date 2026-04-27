# 📊 Consolidador de Planilhas CSV

Aplicação em Streamlit para consolidar vários arquivos CSV com a mesma estrutura em uma única planilha Excel.

## ✅ O que já está pronto

- Upload de `.zip` com vários `.csv`
- Upload direto de múltiplos `.csv`
- Configuração de separador e codificação
- Consolidação em uma aba única de Excel
- Coluna `arquivo_origem` para rastrear de qual arquivo cada linha veio
- Reordenação opcional das colunas (informada pelo usuário)

## 🚀 Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Depois, acesse `http://localhost:8501`.

## 🧭 Próximos passos sugeridos

- Filtros antes da exportação
- Remoção de duplicatas
- Validação de tipos de dados
- Regras salvas para ordem de colunas
