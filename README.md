# 📊 Consolidador de Planilhas CSV

Aplicação em Streamlit para consolidar vários arquivos CSV com a mesma estrutura em uma única planilha Excel.

## ✅ O que já está pronto

- Upload de `.zip` com vários `.csv`
- Upload direto de múltiplos `.csv`
- Configuração de separador e codificação
- Consolidação em uma aba única de Excel
- Coluna `arquivo_origem` para rastrear de qual arquivo cada linha veio
- Reordenação opcional das colunas (informada pelo usuário)
- Geração opcional da coluna `ifConcessora` no formato `codigo - descricao` (ex.: `01 - Nubank`)
- Ordem padrão do relatório com priorização automática das colunas principais
- Limpeza automática de `inscricaoEmpregador.codigo` (coluna exportada em branco no consolidado)
- Opção para ocultar no Excel as colunas após `contrato` (mantendo visíveis as 12 primeiras)

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


## ☁️ Deploy (Streamlit Cloud)

Para evitar incompatibilidades de build no ambiente cloud, este projeto fixa o runtime em **Python 3.12.10** via `runtime.txt`.

Se o deploy ficar “in the oven” por muito tempo, normalmente é por compilação de dependências em versão de Python sem wheel pronta.


## 🧯 Troubleshooting (Streamlit Cloud)

Se aparecer *"Your app is in the oven"* por muito tempo:

1. Vá em **App settings → Advanced settings** e confirme `Python 3.12`.
2. Clique em **Reboot app** e marque **Clear cache**.
3. Confirme se o branch selecionado contém este `requirements.txt` sem versões antigas fixas.

O erro de `pillow==10.4.0` + `Python 3.14` acontece quando o deploy usa dependências antigas sem wheel compatível.
