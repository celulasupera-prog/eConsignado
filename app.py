import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Consolidador de Planilhas CSV",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Consolidador de Planilhas CSV")
st.markdown(
    """
Consolide vários arquivos CSV (mesma estrutura) em uma única planilha Excel.
Você pode enviar um `.zip` com vários CSVs ou selecionar múltiplos CSVs diretamente.
"""
)

with st.sidebar:
    st.header("ℹ️ Informações")
    st.markdown(
        """
**Como usar:**
1. Faça upload de um arquivo ZIP com CSVs ou selecione múltiplos CSVs
2. Defina separador e codificação
3. (Opcional) informe a ordem final das colunas
4. Clique em **Consolidar Arquivos**
5. Baixe o Excel consolidado

**Formatos aceitos:**
- Arquivos CSV (`.csv`)
- Arquivos ZIP contendo CSVs (`.zip`)
"""
    )

st.header("1️⃣ Upload dos Arquivos")
upload_option = st.radio(
    "Escolha o tipo de upload:",
    ["Arquivo ZIP", "Múltiplos arquivos CSV"],
    horizontal=True,
)

st.header("2️⃣ Configurações")
col1, col2 = st.columns(2)

with col1:
    separador = st.selectbox(
        "Separador do CSV:",
        [";", ",", "\t", "|"],
        index=0,
        help="Caractere que separa as colunas nos CSVs",
    )

with col2:
    encoding = st.selectbox(
        "Codificação do arquivo:",
        ["utf-8", "latin1", "iso-8859-1", "cp1252"],
        index=0,
        help="Codificação de texto usada nos arquivos",
    )

create_bank_column = st.checkbox(
    "Criar coluna de banco no formato `codigo - descricao`",
    value=True,
    help=(
        "Quando existirem as colunas `ifConcessora.codigo` e "
        "`ifConcessora.descricao`, será criada a coluna `ifConcessora`."
    ),
)

st.header("3️⃣ Ordem das Colunas (Opcional)")
column_order_text = st.text_area(
    "Informe a ordem desejada das colunas (uma por linha)",
    placeholder="Exemplo:\nCPF\nNOME\nVALOR\nDATA",
    help=(
        "As colunas informadas serão posicionadas primeiro e na ordem exata definida. "
        "As demais colunas serão adicionadas automaticamente ao final."
    ),
)

preferred_columns = [line.strip() for line in column_order_text.splitlines() if line.strip()]

default_report_columns = [
    "matricula",
    "nomeTrabalhador",
    "inscricaoEmpregador.codigo",
    "nomeEmpregador",
    "valorParcela",
    "valorEmprestimo",
    "totalParcelas",
    "dataInicioContrato",
    "dataFimContrato",
    "competenciaInicioDesconto",
    "ifConcessora",
    "contrato",
    "inscricaoEmpregador.descricao",
    "numeroInscricaoEmpregador",
    "cpf",
    "competenciaFimDesconto",
    "valorLiberado",
    "qtdPagamentos",
    "qtdEscrituracoes",
    "categoriaTrabalhador.codigo",
    "categoriaTrabalhador.descricao",
    "competencia",
    "inscricaoEstabelecimento.codigo",
    "inscricaoEstabelecimento.descricao",
    "numeroInscricaoEstabelecimento",
    "dataAdmissao",
    "arquivo_origem",
]

use_default_report_order = st.checkbox(
    "Usar ordem padrão de colunas do relatório",
    value=True,
    help="Aplica automaticamente a ordem de colunas solicitada para o arquivo final.",
)

hide_columns_after_contract = st.checkbox(
    "Ocultar no Excel as colunas após `contrato`",
    value=True,
    help=(
        "Mantém visíveis apenas as 12 primeiras colunas do relatório "
        "(de `matricula` até `contrato`)."
    ),
)

uploaded_files = None

if upload_option == "Arquivo ZIP":
    uploaded_file = st.file_uploader(
        "Faça upload do arquivo ZIP contendo os CSVs:",
        type=["zip"],
        help="Selecione um ZIP que contenha os CSVs",
    )

    if uploaded_file:
        try:
            with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
                csv_files = sorted(
                    [f for f in zip_ref.namelist() if f.lower().endswith(".csv")]
                )

                if not csv_files:
                    st.warning("⚠️ Nenhum arquivo CSV foi encontrado dentro do ZIP.")
                else:
                    st.success(f"✅ {len(csv_files)} arquivos CSV encontrados no ZIP")
                    uploaded_files = [
                        {"name": csv_file, "content": zip_ref.read(csv_file)}
                        for csv_file in csv_files
                    ]

        except Exception as error:
            st.error(f"❌ Erro ao processar o arquivo ZIP: {error}")
else:
    uploaded_csvs = st.file_uploader(
        "Faça upload dos arquivos CSV:",
        type=["csv"],
        accept_multiple_files=True,
        help="Selecione um ou mais CSVs para consolidar",
    )

    if uploaded_csvs:
        st.success(f"✅ {len(uploaded_csvs)} arquivo(s) CSV selecionado(s)")
        uploaded_files = [
            {"name": csv_file.name, "content": csv_file.getvalue()} for csv_file in uploaded_csvs
        ]

if uploaded_files:
    st.header("4️⃣ Consolidação")

    with st.expander("📋 Ver lista de arquivos"):
        for index, file_info in enumerate(uploaded_files, start=1):
            st.text(f"{index}. {file_info['name']}")

    if st.button("🚀 Consolidar Arquivos", type="primary", use_container_width=True):
        try:
            with st.spinner("Processando arquivos..."):
                dataframes = []
                progress_bar = st.progress(0)
                status_text = st.empty()

                for idx, file_info in enumerate(uploaded_files):
                    status_text.text(f"Processando: {file_info['name']}")
                    df = pd.read_csv(
                        io.BytesIO(file_info["content"]),
                        sep=separador,
                        encoding=encoding,
                    )
                    df["arquivo_origem"] = file_info["name"]
                    dataframes.append(df)
                    progress_bar.progress((idx + 1) / len(uploaded_files))

                status_text.text("Consolidando dados...")
                df_consolidado = pd.concat(dataframes, ignore_index=True, sort=False)

                if "inscricaoEmpregador.codigo" in df_consolidado.columns:
                    df_consolidado["inscricaoEmpregador.codigo"] = ""

                if create_bank_column:
                    bank_code_column = "ifConcessora.codigo"
                    bank_description_column = "ifConcessora.descricao"

                    if (
                        bank_code_column in df_consolidado.columns
                        and bank_description_column in df_consolidado.columns
                    ):
                        bank_code = (
                            df_consolidado[bank_code_column]
                            .astype("string")
                            .fillna("")
                            .str.strip()
                            .str.replace(r"\\.0+$", "", regex=True)
                        )
                        bank_description = (
                            df_consolidado[bank_description_column]
                            .astype("string")
                            .fillna("")
                            .str.strip()
                        )

                        one_digit_mask = bank_code.str.fullmatch(r"\\d")
                        bank_code = bank_code.where(~one_digit_mask, bank_code.str.zfill(2))

                        combined_bank = (bank_code + " - " + bank_description).str.strip()
                        combined_bank = combined_bank.str.replace(r"^\\s*-\\s*", "", regex=True)
                        combined_bank = combined_bank.str.replace(r"\\s*-\\s*$", "", regex=True)
                        df_consolidado["ifConcessora"] = combined_bank
                    else:
                        st.warning(
                            "⚠️ Não foi possível criar `ifConcessora`: "
                            "colunas `ifConcessora.codigo` e/ou "
                            "`ifConcessora.descricao` não encontradas."
                        )

                if use_default_report_order:
                    existing_default_columns = [
                        column for column in default_report_columns if column in df_consolidado.columns
                    ]
                    missing_default_columns = [
                        column for column in default_report_columns if column not in df_consolidado.columns
                    ]
                    remaining_columns = [
                        column
                        for column in df_consolidado.columns
                        if column not in existing_default_columns
                    ]
                    final_columns = existing_default_columns + remaining_columns
                    df_consolidado = df_consolidado[final_columns]

                    if missing_default_columns:
                        st.warning(
                            "⚠️ Algumas colunas da ordem padrão não foram encontradas: "
                            + ", ".join(missing_default_columns)
                        )
                elif preferred_columns:
                    existing_preferred = [
                        column for column in preferred_columns if column in df_consolidado.columns
                    ]
                    missing_preferred = [
                        column for column in preferred_columns if column not in df_consolidado.columns
                    ]
                    remaining_columns = [
                        column
                        for column in df_consolidado.columns
                        if column not in existing_preferred
                    ]
                    final_columns = existing_preferred + remaining_columns
                    df_consolidado = df_consolidado[final_columns]

                    if missing_preferred:
                        st.warning(
                            "⚠️ Algumas colunas informadas não foram encontradas: "
                            + ", ".join(missing_preferred)
                        )

                progress_bar.empty()
                status_text.empty()

                st.success("✅ Consolidação concluída com sucesso!")
                m1, m2, m3 = st.columns(3)
                m1.metric("Arquivos Processados", len(uploaded_files))
                m2.metric("Total de Linhas", len(df_consolidado))
                m3.metric("Total de Colunas", len(df_consolidado.columns))

                st.subheader("👀 Prévia dos Dados Consolidados")
                st.dataframe(df_consolidado.head(10), use_container_width=True)

                with st.expander("📊 Informações das Colunas"):
                    st.write(f"**Colunas encontradas ({len(df_consolidado.columns)}):**")
                    for column in df_consolidado.columns:
                        st.text(f"• {column}")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name="Dados Consolidados")
                    if hide_columns_after_contract:
                        worksheet = writer.sheets["Dados Consolidados"]
                        first_hidden_column = 13
                        for column_idx in range(first_hidden_column, worksheet.max_column + 1):
                            column_letter = get_column_letter(column_idx)
                            worksheet.column_dimensions[column_letter].hidden = True

                output_filename = f"consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.download_button(
                    label="📥 Baixar Planilha Consolidada (Excel)",
                    data=output.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as error:
            st.error(f"❌ Erro ao processar os arquivos: {error}")
            st.exception(error)
else:
    st.info("👆 Faça upload dos arquivos para começar")

st.markdown("---")
st.markdown(
    """
<div style='text-align: center'>
    <p>Desenvolvido com ❤️ usando Streamlit</p>
</div>
""",
    unsafe_allow_html=True,
)
