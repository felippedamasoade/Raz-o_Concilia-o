import pandas as pd
import streamlit as st
from io import BytesIO

# Função para obter o saldo de conciliação
def obter_saldo_conciliacao(df):
    if 'Unnamed: 5' in df.columns:
        df.drop(columns=['Unnamed: 5'], inplace=True)

    colunas = ['filial', 'documento', 'bordero', 'ano', 'tipo', 'historico', 'debito', 'credito', 'saldo']
    
    if len(df.columns) == len(colunas):
        df.columns = colunas
    else:
        st.error("Erro: O número de colunas no arquivo não corresponde ao esperado.")
        return None

    df['Numero'] = df['historico'].str.extract(r'(\d+)')

    conciliacao = df.groupby('Numero').agg({
        'debito': 'sum',
        'credito': 'sum'
    }).reset_index()

    conciliacao['saldo'] = conciliacao['credito'] - conciliacao['debito']
    conciliacao['saldo'] = conciliacao['saldo'].round(3)
    conciliacao = conciliacao[conciliacao['saldo'] != 0]

    return conciliacao[['Numero', 'saldo']]

# Função para converter o DataFrame em Excel em memória
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# Interface do Streamlit
st.title("Conciliador de Saldo")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Escolha um arquivo Excel (formato .xls)", type="xls")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine='xlrd')
        resultado = obter_saldo_conciliacao(df)

        if resultado is not None:
            st.write("Resultado da conciliação:")
            st.dataframe(resultado)

            # Converter e baixar o resultado em Excel
            resultado_excel = to_excel(resultado)
            st.download_button(
                label="Baixar Resultado em Excel",
                data=resultado_excel,
                file_name='resultado_conciliacao.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
