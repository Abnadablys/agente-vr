import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from calendar import monthrange
import zipfile
import os
import tempfile
from io import BytesIO
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.agents import initialize_agent, Tool
from langchain.agents import AgentType
from langchain_experimental.utilities import PythonREPL
import warnings  # Para suprimir warnings depreciados

# Suprime warnings depreciados do LangChain
warnings.filterwarnings("ignore")

# Instruções: Obtenha uma API key gratuita em https://aistudio.google.com
# Coloque em st.secrets ou como env var: GEMINI_API_KEY
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", "")

if not GEMINI_API_KEY:
    st.error("Por favor, configure a GEMINI_API_KEY em st.secrets. Registre-se gratuitamente em https://aistudio.google.com para obter uma.")
    st.stop()

# Inicializa LLM com Gemini 2.0 Flash
llm = ChatGoogleGenerativeAI(
    model="gemini-2.0-flash",
    temperature=0,
    google_api_key=GEMINI_API_KEY
)

# Tool para execução de código Python (REPL seguro via LangChain)
python_repl = PythonREPL()
tools = [
    Tool(
        name="python_repl",
        func=python_repl.run,
        description="Uma ferramenta de REPL Python. Use para executar código Python. Entrada deve ser código válido."
    )
]

# Inicializa agente LangChain
agent = initialize_agent(
    tools=tools,
    llm=llm,
    agent=AgentType.ZERO_SHOT_REACT_DESCRIPTION,
    verbose=True,
    handle_parsing_errors=True
)

# Prompt template para o agente
system_prompt = """
Você é um agente de IA que automatiza o processo de compra de VR/VA com base em planilhas fornecidas.
Objetivo: Consolidar dados, aplicar exclusões (diretores, estagiários, aprendizes, afastados, exterior), tratar admissões/desligamentos/férias proporcionalmente, calcular dias úteis por sindicato, gerar planilha final com custo 80% empresa / 20% profissional.

Planilhas disponíveis nos caminhos: {file_paths}

Regras detalhadas:
- Base única: Consolidar ATIVOS, FÉRIAS, DESLIGADOS, ADMISSÃO, sindicato x valor, dias uteis.
- Exclusões: Remover por matrícula.
- Cálculos: Dias úteis por sindicato, subtrair férias proporcionais, proporcional para admissões.
- Desligamentos: Excluir se OK até dia 15; integral após.
- Saída: Gerar Excel 'VR MENSAL 05.2025.xlsx' com abas 'VR MENSAL 05.2025' e 'Validações'.

⚠️ Muito importante:
- Gere e execute código para processar as planilhas.
- O resultado final deve ser salvo em uma variável chamada 'output_excel',
que deve ser um objeto BytesIO contendo o binário do Excel gerado.
"""

# Interface Streamlit
st.title("Automação de VR/VA com Agente de IA (LangChain + Gemini)")

st.write("Suba um ZIP com as planilhas ou individualmente. O agente de IA processará via prompt, decidindo ferramentas e código.")

# Upload ZIP ou individuais
zip_file = st.file_uploader("Subir ZIP", type="zip")

st.subheader("Ou suba cada planilha individualmente:")
files = {
    'ATIVOS.xlsx': st.file_uploader("ATIVOS.xlsx", type="xlsx"),
    'AFASTAMENTOS.xlsx': st.file_uploader("AFASTAMENTOS.xlsx", type="xlsx"),
    'APRENDIZ.xlsx': st.file_uploader("APRENDIZ.xlsx", type="xlsx"),
    'DESLIGADOS.xlsx': st.file_uploader("DESLIGADOS.xlsx", type="xlsx"),
    'ESTÁGIO.xlsx': st.file_uploader("ESTÁGIO.xlsx", type="xlsx"),
    'EXTERIOR.xlsx': st.file_uploader("EXTERIOR.xlsx", type="xlsx"),
    'Base sindicato x valor.xlsx': st.file_uploader("Base sindicato x valor.xlsx", type="xlsx"),
    'Base dias uteis.xlsx': st.file_uploader("Base dias uteis.xlsx", type="xlsx"),
    'ADMISSÃO ABRIL.xlsx': st.file_uploader("ADMISSÃO ABRIL.xlsx", type="xlsx"),
    'FÉRIAS.xlsx': st.file_uploader("FÉRIAS.xlsx", type="xlsx"),
}

if st.button("Processar com Agente de IA"):
    with tempfile.TemporaryDirectory() as tmpdir:
        file_paths = {}
        required_files = list(files.keys())  # Lista dos 10 nomes necessários
        if zip_file:
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
            extracted_files = [f for f in os.listdir(tmpdir) if f.endswith('.xlsx')]
            missing_files = [f for f in required_files if f not in extracted_files]
            if missing_files:
                st.error(f"Planilhas faltando no ZIP: {', '.join(missing_files)}")
                st.stop()
            for fname in extracted_files:
                file_paths[fname] = os.path.join(tmpdir, fname)
        else:
            missing_files = []
            for fname, uploaded in files.items():
                if uploaded:
                    path = os.path.join(tmpdir, fname)
                    with open(path, 'wb') as f:
                        f.write(uploaded.read())
                    file_paths[fname] = path
                else:
                    missing_files.append(fname)
            if missing_files:
                st.error(f"Planilhas faltando: {', '.join(missing_files)}")
                st.stop()

        # Prompt para o agente
        prompt = system_prompt.format(file_paths=", ".join(file_paths.values()))

        # Chama o agente
        with st.spinner("Agente de IA processando (decidindo ferramentas e código)..."):
            try:
                response = agent.run(prompt)

                # Executa o código gerado pelo agente em um sandbox local
                exec_locals = {}
                try:
                    exec(response, {}, exec_locals)
                except Exception as e:
                    st.error(f"Erro ao executar código gerado: {e}")
                    st.stop()

                # Verifica se a saída foi criada
                if "output_excel" not in exec_locals:
                    st.error("O agente não gerou a variável 'output_excel'.")
                    st.stop()

                output_excel = exec_locals["output_excel"]  # BytesIO real

                st.success("✅ Processado com sucesso!")

                st.download_button(
                    "Baixar Planilha Final",
                    data=output_excel,
                    file_name="VR MENSAL 05.2025.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro no agente: {e}")
