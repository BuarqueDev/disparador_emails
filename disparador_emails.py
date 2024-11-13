import pandas as pd
import win32com.client as win32
import pythoncom
import streamlit as st
import time

# Inicializa o COM
pythoncom.CoInitialize()

# Título no navegador
st.title("Disparador de E-mails para Prestadores")

# Carregar a planilha via upload do usuário
uploaded_file = st.file_uploader("Carregue sua planilha em Excel", type=["xlsx"])

if uploaded_file is not None:
    # Carregar a planilha
    df = pd.read_excel(uploaded_file)
    st.write("Dados da Planilha:", df.head())  # Exibe as primeiras linhas para conferência

    # Carregar o HTML da assinatura
    with open("ass.html", "r", encoding="utf-8") as file:
        assinatura_html = file.read()

    # Botão para enviar os e-mails
    if st.button("Enviar E-mails"):
        # Cria uma instância do Outlook
        outlook = win32.Dispatch("outlook.application")

        # Contador de e-mails enviados
        email_count = 0

        # Agrupa os funcionários por prestador
        for nome_prestador in df["Nome Prestador"].unique():
            prestador_df = df[df["Nome Prestador"] == nome_prestador]
            email_prestador = prestador_df["Email Prestador"].iloc[0]
            assunto = f"{nome_prestador} - COMPARECIMENTO DO EXAME - ASO - URGENTE"

            # Cria o corpo do e-mail com formatação em HTML
            if "*" in nome_prestador:
                corpo = (
                    "Olá, prezados!<br><br>"
                    "Poderiam por gentileza verificar se os colaboradores abaixo compareceram para realização de exame, "
                    "se sim preencher o ASO com as informações necessárias no SOC.<br><br>"
                )
            else:
                corpo = (
                    "Olá, prezados!<br><br>"
                    "Poderiam por gentileza verificar se os colaboradores abaixo compareceram para realização de exame, "
                    "se sim encaminhar a cópia do ASO.<br><br>"
                )

            # Adiciona as informações dos funcionários com quebras de linha em HTML
            for _, row in prestador_df.iterrows():
                corpo += (
                    f"Empresa: {row['Empresa']}<br>"
                    f"CPF Funcionário: {row['CPF Funcionário']}<br>"
                    f"Funcionário: {row['Funcionário']}<br><br>"
                )

            # Envia o e-mail com corpo HTML e assinatura carregada
            mail = outlook.CreateItem(0)
            mail.To = email_prestador
            mail.Subject = assunto
            mail.HTMLBody = corpo + "<br><br>" + assinatura_html  # Anexa a assinatura carregada
            mail.Send()
            st.write(f"E-mail enviado para {nome_prestador}")

            # Marca o status de envio na planilha
            df.loc[prestador_df.index, 'Enviado'] = 'Sim'
            email_count += 1

            # Verifica se atingiu o limite de 50 e-mails
            if email_count % 50 == 0:
                st.write("Aguardando 5 minutos para o próximo lote de envios...")
                time.sleep(300)  # Pausa por 5 minutos (300 segundos)

        # Salva a planilha atualizada após enviar todos os e-mails
        df.to_excel("planilha_atualizada.xlsx", index=False)
        st.write("Todos os e-mails foram enviados! A planilha está pronta para download.")
        
        # Disponibiliza a planilha para download
        with open("planilha_atualizada.xlsx", "rb") as file:
            st.download_button("Baixar Planilha Atualizada", data=file, file_name="planilha_atualizada.xlsx")
