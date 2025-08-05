import pandas as pd  # Certifique-se de que o pandas está importado
import smtplib
import schedule
import time
from datetime import datetime, timedelta
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do e-mail
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS', 'diego.arinaldo@senff.com.br')  # Usar variáveis de ambiente para segurança
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', "Meteora549@")  # Usar variáveis de ambiente para segurança
TO_EMAIL = 'diego.santos@bsd.com.br'

# Função para enviar e-mail com codificação UTF-8
def send_email(subject, body):
    try:
        # Criar o objeto MIME
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = TO_EMAIL
        msg['Subject'] = subject
        
        # Definir o corpo do e-mail com codificação UTF-8
        body = MIMEText(body, 'plain', 'utf-8')
        msg.attach(body)
        
        # Usando SMTP_SSL para conexão segura
        with smtplib.SMTP_SSL('smtp.emailemnuvem.com.br', 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, TO_EMAIL, msg.as_string())
        
        print("E-mail de lembrete enviado com sucesso.")
    
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Função para verificar notas fiscais
def check_invoices():
    try:
        print("Lendo o arquivo Excel...")
        df = pd.read_excel(r'C:\Users\diego.arinaldo\Documents\python\Auto\Lançamentos esperados (1).xlsx')
        print("Arquivo lido com sucesso. Colunas:", df.columns)

        # Verifique o nome exato da coluna "Descritivo do serviço"
        if 'Descritivo do serviço' not in df.columns:
            print("Erro: A coluna 'Descritivo do serviço' não foi encontrada no arquivo.")
            print("Colunas disponíveis:", df.columns)
            return

        # Garantir que a coluna "Vencimento esperado" seja lida como número (dia do mês)
        df['Vencimento esperado'] = pd.to_numeric(df['Vencimento esperado'], errors='coerce').astype('Int64')
        print('Valores e tipos da coluna Vencimento esperado:')
        print(df[['Vencimento esperado', 'Descritivo do serviço']])
        print(df['Vencimento esperado'].apply(type))

        today = datetime.now().date()
        reminder_date = today + timedelta(days=2)  # Data de 2 dias após a data atual
        print(f"Hoje: {today}, Data de lembrete: {reminder_date}, Dia do lembrete: {reminder_date.day}")

        # Filtrar notas fiscais com vencimento em dois dias
        due_invoices = df[df['Vencimento esperado'] == int(reminder_date.day)]
        print(f"Notas fiscais encontradas para o dia {reminder_date.day}: {len(due_invoices)}")
        print(due_invoices)

        if not due_invoices.empty:
            subject = 'Lembrete: Notas Fiscais a Vencer'
            body = 'As seguintes notas fiscais estão prestes a vencer em dois dias:\n\n'
            for index, row in due_invoices.iterrows():
                body += f"- {row['Descritivo do serviço']} (Vencimento: {row['Vencimento esperado']})\n"
            print("Enviando e-mail...")
            send_email(subject, body)
        else:
            print("Nenhuma nota fiscal a vencer em dois dias.")

    except FileNotFoundError:
        print("O arquivo Excel não foi encontrado.")
    except Exception as e:
        print(f"Erro ao verificar notas fiscais: {e}")

# Agendar a verificação diária
schedule.every().day.at("08:34").do(check_invoices)

# Loop para manter o script em execução
while True:
    print("Aguardando próxima execução...")
    schedule.run_pending()
    time.sleep(60)  # Espera um minuto para a próxima execução





