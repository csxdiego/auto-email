# auto-email
import pandas as pd  # Certifique-se de que o pandas está importado
import smtplib
import schedule
import time
from datetime import datetime, timedelta
import os

# Configurações do e-mail
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS', 'diego.arinaldo@senff.com.br')  # Usar variáveis de ambiente para segurança
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD', "Meteora549@")  # Usar variáveis de ambiente para segurança
TO_EMAIL = 'diego.santos@bsd.com.br'

# Função para enviar e-mail
def send_email(subject, body):
    try:
        with smtplib.SMTP('smtp.emailemnuvem.com.br', 587) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            message = f'Subject: {subject}\n\n{body}'
            server.sendmail(EMAIL_ADDRESS, TO_EMAIL, message)
        print("E-mail de lembrete enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Função para verificar notas fiscais
def check_invoices():
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(r'C:\Users\diego.arinaldo\Documents\python\Auto\Lançamentos esperados (1).xlsx')

        # Garantir que a coluna "Vencimento esperado" seja lida como número (dia do mês)
        df['Vencimento esperado'] = pd.to_numeric(df['Vencimento esperado'], errors='coerce')

        # Obter o dia atual
        today = datetime.now().date()
        reminder_date = today + timedelta(days=2)  # Data dois dias depois de hoje

        # Filtrar as notas fiscais que vencem dois dias após a data atual
        due_invoices = df[df['Vencimento esperado'] == reminder_date.day]

        if not due_invoices.empty:
            subject = 'Lembrete: Notas Fiscais a Vencer'
            body = 'As seguintes notas fiscais estão prestes a vencer em dois dias:\n\n'
            for index, row in due_invoices.iterrows():
                body += f"- {row['descritivo do serviço']} (Vencimento: {row['Vencimento esperado']})\n"
            send_email(subject, body)
        else:
            print("Nenhuma nota fiscal a vencer em dois dias.")

    except FileNotFoundError:
        print("O arquivo Excel não foi encontrado.")
    except Exception as e:
        print(f"Erro ao verificar notas fiscais: {e}")

# Agendar a verificação diária
schedule.every().day.at("09:00").do(check_invoices)

# Loop para manter o script em execução
while True:
    schedule.run_pending()
    time.sleep(60)  # Espera um minuto para a próxima execução
