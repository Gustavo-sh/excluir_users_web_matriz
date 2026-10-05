import win32com.client
import gspread
import pyodbc
from telegram_config import notify_telegram
from utils import write_log

CONNECTION_STRING = "Driver={SQL Server};Server=primno4;Database=Robbyson;Trusted_Connection=yes;"
CONN = pyodbc.connect(CONNECTION_STRING)
CURSOR = CONN.cursor()

def excluir_senha(matricula):
    CURSOR.execute(
        """
        DELETE FROM robbysonmatriz.dbo.acessos_sistema_matriz
        WHERE username = ?
        """,
        (matricula,)
    )
    CONN.commit()
    write_log(f"Senha de {matricula} excluida")


def gerar_pendentes():

    pendentes = {}
    client = gspread.service_account("credenciais.json")
    write_log("cliente coletado")
    planilha = client.open("Solicitação de exclusão de usuário - Sistema Matriz (respostas)")
    write_log("planilha coletada")
    aba = planilha.worksheet("Respostas")
    dados = aba.get_all_records()
    write_log("dados coletados")

    for numero_linha, linha in enumerate(dados, start=2):

        if linha["Tratado"] != "Sim":
            matricula = linha["Sua matricula AeC"]
            nome = linha["Seu nome completo (conforme hominum)"]
            email = linha["Seu e-mail AeC"]
            pendentes[matricula] = {
                "linha": numero_linha,
                "nome": nome,
                "email": email
            }

    write_log("aba e pendentes gerados")
    return aba, pendentes


def enviar_email(email):

    outlook = win32com.client.Dispatch("Outlook.Application")
    mensagem = outlook.CreateItem(0)
    mensagem.To = email
    mensagem.Subject = "Exclusão de Senha - WebMatriz"
    mensagem.HTMLBody = f"""
        <p>Olá,</p>

        <p>
            Sua senha de acesso ao WebMatriz foi excluída conforme solicitado.
        </p>

        <p>
            Você já pode efetuar um novo cadastro em
            <a href="https://matriz.metas.grupoaec.com.br/register">
                https://matriz.metas.grupoaec.com.br/register
            </a>.
        </p>

        <p>
            Atenciosamente,<br>
            Suporte WM.
        </p>
    """
    mensagem.Send()
    write_log(f"Email enviado para {email}")


def main():

    try:
        aba, pendentes = gerar_pendentes()
        if not pendentes:
            write_log("Nenhum pendente encontrado")
            return
        for matricula, dados in pendentes.items():
            excluir_senha(matricula)
            enviar_email(dados["email"])
            aba.update_cell(
                dados["linha"],
                6,
                "Sim"
            )

    except Exception as e:
        notify_telegram(f"Erro no processo de exclusão de senha: {e}")

    finally:
        CURSOR.close()
        CONN.close()


if __name__ == "__main__":
    main()