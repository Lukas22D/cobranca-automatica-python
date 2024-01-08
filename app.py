import win32com.client as client
import pandas as pd
import datetime as dt


# Lendo o Arquivo excel
tabela = pd.read_excel('Contas a Receber.xlsx')

# Validação da Data Atual
data_atual = dt.datetime.now()

# Coletando dados dos clientes em atraso
tabela_devedores = tabela.loc[tabela['Status'] == 'Em aberto']
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento'] < data_atual]

try:
    # Criando objeto Outlook
    outlook = client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Obtendo a conta de e-mail
    emissor = outlook.session.Accounts['Seuemail@gmail.com']
    dados = tabela_devedores[['Valor em aberto', 'Data Prevista para pagamento', 'E-mail', 'NF']].values.tolist()

    for dado in dados:
        destinatario = dado[2]
        nf = dado[3]
        prazo = dado[1]
        prazo = prazo.strftime("%d/%m/%Y")
        valor = dado[0]
        assunto = f'Atraso de pagamento - NF {nf}'
        
        # Criando um novo item de e-mail
        mensagem = outlook.CreateItem(0)
        
        # Configurando destinatário, assunto e corpo do e-mail
        mensagem.To = destinatario
        mensagem.Subject = assunto
        corpo_mensagem = f'''
        Prezado Cliente,

        Verificamos um atraso no pagamento referente a NF {nf} com vencimento em {prazo} e valor total de R${valor:.2f}.
        Gostaríamos de verificar se há algum problema que necessite de auxílio de nossa equipe. 

        Em caso de dúvidas, é só entrar em contato com nosso time através do e-mail tecobreiautomaticamentecompython@gmail.com

        Att,
        Candiotto da Hashtag
        '''
        mensagem.Body = corpo_mensagem

        # Enviando a mensagem
        mensagem.Send()

except Exception as e:
    print(f"Erro ao criar objeto Outlook: {e}")
