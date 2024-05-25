# Import for the Desktop Bot
from botcity.core import DesktopBot, Backend

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = DesktopBot()

    app_path = r'D:\Sicalc Auto Atendimento\SicalcAA.exe'
    
    # Implement here your logic...
    bot.execute(app_path)
    
    bot.connect_to_app(backend=Backend.WIN_32, path=app_path)
    
    #Encontrando pop-up pelo titulo da janela
    janela_esclarecimento = bot.find_app_window(title='Esclarecimento ao Contribuinte')
    
    #Aguarda 1s
    bot.wait(1000)
    
    #Encontra o botão de continuar a partir da janela encontrada anteriormente 
    btn_continuar = bot.find_app_element(
        from_parent_window=janela_esclarecimento, 
        title='&Continuar', 
        class_name='ThunderRT6CommandButton')
    btn_continuar.click()
    
    
    #Encontrando referência da janela principal pelo titulo
    janela_principal = bot.find_app_window(
        title='Sicalc Auto Atendimento',
        class_name='ThunderRT6MDIForm'
    )
    
    #Selecionando a opção no menu
    janela_principal.menu_select('Funções -> Preenchimento de DARF')
    
    #Formulário DARF
    darf = bot.find_app_element(
        from_parent_window=janela_principal,
        title='Preenchimento de DARF',
        class_name='ThunderRT6FormDC')
    
    #Acessando campo 'Edit' referente ao 'Cód. Receita'
    #A partir da janela do formulário enontrada anteriomente
    darf.Edit3.type_keys('5629')
    
    #Tecla tab para avançar o processo
    bot.tab()
    
    #Encontrando novamente a referência da janela principal
    janela_principal = bot.find_app_window(
        title_re='Sicalc Auto Atendimento',
        class_name='ThunderRT6MDIForm'
    )
    
    #Encontrando a referência do segundo formulário
    form_darf = bot.find_app_element(from_parent_window=janela_principal,title='Receita',class_name='ThunderRT6Frame')
    bot.wait(2000)
    form_darf.type_keys("{TAB}")
    
    #Período de apuração
    form_darf.Edit4.type_keys('310120')
    bot.wait(2000)
    form_darf.type_keys("{TAB}")
    
    #Valor em reais
    bot.wait(2000)
    form_darf.Edit5.type_keys('10000')
    
    #Botão Calcular
    form_darf.type_keys("{ENTER}")
    
    #Atalho para o botão DARF
    #Nessa situação, a construção "%{<tecla>}" corresponde aos atalhos que utilizam ALT
    form_darf.type_keys("%{f}")
    
    #Preenchendo ultimo formulario
    form_darf = bot.find_app_window(title="Preenchimento DARF Auto Atendimento", class_name="ThunderRT6FormDC")

    #Nome
    form_darf.Edit5.type_keys("Petrobras")

    #Telefone
    form_darf.Edit6.type_keys("1199991234")

    #CNPJ
    form_darf.Edit11.type_keys("33000167000101")

    #Referencia
    form_darf.Edit10.type_keys("0")
    
    #Imprimir
    btn_imprimir = bot.find_app_element(from_parent_window=form_darf, title="&Imprimir", class_name="ThunderRT6CommandButton")
    btn_imprimir.click()
    
    #Salvando arquivo PDF
    save = bot.find_app_window(title="Salvar Saída de Impressão como")
    save.type_keys(r"C:\Users\55929\Documents\DARF.pdf")
    save.type_keys("{ENTER}")

    # Fechando janela do formulário
    form_darf.type_keys("%{F4}")
    
    
    
    

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finished OK."
    )

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()