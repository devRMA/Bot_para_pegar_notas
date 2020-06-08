from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from openpyxl import load_workbook
from os.path import isfile # isfile('diretorio/arquivo.txt') verifica se existe um arquivo, retorna um valor bool
from datetime import datetime

# Função que vai abrir a conexão com o excel
def abrir_conexao_com_excel(local_planilha):
    try:
        return load_workbook(local_planilha)
    except:
        return False


# Função que vai ficar responsavel por adicionar as notas, no excel
def adicionar_nota(trimestre, notas, materia_para_adicionar):
    global materias, arquivo, driver
    colunas = ['C', 'D', 'E', 'F', 'G']
    linhas = list(range(5, 20))  # linhas, onde ficam as materias
    if arquivo:
        try:
            planilha = arquivo.get_sheet_by_name(trimestre)
        except:
            driver.close()
            raise SystemExit(f"""
A planilha "{trimestre}" não foi encontrada!
Entre em Boletim.xlsx e altere o nome da planilha para {trimestre} ou
entre em https://github.com/rafaelalves271/Pegar_nota_no_portal/blob/master/Boletim.xlsx e 
baixe o arquivo novamente
""")
        else:
            if len(notas) == len(colunas):
                qual_linha_adicionar = linhas[materias.index(materia_para_adicionar)]
                for c in range(0, len(colunas)):
                    planilha[f'{colunas[c]}{qual_linha_adicionar}'].value = int(notas[c])
    else:
        driver.close()
        raise SystemExit("""
O arquivo Boletim.xlsx não foi encontrado!
Entre em https://github.com/rafaelalves271/Pegar_nota_no_portal e 
baixe o arquivo Boletim.xlsx novamente""")


def fazer_enquanto_der_erro(codigo):
    while True:
        try:
            exec(codigo)
            break
        except:
            sleep(0.01)

while True:
    escolha = int(input("Qual trimestre você está fazendo?\n1 - 1° Trimestre\n2 - 2° Trimestre\n3 - 3° Trimestre\n>"))
    if escolha == 1:
        trimestre_usado = "1 Tri"
        break
    elif escolha == 2:
        trimestre_usado = "2 Tri"
        break
    elif escolha == 3:
        trimestre_usado = "3 Tri"
        break
existe_arquivo_config = isfile("config.txt")
if not existe_arquivo_config:
    print("Arquivo de configurações não encontrado!")
    user = str(input("Digite o CPF que você usa para logar no portal oficinas(sem '.' e sem '-'): "))
    if len(user) != 11:
        raise SystemExit("CPF inválido!")
    password = str(input("Digite sua senha: "))
    print("\n" * 40)
    ano_atual = int(datetime.today().strftime('%Y'))
    oficina = str(input("Digite, em qual oficina você está: ")).upper()
    oficina = f"{oficina} / {ano_atual} "
    sala = str(input("Digite o código e a sala que você está (exatamente como está no portal): "))
    arquivo_config = open('config.txt', 'w')
    conteudo = f"""# Não altere nada daqui, se não souber o que está fazendo
user = '{user}'
password = '{password}'
oficina = '{oficina}'
sala = '{sala}'"""
    arquivo_config.writelines(conteudo)
    arquivo_config.close()
else:
    print("Lendo configurações...")
    arquivo_config = open('config.txt', 'r')
    for c in arquivo_config.readlines():
        exec(c.replace("\n", ""))
    arquivo_config.close()
arquivo = abrir_conexao_com_excel('Boletim.xlsx')
materias = ["Arte",
            "Biologia",
            "Ciências Aplicadas",
            "Educação Física",
            "Filosofia",
            "Física",
            "Geografia",
            "História",
            "L.E.M. Inglês",
            "Língua Portuguesa",
            "Matemática",
            "Oficinas Tecnológicas",
            "Produção Textual",
            "Química",
            "Sociologia"
            ]
txt_provas = ["Alunos_0__Avaliacoes_0__Conceito",  # nota 1
              "Alunos_0__Avaliacoes_1__Conceito",  # nota 2
              "Alunos_0__Avaliacoes_2__Conceito",  # nota 3
              "Alunos_0__Avaliacoes_3__Conceito",  # nota 4
              "Alunos_0__Avaliacoes_4__Conceito",  # nota 5
              ]
print("Abrindo o browser.")
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get(
    "https://docentes.sesisenaipr.org.br/Corpore.Net/Login.aspx?ReturnUrl=%2fcorpore.net%2fMain.aspx%3fSelectedMenuID" +
    "Key%3dmnAcessoOficinaCNI%26ActionID%3dCstACessoOficinaActionWeb&SelectedMenuIDKey=mnAcessoOficinaCNI&ActionID" +
    "=CstACessoOficinaActionWeb")
print("Fazendo o login")
fazer_enquanto_der_erro(f"""
user = driver.find_element_by_id("txtUser")
user.send_keys('{user}')""")
fazer_enquanto_der_erro(f"""
password = driver.find_element_by_id("txtPass")
password.send_keys('{password}')
btn_acessar = driver.find_element_by_id("btnLogin")
btn_acessar.click()
""")

driver.get(
    "https://docentes.sesisenaipr.org.br/Corpore.Net/SharedServices/LibPages/ContextLoader.aspx?UnloadedPropertie"+
    "s=CodColigada;CodFilial;CodTipoCurso&Module=S&LoadFull=FALSE")
print("Passando do popup")
fazer_enquanto_der_erro("""
for c in range(0, 3):
    btn_avancar = driver.find_element_by_id("ctbAvancar")
    btn_avancar.click()
    sleep(0.9)
btn_concluir = driver.find_element_by_id("ctbConcluir")
btn_concluir.click()
driver.get("https://docentes.sesisenaipr.org.br/corpore.net/Main.aspx?SelectedMenuIDKey=MainLive")
""")
print("Abrindo a aba de notas")
fazer_enquanto_der_erro("""
driver.find_element_by_id("avaliacao_header").click()
""")
fazer_enquanto_der_erro("""
driver.find_element_by_id("avaliacao_content").click()
""")
print("Configurando a pesquisa das notas")
fazer_enquanto_der_erro(f"""
driver.find_element_by_xpath("//select[@name='IdOficina']/option[text()='{oficina}']").click()
""")
fazer_enquanto_der_erro("""
driver.find_element_by_xpath("//select[@name='CodTurno']/option[text()='(55) Matutino']").click()
""")
fazer_enquanto_der_erro(f"""
driver.find_element_by_xpath("//select[@name='IdSala']/option[text()='{sala}']").click()
""")
fazer_enquanto_der_erro("""
driver.find_element_by_xpath("//select[@name='IdAvaliacao']/option[text()='1 - Avaliação 1']").click()
""")

str_cmb_materia = "//select[@name='CodDisciplina']/option[text()='{}']"

for materia in materias:
    print(f"Pegando as notas de {materia}")
    while True:
        try:
            driver.find_element_by_xpath(str_cmb_materia.format(materia)).click()
            break
        except:
            try:
                driver.find_element_by_id("btnHabilitaFiltroGrid").click()
            except:
                sleep(0.001)
    fazer_enquanto_der_erro("""
driver.find_element_by_id("btnFiltro").click()
    """)
    var_notas = []
    for nota in range(0, len(txt_provas)):
        var_notas.append(driver.find_element_by_id(txt_provas[nota]).get_attribute('value'))
    adicionar_nota(trimestre_usado, var_notas, materia)
print("Salvando o excel"+"\033[31m"+"(não pare a execução enquanto o salvamento não termina)"+"\033[0;0m")
driver.close()
arquivo.save('Boletim.xlsx')
arquivo.close()
print("Salvo!")
