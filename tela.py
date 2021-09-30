import csv
import os
import PySimpleGUI as sg
from datetime import datetime
from pyexcel.cookbook import merge_all_to_a_book
import glob


class Pessoa:
    def __init__(self):
        self.nome = ''
        self.sobrenome = ''
        self.cpf = ''
        self.data_nascimento = ''
        self.idade = ''
        self.maior_idade = ''

    @classmethod
    def cria_nova_pessoa(cls, nome, sobrenome, cpf, data_nascimento):
        pessoa = cls()
        pessoa.nome = nome
        pessoa.sobrenome = sobrenome
        pessoa.cpf = cpf
        pessoa.data_nascimento = data_nascimento
        return pessoa

    def calculo_idade(self):
        nova_data_nasc = datetime.strptime(self.data_nascimento, '%d/%m/%Y')
        data_atual = datetime.now()
        dias_ano = 365.25
        idade = ((data_atual - nova_data_nasc).days / dias_ano)
        self.idade = int(idade)

    def verifica_se_e_maior(self):
        if self.idade >= 18:
            self.maior_idade = 'true'
        else:
            self.maior_idade = 'false'

    def retorna_dicionario(self):
        user = {"Nome": self.nome,
                "Sobrenome": self.sobrenome,
                "CPF": self.cpf,
                "Data de Nascimento": self.data_nascimento,
                "Idade": self.idade,
                "maior_de_idade": self.maior_idade}
        return user


class InterfaceUsuario:
    def __init__(self):
        layout = [
            [sg.Text('Nome', size=(15, 0)), sg.Input(size=(30, 0), key='Nome')],
            [sg.Text('Sobrenome', size=(15, 0)), sg.Input(size=(30, 0), key='Sobrenome')],
            [sg.Text('CPF', size=(15, 0)), sg.Input(size=(30, 0), key='CPF')],
            [sg.Text('Data de Nascimento (dd/mm/aaaa)', size=(15, 0)),
             sg.Input(size=(30, 0), key='Data_de_Nascimento')],
            [sg.Text('Vaga (1,2,3,4,5)', size=(15, 0)), sg.Input(size=(30, 0), key='vaga')],
            [sg.Button('Salvar Dados')],
            [sg.Output(size=(50, 20))],
            [sg.Button('Cadastrar')],
        ]

        self.janela = sg.Window('Informações do Candidato').layout(layout)
        self.fechar_janela = False
        self.list_user = []

    def ler_botao(self):
        self.button, self.values = self.janela.Read()
        if self.button == 'Salvar Dados':
            self.extrair_dados()
        elif self.button == 'Cadastrar':
            self.realiza_cadastro()
        else:
            self.fechar_janela = True

    def extrair_dados(self):
        Nome = self.values['Nome']
        Sobrenome = self.values['Sobrenome']
        CPF = self.values['CPF']
        Data_de_Nascimento = self.values['Data_de_Nascimento']
        vaga = self.values['vaga']

        # Validação de campos vazios
        for k, v in self.values.items():
            if self.values[k] == "":
                print("=====================")
                print(f'ERRO! Favor preencher o campo {k}')
                return

        # Validação de limite de cadastro
        if len(self.list_user) >= 10:
            print("=====================")
            print("Numero maximo candidatos Atingido. Não será cadastrado!")
            return

        pessoa = Pessoa.cria_nova_pessoa(Nome, Sobrenome, CPF, Data_de_Nascimento)
        pessoa.calculo_idade()
        pessoa.verifica_se_e_maior()

        if self.verifica_se_vaga_ainda_disponivel(vaga):
            print("=====================")
            print(f'ERRO!! Todas as posições foram preenchidas para a vaga {vaga}')
            return

        if self.verifica_se_cpf_ja_cadastrado(pessoa):
            print("=====================")
            print('ERRO!! CPF já cadastrado')
            return

        # Adicionar novo usuário a lista depois de validar
        pessoa_dicionario = pessoa.retorna_dicionario()

        print("=====================")
        for k, v in pessoa_dicionario.items():
            if k == "maior_de_idade":
                if v == 'true':
                    print("Maior de idade")
                else:
                    print("Menor de idade")
            else:
                print(f'{k} : {v}')

        pessoa_dicionario['vaga'] = vaga
        self.list_user.append(pessoa_dicionario)

    def verifica_se_vaga_ainda_disponivel(self, vaga):
        contador = 0
        if os.path.exists("candidatos.csv"):
            with open('candidatos.csv', 'r') as arquivo:
                texto = arquivo.readlines()
            for register_user in texto:
                register_user = register_user.split(',')
                try:
                    register_user.index(vaga)
                    contador += 1
                    if contador >= 3:
                        return True
                except:
                    pass

        if len(self.list_user) > 0:
            for user_list in self.list_user:
                if user_list['vaga'] == vaga:
                    contador += 1
                    if contador >= 3:
                        return True
        return False

    def verifica_se_cpf_ja_cadastrado(self, pessoa):
        if os.path.exists("candidatos.csv"):
            with open('candidatos.csv', 'r') as arquivo:
                texto = arquivo.readlines()
            for register_user in texto:
                register_user = register_user.split(',')
                try:
                    register_user.index(pessoa.cpf)
                    return True
                except:
                    pass

        if len(self.list_user) > 0:
            for user_list in self.list_user:
                if user_list["CPF"] == pessoa.cpf:
                    return True

        return False

    def realiza_cadastro(self):
        print("=====================")
        print("CADASTRO REALIZADO COM SUCESSO")
        print("=====================")
        self.salvar_csv(self.list_user)
        self.list_user = []

    def salvar_csv(self, list_user):
        if os.path.exists("candidatos.csv"):
            with open("candidatos.csv", "a") as csvfile:
                field_names = ["Nome", "Sobrenome", "CPF", "Data de Nascimento", "Idade", "maior_de_idade", "vaga"]
                file = csv.DictWriter(csvfile, fieldnames=field_names)
                file.writerows(list_user)
                csvfile.close()
        else:
            with open("candidatos.csv", "w+") as csvfile:
                field_names = ["Nome", "Sobrenome", "CPF", "Data de Nascimento", "Idade", "maior_de_idade", "vaga"]
                file = csv.DictWriter(csvfile, fieldnames=field_names)
                file.writeheader()
                file.writerows(list_user)
                csvfile.close()

        merge_all_to_a_book(glob.glob("*.csv"), "output.xlsx")

    def fecha_janela(self):
        return self.fechar_janela


try:
    tela = InterfaceUsuario()
    while 1:
        tela.ler_botao()
        if tela.fechar_janela:
            break
except Exception as e:
    print(e)
