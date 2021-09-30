import csv
import os
import PySimpleGUI as sg
from datetime import datetime
from pyexcel.cookbook import merge_all_to_a_book
import glob


class Pessoa:
    def __init__(self):
        layout = [
            [sg.Text('Nome', size=(10, 0)), sg.Input(size=(30, 0), key='Nome')],
            [sg.Text('Sobrenome', size=(10, 0)), sg.Input(size=(30, 0), key='Sobrenome')],
            [sg.Text('CPF', size=(10, 0)), sg.Input(size=(30, 0), key='CPF')],
            [sg.Text('Data de Nascimento (dd/mm/aaaa)', size=(10, 0)),
             sg.Input(size=(30, 0), key='Data_de_Nascimento')],
            [sg.Button('Salvar Dados')],
            [sg.Output(size=(40, 20))],
            [sg.Button('Cadastrar')],
        ]

        self.janela = sg.Window('Informações do Candidato').layout(layout)
        self.fechar_janela = False
        self.list_user = []

    def ler_botao(self):
        self.Button, self.values = self.janela.Read()
        if self.Button == 'Salvar Dados':
            self.ExtrairDados()
        elif self.Button == 'Cadastrar':
            print("=====================")
            print("CADASTRO REALIZADO COM SUCESSO")
            print("=====================")
            self.salvar_csv(self.list_user)
            self.list_user = []
        else:
            self.fechar_janela = True

    def ExtrairDados(self):
        Nome = self.values['Nome']
        Sobrenome = self.values['Sobrenome']
        CPF = self.values['CPF']
        Data_de_Nascimento = self.values['Data_de_Nascimento']

        # Validação de campos vazios
        for k, v in self.values.items():
            if self.values[k] == "":
                print(f'ERRO! Favor preencher o campo {k}. O cliente não foi adicionado a lista.')
                return

        # Validação de limite de cadastro
        if len(self.list_user) >= 3:
            print("Numero maximo candidatos Atingido. Não será cadastrado!")
            return

        # Calculo da idade
        nova_data_nasc = datetime.strptime(Data_de_Nascimento, '%d/%m/%Y')
        data_atual = datetime.now()
        dias_ano = 365.25
        idade = ((data_atual - nova_data_nasc).days / dias_ano)
        idade = int(idade)

        user = {"Nome": Nome,
                "Sobrenome": Sobrenome,
                "CPF": CPF,
                "Data de Nascimento": Data_de_Nascimento,
                "Idade": idade}

        if self.confirm_cpf(user):
            print("=====================")
            print('ERRO!! CPF já cadastrado')
            return

        # Adicionando campo de maior de idade
        if user['Idade'] >= 18:
            user["maior_de_idade"] = 'true'
        else:
            user["maior_de_idade"] = 'false'

        # Adicionar novo usuário a lista depois de validar
        self.list_user.append(user)
        print("=====================")
        for k, v in user.items():
            if k == "maior_de_idade":
                if v == 'true':
                    print("Maior de idade")
                else:
                    print("Menor de idade")
            else:
                print(f'{k} : {v}')

    def confirm_cpf(self, user):
        if os.path.exists("candidatos.csv"):
            with open('candidatos.csv', 'r') as arquivo:
                texto = arquivo.readlines()
            for register_user in texto:
                register_user = register_user.split(',')
                try:
                    register_user.index(user['CPF'])
                    return True
                except:
                    pass
        if len(self.list_user) > 0:
            for user_list in self.list_user:
                if user_list["CPF"] == user["CPF"]:
                    return True
        return False

    def salvar_csv(self, list_user):
        if os.path.exists("candidatos.csv"):
            with open('candidatos.csv', 'a') as csvfile:
                field_names = ["Nome", "Sobrenome", "CPF", "Data de Nascimento", "Idade", "maior_de_idade"]
                file = csv.DictWriter(csvfile, fieldnames=field_names)
                file.writerows(list_user)
                csvfile.close()
        else:
            with open('candidatos.csv', 'w+') as csvfile:
                field_names = ["Nome", "Sobrenome", "CPF", "Data de Nascimento", "Idade", "maior_de_idade"]
                file = csv.DictWriter(csvfile, fieldnames=field_names)
                file.writeheader()
                file.writerows(list_user)
                csvfile.close()

        merge_all_to_a_book(glob.glob("*.csv"), "output.xlsx")

    def fecha_janela(self):
        return self.fechar_janela


try:
    tela = Pessoa()
    while 1:
        tela.ler_botao()
        if tela.fechar_janela:
            break

except Exception as e:
    print(e)
