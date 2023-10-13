import requests
import json
from datetime import datetime
import re
import locale
from num2words import num2words
from docxtpl import DocxTemplate as dt

class GeraDocumento:
    def __init__(self) -> None:
        """
            Consulta CNPJ API Client.
            Esta classe utiliza a api Consulta CNPJ Gratis para encontrar
            a razão social e endereço de uma empresa. Necessita criar uma conta
            gratuita para poder utilizar: https://rapidapi.com/cnpja/api/consulta-cnpj-gratis/
        
            Args:
            cnpj: Informe o cnpj que deseja consultar.
        """

        self.cnpj = ""
        self.secrets = "secrets/secrets.txt"
        self.basefile = "secrets/baseFile.docx"
        self.headers = {}
        self.querystring = {"simples":"false"}
        self.response = ""
        self.idPromiss = ""
        self.dVencimento = ""
        self.valorP = ""
        self.descMotivo = ""
        self.nPedido = ""
        self.dataEmiss = ""
        self.rNome = ""
        self.rCNPJ = ""
        self.rStreet = ""
        self.rNumber = ""
        self.rDistrict = ""
        self.rCity = ""
        self.rState = ""
        self.rZip = ""

    def start(self): #Inicia rotina de inputs e validações
        while True:
            self.cnpj = input("Informe o CNPJ (somente números):\n")
            if self.validCNPJ(self.cnpj):
                print("CNPJ no formato válido.")
                break
            else:
                print("Formato de CNPJ inválido.")
        self.confing()
        print("Configuração concluída.")
        self.consulta(self.cnpj)
        print("Consulta concluída")
        self.inputs()
        print("Dados incluídos.")
        self.dateSlipt(self.dVencimento)
        print("Data formatada.")
        self.longText(self.valorP)
        print("Valor por extenso ok.")
        self.valorP = self.addDot(self.valorP)
        print("Valor agora com separador de milhar.")
        self.emitionDate()
        print(f"Data de emissão {self.dataEmiss}")
        self.savedoc(idPromiss=self.idPromiss,
              dVencimento=self.dVencimento,
              valorP=self.valorP,
              vDia=self.vDia,
              vMes=self.vMes,
              vAno=self.vAno,
              vExtenso=self.vExtenso,
              descMotivo=self.descMotivo,
              nPedido=self.nPedido,
              dataEmiss=self.dataEmiss,
              rNome=self.rNome,
              rCNPJ=self.rCNPJ,
              rStreet=self.rStreet,
              rNumber=self.rNumber,
              rDistrict=self.rDistrict,
              rCity=self.rCity,
              rState=self.rState,
              rZip=self.rZip)
        print("Documento salvo!")
        input("Aperte Enter para sair...")
        #return self.headers

    def confing(self): #Adiciona a chave da API para consulta
        with open(self.secrets, 'r') as file:
            for line in file:
                self.headers = json.loads(line)
            return self.headers

    def inputs(self): #Solicita os dados de input para o documento
        self.idPromiss = input("Informe o número da promissória:\n")
        while True:
            self.dVencimento = input("Informe a data do vencimento (dd/mm/AAAA):\n")
            if self.validDate(self.dVencimento):
                print("Formato de data válido.")
                break
            else:
                print("Formato de data inválido.")
        while True:
            self.valorP = input("Informe o valor a ser pago (somente número com vírgula: xxx,xx):\n")
            if self.validValue(self.valorP):
                print("Formato de valor válido.")
                break
            else:
                print("Formato de valor inválido. Lembre-se de colocar os centavos.")
        
        self.descMotivo = input("Informe a descrição breve do motivo da antecipação:\n")
        self.nPedido = input("Informe o número do pedido:\n")
        self.nPedido = self.addCdot(self.nPedido)
        return self.idPromiss, self.dVencimento, self.valorP, self.descMotivo, self.nPedido
  
    def consulta(self, cnpj): #Consulta o CNPJ informado e retorna somente as informações relevantes
        url = f"https://consulta-cnpj-gratis.p.rapidapi.com/office/{cnpj}"
        self.response = requests.get(url, headers=self.headers, params=self.querystring)
        if self.response.status_code == 200:
            data = json.loads(self.response.text)
        
            self.rNome = data['company']['name']
            self.rCNPJ = data['taxId']
            self.rStreet = data['address']['street']
            self.rNumber = data['address']['number']
            self.rDistrict = data['address']['district']
            self.rCity = data['address']['city']
            self.rState = data['address']['state']
            self.rZip = data['address']['zip']

            return self.rNome, self.rCNPJ, self.rStreet, self.rNumber, self.rDistrict, self.rCity, self.rState, self.rZip
        else:
            print(f'Erro: {self.response.status_code}')
            print('Erro de consulta. Verifique o CNPJ informado.')

    def validCNPJ(self, cnpj): #Valida o formato do cnpj inserido
        pattern = r'^(\d{14})$'
        return bool(re.match(pattern, cnpj))

    def validDate(self, dateInput): #Valida a data inserida
        pattern = r'^(0[1-9]|[12]\d|3[01])/(0[1-9]|1[0-2])/(?P<year>\d{4})$'
        match = re.match(pattern, dateInput)
        if match: # Extract day, month, and year from the match
            day = int(match.group(1))
            month = int(match.group(2))
            year = int(match.group('year'))
            # Get the current year
            current_year = datetime.now().year
            # Define the valid year range (from current year to +5 years)
            valid_years = range(current_year, current_year + 6)
            # Check if the day, month, and year are valid
            if (day >= 1 and day <= 31) and (month >= 1 and month <= 12) and (year in valid_years):
                return True
        return False

    def dateSlipt(self, dateInput): #Divide a data em dia, mês e ano
        pattern = r'^(\d{2})/(\d{2})/(\d{4})$'
        match = re.match(pattern, dateInput)
        if match:
            self.vDia = match.group(1)
            #self.vMes = match.group(2)
            self.vAno = match.group(3)
            self.vMes = self.monthName(match.group(2))
            return self.vDia, self.monthName(self.vMes), self.vAno
        else:
            print("Date error")
    
    def validValue(self, valorPago): #Valida o formato do valor inserido
        pattern = r'^(\d{3,6}),(\d{2})$'
        return bool(re.match(pattern, valorPago))
    
    def monthName(self, monthNumber):
        monthNames = {
            '01': 'Janeiro',
            '02': 'Fevereiro',
            '03': 'Março',
            '04': 'Abril',
            '05': 'Maio',
            '06': 'Junho',
            '07': 'Julho',
            '08': 'Agosto',
            '09': 'Setembro',
            '10': 'Outubro',
            '11': 'Novembro',
            '12': 'Dezembro'
        }

        if monthNumber in monthNames:
            return monthNames[monthNumber]

    def emitionDate(self): #Retorna a data atual formatada
        today = datetime.now()
        self.dataEmiss = today.strftime('%d/%m/%Y')
        return self.dataEmiss

    def longText(self, valorPago): #Recebe o valor em número e retorna por extenso.
        pattern = r'^(\d{3,6}),(\d{2})$'
        match = re.match(pattern, valorPago)
        bucks = match.group(1)
        cents = match.group(2)
        if cents != "00":
            centss = num2words(cents, lang='pt_BR')
            centavos = f"e {centss} centavos"
        else:
            centavos = ""
        reais = num2words(bucks, lang='pt_BR')
        
        self.vExtenso = f"{reais} reais {centavos}"
        return self.vExtenso

    def addDot(self, valorP): #Recebe o valor e retorna com a pontuação de milhar.
        try:
            # Set the locale for number formatting to 'pt_BR'
            locale.setlocale(locale.LC_NUMERIC, 'pt_BR')

            # Replace a comma with a period for proper float conversion
            number = float(valorP.replace(',', '.'))

            # Format the number using locale.format
            formattedNumber = locale.format("%.2f", number, grouping=True)
            return formattedNumber
        
        except ValueError:
            print("Erro de formato")
            return valorP

    def addCdot(self, pedido):
        number_str = str(pedido)
        mid_index = len(number_str) // 2
        fnumber = number_str[:mid_index] + "." + number_str[mid_index:]
        return fnumber

    def savedoc(self, **kwargs): #Recebe todos os dados e salva no documento.
        """
        Esta função recebe os valores que serão preenchidos no documento final.
    
        Parameters:
        - idPromiss: Número da Promissória
        - dVencimento: Data do vencimento em dd/mm/AAAA
        - valorP: R$ X00.000,00
        - vDia: Dia do vencimento, número
        - vMes: Mês do vencimento, número
        - vAno: Ano do vencimento, número
        - vExtenso: Valor pago por extenso: xis mil reais
        - descMotivo: Descrição breve do motivo da antecipação
        - nPedido: Número do Pedido
        - dataEmiss: Data da emissão em dd/mm/AAAA
        - rNome: Razão Social
        - rCNPJ: CNPJ
        - rStreet: Nome da Rua
        - rNumber: nº
        - rDistrict: bairro
        - rCity: cidade
        - rState: UF
        - rZip: CEP
        """
        idPromiss = kwargs['idPromiss']
        dvenc = kwargs['dVencimento']
        valorP = kwargs['valorP']
        vDia = kwargs['vDia']
        vMes = kwargs['vMes']
        vAno = kwargs['vAno']
        vExtemsp = kwargs['vExtenso']
        desMotivo = kwargs['descMotivo']
        nPedido = kwargs['nPedido']
        dataEmiss = kwargs['dataEmiss']
        rNome = kwargs['rNome']
        rCNPJ = kwargs['rCNPJ']
        rStreet = kwargs['rStreet']
        rNumber = kwargs['rNumber']
        rDistrict = kwargs['rDistrict']
        rCity = kwargs['rCity']
        rState = kwargs['rState']
        rZip = kwargs['rZip']

        doc = dt(self.basefile)
        context = {
            'idPromiss' : idPromiss,
            'dVencimento' : dvenc,
            'valorP' : valorP,
            'vDia' : vDia,
            'vMes' : vMes,
            'vAno' : vAno,
            'vExtenso' : vExtemsp,
            'descMotivo' : desMotivo,
            'nPedido' : nPedido,
            'dataEmiss' : dataEmiss,
            'rNome' : rNome,
            'rCNPJ' : rCNPJ,
            'rStreet' : rStreet,
            'rNumber' : rNumber,
            'rDistrict' : rDistrict,
            'rCity' : rCity,
            'rState' : rState,
            'rZip' : rZip
        }
        doc.render(context)
        doc.save("generated_doc.docx")
