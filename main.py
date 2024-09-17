import pandas as pd
from datetime import datetime
import os
import sys

# Função para obter o diretório base (funciona tanto para script quanto para exe)
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Configuração do sistema de log
def setup_logger():
    base_dir = get_base_dir()
    log_dir = os.path.join(base_dir, 'log')
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%d%m%Y-%H-%M-%S")
    log_file = os.path.join(log_dir, f"{timestamp}.log")
    
    class Logger(object):
        def __init__(self, file):
            self.terminal = sys.stdout
            self.log = open(file, "w", encoding='utf-8')
   
        def write(self, message):
            self.terminal.write(message)
            self.log.write(message)
            self.log.flush()

        def flush(self):
            pass
    
    sys.stdout = Logger(log_file)
    print(f"Log iniciado em: {log_file}")

# Função para limpar a data
def limpar_data(data_str):
    if isinstance(data_str, str):
        data_str = data_str.split('às')[0].strip()
        try:
            return datetime.strptime(data_str, '%d/%m/%Y')
        except ValueError:
            return None
    elif isinstance(data_str, datetime):
        return data_str
    return None

def processar_edenred(file_path):
    try:
        df = pd.read_excel(file_path, header=None)
        print(f"Total de linhas no DataFrame Edenred: {len(df)}")
        
        nome_vr = str(df.iloc[0, 0]).strip()
        movimentos = []

        for i in range(1, len(df), 4):
            if i+3 < len(df):
                data = limpar_data(df.iloc[i, 0])
                descricao = df.iloc[i+1, 0]
                gasto = df.iloc[i+2, 0]
                
                descricao_limpa = descricao.replace("Compra: ", "") if isinstance(descricao, str) else descricao

                if data:
                    movimento = {
                        'data': data,
                        'descricao': descricao_limpa,
                        'valor': float(str(gasto).replace(',', '.')),
                        'nome': nome_vr,
                        'categoria': 'Alimentação'
                    }
                    movimentos.append(movimento)
        
        return movimentos
    except Exception as e:
        print(f"Erro ao processar arquivo Edenred: {e}")
        return []

def processar_activobank(file_path):
    try:
        df = pd.read_excel(file_path)
        print(f"Total de linhas no DataFrame ActivoBank: {len(df)}")
        
        movimentos = []

        for _, row in df.iterrows():
            data = limpar_data(row.iloc[0])
            descricao = row.iloc[2]
            valor = row.iloc[3]
            
            if data and not pd.isna(valor):
                # Remove "COMPRA" da descrição, se existir
                descricao_limpa = descricao.replace("COMPRA ", "").strip() if isinstance(descricao, str) else descricao
                
                movimento = {
                    'data': data,
                    'descricao': descricao_limpa,
                    'valor': float(valor),
                    'nome': 'Conjunta',
                    'categoria': ''
                }
                movimentos.append(movimento)
        
        return movimentos
    except Exception as e:
        print(f"Erro ao processar arquivo ActivoBank: {e}")
        return []

# Configuração do logger
setup_logger()

# Obtém o diretório base
base_dir = get_base_dir()

# Define os caminhos dos arquivos
edenred_path = os.path.join(base_dir, 'edenred.xlsx')
activobank_path = os.path.join(base_dir, 'activobank.xlsx')

# Verifica se os arquivos existem
if not os.path.exists(edenred_path):
    print(f"Arquivo não encontrado: {edenred_path}")
if not os.path.exists(activobank_path):
    print(f"Arquivo não encontrado: {activobank_path}")

# Processamento dos arquivos
movimentos_edenred = processar_edenred(edenred_path)
movimentos_activobank = processar_activobank(activobank_path)

# Combina os movimentos
movimentos = movimentos_edenred + movimentos_activobank

if not movimentos:
    print("Nenhum movimento foi processado. Verifique os arquivos de entrada.")
else:
    # Ordena os movimentos por data
    movimentos.sort(key=lambda x: x['data'])

    print(f"Total de movimentos encontrados: {len(movimentos)}")
    
    # Imprime todos os movimentos no log
    print("\nDetalhes de todos os movimentos:")
    for i, movimento in enumerate(movimentos, 1):
        print(f"\nMovimento {i}:")
        print(f"Data: {movimento['data'].strftime('%d/%m/%Y')}")
        print(f"Descrição: {movimento['descricao']}")
        print(f"Valor: {movimento['valor']:.2f}")
        print(f"Nome: {movimento['nome']}")
        print(f"Categoria: {movimento['categoria']}")

    # Cria um novo DataFrame com os movimentos processados
    df_output = pd.DataFrame([
        {
            'Descrição': m['descricao'],
            'Nome': m['nome'],
            'Categoria': m['categoria'],
            'Subcategoria': '',
            'Data': m['data'].strftime('%d/%m/%Y'),
            'Valor': m['valor']
        } for m in movimentos
    ])

    # Salva o DataFrame como um arquivo Excel
    output_path = os.path.join(base_dir, 'final_output.xlsx')
    df_output.to_excel(output_path, index=False)

    print(f"\nDados salvos em {output_path}")
