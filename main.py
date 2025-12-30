import os
import pandas as pd
from docx import Document

def extrair_dados():
  pasta = "./documentos"
  # Lista vazia para guardar os dados de cada arquivo
  lista_dados = []

  if not os.path.exists(pasta):
    print(f" A pasta '{pasta}' não existia. Criando-a agora...")
    os.makedirs(pasta)
    return

  # Lista todos os arquivos da pasta 'documentos'
  arquivos = [f for f in os.listdir(pasta) if f.endswith(".docx")]

  print(f"--- Processando {len(arquivos)} arquivos ---\n")
        
  for nome in arquivos:
    caminho_completo = os.path.join(pasta, nome)

    # Tratamento de Erro ao abrir o arquivo
    try:
      # Abre o arquivo sem o usuário ver
      doc = Document(caminho_completo)

      # Cria um "Dicionário" com os dados do arquivo atual
      dados_arquivo = {
        "Arquivo": nome,
        "Autor": doc.core_properties.author or "Não informado",
        "Título": doc.core_properties.title or "Sem título"
      }

      # Adiciona o dicionário na lista principal
      lista_dados.append(dados_arquivo)
      print(f" Lido: {nome}")
    
    except Exception as e:
      print(f" Erro em {nome}: {e}")

  # PANDAS
  if lista_dados:
    # Transforma a lista em uma tabela (DataFrame)
    df = pd.DataFrame(lista_dados)

    # Salva em um arquivo Excel
    nome_excel = "relatorio_documentos.xlsx"
    df.to_excel(nome_excel, index=False)
    print(f"\n SUCESSO! Planilha '{nome_excel}' gerada com sucesso.")
  else:
    print("Nenhum dado extraído para gerar o Excel.")

if __name__ == "__main__":
  extrair_dados()