import os
from docx import Document

def extrair_dados():
  pasta = "./documentos"

  if not os.path.exists(pasta):
    print("Pasta não encontrada!")
    return

  # Lista todos os arquivos da pasta 'documentos'
  arquivos = [f for f in os.listdir(pasta) if f.endswith((".docx"))]

  print(f"--- Iniciando a leitura de {len(arquivos)} arquivos ---\n")
        
  for nome in arquivos:
    caminho_completo = os.path.join(pasta, nome)

    # Tratamento de Erro ao abrir o arquivo
    try:
      # Abre o arquivo sem o usuário ver
      doc = Document(caminho_completo)

      # Le o autor que você definiu no gerador.py
      autor = doc.core_properties.author
      titulo = doc.core_properties.title

      print(f" Arquivo: {nome}")
      print(f" Autor: {autor or 'Desconhecido'}")
      print(f" Título: {titulo or 'Sem título'}")
      print("-" * 30)
    
    except Exception as e:
      print(f" Erro ao ler {nome}")
      print(f" Detalhe técnico: {e}")
      print("-" * 30)

if __name__ == "__main__":
  extrair_dados()