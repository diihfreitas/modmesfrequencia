from docx import Document
import re
from datetime import datetime

# Caminhos dos arquivos
arquivo_original = "FICHA_ATUALIZADA.docx"
novo_arquivo = "FICHA_ATUALIZADA_05-25.docx"

# Mês e ano para atualizar
novo_mes = "05"  # Maio
ano_completo = "2025"

# Regex para encontrar datas no formato dd/mm/yy
padrao_data = re.compile(r"\b(\d{2})/(\d{2})/(\d{2})\b")

# Abre o documento
doc = Document(arquivo_original)

# Percorre tabelas (a ficha está em formato tabular)
for tabela in doc.tables:
    # Pula a primeira linha se for cabeçalho
    for linha in tabela.rows[1:]:
        celulas = linha.cells
        if len(celulas) >= 3:  # Verifica se tem colunas suficientes: DATA | ASSINATURA | CARIMBO
            data_texto = celulas[0].text.strip()  # Coluna "DATA"
            
            # Verifica e substitui a data com o novo mês
            nova_data = padrao_data.sub(
                lambda m: f"{m.group(1)}/{novo_mes}/{m.group(3)}", data_texto
            )
            celulas[0].text = nova_data

            # Converte para datetime para checar o dia da semana
            try:
                data_dt = datetime.strptime(nova_data, "%d/%m/%y")
                dia_semana = data_dt.strftime("%A").upper()
                if dia_semana == "SATURDAY":
                    celulas[1].text = "SÁBADO"
                elif dia_semana == "SUNDAY":
                    celulas[1].text = "DOMINGO"
                else:
                    celulas[1].text = ""  # Limpa a célula se não for fim de semana
            except ValueError:
                pass  # Se a data for inválida, ignora

# Salva novo arquivo
doc.save(novo_arquivo)
print(f"Arquivo atualizado salvo como: {novo_arquivo}")