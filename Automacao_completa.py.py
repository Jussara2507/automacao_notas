import pypdf
import pandas as pd
import re
import openpyxl
import unicodedata
from datetime import datetime, timedelta

def normalizar_texto(texto):
    """Remove acentos e padroniza o texto para comparação"""
    if not isinstance(texto, str):
        return ""
    nfkd = unicodedata.normalize('NFKD', texto)
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^a-zA-Z0-9]', ' ', sem_acento).upper().strip()

def obter_nome_aba_por_data(data_str):
    """Retorna o nome da aba formatado com base na data da nota (ex: Faturamento_Setembro26)"""
    meses_pt = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    dt = datetime.strptime(data_str, '%d/%m/%Y')
    nome_mes = meses_pt[dt.month]
    ano_curto = dt.strftime('%y')
    return f"Faturamento_{nome_mes}{ano_curto}"

def processar_automacao_completa(caminho_pdf, caminho_excel):
    print("1. A ler o relatório PDF da Mastermaq...")
    leitor = pypdf.PdfReader(caminho_pdf)
    texto = ""
    for pagina in leitor.pages:
        texto += pagina.extract_text() + "\n"
        
    linhas = texto.split('\n')
    notas_pdf = []
    
    for linha in linhas:
        linha = linha.strip()
        if not linha or "NOTAS PROCESSADAS" in linha or "Pág.:" in linha or "GR EFICAZ" in linha or "TOTAL" in linha or "Número" in linha:
            continue
            
        if "Autorizado o uso da NF-e" in linha:
            match = re.match(r'^(\d+)\s+(\d+)\s+([\d\.]+,\s*\d{2})\s+(\d{2}/\d{2}/\d{4})\s+(.+)', linha)
            if match:
                num_nota = int(match.group(1))
                serie = int(match.group(2))
                valor_str = match.group(3).replace('.', '').replace(',', '.')
                valor = float(valor_str)
                emissao_str = match.group(4)
                resto = match.group(5)
                
                nome_cliente_pdf = re.sub(r'\s*\d{2}[\.,]?\d{3}[\.,]?\d{3}[/\.]?\d{4}[-]?\d{2}.*', '', resto).strip()
                
                notas_pdf.append({
                    'Numero_Nota': num_nota,
                    'Serie': serie,
                    'Data_Emissao': emissao_str,
                    'Valor': valor,
                    'Nome_Cliente': nome_cliente_pdf
                })

    print(f"   -> Encontradas {len(notas_pdf)} notas autorizadas no PDF.")

    print("2. A abrir a folha de cálculo oficial de controlo...")
    wb = openpyxl.load_workbook(caminho_excel)
    ws_clientes = wb['Cadastro_Clientes']

    # Mapear clientes, prazos e formas de pagamento
    clientes_cadastrados = {}
    maior_id = 0
    
    for row in ws_clientes.iter_rows(min_row=2, min_col=1, max_col=13):
        cli_id = row[0].value
        cli_nome = row[1].value
        cli_prazo = row[11].value if len(row) > 11 else 30
        cli_forma = row[12].value if len(row) > 12 else "BOLETO"
        
        if cli_id is not None:
            try:
                cli_id_int = int(cli_id)
                if cli_id_int > maior_id:
                    maior_id = cli_id_int
            except ValueError:
                pass
                
        if cli_nome and cli_id is not None:
            nome_oficial = str(cli_nome).strip()
            chave_norm = normalizar_texto(nome_oficial)
            try:
                prazo_int = int(cli_prazo) if pd.notna(cli_prazo) else 30
            except ValueError:
                prazo_int = 30
            forma_str = str(cli_forma).strip() if pd.notna(cli_forma) else "BOLETO"
                
            clientes_cadastrados[chave_norm] = {
                'id': int(cli_id),
                'nome_oficial': nome_oficial,
                'prazo': prazo_int,
                'forma': forma_str
            }

    print("3. A processar notas e a direcionar para a aba correta de cada mês...")
    
    for nota in notas_pdf:
        nome_aba = obter_nome_aba_por_data(nota['Data_Emissao'])
        
        if nome_aba not in wb.sheetnames:
            ws_nova = wb.create_sheet(title=nome_aba)
            cabecalho = ["ID_Cliente", "Razão Social", "DT FATUR.", "DT VENCTO", "VALOR", "NOTA FISCAL", "FORMA PAGTO"]
            ws_nova.append([]) 
            ws_nova.append([]) 
            ws_nova.row_dimensions[2].height = 20
            for col_idx, text_head in enumerate(cabecalho, 1):
                ws_nova.cell(row=2, column=col_idx, value=text_head)
            print(f"   [NOVA ABA] A aba '{nome_aba}' foi criada automaticamente no Excel.")
            
        ws_faturamento = wb[nome_aba]

        notas_existentes = []
        for row in ws_faturamento.iter_rows(min_row=3, min_col=6, max_col=6):
            if row[0].value is not None:
                try:
                    notas_existentes.append(int(row[0].value))
                except ValueError:
                    pass

        if nota['Numero_Nota'] in notas_existentes:
            continue 

        nome_pdf_norm = normalizar_texto(nota['Nome_Cliente'])
        id_cliente = None
        razao_social_oficial = nota['Nome_Cliente']
        prazo_dias = 30
        forma_pagto = "BOLETO"
        
        primeira_palavra = nome_pdf_norm.split()[0] if nome_pdf_norm else ""
        
        for cad_norm, dados in clientes_cadastrados.items():
            if nome_pdf_norm in cad_norm or cad_norm in nome_pdf_norm or (primeira_palavra and primeira_palavra in cad_norm):
                id_cliente = dados['id']
                razao_social_oficial = dados['nome_oficial']
                prazo_dias = dados['prazo']
                forma_pagto = dados['forma']
                break
                
        if id_cliente is None:
            maior_id += 1
            id_cliente = maior_id
            razao_social_oficial = nota['Nome_Cliente']
            prazo_dias = 30
            forma_pagto = "BOLETO"
            
            proxima_linha_cli = 2
            while ws_clientes.cell(row=proxima_linha_cli, column=1).value is not None:
                proxima_linha_cli += 1
                
            ws_clientes.cell(row=proxima_linha_cli, column=1, value=id_cliente)
            ws_clientes.cell(row=proxima_linha_cli, column=2, value=razao_social_oficial)
            ws_clientes.cell(row=proxima_linha_cli, column=12, value=prazo_dias)
            ws_clientes.cell(row=proxima_linha_cli, column=13, value=forma_pagto)
            
            clientes_cadastrados[normalizar_texto(razao_social_oficial)] = {
                'id': id_cliente,
                'nome_oficial': razao_social_oficial,
                'prazo': prazo_dias,
                'forma': forma_pagto
            }
            print(f"   [NOVO CLIENTE] '{razao_social_oficial}' auto-cadastrado com ID {id_cliente}.")

        dt_fatur = datetime.strptime(nota['Data_Emissao'], '%d/%m/%Y')
        dt_vencto = dt_fatur + timedelta(days=prazo_dias)

        proxima_linha_fat = 3
        while ws_faturamento.cell(row=proxima_linha_fat, column=2).value is not None:
            proxima_linha_fat += 1

        ws_faturamento.cell(row=proxima_linha_fat, column=1, value=id_cliente)
        ws_faturamento.cell(row=proxima_linha_fat, column=2, value=razao_social_oficial)
        ws_faturamento.cell(row=proxima_linha_fat, column=3, value=dt_fatur.strftime('%d/%m/%Y'))
        ws_faturamento.cell(row=proxima_linha_fat, column=4, value=dt_vencto.strftime('%d/%m/%Y'))
        ws_faturamento.cell(row=proxima_linha_fat, column=5, value=nota['Valor'])
        ws_faturamento.cell(row=proxima_linha_fat, column=6, value=nota['Numero_Nota'])
        ws_faturamento.cell(row=proxima_linha_fat, column=7, value=forma_pagto)

        print(f"   -> Nota {nota['Numero_Nota']} ({razao_social_oficial}) lançada na aba [{nome_aba}].")

    wb.save(caminho_excel)
    print("Sucesso total! Faturamento processado e datas formatadas corretamente.")

if __name__ == "__main__":
    # Caminhos absolutos para garantir que o script ache os arquivos em qualquer situação
    caminho_pdf = r"C:\Users\jussa\Documents\Faturamento_GR_Eficaz\Listagem_Processo.pdf"
    caminho_excel = r"C:\Users\jussa\Documents\Faturamento_GR_Eficaz\SISTEMA DE CLIENTES E VENDAS - GR EFICAZ V3 MÊS 09.26.xlsx"
    
    processar_automacao_completa(caminho_pdf, caminho_excel)
    