import re
import argparse
import pandas as pd
import pypdf

def extract_base_data(pdf_path, output_excel_path, target_bases=None):
    """
    Extrai dados do Relatório de Manuseio da CPFL para bases operacionais específicas.
    """
    if target_bases is None:
        target_bases = ['1223', '1247', '1262']
    
    print(f"Iniciando a leitura do arquivo PDF: {pdf_path}")
    reader = pypdf.PdfReader(pdf_path)
    total_pages = len(reader.pages)
    print(f"Total de páginas encontradas: {total_pages}")
    
    # Expressões regulares ajustadas ao padrão do PDF CPFL
    city_pattern = re.compile(r'^#{1,4}\s*(.+)')
    ul_pattern = re.compile(r'Unidade de Leitura \(UL\):\s*(.+)', re.IGNORECASE)
    data_row_pattern = re.compile(r'^(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)$')
    
    rows = []
    curr_base = None
    curr_city = None
    curr_ul = None
    
    # Varre todas as páginas do PDF
    for page_idx, page in enumerate(reader.pages):
        page_num = page_idx + 1
        text = page.extract_text()
        if not text:
            continue
            
        lines = text.split('\n')
        for line in lines:
            line_s = line.strip()
            if not line_s:
                continue
            
            # Identifica a Base Operacional (ex: 1223 - RIBEIRÃO PRETO)
            if re.match(r'^\d{4}\s*-', line_s):
                curr_base = line_s
                continue
                
            # Identifica a Cidade / Município
            m_city = city_pattern.match(line_s)
            if m_city:
                curr_city = m_city.group(1).strip()
                continue
                
            # Identifica a Unidade de Leitura (UL)
            m_ul = ul_pattern.search(line_s)
            if m_ul:
                curr_ul = m_ul.group(1).strip()
                continue
                
            # Identifica a linha numérica de dados do pacote
            m_data = data_row_pattern.match(line_s)
            if m_data:
                # Filtra apenas se a base for uma das informadas em target_bases
                if curr_base and any(base_code in curr_base for base_code in target_bases):
                    cod_entr, nro_pacote, nu_inicial, nu_final, qtde_contas = m_data.groups()
                    rows.append({
                        'Base Operacional': curr_base,
                        'Cidade': curr_city,
                        'Cod Entr': cod_entr,
                        'Nro Pacote': nro_pacote,
                        'Nu Impr Inicial': int(nu_inicial),
                        'Nu Impr Final': int(nu_final),
                        'Qtde de Contas': int(qtde_contas),
                        'Unidade de Leitura (UL)': curr_ul,
                        'Página PDF': page_num
                    })
                    
    print(f"Processamento concluído! Registros extraídos: {len(rows)}")
    
    if len(rows) == 0:
        print("Nenhum registro foi localizado para as bases informadas.")
        return
        
    # Gera o DataFrame e salva em Excel com colunas autorreguladas
    df = pd.DataFrame(rows)
    
    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Bases Operacionais')
        
        worksheet = writer.sheets['Bases Operacionais']
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    print(f"Arquivo Excel gerado com sucesso em: {output_excel_path}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extrai dados de manuseio CPFL por base operacional.')
    parser.add_argument('--pdf', required=True, help='Caminho do arquivo PDF')
    parser.add_argument('--output', default='relatorio_bases.xlsx', help='Caminho de saída para o arquivo Excel')
    
    args = parser.parse_args()
    extract_base_data(args.pdf, args.output, target_bases=['1223', '1247', '1262'])
