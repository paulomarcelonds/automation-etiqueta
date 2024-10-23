import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image, PageBreak
from reportlab.lib.units import inch
import qrcode
from io import BytesIO
from datetime import datetime

def format_date(date_str):
    """Converte a data para o formato dd/mm/yyyy"""
    try:
        date_obj = pd.to_datetime(date_str)
        return date_obj.strftime('%d/%m/%Y')
    except:
        return date_str

def create_qr_code(serials):
    """Cria um QR code a partir de uma lista de números seriais"""
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    # Junta os seriais com quebra de linha (\n) ao invés de vírgula
    serial_text = '\n'.join(str(serial) for serial in serials)
    qr.add_data(serial_text)
    qr.make(fit=True)
    
    # Cria uma imagem do QR code
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Converte a imagem para bytes
    img_buffer = BytesIO()
    img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    return img_buffer

def generate_label_pdf(excel_file, output_pdf):
    """Gera um PDF com etiquetas a partir de um arquivo Excel"""
    
    # Lê o arquivo Excel
    df = pd.read_excel(excel_file)
    
    # Cria o documento PDF
    doc = SimpleDocTemplate(output_pdf, 
                          pagesize=letter,
                          rightMargin=36,
                          leftMargin=36,
                          topMargin=36,
                          bottomMargin=36)
    elements = []
    
    # Agrupa os dados por caixa
    for idx, caixa in enumerate(df['CAIXA'].unique()):
        # Filtra dados para a caixa atual
        caixa_df = df[df['CAIXA'] == caixa]
        
        # Formata a data
        data_formatada = format_date(caixa_df['DATA'].iloc[0])
        
        # Prepara os dados para a etiqueta
        data = [
            ['CLARO', 'ETIQUETA DE RETORNO REVERSA'],
            ['NOME', caixa_df['NOME'].iloc[0]],
            ['DATA', data_formatada],
            ['CD', caixa_df['CD'].iloc[0]],
            ['CIDADE', caixa_df['CIDADE'].iloc[0]],
            ['COD._ITEM', caixa_df['COD._ITEM'].iloc[0]],
            ['DESCRICAO', caixa_df['DESCRICAO'].iloc[0]],
            ['QUANTIDADE', len(caixa_df)],
            ['N._Nfe', caixa_df['N._Nfe'].iloc[0]],
            ['CAIXA', caixa_df['CAIXA'].iloc[0]],
            ['LOTE', caixa_df['LOTE'].iloc[0]]
        ]
        
        # Cria a tabela com tamanho ajustado
        table = Table(data, colWidths=[3*inch, 4*inch])
        
        # Estilo da tabela
        table_style = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
            ('BACKGROUND', (1, 0), (1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWHEIGHT', (0, 0), (-1, -1), 25),
        ])
        
        table.setStyle(table_style)
        
        # Adiciona espaço no topo da página
        elements.append(Spacer(1, 30))
        elements.append(table)
        
        # Gera QR code com os seriais da caixa
        serials = caixa_df['SERIAL'].tolist()
        qr_buffer = create_qr_code(serials)
        
        # Cria a imagem do QR code e adiciona ao PDF
        qr_img = Image(qr_buffer, width=2*inch, height=2*inch)
        elements.append(Spacer(1, 20))
        elements.append(qr_img)
        
        # Adiciona quebra de página após cada etiqueta (exceto a última)
        if idx < len(df['CAIXA'].unique()) - 1:
            elements.append(PageBreak())
    
    # Gera o PDF
    doc.build(elements)

def main():
    # Exemplo de uso
    excel_file = r"C:\Users\working_09062024\Documents\automation-etiqueta\reversa.xlsx"
    output_pdf = "etiquetas_saida.pdf"
    
    try:
        generate_label_pdf(excel_file, output_pdf)
        print(f"PDF gerado com sucesso: {output_pdf}")
    except Exception as e:
        print(f"Erro ao gerar PDF: {str(e)}")

if __name__ == "__main__":
    main()