from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import json
import os

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/importar", methods=["POST"])
def importar():
    """Importa dados de uma planilha existente para edição"""
    try:
        if 'arquivo' not in request.files:
            return jsonify({'erro': 'Nenhum arquivo foi enviado'}), 400
        
        arquivo = request.files['arquivo']
        if arquivo.filename == '':
            return jsonify({'erro': 'Nenhum arquivo foi selecionado'}), 400
        
        # Verifica se é um arquivo Excel
        if not arquivo.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({'erro': 'Arquivo deve ser uma planilha Excel (.xlsx ou .xls)'}), 400
        
        # Carrega a planilha
        wb = load_workbook(arquivo)
        ws = wb.active
        
        dados_importados = []
        
        # Lê os dados a partir da linha 3 (pula cabeçalho e totais)
        for linha in range(3, ws.max_row + 1):
            # Verifica se a linha tem dados (pelo menos o código do produto)
            codigo = ws.cell(row=linha, column=2).value
            if not codigo:
                continue
                
            item = {
                'codigo': str(codigo) if codigo else '',
                'cfop': str(ws.cell(row=linha, column=3).value) if ws.cell(row=linha, column=3).value else '',
                'quantidade': float(ws.cell(row=linha, column=4).value) if ws.cell(row=linha, column=4).value else 0,
                'valorUnitario': float(ws.cell(row=linha, column=5).value) if ws.cell(row=linha, column=5).value else 0,
                'seguro': float(ws.cell(row=linha, column=7).value) if ws.cell(row=linha, column=7).value else 0,
                'frete': float(ws.cell(row=linha, column=8).value) if ws.cell(row=linha, column=8).value else 0,
                'desconto': float(ws.cell(row=linha, column=9).value) if ws.cell(row=linha, column=9).value else 0,
                'outrasDespesas': float(ws.cell(row=linha, column=10).value) if ws.cell(row=linha, column=10).value else 0,
                'numeroDI': str(ws.cell(row=linha, column=11).value) if ws.cell(row=linha, column=11).value else '',
                'dataRegistro': str(ws.cell(row=linha, column=12).value) if ws.cell(row=linha, column=12).value else '',
                'codigoExportador': str(ws.cell(row=linha, column=13).value) if ws.cell(row=linha, column=13).value else '',
                'viaTransporte': int(ws.cell(row=linha, column=14).value) if ws.cell(row=linha, column=14).value else 0,
                'valorAFRMM': float(ws.cell(row=linha, column=15).value) if ws.cell(row=linha, column=15).value else 0,
                'desembaracoUF': str(ws.cell(row=linha, column=16).value) if ws.cell(row=linha, column=16).value else '',
                'desembaracoLocal': str(ws.cell(row=linha, column=17).value) if ws.cell(row=linha, column=17).value else '',
                'desembaracoData': str(ws.cell(row=linha, column=18).value) if ws.cell(row=linha, column=18).value else '',
                'adicao': int(ws.cell(row=linha, column=19).value) if ws.cell(row=linha, column=19).value else 0,
                'itemAdicao': int(ws.cell(row=linha, column=20).value) if ws.cell(row=linha, column=20).value else 0,
                'codigoFabricante': str(ws.cell(row=linha, column=21).value) if ws.cell(row=linha, column=21).value else '',
                'percentualII': float(ws.cell(row=linha, column=22).value) if ws.cell(row=linha, column=22).value else 0,
                'baseII': float(ws.cell(row=linha, column=23).value) if ws.cell(row=linha, column=23).value else 0,
                'despesasAduaneiras': float(ws.cell(row=linha, column=25).value) if ws.cell(row=linha, column=25).value else 0,
                'valorIOF': float(ws.cell(row=linha, column=26).value) if ws.cell(row=linha, column=26).value else 0,
                'cstIPI': int(ws.cell(row=linha, column=27).value) if ws.cell(row=linha, column=27).value else 0,
                'percentualIPI': float(ws.cell(row=linha, column=28).value) if ws.cell(row=linha, column=28).value else 0,
                'baseIPI': float(ws.cell(row=linha, column=29).value) if ws.cell(row=linha, column=29).value else 0,
                'cstPIS': int(ws.cell(row=linha, column=31).value) if ws.cell(row=linha, column=31).value else 0,
                'percentualPIS': float(ws.cell(row=linha, column=32).value) if ws.cell(row=linha, column=32).value else 0,
                'basePIS': float(ws.cell(row=linha, column=33).value) if ws.cell(row=linha, column=33).value else 0,
                'cstCOFINS': int(ws.cell(row=linha, column=35).value) if ws.cell(row=linha, column=35).value else 0,
                'percentualCOFINS': float(ws.cell(row=linha, column=36).value) if ws.cell(row=linha, column=36).value else 0,
                'baseCOFINS': float(ws.cell(row=linha, column=37).value) if ws.cell(row=linha, column=37).value else 0,
                'cstICMS': int(ws.cell(row=linha, column=39).value) if ws.cell(row=linha, column=39).value else 0,
                'percentualICMS': float(ws.cell(row=linha, column=40).value) if ws.cell(row=linha, column=40).value else 0,
                'percentualRedICMS': float(ws.cell(row=linha, column=41).value) if ws.cell(row=linha, column=41).value else 0,
                'baseICMS': float(ws.cell(row=linha, column=42).value) if ws.cell(row=linha, column=42).value else 0,
                'valorICMSST': float(ws.cell(row=linha, column=44).value) if ws.cell(row=linha, column=44).value else 0
            }
            
            dados_importados.append(item)
        
        return jsonify({'dados': dados_importados, 'sucesso': True})
        
    except Exception as e:
        return jsonify({'erro': f'Erro ao importar planilha: {str(e)}'}), 500

@app.route("/exportar", methods=["POST"])
def exportar():
    dados = request.get_json()
    wb = Workbook()
    ws = wb.active
    ws.title = "Produtos"

    # Define os cabeçalhos conforme especificação
    cabecalhos = [
        "item", "codigo", "CFOP", "quantidade", "valor unitário", "valor total",
        "seguro", "frete", "desconto", "outras despesas", "Número DI/DSI/DA",
        "Data registro", "Código do Exportador", "Via de transporte", "Valor AFRMM",
        "Desembaraço (UF)", "Desembaraço (Local)", "Desembaraço (Data)", "adição",
        "item adição", "Código Fabricante", "%II", "Base II", "Valor II",
        "Despesas aduaneiras", "Valor IOF", "CST IPI", "%IPI", "Base IPI",
        "Valor IPI", "CST PIS", "%PIS", "Base PIS", "Valor PIS", "CST COFINS",
        "%COFINS", "Base COFINS", "Valor COFINS", "CST ICMS", "%ICMS", "%Red ICMS",
        "Base ICMS", "Valor ICMS", "Valor ICMS ST"
    ]

    # Linha 1: Cabeçalhos
    for col, cabecalho in enumerate(cabecalhos, 1):
        cell = ws.cell(row=1, column=col, value=cabecalho)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0056A4", end_color="0056A4", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    # Linha 2: Totais (fórmulas de soma para colunas específicas)
    colunas_soma = {
        6: "F",   # valor total
        7: "G",   # seguro
        8: "H",   # frete
        9: "I",   # desconto
        10: "J",  # outras despesas
        15: "O",  # Valor AFRMM
        23: "W",  # Base II
        24: "X",  # Valor II
        25: "Y",  # Despesas aduaneiras
        26: "Z",  # Valor IOF
        29: "AC", # Base IPI
        30: "AD", # Valor IPI
        33: "AG", # Base PIS
        34: "AH", # Valor PIS
        37: "AK", # Base COFINS
        38: "AL", # Valor COFINS
        42: "AP", # Base ICMS
        43: "AQ", # Valor ICMS
        44: "AR"  # Valor ICMS ST
    }

    for col in range(1, len(cabecalhos) + 1):
        if col in colunas_soma:
            letra_col = colunas_soma[col]
            formula = f"=SUM({letra_col}3:{letra_col}16)"
            cell = ws.cell(row=2, column=col, value=formula)
        else:
            cell = ws.cell(row=2, column=col, value="")
        
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="E8F4FD", end_color="E8F4FD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    # A partir da linha 3: Dados dos itens
    if dados:
        for linha_idx, item in enumerate(dados, 3):
            # Mapear os dados do formulário para as colunas corretas
            dados_linha = [
                linha_idx - 2,  # item (numeração sequencial)
                item.get('codigo', ''),
                item.get('cfop', ''),
                float(item.get('quantidade', 0)) if item.get('quantidade') else 0,
                float(item.get('valorUnitario', 0)) if item.get('valorUnitario') else 0,
                f"=D{linha_idx}*E{linha_idx}",  # valor total (fórmula)
                float(item.get('seguro', 0)) if item.get('seguro') else 0,
                float(item.get('frete', 0)) if item.get('frete') else 0,
                float(item.get('desconto', 0)) if item.get('desconto') else 0,
                float(item.get('outrasDespesas', 0)) if item.get('outrasDespesas') else 0,
                item.get('numeroDI', ''),
                item.get('dataRegistro', ''),
                item.get('codigoExportador', ''),
                int(item.get('viaTransporte', 0)) if item.get('viaTransporte') else 0,
                float(item.get('valorAFRMM', 0)) if item.get('valorAFRMM') else 0,
                item.get('desembaracoUF', ''),
                item.get('desembaracoLocal', ''),
                item.get('desembaracoData', ''),
                int(item.get('adicao', 0)) if item.get('adicao') else 0,
                int(item.get('itemAdicao', 0)) if item.get('itemAdicao') else 0,
                item.get('codigoFabricante', ''),
                float(item.get('percentualII', 0)) if item.get('percentualII') else 0,
                float(item.get('baseII', 0)) if item.get('baseII') else 0,  # Base II agora é valor numérico
                f"=ROUND(W{linha_idx}*V{linha_idx}/100,2)",  # Valor II (fórmula corrigida)
                float(item.get('despesasAduaneiras', 0)) if item.get('despesasAduaneiras') else 0,
                float(item.get('valorIOF', 0)) if item.get('valorIOF') else 0,
                int(item.get('cstIPI', 0)) if item.get('cstIPI') else 0,
                float(item.get('percentualIPI', 0)) if item.get('percentualIPI') else 0,
                float(item.get('baseIPI', 0)) if item.get('baseIPI') else 0,
                f"=ROUND(AC{linha_idx}*AB{linha_idx}/100,2)",  # Valor IPI (fórmula)
                int(item.get('cstPIS', 0)) if item.get('cstPIS') else 0,
                float(item.get('percentualPIS', 0)) if item.get('percentualPIS') else 0,
                float(item.get('basePIS', 0)) if item.get('basePIS') else 0,
                f"=ROUND(AG{linha_idx}*AF{linha_idx}/100,2)",  # Valor PIS (fórmula)
                int(item.get('cstCOFINS', 0)) if item.get('cstCOFINS') else 0,
                float(item.get('percentualCOFINS', 0)) if item.get('percentualCOFINS') else 0,
                float(item.get('baseCOFINS', 0)) if item.get('baseCOFINS') else 0,
                f"=ROUND(AK{linha_idx}*AJ{linha_idx}/100,2)",  # Valor COFINS (fórmula)
                int(item.get('cstICMS', 0)) if item.get('cstICMS') else 0,
                float(item.get('percentualICMS', 0)) if item.get('percentualICMS') else 0,
                float(item.get('percentualRedICMS', 0)) if item.get('percentualRedICMS') else 0,
                float(item.get('baseICMS', 0)) if item.get('baseICMS') else 0,
                f"=ROUND(AP{linha_idx}*AN{linha_idx}/100,2)",  # Valor ICMS (fórmula)
                float(item.get('valorICMSST', 0)) if item.get('valorICMSST') else 0
            ]

            for col, valor in enumerate(dados_linha, 1):
                cell = ws.cell(row=linha_idx, column=col, value=valor)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # Formatação especial para números
                if isinstance(valor, (int, float)) and valor != 0:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif isinstance(valor, str) and valor.startswith('='):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # Ajustar largura das colunas
    for col in range(1, len(cabecalhos) + 1):
        column_letter = get_column_letter(col)
        ws.column_dimensions[column_letter].width = 15

    # Salva em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Retorna o arquivo para download
    filename = f"planilha_sisloc_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xls"
    return send_file(
        output, 
        as_attachment=True, 
        download_name=filename, 
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)

