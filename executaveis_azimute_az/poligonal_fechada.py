import ezdxf
import math
import csv
import os
import re
import glob
import locale
from docx import Document
from docx.shared import Inches
from datetime import datetime
from decimal import Decimal, getcontext
import pandas as pd
import locale
import openpyxl
from openpyxl.styles import Alignment, Font
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


getcontext().prec = 28  # Define a precisão para 28 casas decimais

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Para Render (Linux)
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')  # Para Windows
    except locale.Error:
        locale.setlocale(locale.LC_TIME, 'C')  # Fallback universal

# Obter data atual formatada
data_atual = datetime.now().strftime("%d de %B de %Y")

def limpar_dxf(original_path, saida_path):
    try:
        doc_antigo = ezdxf.readfile(original_path)
        msp_antigo = doc_antigo.modelspace()
        doc_novo = ezdxf.new(dxfversion='R2010')
        msp_novo = doc_novo.modelspace()

        pontos_polilinha = None

        # Copiar polilinha fechada
        for entity in msp_antigo.query('LWPOLYLINE'):
            if entity.closed:
                pontos_polilinha = [point[:2] for point in entity.get_points('xy')]
                
                # Remover pontos duplicados consecutivos
                pontos_unicos = []
                tolerancia = 1e-6
                for pt in pontos_polilinha:
                    if not pontos_unicos or math.hypot(pt[0] - pontos_unicos[-1][0], pt[1] - pontos_unicos[-1][1]) > tolerancia:
                        pontos_unicos.append(pt)

                if math.hypot(pontos_unicos[0][0] - pontos_unicos[-1][0], pontos_unicos[0][1] - pontos_unicos[-1][1]) < tolerancia:
                    pontos_unicos.pop()

                # Inserir polilinha limpa no DXF
                msp_novo.add_lwpolyline(
                    pontos_unicos,
                    close=True,
                    dxfattribs={'layer': 'DIVISA_PROJETADA'}
                )
                break

        # Copiar Ponto Az do arquivo original (TEXT, INSERT ou POINT)
        ponto_az_copiado = False

        # Copiar TEXT
        for entity in msp_antigo.query('TEXT'):
            if "Az" in entity.dxf.text:
                msp_novo.add_text(
                    entity.dxf.text,
                    dxfattribs={
                        'insert': (entity.dxf.insert.x, entity.dxf.insert.y),
                        'height': entity.dxf.height,
                        'rotation': entity.dxf.rotation,
                        'layer': entity.dxf.layer
                    }
                )
                ponto_az_copiado = True

        # Copiar INSERT (blocos com nome contendo Az)
        for entity in msp_antigo.query('INSERT'):
            if "Az" in entity.dxf.name:
                msp_novo.add_blockref(
                    entity.dxf.name,
                    insert=(entity.dxf.insert.x, entity.dxf.insert.y),
                    dxfattribs={'layer': entity.dxf.layer}
                )
                ponto_az_copiado = True

        # Copiar POINT (ponto simples chamado Az)
        for entity in msp_antigo.query('POINT'):
            msp_novo.add_point(
                (entity.dxf.location.x, entity.dxf.location.y),
                dxfattribs={'layer': entity.dxf.layer}
            )
            ponto_az_copiado = True

        if not ponto_az_copiado:
            print("⚠️ Atenção: Ponto Az não foi encontrado para copiar!")

        doc_novo.saveas(saida_path)
        print(f"✅ DXF limpo salvo em: {saida_path}")
        return saida_path

    except Exception as e:
        print(f"❌ Erro ao limpar DXF: {e}")
        return original_path



def get_document_info_from_dxf(dxf_file_path):
    try:
        doc = ezdxf.readfile(dxf_file_path)  
        msp = doc.modelspace()  

        lines = []
        perimeter_dxf = 0
        area_dxf = 0
        ponto_az = None
        area_poligonal = None

        for entity in msp.query('LWPOLYLINE'):
            if entity.closed:
                points = entity.get_points('xy')
                
                # Verifica e remove vértice repetido no final, se houver
                if points[0] == points[-1]:
                    points.pop()
                
                num_points = len(points)

                for i in range(num_points):
                    start_point = (points[i][0], points[i][1])
                    end_point = (points[(i + 1) % num_points][0], points[(i + 1) % num_points][1])
                    lines.append((start_point, end_point))

                    segment_length = ((end_point[0] - start_point[0]) ** 2 + 
                                      (end_point[1] - start_point[1]) ** 2) ** 0.5
                    perimeter_dxf += segment_length

                x = [point[0] for point in points]
                y = [point[1] for point in points]
                area_dxf = abs(sum(x[i] * y[(i + 1) % num_points] - x[(i + 1) % num_points] * y[i] for i in range(num_points)) / 2)

                break  

        if not lines:
            print("Nenhuma polilinha encontrada no arquivo DXF.")
            return None, [], 0, 0, None, None

        for entity in msp.query('TEXT'):
            if "Az" in entity.dxf.text:
                ponto_az = (entity.dxf.insert.x, entity.dxf.insert.y, 0)
                print(f"Ponto Az encontrado em texto: {ponto_az}")

        for entity in msp.query('INSERT'):
            if "Az" in entity.dxf.name:
                ponto_az = (entity.dxf.insert.x, entity.dxf.insert.y, 0)
                print(f"Ponto Az encontrado no bloco: {ponto_az}")

        for entity in msp.query('POINT'):
            ponto_az = (entity.dxf.location.x, entity.dxf.location.y, 0)
            print(f"Ponto Az encontrado como ponto: {ponto_az}")

        if not ponto_az:
            print("Ponto Az não encontrado no arquivo DXF.")
            return None, lines, 0, 0, None, None

        print(f"Linhas processadas: {len(lines)}")
        print(f"Perímetro do DXF: {perimeter_dxf:.2f} metros")
        print(f"Área do DXF: {area_dxf:.2f} metros quadrados")

        return doc, lines, perimeter_dxf, area_dxf, ponto_az, area_poligonal

    except Exception as e:
        print(f"Erro ao obter informações do documento: {e}")
        return None, [], 0, 0, None, None


# 🔹 Função para definir a fonte padrão
def set_default_font(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)

    
def calculate_point_on_line(start, end, distance):
    """
    Calcula um ponto a uma determinada distância sobre uma linha entre dois pontos.
    :param start: Coordenadas do ponto inicial (x, y).
    :param end: Coordenadas do ponto final (x, y).
    :param distance: Distância do ponto inicial ao longo da linha.
    :return: Coordenadas do ponto calculado (x, y).
    """
    dx, dy = end[0] - start[0], end[1] - start[1]
    length = math.hypot(dx, dy)  # Calcula o comprimento da linha
    if length == 0:
        raise ValueError("Ponto inicial e final são iguais, não é possível calcular um ponto na linha.")
    return (
        start[0] + (dx / length) * distance,
        start[1] + (dy / length) * distance
    )


def calculate_azimuth(p1, p2):
    """
    Calcula o azimute entre dois pontos em graus.
    Azimute é o ângulo medido no sentido horário a partir do Norte.
    """
    delta_x = p2[0] - p1[0]  # Diferença em X (Leste/Oeste)
    delta_y = p2[1] - p1[1]  # Diferença em Y (Norte/Sul)

    # Calcular o ângulo em radianos
    azimuth_rad = math.atan2(delta_x, delta_y)

    # Converter para graus
    azimuth_deg = math.degrees(azimuth_rad)

    # Garantir que o azimute esteja no intervalo [0, 360)
    if azimuth_deg < 0:
        azimuth_deg += 360

    return azimuth_deg

def calculate_distance(point1, point2):
    """
    Calcula a distância entre dois pontos em um plano 2D.
    :param point1: Tupla (x1, y1) representando o primeiro ponto.
    :param point2: Tupla (x2, y2) representando o segundo ponto.
    :return: Distância entre os pontos.
    """
    dx = point2[0] - point1[0]
    dy = point2[1] - point1[1]
    return math.sqrt(dx**2 + dy**2)




# Função para calcular azimute e distância
def calculate_azimuth_and_distance(start_point, end_point):
    dx = end_point[0] - start_point[0]
    dy = end_point[1] - start_point[1]
    distance = math.hypot(dx, dy)
    azimuth = math.degrees(math.atan2(dx, dy))
    if azimuth < 0:
        azimuth += 360
    return azimuth, distance


def add_azimuth_arc(doc, msp, ponto_az, v1, azimuth):
    """
    Adiciona o arco do azimute no ModelSpace.
    """
    try:
        if 'LAYOUT_AZIMUTES' not in doc.layers:
            doc.layers.new(name='LAYOUT_AZIMUTES', dxfattribs={'color': 5})

        # Traçar segmento entre Az e V1
        msp.add_line(start=ponto_az, end=v1, dxfattribs={'layer': 'LAYOUT_AZIMUTES'})

        # Adicionar rótulo do azimute
        azimuth_label = f"Azimute = {convert_to_dms(azimuth)}"
        label_position = (
            ponto_az[0] + 1.5 * math.cos(math.radians(azimuth / 2)),
            ponto_az[1] + 1.5 * math.sin(math.radians(azimuth / 2))
        )
        msp.add_text(
            azimuth_label,
            dxfattribs={'height': 0.5, 'layer': 'LAYOUT_AZIMUTES', 'insert': label_position}
        )

        print(f"Rótulo do azimute ({azimuth_label}) adicionado com sucesso em {label_position}")

    except Exception as e:
        print(f"Erro ao adicionar arco do azimute: {e}")


# Função para converter graus decimais para DMS
def convert_to_dms(decimal_degrees):
    degrees = int(decimal_degrees)
    minutes = int(abs(decimal_degrees - degrees) * 60)
    seconds = abs((decimal_degrees - degrees - minutes / 60) * 3600)
    return f"{degrees}° {minutes}' {seconds:.2f}\""

# Função para calcular a área de uma poligonal
def calculate_polygon_area(points):
    n = len(points)
    area = 0.0
    for i in range(n):
        x1, y1 = points[i][0], points[i][1]
        x2, y2 = points[(i + 1) % n][0], points[(i + 1) % n][1]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def add_label_and_distance(doc, msp, start_point, end_point, label, distance):
    """
    Adiciona um rótulo no vértice e a distância corretamente alinhada à linha no arquivo DXF.
    
    :param doc: Objeto Drawing do ezdxf.
    :param msp: ModelSpace do ezdxf.
    :param start_point: Coordenadas do ponto inicial (x, y).
    :param end_point: Coordenadas do ponto final (x, y).
    :param label: Nome do vértice (ex: V1, V2).
    :param distance: Distância entre os pontos (em metros).
    """
    try:
        msp = doc.modelspace()

        # Criar camadas necessárias (sem alterar as que não precisam)
        for layer_name, color in [
            ("LAYOUT_VERTICES", 2),  # Vermelho para vértices
            ("LAYOUT_DISTANCIAS", 4),  # Azul para distâncias
            ("LAYOUT_AZIMUTES", 5)  # Magenta para azimutes
        ]:
            if layer_name not in doc.layers:
                doc.layers.new(name=layer_name, dxfattribs={"color": color})

        # 🔹 Adicionar círculo no ponto inicial (Vértices)
        msp.add_circle(center=start_point[:2], radius=1.0, dxfattribs={'layer': 'LAYOUT_VERTICES'})

        # 🔹 Adicionar rótulo do vértice
        text_point = (start_point[0] + 1, start_point[1])  # Posição deslocada
        msp.add_text(
            label,
            dxfattribs={'height': 0.5, 'layer': 'LAYOUT_VERTICES', 'insert': text_point}
        )

        # 🔹 Calcular o ponto médio da linha
        mid_x = (start_point[0] + end_point[0]) / 2
        mid_y = (start_point[1] + end_point[1]) / 2

        # 🔹 Vetor da linha
        dx = end_point[0] - start_point[0]
        dy = end_point[1] - start_point[1]
        length = math.hypot(dx, dy)

        # Evitar erro de divisão por zero
        if length == 0:
            return

        # 🔹 Ângulo da linha
        angle = math.degrees(math.atan2(dy, dx))

        # 🔹 Ajuste de ângulo para manter leitura correta
        if angle < -90 or angle > 90:
            angle += 180  

        # 🔹 Afastar o rótulo da linha
        offset = 0.3  # Ajuste para evitar sobreposição
        perp_x = -dy / length * offset
        perp_y = dx / length * offset
        displaced_mid_point = (mid_x + perp_x, mid_y + perp_y)

        # 🔹 Formatar a distância corretamente
        distancia_formatada = f"{distance:.2f}".replace(".", ",")

        # 🔹 Adicionar rótulo da distância corretamente alinhado
        msp.add_text(
            f"{distancia_formatada} ",
            dxfattribs={
                "height": 1.0,  # Aumenta a altura do texto
                "layer": "LAYOUT_DISTANCIAS",
                "rotation": angle,  # Alinhamento correto à linha
                "insert": displaced_mid_point
            }
        )

        print(f"✅ Distância {distancia_formatada} m adicionada corretamente em {displaced_mid_point}")

    except Exception as e:
        print(f"❌ Erro ao adicionar rótulo de distância: {e}")




#     return confrontantes
def sanitize_filename(filename):
    # Substitui os caracteres inválidos por um caractere válido (ex: espaço ou underline)
    sanitized_filename = re.sub(r'[\\/*?:"<>|]', "_", filename)  # Substitui caracteres inválidos por "_"
    return sanitized_filename
        
        


# Função para criar memorial descritivo
def create_memorial_descritivo(doc,msp, lines, proprietario, matricula, caminho_salvar, excel_file_path, ponto_az,distance_az_v1, azimute_az_v1, tipo, encoding='ISO-8859-1'):
    """
    Cria o memorial descritivo diretamente no arquivo DXF e salva os dados em uma planilha Excel.
    """
    # Carregar a planilha de confrontantes
    confrontantes_df = pd.read_excel(excel_file_path)

    # Número de registros no arquivo
    num_registros = len(confrontantes_df)

    # Transformar o dataframe em um dicionário de código -> confrontante
    confrontantes_dict = dict(zip(confrontantes_df['Código'], confrontantes_df['Confrontante']))

    # Verificar se a planilha foi carregada corretamente
    if not confrontantes_dict:
        print("Erro ao carregar confrontantes.")
        return None

    if not lines:
        print("Nenhuma linha disponível para criar o memorial descritivo.")
        return None



    # Coletar os pontos diretamente na ordem dos vértices no DXF
    ordered_points = [line[0] for line in lines] + [lines[-1][1]]  # Fechando a poligonal
    num_vertices = len(ordered_points)
    
    # Calcular a área da poligonal
    area = calculate_polygon_area(ordered_points)

    # Ajustar para o sentido horário se necessário
    if area < 0:  # Sentido horário ou antihorário
        ordered_points.reverse()  # Reorganizar para sentido horário
        print(f"Pontos reorganizados para sentido horário: {ordered_points}")
       
    # Preparar os dados para o Excel
    data = []
    for i in range(len(ordered_points) - 1):
        start_point = ordered_points[i]
        end_point = ordered_points[i + 1]

        # Calcular azimute e distância
        azimuth, distance = calculate_azimuth_and_distance(start_point, end_point)
        azimuth_dms = convert_to_dms(azimuth)

        # Buscar o confrontante
        confrontante = confrontantes_df.iloc[i]['Confrontante'] if i < len(confrontantes_df) else "Desconhecido"

        # Adicionar as coordenadas do ponto Az apenas na primeira linha
        coord_e_ponto_az = f"{ponto_az[0]:.3f}".replace('.', ',') if i == 0 else ""
        coord_n_ponto_az = f"{ponto_az[1]:.3f}".replace('.', ',') if i == 0 else ""


        # Adicionar linha ao conjunto de dados
        data.append({
            "V": f"V{i + 1}",
            "E": f"{start_point[0]:.3f}".replace('.', ','),
            "N": f"{start_point[1]:.3f}".replace('.', ','),
            "Z": "0.000",
            "Divisa": f"V{i + 1}_V{1 if (i + 1) == len(ordered_points) - 1 else i + 2}", 
            "Azimute": azimuth_dms,
            "Distancia(m)": f"{distance:.2f}".replace('.', ','),
            "Confrontante": confrontante,
            "Coord_E_ponto_Az": coord_e_ponto_az,
            "Coord_N_ponto_Az": coord_n_ponto_az,
            "distancia_Az_V1": f"{distance_az_v1:.2f}".replace('.', ',') if i == 0 else "",  # Adicionar apenas na primeira linha
            "Azimute Az_V1":convert_to_dms(azimute_az_v1)  if i == 0 else ""  # Adicionar apenas na primeira linha

        })

        # Adicionar rótulos e distância ao DXF
        add_label_and_distance(doc, msp, start_point, end_point, f"V{i + 1}", distance)


    # Criar DataFrame e salvar em Excel
    df = pd.DataFrame(data, dtype=str)
    excel_output_path = os.path.join(caminho_salvar, f"{tipo}_Memorial_{matricula}.xlsx")

    df.to_excel(excel_output_path, index=False)

    # Formatar o Excel
    wb = openpyxl.load_workbook(excel_output_path)
    ws = wb.active

    # Formatar cabeçalho
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Ajustar largura das colunas
    column_widths = {
        "A": 8,   # V
        "B": 15,  # E
        "C": 15,  # N
        "D": 10,  # Z
        "E": 20,  # Divisa
        "F": 15,  # Azimute
        "G": 15,  # Distancia(m)
        "H": 30,  # Confrontante
        "I": 20,  # Coord_E_ponto_Az
        "J": 20,   # Coord_N_ponto_Az
        "K": 15,  # Coluna distancia_Az_V1  (Nova coluna)
        "L": 15,  # Coluna Azimute Az_V1   (Nova coluna)

    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Centralizar o conteúdo das células
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Salvar o arquivo formatado
    wb.save(excel_output_path)
    print(f"Arquivo Excel salvo e formatado em: {excel_output_path}")

    # Adicionar o arco de azimute e o segmento ao desenho
    try:
        msp = doc.modelspace()  # Obtenha o ModelSpace do documento
        v1 = ordered_points[0]  # Primeiro vértice
        azimuth = calculate_azimuth(ponto_az, v1)
        add_azimuth_arc(doc,msp, ponto_az, v1, azimuth)  # Use msp diretamente
        print("Ângulo de Azimute adicionado ao arquivo DXF com sucesso.")
    except Exception as e:
        print(f"Erro ao adicionar Azimute ao arquivo DXF: {e}")

    # Calcular a distância entre ponto Az e V1 e adicionar ao DXF
    try:
        distance_az_v1 = calculate_distance(ponto_az, v1)
        add_label_and_distance(doc,msp, ponto_az, v1, "", distance_az_v1)
        print(f"Distância entre ponto Az e V1: {distance_az_v1:.2f} m adicionada ao DXF.")
    except Exception as e:
        print(f"Erro ao adicionar a distância entre ponto Az e V1 ao DXF: {e}")

    # Salvar o arquivo DXF com as alterações
    # Salvar o arquivo DXF com as alterações
    try:
        dxf_output_path = os.path.join(caminho_salvar, f"{tipo}_Memorial_{matricula}.dxf")
        doc.saveas(dxf_output_path)
        print(f"Arquivo DXF atualizado salvo em: {dxf_output_path}")

    except Exception as e:
        print(f"Erro ao salvar o arquivo DXF final: {e}")

    return excel_output_path




def create_memorial_document(
    proprietario, matricula, descricao, excel_file_path, template_path, output_path,
    perimeter_dxf, area_dxf, desc_ponto_Az, Coorde_E_ponto_Az, Coorde_N_ponto_Az,
    azimuth, distance, uso_solo, area_imovel, cidade, rua, comarca, RI, caminho_salvar,tipo
):



    try:
        # Ler o arquivo Excel
        df = pd.read_excel(excel_file_path)
        df['N'] = pd.to_numeric(df['N'].astype(str).str.replace(',', '.'), errors='coerce')
        df['E'] = pd.to_numeric(df['E'].astype(str).str.replace(',', '.'), errors='coerce')
        df['Distancia(m)'] = pd.to_numeric(df['Distancia(m)'].astype(str).str.replace(',', '.'), errors='coerce')
        # Criar o documento Word
        doc_word = Document(template_path)
        set_default_font(doc_word)  # 🔹 Aplica a fonte Arial 12 ao documento
        # Adiciona o preâmbulo centralizado com texto em negrito
        p1 = doc_word.add_paragraph(style='Normal')
        run = p1.add_run("MEMORIAL DESCRITIVO")
        run.bold = True
        p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        # Método recomendado para garantir espaço visível:
        p_espaco = doc_word.add_paragraph()
        p_espaco.add_run('\n')  # Garantindo espaçamento vertical.


        #doc_word.add_paragraph(f"Imóvel: Área da matrícula {matricula} destinada a {descricao} ", style='Normal')
        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Imóvel: ")

        if tipo == "ETE":
            texto = f"Área da matrícula {matricula} destinada a {descricao} - SES de {cidade}"
        elif tipo == "REM":
            texto = f"Área remanescente da matrícula {matricula} destinada a {descricao}"
        elif tipo == "SER":
            texto = f"Área da matrícula {matricula} destinada a SERVIDÃO ADMINISTRATIVA ACESSO À {descricao}"
        elif tipo == "ACE":
            texto = f"Área da matrícula {matricula} destinada ao ACESSO DA SERVIDÃO ADMINISTRATIVA DA {descricao}"
        else:
            texto = "Tipo não especificado"

        run_bold = p.add_run(texto)
        run_bold.bold = True

        doc_word.add_paragraph(f"Matrícula: Número - {matricula} do {RI} de {comarca} ", style='Normal')
        doc_word.add_paragraph(f"Proprietário: {proprietario}", style='Normal')
        doc_word.add_paragraph(f"Local: {rua} - {cidade}", style='Normal')
        
        doc_word.add_paragraph(f"Área: {area_dxf:,.2f} m²".replace(",", "X").replace(".", ",").replace("X", "."), style='Normal')
        doc_word.add_paragraph(f"Perímetro: {perimeter_dxf:,.2f} m".replace(",", "X").replace(".", ",").replace("X", "."), style='Normal')
        # Pula uma linha antes deste parágrafo
        doc_word.add_paragraph()
        
        
        # Primeiro, formate corretamente a área uma única vez antes do parágrafo:
        area_dxf_formatada = f"{area_dxf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        # Cria o parágrafo com variáveis formatadas (em uma única linha no Python para clareza)
        texto_paragrafo = (f"Área {uso_solo} com {area_dxf_formatada} m², parte de um todo maior da Matrícula N° {matricula} com {area_imovel} do {RI} de {comarca}, localizada na {rua}, na cidade de {cidade}, definida através do seguinte levantamento topográfico, onde os ângulos foram medidos no sentido horário.")
        
        # Cria o parágrafo e remove qualquer indentação especial (recuo pendente ou primeira linha)
        p = doc_word.add_paragraph(texto_paragrafo, style='Normal')
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY  # alinhamento justificado, se desejar
        doc_word.add_paragraph()  # Linha em branco após o parágrafo
        
        # Remove indentação/recuos
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(0)
        p.paragraph_format.right_indent = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_together = True
        
        # Removendo explicitamente recuo pendente (esse é o ajuste essencial!)
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.hanging_indent = None

        # Formata coordenadas individualmente (sem milhar)
        coord_E_ponto_Az = f"{Coorde_E_ponto_Az:.3f}".replace(".", ",")
        coord_N_ponto_Az = f"{Coorde_N_ponto_Az:.3f}".replace(".", ",")
        
        # Primeiro parágrafo ajustado corretamente
        p = doc_word.add_paragraph(
            f"O Ponto Az, ponto de amarração, está localizado na {desc_ponto_Az} nas coordenadas "
            f"E(X) {coord_E_ponto_Az} e N(Y) {coord_N_ponto_Az}.",
            style='Normal'
        )
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        doc_word.add_paragraph()  # Linha em branco após o primeiro parágrafo
        
        # Formatação correta da distância (com ponto do milhar)
        distance_formatada = f"{distance:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # Segundo parágrafo ajustado corretamente
        # Cria o parágrafo vazio inicialmente
        p = doc_word.add_paragraph(style='Normal')
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        # Primeira parte sem negrito até "Vértice"
        p.add_run(f"Daí, com Azimute de {convert_to_dms(azimuth)} e distância de {distance_formatada} m, chega-se ao Vértice ")
        
        # Insere o vértice V1 em negrito
        run_v1 = p.add_run("V1")
        run_v1.bold = True  # V1 em negrito
        
        # Restante do texto normal
        p.add_run(", origem da descrição desta área.")

        doc_word.add_paragraph()  # Linha em branco após o primeiro parágrafo
        
        # Início da descrição do perímetro
        initial = df.iloc[0]
        coord_N_inicial = f"{initial['N']:.3f}".replace(".", ",")
        coord_E_inicial = f"{initial['E']:.3f}".replace(".", ",")
        
        # Primeiro parágrafo
        p1 = doc_word.add_paragraph("Pontos definidos pelas Coordenadas Planas no Sistema U.T.M. – SIRGAS 2000.", style='Normal')
        p1.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        # Pula uma linha
        doc_word.add_paragraph()
        
        # Segundo parágrafo
        # Cria o parágrafo vazio inicialmente
        p2 = doc_word.add_paragraph(style='Normal')
        p2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
        # Texto inicial sem negrito até "vértice"
        p2.add_run("Inicia-se a descrição deste perímetro no vértice ")
        
        # Insere o vértice inicial em negrito
        run_v_inicial = p2.add_run(f"{initial['V']}")
        run_v_inicial.bold = True  # Define negrito
        
        # Restante do texto sem negrito
        p2.add_run(
            f", de coordenadas N(Y) {coord_N_inicial} e E(X) {coord_E_inicial}, "
            f"situado no limite com {initial['Confrontante']}."
        )


        doc_word.add_paragraph()  # Linha em branco após o parágrafo

        # Descrição dos segmentos
        num_points = len(df)
        for i in range(num_points):
            current = df.iloc[i]
            next_index = (i + 1) % num_points
            next_point = df.iloc[next_index]
        
            distancia_formatada = f"{current['Distancia(m)']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            coord_N_formatada = f"{next_point['N']:.3f}".replace(".", ",")
            coord_E_formatada = f"{next_point['E']:.3f}".replace(".", ",")
        
            # Checa se o próximo vértice é V1, para inserir texto especial
            if next_point['V'] == 'V1':
                complemento = ", origem desta descrição,"
            else:
                complemento = ""
        
            # Criação do parágrafo inicialmente vazio
            p = doc_word.add_paragraph(style='Normal')
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        
            # Primeira parte do texto sem negrito
            p.add_run(
                f"Deste, segue com azimute de {current['Azimute']} e distância de {distancia_formatada} m, "
                f"confrontando neste trecho com área pertencente à {current['Confrontante']}, até o vértice "
            )
        
            # Adiciona o rótulo do vértice em negrito
            run_vertice = p.add_run(f"{next_point['V']}")
            run_vertice.bold = True  # Aqui é definido o negrito para o vértice
        
            # Completa o restante do texto sem negrito
            p.add_run(
                f"{complemento} de coordenadas N(Y) {coord_N_formatada} e E(X) {coord_E_formatada};"
            )
        
            # Adiciona uma linha em branco após cada parágrafo
            doc_word.add_paragraph()

        # Adicionar data e assinatura
        data_atual = datetime.now().strftime("%d de %B de %Y")
        doc_word.add_paragraph(f"\n Porto Alegre, RS, {data_atual}.", style='Normal')
        doc_word.add_paragraph("\n\n")
        output_path = os.path.normpath(os.path.join(caminho_salvar, f"{tipo}_Memorial_MAT_{matricula}.docx"))

        doc_word.save(output_path)
        print(f"Memorial descritivo salvo em: {output_path}")
    except Exception as e:
        print(f"Erro ao criar o documento memorial: {e}")

        
# Função principal
def main_poligonal_fechada(arquivo_excel_recebido, arquivo_dxf_recebido, diretorio_preparado, diretorio_concluido, caminho_template):
    # Carrega arquivo Excel com os dados do imóvel
    #dados_imovel_excel_path = input("Digite o caminho completo do arquivo Excel com Dados do Imóvel: ").strip('"')
    dados_imovel_excel_path = arquivo_excel_recebido
    # Ler especificamente a aba "Dados_do_Imóvel", sem cabeçalho
    dados_imovel_df = pd.read_excel(dados_imovel_excel_path, sheet_name='Dados_do_Imóvel', header=None)
    
    # Converter diretamente colunas em dicionário para extração direta dos dados
    dados_imovel = dict(zip(dados_imovel_df.iloc[:, 0], dados_imovel_df.iloc[:, 1]))
    
    # Carregar variáveis conforme correspondência solicitada
    proprietario = dados_imovel.get("NOME DO PROPRIETÁRIO", "").strip()
    matricula = dados_imovel.get("DOCUMENTAÇÃO DO IMÓVEL", "").strip()
    matricula = sanitize_filename(matricula)  # preserva lógica original
    descricao = dados_imovel.get("OBRA", "").strip()
    uso_solo = dados_imovel.get("ZONA", "").strip()
    area_imovel = dados_imovel.get("ÁREA TOTAL DO TERRENO DOCUMENTADA", "").replace("\t", "").replace("\n", "").strip()
    cidade = dados_imovel.get("CIDADE", "").strip()
    rua = dados_imovel.get("LOCAL", "").strip()
    comarca = dados_imovel.get("COMARCA", "").strip()
    RI = dados_imovel.get("RI", "").strip()
    desc_ponto_Az = dados_imovel.get("AZ", "").strip()

    #caminho_salvar = input("Digite o caminho de salvamento: ").strip('"') # remove aspas
    caminho_salvar = diretorio_concluido
    os.makedirs(caminho_salvar, exist_ok=True)

    # Pedir o caminho do arquivo DXF
    #dxf_file_path = input("Digite o caminho completo do arquivo DXF: ").strip('"')  # Remove aspas ao redor do caminho
    # Pedir o caminho do arquivo DXF original
    #original_dxf = input("Digite o caminho completo do arquivo DXF: ").strip('"')
    original_dxf = arquivo_dxf_recebido
    # Define o caminho do arquivo limpo usando a pasta indicada pelo usuário
    limpo_dxf = os.path.join(caminho_salvar, "arquivo_limpo.dxf")
    # Pedir o caminho do arquivo EXCEL com os codigos dos vertices e confrontantes
    #exc_file_path = input("Digite o caminho completo do arquivo Excel(codigos e confrontantes): ").strip('"')  # Remove aspas ao redor do caminho
    exc_file_path = diretorio_preparado
    dxf_file_path = limpar_dxf(original_dxf, limpo_dxf)
    # Extração automática do tipo (ETE, REM, SER, ACE) a partir do nome DXF
    dxf_filename = os.path.basename(original_dxf).upper()

    if "ETE" in dxf_filename:
        tipo = "ETE"
    elif "REM" in dxf_filename:
        tipo = "REM"
    elif "SER" in dxf_filename:
        tipo = "SER"
    elif "ACE" in dxf_filename:
        tipo = "ACE"
    else:
        print("❌ Não foi possível determinar automaticamente o tipo (ETE, REM, SER ou ACE).")
        return

    diretorio_confrontantes = diretorio_preparado  # definir corretamente antes
    padrao_busca = os.path.join(diretorio_confrontantes, f"FECHADA_*_{tipo}.xlsx")
    arquivos_encontrados = glob.glob(padrao_busca)

    if not arquivos_encontrados:
        print(f"❌ Arquivo de confrontantes não encontrado com o padrão: {padrao_busca}")
        return

    exc_file_path = arquivos_encontrados[0]

    doc,lines, perimeter_dxf, area_dxf, ponto_az, area_poligonal = get_document_info_from_dxf(dxf_file_path)

    try:
        doc_dxf = ezdxf.readfile(dxf_file_path)
        msp = doc_dxf.modelspace()  # Acessar o espaço de modelo
    except Exception as e:
        print(f"Erro ao abrir o arquivo DXF para edição: {e}")
        return None
    
    if not doc or not ponto_az:
        print("Erro ao processar o arquivo DXF.")
        return
    if ponto_az is None:
        print("Erro: O ponto Az não foi encontrado no arquivo DXF.")
        return
    else:
        print(f"Ponto Az identificado: {ponto_az}")

    # Desenhar a linha entre ponto Az e V1
    v1 = lines[0][0]  # Primeiro vértice da poligonal
    distance_az_v1 = calculate_distance(ponto_az, v1)
    azimute_az_v1 = calculate_azimuth(ponto_az, v1)
    distance = math.hypot(v1[0] - ponto_az[0], v1[1] - ponto_az[1])
    
    # Calcular o azimute entre Az e V1 e adicionar arco do azimute
    azimuth = calculate_azimuth(ponto_az, v1)
    print(f"Azimute do ponto Az para V1: {azimuth:.2f}°")

    # (Opcional) Adicionar o arco do azimute (se necessário)
    add_azimuth_arc(doc,msp, ponto_az, v1, azimuth)
    
    if doc and lines:

        print(f"Nome do documento: {doc}")
        print(f"Número de linhas: {len(lines)}")
        
        # Criar o memorial descritivo diretamente (coleta de confrontantes interna)
        excel_output_path = create_memorial_descritivo(
            doc=doc,
            msp=msp,
            lines=lines,
            proprietario=proprietario,
            matricula=matricula,
            caminho_salvar=caminho_salvar,
            excel_file_path=exc_file_path,
            ponto_az=ponto_az,distance_az_v1=distance_az_v1,
            azimute_az_v1=azimute_az_v1,
            tipo=tipo
            )

        if excel_output_path:
            # Caminhos para o template e saída do documento
            template_path = caminho_template
            output_path = os.path.normpath(os.path.join(caminho_salvar, f"{tipo}_Memorial_MAT_{matricula}.docx"))
            #pdf_file_path = os.path.normpath(os.path.join(caminho_salvar, f"{tipo}_Memorial_MAT_{matricula}.pdf"))


            create_memorial_document(
                proprietario=proprietario,
                matricula=matricula,
                descricao=descricao,
                excel_file_path=excel_output_path,
                template_path=template_path,
                output_path=output_path,
                perimeter_dxf=perimeter_dxf,
                area_dxf=area_dxf,
                desc_ponto_Az=desc_ponto_Az,
                Coorde_E_ponto_Az=ponto_az[0],
                Coorde_N_ponto_Az=ponto_az[1],
                azimuth=azimuth,
                distance=distance,
                uso_solo=uso_solo,
                area_imovel=area_imovel,
                cidade=cidade,
                rua=rua,
                comarca=comarca,
                RI=RI,
                caminho_salvar=caminho_salvar,
                tipo=tipo# <-- adicionado aqui
            )

                                
           
            
            # Fechar o documento do AutoCAD (se necessário)
            
            print("Processamento concluído com sucesso.")


    else:
        print("Erro ao processar o arquivo DXF.")




