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
import logging


# Diretório para logs
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
LOG_DIR = os.path.join(BASE_DIR, 'static', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)

# Arquivo de log específico para poligonal_fechada
log_file = os.path.join(LOG_DIR, f'poligonal_fechada_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

# Configuração básica do logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))

# Verificar se já não existem handlers para não duplicar
if not logger.handlers:
    logger.addHandler(file_handler)

getcontext().prec = 28  # Define a precisão para 28 casas decimais

MESES_PT_BR = {
    'January': 'janeiro',
    'February': 'fevereiro',
    'March': 'março',
    'April': 'abril',
    'May': 'maio',
    'June': 'junho',
    'July': 'julho',
    'August': 'agosto',
    'September': 'setembro',
    'October': 'outubro',
    'November': 'novembro',
    'December': 'dezembro'
}
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
def add_azimuth_arc_to_dxf(msp, ponto_az, v1, azimute):
    """
    Adiciona o arco do azimute ao DXF usando ezdxf.
    """
    try:
        logger.info(f"Iniciando a adição do arco de azimute. Azimute: {azimute}°")

        # Criar camada 'Azimute', se não existir
        if 'Azimute' not in msp.doc.layers:
            msp.doc.layers.new(name='Azimute', dxfattribs={'color': 1})
            logger.info("Camada 'Azimute' criada com sucesso.")

        # Traçar segmento entre Az e V1
        msp.add_line(start=ponto_az, end=v1, dxfattribs={'layer': 'Azimute'})
        logger.info(f"Segmento entre Az e V1 desenhado de {ponto_az} para {v1}")

        # Traçar segmento para o norte
        north_point = (ponto_az[0], ponto_az[1] + 2)
        msp.add_line(start=ponto_az, end=north_point, dxfattribs={'layer': 'Azimute'})
        logger.info(f"Linha para o norte desenhada com sucesso de {ponto_az} para {north_point}")

        # Calcular o ponto inicial (1 metro de Az para V1)
        # Calcular distância entre ponto Az e V1 para definir raio adaptativo
        dist = calculate_distance(ponto_az, v1)
        radius = 0.4 if dist <= 0.5 else 1.0

        # Calcular os pontos do arco com esse raio
        start_arc = calculate_point_on_line(ponto_az, v1, radius)
        end_arc = calculate_point_on_line(ponto_az, north_point, radius)

        # Traçar o arco do azimute
        msp.add_arc(
            center=ponto_az,
            radius=radius,
            start_angle=math.degrees(math.atan2(start_arc[1] - ponto_az[1], start_arc[0] - ponto_az[0])),
            end_angle=math.degrees(math.atan2(end_arc[1] - ponto_az[1], end_arc[0] - ponto_az[0])),
            dxfattribs={'layer': 'Azimute'}
        )
        logger.info(f"Arco do azimute desenhado com sucesso com valor de {azimute}° no ponto {ponto_az}")

       # Adicionar rótulo do azimute diretamente com o texto "Azimute:"
        azimuth_label = f"Azimute: {convert_to_dms(azimute)}"  # Incluir o prefixo "Azimute:"

        # Calcular a posição do rótulo
        label_position = (
            ponto_az[0] + 1.0 * math.cos(math.radians(azimute / 2)),
            ponto_az[1] + 1.0 * math.sin(math.radians(azimute / 2))
        )

        # Adicionar o texto ao desenho
        msp.add_text(
            azimuth_label,
            dxfattribs={
                'height': 0.25,
                'layer': 'Azimute',
                'insert': label_position  # Define a posição diretamente
            }
        )

        logger.info(f"Rótulo do azimute adicionado com sucesso: '{azimuth_label}' em {label_position}")


    except Exception as e:
        logger.error(f"Erro na função `add_azimuth_arc_to_dxf`: {e}")

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

def add_north_arrow(msp, base_point, length=10):
    """
    Adiciona uma seta (linha) apontando para o Norte a partir do ponto base (Az).
    """
    # Linha vertical representando o Norte
    north_end = (base_point[0], base_point[1] + length)
    msp.add_line(start=base_point, end=north_end, dxfattribs={'layer': 'LAYOUT_AZIMUTES'})

    # Adiciona o texto "N"
    msp.add_text("N", dxfattribs={
        'height': 1.0,
        'insert': (north_end[0], north_end[1] + 1),
        'layer': 'LAYOUT_AZIMUTES'
    })



# Função para calcular azimute e distância
def calculate_azimuth_and_distance(start_point, end_point):
    dx = end_point[0] - start_point[0]
    dy = end_point[1] - start_point[1]
    distance = math.hypot(dx, dy)
    azimuth = math.degrees(math.atan2(dx, dy))
    if azimuth < 0:
        azimuth += 360
    return azimuth, distance


def add_azimuth_arc(doc, msp, ponto_az, v1, azimuth, radius=8):
    """
    Adiciona o arco geométrico representando o ângulo de Azimute entre o norte e a linha Az→V1.
    """
    try:
        # Cria a camada específica caso não exista
        if 'LAYOUT_AZIMUTES' not in doc.layers:
            doc.layers.new(name='LAYOUT_AZIMUTES', dxfattribs={'color': 5})

        # Ângulo inicial sempre aponta para o norte (90° na convenção CAD)
        start_angle = 90.0

        # O ângulo final é obtido subtraindo do azimute (pois CAD mede no sentido anti-horário)
        end_angle = 90.0 - azimuth

        # Garante que os ângulos estejam no intervalo 0-360
        if end_angle < 0:
            end_angle += 360

        # Adiciona o arco geométrico ao DXF
        msp.add_arc(
            center=ponto_az,
            radius=radius,
            start_angle=end_angle,
            end_angle=start_angle,
            dxfattribs={'layer': 'LAYOUT_AZIMUTES'}
        )

        # Adiciona o texto de rótulo próximo ao arco (já está correto)
        mid_angle_rad = math.radians((start_angle + end_angle) / 2)
        label_position = (
            ponto_az[0] + (radius + 1.5) * math.cos(mid_angle_rad),
            ponto_az[1] + (radius + 1.5) * math.sin(mid_angle_rad)
        )
        azimuth_label = f"Azimute = {convert_to_dms(azimuth)}"
        msp.add_text(
            azimuth_label,
            dxfattribs={
                'height': 1.0,
                'layer': 'LAYOUT_AZIMUTES',
                'insert': label_position
            }
        )

        logger.info(f"✅ Arco do azimute ({azimuth_label}) adicionado com sucesso.")

    except Exception as e:
        logger.error(f"❌ Erro ao adicionar arco do azimute: {e}")



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
def create_memorial_descritivo(
    uuid_str, doc, msp, lines, proprietario, matricula, caminho_salvar,
    excel_file_path, ponto_az, distance_az_v1, azimute_az_v1, tipo, diretorio_concluido=None, encoding='ISO-8859-1'
):

    # Carregar confrontantes da planilha FECHADA
    confrontantes_df = pd.read_excel(excel_file_path)
    confrontantes_dict = dict(zip(confrontantes_df['Código'], confrontantes_df['Confrontante']))

    if confrontantes_df.empty:
        logger.error("❌ Planilha de confrontantes está vazia.")
        return None

    if not lines:
        logger.error("❌ Sem linhas disponíveis no DXF.")
        return None

    ordered_points = [line[0] for line in lines] + [lines[-1][1]]

    area = calculate_polygon_area(ordered_points)

    if area < 0:
        ordered_points.reverse()
        logger.info("Pontos reorganizados para sentido horário.")

    data = []
    total_vertices = len(ordered_points) - 1

    for i in range(total_vertices):
        start_point = ordered_points[i]
        end_point = ordered_points[i + 1]

        azimuth, distance = calculate_azimuth_and_distance(start_point, end_point)
        azimuth_dms = convert_to_dms(azimuth)

        confrontante = confrontantes_df.iloc[i]['Confrontante'] if i < len(confrontantes_df) else "Desconhecido"

        coord_e_ponto_az = f"{ponto_az[0]:.3f}".replace('.', ',') if i == 0 else ""
        coord_n_ponto_az = f"{ponto_az[1]:.3f}".replace('.', ',') if i == 0 else ""

        data.append({
            "V": f"V{i + 1}",
            "E": f"{start_point[0]:.3f}".replace('.', ','),
            "N": f"{start_point[1]:.3f}".replace('.', ','),
            "Z": "0,000",
            "Divisa": f"V{i + 1}_V{1 if (i + 1) == total_vertices else i + 2}",
            "Azimute": azimuth_dms,
            "Distancia(m)": f"{distance:.2f}".replace('.', ','),
            "Confrontante": confrontante,
            "Coord_E_ponto_Az": coord_e_ponto_az,
            "Coord_N_ponto_Az": coord_n_ponto_az,
            "distancia_Az_V1": f"{distance_az_v1:.2f}".replace('.', ',') if i == 0 else "",
            "Azimute Az_V1": convert_to_dms(azimute_az_v1) if i == 0 else ""
        })

        # Adicionar labels no DXF
        add_label_and_distance(doc, msp, start_point, end_point, f"V{i + 1}", distance)

    # Caminho padronizado do Excel de saída
    excel_output_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_{matricula}.xlsx")

    # Salvar no Excel
    df = pd.DataFrame(data, dtype=str)
    df.to_excel(excel_output_path, index=False)

    # Formatar Excel
    wb = openpyxl.load_workbook(excel_output_path)
    ws = wb.active

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    column_widths = {
        "A": 8, "B": 15, "C": 15, "D": 10, "E": 20,
        "F": 15, "G": 15, "H": 40, "I": 20, "J": 20,
        "K": 18, "L": 18,
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(excel_output_path)
    logger.info(f"✅ Excel salvo e formatado: {excel_output_path}")

    # Adicionar arco de Azimute ao DXF
    try:
        msp = doc.modelspace()
        v1 = ordered_points[0]
        azimuth = calculate_azimuth(ponto_az, v1)
        add_azimuth_arc(doc, msp, ponto_az, v1, azimuth)
        logger.info("✅ Arco de azimute adicionado ao DXF.")
    except Exception as e:
        logger.error(f"❌ Erro ao adicionar arco de azimute: {e}")
    
    # Adicionar linha entre ponto Az e V1 (parte faltante adicionada aqui)
    try:
        msp = doc.modelspace()
        msp.add_line(start=ponto_az, end=v1, dxfattribs={'layer': 'LAYOUT_AZIMUTES'})
        logger.info("✅ Linha Az→V1 adicionada ao DXF.")
    except Exception as e:
        logger.error(f"❌ Erro ao adicionar linha Az→V1: {e}")


    # Adicionar distância entre Az e V1 no DXF
    try:
        msp = doc.modelspace()
        add_label_and_distance(doc, msp, ponto_az, v1, "Az-V1", distance_az_v1)
        logger.info(f"✅ Distância Az-V1 ({distance_az_v1:.2f} m) adicionada ao DXF.")
    except Exception as e:
        logger.error(f"❌ Erro ao adicionar distância Az-V1: {e}")

    # Adicionar linha apontando para o Norte no ponto Az
    try:
        msp = doc.modelspace()  # É importante garantir o msp atualizado aqui também
        add_north_arrow(msp, ponto_az)
        logger.info("✅ Linha Norte adicionada ao DXF.")
    except Exception as e:
        logger.error(f"❌ Erro ao adicionar linha Norte: {e}")

    # Salvar o DXF com as alterações
    try:
        dxf_output_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_{matricula}.dxf")
        doc.saveas(dxf_output_path)
        logger.info(f"✅ DXF atualizado salvo: {dxf_output_path}")
    except Exception as e:
        logger.error(f"❌ Erro ao salvar DXF atualizado: {e}")

    return excel_output_path





def create_memorial_document(
    uuid_str, proprietario, matricula, descricao, excel_file_path, template_path, 
    output_path, perimeter_dxf, area_dxf, desc_ponto_Az, Coorde_E_ponto_Az, Coorde_N_ponto_Az,
    azimuth, distance, uso_solo, area_imovel, cidade, rua, comarca, RI, caminho_salvar, tipo
):
    try:
        # Ler arquivo Excel
        df = pd.read_excel(excel_file_path)
        df['N'] = pd.to_numeric(df['N'].astype(str).str.replace(',', '.'), errors='coerce')
        df['E'] = pd.to_numeric(df['E'].astype(str).str.replace(',', '.'), errors='coerce')
        df['Distancia(m)'] = pd.to_numeric(df['Distancia(m)'].astype(str).str.replace(',', '.'), errors='coerce')

        # Criar documento Word
        doc_word = Document(template_path)
        set_default_font(doc_word)

        p1 = doc_word.add_paragraph("MEMORIAL DESCRITIVO", style='Normal')
        p1.runs[0].bold = True
        p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc_word.add_paragraph()

        texto_tipo = {
            "ETE": f"Área da matrícula {matricula} destinada a {descricao} - SES de {cidade}",
            "REM": f"Área remanescente da matrícula {matricula} destinada a {descricao}",
            "SER": f"Área da matrícula {matricula} destinada à SERVIDÃO ADMINISTRATIVA DE ACESSO À {descricao}",
            "ACE": f"Área da matrícula {matricula} destinada ao ACESSO DA SERVIDÃO ADMINISTRATIVA DA {descricao}",
        }.get(tipo, "Tipo não especificado")

        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Imóvel: ")
        p.add_run(texto_tipo).bold = True

        doc_word.add_paragraph(f"Matrícula: Número - {matricula} do {RI} de {comarca}", style='Normal')
        doc_word.add_paragraph(f"Proprietário: {proprietario}", style='Normal')
        doc_word.add_paragraph(f"Local: {rua} - {cidade}", style='Normal')
        doc_word.add_paragraph(f"Área: {area_dxf:,.2f} m²".replace(",", "X").replace(".", ",").replace("X", "."), style='Normal')
        doc_word.add_paragraph(f"Perímetro: {perimeter_dxf:,.2f} m".replace(",", "X").replace(".", ",").replace("X", "."), style='Normal')
        doc_word.add_paragraph()

        area_dxf_formatada = f"{area_dxf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        texto_paragrafo = (f"Área {uso_solo} com {area_dxf_formatada} m², parte de um todo maior da Matrícula Nº {matricula} com {area_imovel} "
                           f"do {RI} de {comarca}, localizada na {rua}, na cidade de {cidade}, definida através do seguinte levantamento "
                           "topográfico, onde os ângulos foram medidos no sentido horário.")
        p = doc_word.add_paragraph(texto_paragrafo, style='Normal')
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        doc_word.add_paragraph()

        coord_E_ponto_Az = f"{Coorde_E_ponto_Az:.3f}".replace(".", ",")
        coord_N_ponto_Az = f"{Coorde_N_ponto_Az:.3f}".replace(".", ",")
        doc_word.add_paragraph(
            f"O Ponto Az, ponto de amarração, está localizado na {desc_ponto_Az} nas coordenadas "
            f"E(X) {coord_E_ponto_Az} e N(Y) {coord_N_ponto_Az}.", style='Normal'
        ).alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        doc_word.add_paragraph()

        distance_formatada = f"{distance:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        p = doc_word.add_paragraph(style='Normal')
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p.add_run(f"Daí, com Azimute de {convert_to_dms(azimuth)} e distância de {distance_formatada} m, chega-se ao Vértice ")
        p.add_run("V1").bold = True
        p.add_run(", origem da descrição desta área.")
        doc_word.add_paragraph()

        initial = df.iloc[0]
        coord_N_inicial = f"{initial['N']:.3f}".replace(".", ",")
        coord_E_inicial = f"{initial['E']:.3f}".replace(".", ",")
        doc_word.add_paragraph("Pontos definidos pelas Coordenadas Planas no Sistema U.T.M. – SIRGAS 2000.", style='Normal')
        doc_word.add_paragraph()

        p2 = doc_word.add_paragraph(style='Normal')
        p2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p2.add_run("Inicia-se a descrição deste perímetro no vértice ")
        p2.add_run(f"{initial['V']}").bold = True
        p2.add_run(f", de coordenadas N(Y) {coord_N_inicial} e E(X) {coord_E_inicial}, situado no limite com {initial['Confrontante']}.")
        doc_word.add_paragraph()

        for i in range(len(df)):
            current = df.iloc[i]
            next_point = df.iloc[(i + 1) % len(df)]

            distancia_formatada = f"{current['Distancia(m)']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            coord_N_formatada = f"{next_point['N']:.3f}".replace(".", ",")
            coord_E_formatada = f"{next_point['E']:.3f}".replace(".", ",")

            complemento = ", origem desta descrição," if next_point['V'] == 'V1' else ""

            p = doc_word.add_paragraph(style='Normal')
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            p.add_run(f"Deste, segue com azimute de {current['Azimute']} e distância de {distancia_formatada} m, "
                      f"confrontando neste trecho com área pertencente à {current['Confrontante']}, até o vértice ")
            p.add_run(f"{next_point['V']}").bold = True
            p.add_run(f"{complemento} de coordenadas N(Y) {coord_N_formatada} e E(X) {coord_E_formatada};")
            doc_word.add_paragraph()

        
        data_atual = datetime.now().strftime("%d de %B de %Y")

        # converte mês para português
        for ingles, portugues in MESES_PT_BR.items():
            if ingles in data_atual:
                data_atual = data_atual.replace(ingles, portugues)
                break
        doc_word.add_paragraph(f"\nPorto Alegre, RS, {data_atual}.", style='Normal')
        doc_word.add_paragraph("\n\n")

        output_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_{matricula}.docx")
        doc_word.save(output_path)
        logger.info(f"✅ Memorial descritivo salvo em: {output_path}")

    except Exception as e:
        logger.error(f"❌ Erro ao criar memorial descritivo: {e}")


        
# Função principal
def main_poligonal_fechada(uuid_str, excel_path, dxf_path, diretorio_preparado, diretorio_concluido, caminho_template):

    caminho_salvar = diretorio_concluido 
    os.makedirs(caminho_salvar, exist_ok=True)

    # 🔹 Carrega dados do imóvel
    dados_imovel_df = pd.read_excel(excel_path, sheet_name='Dados_do_Imóvel', header=None)
    dados_imovel = dict(zip(dados_imovel_df.iloc[:, 0], dados_imovel_df.iloc[:, 1]))

    # 🔹 Extrai variáveis necessárias
    proprietario = dados_imovel.get("NOME DO PROPRIETÁRIO", "").strip()
    matricula = sanitize_filename(dados_imovel.get("DOCUMENTAÇÃO DO IMÓVEL", "").strip())
    descricao = dados_imovel.get("OBRA", "").strip()
    uso_solo = dados_imovel.get("ZONA", "").strip()
    area_imovel = dados_imovel.get("ÁREA TOTAL DO TERRENO DOCUMENTADA", "").replace("\t", "").replace("\n", "").strip()
    cidade = dados_imovel.get("CIDADE", "").strip()
    rua = dados_imovel.get("LOCAL", "").strip()
    comarca = dados_imovel.get("COMARCA", "").strip()
    RI = dados_imovel.get("RI", "").strip()
    desc_ponto_Az = dados_imovel.get("AZ", "").strip()

    # 🔹 Define tipo pela nomenclatura do DXF
    dxf_filename = os.path.basename(dxf_path).upper()

    if "ETE" in dxf_filename:
        tipo = "ETE"
    elif "REM" in dxf_filename:
        tipo = "REM"
    elif "SER" in dxf_filename:
        tipo = "SER"
    elif "ACE" in dxf_filename:
        tipo = "ACE"
    else:
        logger.error("❌ Tipo (ETE, REM, SER ou ACE) não identificado no nome do DXF.")
        return

    # 🔹 Busca planilha FECHADA correta com uuid_str
    padrao_busca = os.path.join(diretorio_preparado, f"{uuid_str}_FECHADA_{tipo}.xlsx")
    arquivos_encontrados = glob.glob(padrao_busca)

    if not arquivos_encontrados:
        logger.error(f"❌ Planilha confrontantes FECHADA não encontrada: {padrao_busca}")
        return

    excel_confrontantes = arquivos_encontrados[0]

    # 🔹 Limpa DXF
    dxf_limpo_path = os.path.join(caminho_salvar, f"{uuid_str}_DXF_LIMPO_{matricula}.dxf")
    dxf_file_path = limpar_dxf(dxf_path, dxf_limpo_path)

    # 🔹 Extrair geometria e ponto Az automaticamente do DXF
    doc, lines, perimeter_dxf, area_dxf, ponto_az, _ = get_document_info_from_dxf(dxf_file_path)

    if not doc or not ponto_az:
        logger.error("❌ Documento inválido ou ponto Az não encontrado no DXF.")
        return

    v1 = lines[0][0]
    distance_az_v1 = calculate_distance(ponto_az, v1)
    azimute_az_v1 = calculate_azimuth(ponto_az, v1)

    logger.info(f"📌 Azimute Az→V1: {azimute_az_v1:.4f}°, Distância: {distance_az_v1:.2f} m")

    try:
        doc_dxf = ezdxf.readfile(dxf_file_path)
        msp = doc_dxf.modelspace()
    except Exception as e:
        logger.error(f"Erro ao abrir DXF limpo: {e}")
        return

    # 🔹 Criar memorial descritivo (planilha Excel final)
    excel_file_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}.xlsx")

   
    excel_resultado = create_memorial_descritivo(
        uuid_str=uuid_str,
        doc=doc,
        msp=msp,
        lines=lines,
        proprietario=proprietario,
        matricula=matricula,
        caminho_salvar=caminho_salvar,
        excel_file_path=excel_file_path,
        ponto_az=ponto_az,
        distance_az_v1=distance_az_v1,
        azimute_az_v1=azimute_az_v1,
        tipo=tipo,
        diretorio_concluido=caminho_salvar
    )

    if excel_resultado:
        # 🔹 Gerar Memorial DOCX final
        output_docx_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_{matricula}.docx")

        create_memorial_document(
            uuid_str=uuid_str,
            proprietario=proprietario,
            matricula=matricula,
            descricao=descricao,
            excel_file_path=excel_output_path,
            template_path=caminho_template,
            output_path=output_docx_path,
            perimeter_dxf=perimeter_dxf,
            area_dxf=area_dxf,
            desc_ponto_Az=desc_ponto_Az,
            Coorde_E_ponto_Az=ponto_az[0],
            Coorde_N_ponto_Az=ponto_az[1],
            azimuth=azimute_az_v1,
            distance=distance_az_v1,
            uso_solo=uso_solo,
            area_imovel=area_imovel,
            cidade=cidade,
            rua=rua,
            comarca=comarca,
            RI=RI,
            caminho_salvar=caminho_salvar,
            tipo=tipo
        )

        logger.info("🔵 [main_poligonal_fechada] Processamento concluído com sucesso.")
        print("Processamento concluído com sucesso.")

    else:
        logger.error("❌ Falha ao gerar memorial descritivo.")
        print("Erro ao processar o arquivo DXF.")



