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

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

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

# Função que processa as linhas da poligonal
def get_document_info_from_dxf(dxf_file_path):
    try:
        doc = ezdxf.readfile(dxf_file_path)  # Carregar o arquivo DXF com ezdxf
        msp = doc.modelspace()  # Acessar o ModelSpace do DXF

        lines = []
        ponto_az = None
        area_poligonal = None

        # Iterar pelas entidades no ModelSpace
        for entity in msp.query('LWPOLYLINE'):  # Trabalhar com LWPOLYLINE
            if entity.closed:
                points = entity.get_points('xy')  # Obter os pontos da polilinha
                num_points = len(points)
                
                # Processar os pontos da polilinha, evitando duplicação do último ponto
                for i in range(num_points - 1):
                    start_point = (points[i][0], points[i][1], 0)
                    end_point = (points[i + 1][0], points[i + 1][1], 0)
                    lines.append((start_point, end_point))

                # Conectar o último ponto ao primeiro
                start_point = (points[-1][0], points[-1][1], 0)
                end_point = (points[0][0], points[0][1], 0)
                lines.append((start_point, end_point))
                
                # Calcular a área
                x = [point[0] for point in points]
                y = [point[1] for point in points]
                area_poligonal = abs(sum(x[i] * y[(i + 1) % num_points] - x[(i + 1) % num_points] * y[i] for i in range(num_points)) / 2)

        # Verificar se nenhuma linha foi encontrada
        if not lines:
            print("Nenhuma polilinha encontrada no arquivo DXF.")
            return None, [], None, None

        # Verificar Ponto Az em textos e blocos
        for entity in msp.query('TEXT'):
            if "Az" in entity.dxf.text:
                ponto_az = (entity.dxf.insert.x, entity.dxf.insert.y, 0)
                print(f"Ponto Az encontrado em texto: {ponto_az}")

        for entity in msp.query('INSERT'):  # Verificar blocos
            if "Az" in entity.dxf.name:
                ponto_az = (entity.dxf.insert.x, entity.dxf.insert.y, 0)
                print(f"Ponto Az encontrado no bloco: {ponto_az}")

        for entity in msp.query('POINT'):  # Verificar pontos
            ponto_az = (entity.dxf.location.x, entity.dxf.location.y, 0)
            print(f"Ponto Az encontrado como ponto: {ponto_az}")

        # Se não encontrou o Ponto Az
        if not ponto_az:
            print("Ponto Az não encontrado no arquivo DXF.")
            return None, lines, None, None

        print(f"Linhas processadas: {len(lines)}")
        print(f"Área da poligonal: {area_poligonal:.6f} m²")

        return doc, lines, ponto_az, area_poligonal

    except Exception as e:
        print(f"Erro ao obter informações do documento: {e}")
        return None, [], None, None

def is_clockwise(points):
    """
    Verifica se a poligonal está no sentido horário.
    Retorna True se for horário, False se anti-horário.
    """
    area = 0.0
    for i in range(len(points)):
        j = (i + 1) % len(points)
        area += points[i][0] * points[j][1]
        area -= points[j][0] * points[i][1]
    return area < 0

def ensure_counterclockwise(points):
    """
    Garante que a lista de pontos esteja no sentido anti-horário.
    Se estiver no sentido horário, inverte a ordem dos pontos.
    """
    if is_clockwise(points):
        points.reverse()
    return points

    
# 🔹 Função para definir a fonte padrão
def set_default_font(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    
def calculate_point_on_line(start, end, distance):
    """
    Calcula um ponto a uma determinada distância sobre uma linha.
    """
    dx, dy = end[0] - start[0], end[1] - start[1]
    length = math.hypot(dx, dy)
    return (
        start[0] + dx / length * distance,
        start[1] + dy / length * distance
    )

def add_azimuth_arc_to_dxf(msp, ponto_az, v1, azimute):
    """
    Adiciona o arco do azimute ao DXF usando ezdxf.
    """
    try:
        print(f"Iniciando a adição do arco de azimute. Azimute: {azimute}°")

        # Criar camada 'Azimute', se não existir
        if 'Azimute' not in msp.doc.layers:
            msp.doc.layers.new(name='Azimute', dxfattribs={'color': 1})
            print("Camada 'Azimute' criada com sucesso.")

        # Traçar segmento entre Az e V1
        msp.add_line(start=ponto_az, end=v1, dxfattribs={'layer': 'Azimute'})
        print(f"Segmento entre Az e V1 desenhado de {ponto_az} para {v1}")

        # Traçar segmento para o norte
        north_point = (ponto_az[0], ponto_az[1] + 2)
        msp.add_line(start=ponto_az, end=north_point, dxfattribs={'layer': 'Azimute'})
        print(f"Linha para o norte desenhada com sucesso de {ponto_az} para {north_point}")

        # Calcular o ponto inicial (1 metro de Az para V1)
        start_arc = calculate_point_on_line(ponto_az, v1, 1)

        # Calcular o ponto final (1 metro de Az para o Norte)
        end_arc = calculate_point_on_line(ponto_az, north_point, 1)

        # Traçar o arco do azimute no sentido horário
        msp.add_arc(
            center=ponto_az,
            radius=1,
            start_angle=math.degrees(math.atan2(start_arc[1] - ponto_az[1], start_arc[0] - ponto_az[0])),
            end_angle=math.degrees(math.atan2(end_arc[1] - ponto_az[1], end_arc[0] - ponto_az[0])),
            dxfattribs={'layer': 'Azimute'}
        )
        print(f"Arco do azimute desenhado com sucesso com valor de {azimute}° no ponto {ponto_az}")

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

        print(f"Rótulo do azimute adicionado com sucesso: '{azimuth_label}' em {label_position}")


    except Exception as e:
        print(f"Erro na função `add_azimuth_arc_to_dxf`: {e}")

def calculate_polygon_area(points):
    """
    Calcula a área de uma poligonal fechada utilizando o método do produto cruzado com alta precisão.
    :param points: Lista de coordenadas [(x1, y1), (x2, y2), ...].
    :return: Área da poligonal.
    """
    n = len(points)
    if n < 3:
        return Decimal(0)  # Não é uma poligonal válida

    area = Decimal(0)
    for i in range(n):
        x1, y1 = Decimal(points[i][0]), Decimal(points[i][1])
        x2, y2 = Decimal(points[(i + 1) % n][0]), Decimal(points[(i + 1) % n][1])
        area += x1 * y2 - y1 * x2

    return abs(area) / 2
    
def convert_to_dms(decimal_degrees):
    """
    Converte graus decimais para o formato graus, minutos e segundos (DMS).
    """
    try:
        # Verificar se o valor é NaN
        if math.isnan(decimal_degrees):
            raise ValueError("Valor de entrada é NaN")

        degrees = int(decimal_degrees)
        minutes = int((abs(decimal_degrees) - abs(degrees)) * 60)
        seconds = round((abs(decimal_degrees) - abs(degrees) - minutes / 60) * 3600, 2)
        return f"{degrees}°{minutes}'{seconds}\""
    except Exception as e:
        print(f"Erro na conversão para DMS: {e}")
        return "0°0'0.00\""  # Valor padrão em caso de erro

def calculate_distance(p1, p2):
    return math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2)


def degrees_to_dms(angle):
    """
    Converte um ângulo em graus decimais para o formato graus, minutos e segundos (° ' ").
    """
    degrees = int(angle)
    minutes = int((angle - degrees) * 60)
    seconds = round((angle - degrees - minutes / 60) * 3600, 2)
    return f"{degrees}°{minutes}'{seconds}\""
import math

def calculate_azimuth(p1, p2):
    """
    Calcula o azimute entre dois pontos p1 e p2.
    
    :param p1: Coordenadas do ponto 1 (E, N)
    :param p2: Coordenadas do ponto 2 (E, N)
    :return: Azimute em graus decimais
    """
    delta_x = p2[0] - p1[0]
    delta_y = p2[1] - p1[1]
    
    azimuth = math.degrees(math.atan2(delta_x, delta_y)) % 360  # Garantir valor entre 0° e 360°
    
    return azimuth


def create_arrow_block(doc, block_name="ARROW"):
    """
    Cria um bloco no DXF representando uma seta sólida como um triângulo.
    """
    if block_name in doc.blocks:
        return  # O bloco já existe

    block = doc.blocks.new(name=block_name)

    # Definir o triângulo da seta
    length = 0.5  # Comprimento da seta
    base_half_length = length / 2

    tip = (0, 0)  # Ponta da seta no eixo de coordenadas
    base1 = (-base_half_length, -length)
    base2 = (base_half_length, -length)

    block.add_solid([base1, base2, tip])
import math

def add_giro_angular_arc_to_dxf(doc_dxf, v1, az, v2, radius=1.0):
    """
    Adiciona um arco representando o giro angular horário no espaço de modelo do DXF já aberto.
    """
    try:
        msp = doc_dxf.modelspace()

        # Traçar a reta entre V1 e Az
        msp.add_line(start=v1[:2], end=az[:2])
        print(f"Linha entre V1 e Az traçada com sucesso.")

        # Definir os pontos de apoio
        def calculate_displacement(point1, point2, distance):
            dx = point2[0] - point1[0]
            dy = point2[1] - point1[1]
            magnitude = math.hypot(dx, dy)
            return (
                point1[0] + (dx / magnitude) * distance,
                point1[1] + (dy / magnitude) * distance,
            )

        # Calcular os pontos de apoio
        ponto_inicial = calculate_displacement(v1, v2, radius)  # 2m na reta V1-V2
        ponto_final = calculate_displacement(v1, az, radius)   # 2m na reta V1-Az

        # Calcular os ângulos dos vetores
        angle_v2 = math.degrees(math.atan2(ponto_inicial[1] - v1[1], ponto_inicial[0] - v1[0]))
        angle_az = math.degrees(math.atan2(ponto_final[1] - v1[1], ponto_final[0] - v1[0]))

        # Calcular o giro angular no sentido horário
        giro_angular = (angle_az - angle_v2) % 360  # Garantir que o ângulo esteja no intervalo [0, 360)
        if giro_angular < 0:  # Caso negativo, ajustar para o sentido horário
            giro_angular += 360

        print(f"Giro angular calculado corretamente: {giro_angular:.2f}°")

        # Traçar o arco
        msp.add_arc(center=v1[:2], radius=radius, start_angle=angle_v2, end_angle=angle_az)
        print(f"Arco do giro angular traçado com sucesso.")

        # Adicionar rótulo ao arco
        label_offset = 3.0
        deslocamento_x=3
        deslocamento_y=-3
        angle_middle = math.radians((angle_v2 + angle_az) / 2)
        label_position = (
            v1[0] + (label_offset+deslocamento_x) * math.cos(angle_middle),
            v1[1] + (label_offset+deslocamento_y) * math.sin(angle_middle),
        )
        # Converter o ângulo para DMS e exibir no rótulo
        giro_angular_dms = f"Giro Angular:{convert_to_dms(giro_angular)}"
        msp.add_text(
            giro_angular_dms,
            dxfattribs={
                'height': 0.3,
                'layer': 'Labels',
                'insert': label_position  # Define a posição do texto
            }
        )
        print(f"Rótulo do giro angular ({giro_angular_dms}) adicionado com sucesso.")

    except Exception as e:
        print(f"Erro ao adicionar o arco do giro angular ao DXF: {e}") 



def calculate_arc_angles(p1, p2, p3):
    """
    Calcula os ângulos de início e fim do arco para representar o ângulo interno da poligonal.
    O arco deve sempre ser desenhado dentro da poligonal.
    """
    try:
        # Vetores a partir de p2
        dx1, dy1 = p1[0] - p2[0], p1[1] - p2[1]  # Vetor de p2 para p1
        dx2, dy2 = p3[0] - p2[0], p3[1] - p2[1]  # Vetor de p2 para p3

        # Ângulos dos vetores em relação ao eixo X
        angle1 = math.degrees(math.atan2(dy1, dx1)) % 360
        angle2 = math.degrees(math.atan2(dy2, dx2)) % 360

        # Determinar o ângulo interno
        internal_angle = (angle2 - angle1) % 360
#         if internal_angle > 180:
#             internal_angle = 360 - internal_angle  # Complementar para 360°

        # Ajustar os ângulos de início e fim para sempre formar o ângulo interno dentro da poligonal
        start_angle = angle1
        end_angle = angle1 + internal_angle  # Sempre desenha o arco no sentido correto

        # Se o ângulo interno for maior que 180°, inverter os ângulos para garantir o lado interno
        if internal_angle > 180:
            start_angle, end_angle = end_angle, start_angle

        return start_angle % 360, end_angle % 360

    except Exception as e:
        print(f"Erro ao calcular ângulos do arco: {e}")
        return 0, 0  # Retorno seguro em caso de erro




def insert_and_rotate_arrow(msp, center, radius, angle, rotation_adjustment, block_name="ARROW"):
    """
    Insere um bloco da seta alinhado ao arco e rotacionado conforme especificado.
    """
    arrow_tip_x = center[0] + radius * math.cos(angle)
    arrow_tip_y = center[1] + radius * math.sin(angle)
    arrow_tip = (arrow_tip_x, arrow_tip_y)

    rotation_angle = math.degrees(angle) + rotation_adjustment

    msp.add_blockref(
        block_name,
        insert=arrow_tip,
        dxfattribs={"rotation": rotation_angle}
    )
import math





def add_angle_visualization_to_dwg(msp, ordered_points, angulos_excel):
    try:
        total_points = len(ordered_points)
        
        for i, p2 in enumerate(ordered_points):
            if i == 0:
                print("⏩ Ignorando arco e rótulo para V1")
                continue
            p1 = ordered_points[i - 1] if i > 0 else ordered_points[-1]
            p3 = ordered_points[(i + 1) % total_points]

            def calculate_displacement(base_point, direction_point, distance):
                dx = direction_point[0] - base_point[0]
                dy = direction_point[1] - base_point[1]
                magnitude = math.hypot(dx, dy)
                return (
                    base_point[0] + (dx / magnitude) * distance,
                    base_point[1] + (dy / magnitude) * distance,
                )

            lado_antes = math.hypot(p2[0] - p1[0], p2[1] - p1[1])
            lado_depois = math.hypot(p3[0] - p2[0], p3[1] - p2[1])
            lado_menor = min(lado_antes, lado_depois)

            radius = lado_menor * 0.8 if lado_menor <= 0.5 else 1.0

            ponto_inicial = calculate_displacement(p2, p3, radius)
            ponto_final = calculate_displacement(p2, p1, radius)

            start_angle = math.degrees(math.atan2(ponto_inicial[1] - p2[1], ponto_inicial[0] - p2[0]))
            end_angle = math.degrees(math.atan2(ponto_final[1] - p2[1], ponto_final[0] - p2[0]))

            if end_angle < start_angle:
                end_angle += 360

            internal_angle_dms = angulos_excel[i]

            msp.add_arc(
                center=p2,
                radius=radius,
                start_angle=start_angle,
                end_angle=end_angle,
                dxfattribs={'layer': 'Internal_Arcs'}
            )

            label_offset = 1
            label_position = (
                p2[0] + label_offset * math.cos(math.radians((start_angle + end_angle) / 2)),
                p2[1] + label_offset * math.sin(math.radians((start_angle + end_angle) / 2))
            )

            msp.add_text(
                internal_angle_dms,
                dxfattribs={
                    'height': 0.3,
                    'layer': 'Labels',
                    'insert': label_position
                }
            )

            print(f"Vértice V{i+1}: Ângulo interno {internal_angle_dms}")

    except Exception as e:
        print(f"Erro ao adicionar ângulos internos ao DXF: {e}")




def calculate_internal_angle(p1, p2, p3):
    try:
        # Vetores a partir do ponto central (p2)
        dx1, dy1 = p1[0] - p2[0], p1[1] - p2[1]
        dx2, dy2 = p3[0] - p2[0], p3[1] - p2[1]

        # Ângulos dos vetores em relação ao eixo X
        angle1 = math.atan2(dy1, dx1)
        angle2 = math.atan2(dy2, dx2)

        # Ângulo formado entre os vetores
        internal_angle = (angle2 - angle1) % (2 * math.pi)

        # Se o ângulo for maior que 180°, use o suplementar
        if internal_angle > math.pi:
            internal_angle = (2 * math.pi) - internal_angle

        # Retornar o ângulo em graus
        return math.degrees(internal_angle)

    except Exception as e:
        print(f"Erro inesperado ao calcular o ângulo interno: {e}")
        return 0


def calculate_label_position(p2, start_angle, end_angle, radius=1.8):
    """
    Calcula a posição do rótulo do ângulo interno, deslocando-o para uma posição central no arco.
    
    :param p2: Ponto central do arco (vértice da poligonal)
    :param start_angle: Ângulo de início do arco
    :param end_angle: Ângulo de fim do arco
    :param radius: Raio do deslocamento para evitar sobreposição
    :return: Posição (x, y) onde o rótulo do ângulo interno será colocado
    """
    try:
        # Calcular o ângulo médio entre os dois ângulos do arco
        mid_angle = math.radians((start_angle + end_angle) / 2)

        # Calcular a posição do rótulo deslocado no ângulo médio
        label_x = p2[0] + radius * math.cos(mid_angle)
        label_y = p2[1] + radius * math.sin(mid_angle)

        return (label_x, label_y)

    except Exception as e:
        print(f"Erro ao calcular posição do rótulo do ângulo: {e}")
        return p2  # Retorna o próprio ponto central caso ocorra erro


import math

def add_label_and_distance(msp, start_point, end_point, label, distance):
    """
    Adiciona rótulos e distâncias no espaço de modelo usando ezdxf.
    """
    try:
        # Calcular ponto médio
        mid_point = (
            (start_point[0] + end_point[0]) / 2,
            (start_point[1] + end_point[1]) / 2
        )

        # Vetor da linha
        dx = end_point[0] - start_point[0]
        dy = end_point[1] - start_point[1]
        length = math.hypot(dx, dy)

        # Ângulo da linha
        angle = math.degrees(math.atan2(dy, dx))

        # Corrigir para manter a leitura sempre da esquerda para a direita
        if angle < -90 or angle > 90:
            angle += 180  

        # Afastar o rótulo da linha
        offset = -0.5  # Ajuste o valor para mais afastamento
        perp_x = -dy / length * offset
        perp_y = dx / length * offset
        displaced_mid_point = (mid_point[0] + perp_x, mid_point[1] + perp_y)

        # Criar layer se não existir
        if "Distance_Labels" not in msp.doc.layers:
            msp.doc.layers.new(name="Distance_Labels", dxfattribs={"color": 2})  # Define cor para melhor visualização

        # Adicionar o rótulo da distância no LAYER correto
        msp.add_text(
            f"{distance:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + "",
            dxfattribs={
                "height": 0.3,
                "layer": "Distance_Labels",  # Novo LAYER
                "rotation": angle,
                "insert": displaced_mid_point
            }
        )
        print(f"Distância {distance:.2f} m adicionada corretamente em {displaced_mid_point}")

    except Exception as e:
        print(f"Erro ao adicionar rótulo de distância: {e}")

def calculate_angular_turn(p1, p2, p3):
    """
    Calcula o giro angular no ponto `p2` entre os segmentos `p1-p2` e `p2-p3` no sentido horário.
    Retorna o ângulo em graus.
    """
    import math
    
    dx1, dy1 = p1[0] - p2[0], p1[1] - p2[1]  # Vetor do segmento p1-p2
    dx2, dy2 = p3[0] - p2[0], p3[1] - p2[1]  # Vetor do segmento p2-p3

    angle1 = math.atan2(dy1, dx1)
    angle2 = math.atan2(dy2, dx2)

    # Calcula o ângulo horário
    angular_turn = (angle2 - angle1) % (2 * math.pi)
    angular_turn_degrees = math.degrees(angular_turn)

    return angular_turn_degrees




def create_memorial_descritivo(
    uuid_str, doc, msp, lines, proprietario, matricula, caminho_salvar,
    excel_file_path, ponto_az, distance_az_v1, azimute_az_v1, tipo,
    diretorio_concluido=None, encoding='ISO-8859-1'
):

    """
    Cria o memorial descritivo e o arquivo DXF final para o caso com ponto Az definido no desenho.
    """

    # Carregar confrontantes diretamente da planilha Excel recebida
    confrontantes_df = pd.read_excel(excel_file_path)

    if confrontantes_df.empty:
        logger.error("❌ Planilha de confrontantes está vazia.")
        return None

    confrontantes = confrontantes_df.iloc[:, 1].dropna().tolist()


    if not lines:
        print("Nenhuma linha disponível para criar o memorial descritivo.")
        return None

    dxf_output_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_POLIGONAL_COM_AZ_{matricula}.dxf")


    try:
        doc_dxf = ezdxf.readfile(dxf_file_path)
        msp = doc_dxf.modelspace()
    except Exception as e:
        print(f"Erro ao abrir o arquivo DXF para edição: {e}")
        return None

    # Ordena os pontos da poligonal
    ordered_points = [line[0] for line in lines]
    if ordered_points[-1] != lines[-1][1]:
        ordered_points.append(lines[-1][1])

    ordered_points = ensure_counterclockwise(ordered_points)

    distance_az_v1 = calculate_distance(ponto_az, ordered_points[0])

    # Remove duplicação de ponto final
    tolerancia = 0.001
    if math.isclose(ordered_points[0][0], ordered_points[-1][0], abs_tol=tolerancia) and \
       math.isclose(ordered_points[0][1], ordered_points[-1][1], abs_tol=tolerancia):
        ordered_points.pop()

    try:
        data = []
        total_pontos = len(ordered_points)

        for i in range(total_pontos):
            p1 = ordered_points[i - 1] if i > 0 else ordered_points[-1]
            p2 = ordered_points[i]
            p3 = ordered_points[(i + 1) % total_pontos]

            internal_angle = calculate_internal_angle(p1, p2, p3)
            internal_angle_dms = convert_to_dms(internal_angle)

            description = f"V{i + 1}_V{(i + 2) if i + 1 < total_pontos else 1}"
            dx = p3[0] - p2[0]
            dy = p3[1] - p2[1]
            distance = math.hypot(dx, dy)
            confrontante = confrontantes[i % len(confrontantes)]

            ponto_az_e = f"{ponto_az[0]:,.3f}".replace(",", "").replace(".", ",") if i == 0 else ""
            ponto_az_n = f"{ponto_az[1]:,.3f}".replace(",", "").replace(".", ",") if i == 0 else ""
            distancia_az_v1_str = f"{distance_az_v1:.2f}".replace(".", ",") if i == 0 else ""
            azimute_az_v1_str = convert_to_dms(azimute) if i == 0 else ""
            giro_v1_str = giro_angular_v1_dms if i == 0 else ""

            data.append({
                "V": f"V{i + 1}",
                "E": f"{p2[0]:,.3f}".replace(",", "").replace(".", ","),
                "N": f"{p2[1]:,.3f}".replace(",", "").replace(".", ","),
                "Z": "0,000",
                "Divisa": description,
                "Angulo Interno": internal_angle_dms,
                "Distancia(m)": f"{distance:,.2f}".replace(",", "").replace(".", ","),
                "Confrontante": confrontante,
                "ponto_AZ_E": ponto_az_e,
                "ponto_AZ_N": ponto_az_n,
                "distancia_Az_V1": distancia_az_v1_str,
                "Azimute Az_V1": azimute_az_v1_str,
                "Giro Angular Az_V1_V2": giro_v1_str
            })

            if distance > 0.01:
                add_label_and_distance(msp, p2, p3, f"V{i + 1}", distance)

        # ➕ Salvar Excel
        df = pd.DataFrame(data)
        df.to_excel(excel_file_path, index=False)

        # 📊 Formatação do Excel
        wb = openpyxl.load_workbook(excel_file_path)
        ws = wb.active

        # Cabeçalho
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        col_widths = {
            "A": 8, "B": 15, "C": 15, "D": 0, "E": 15,
            "F": 15, "G": 15, "H": 50, "I": 15,
            "J": 15, "K": 15, "L": 20, "M": 20
        }

        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width

        # Corpo
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        wb.save(excel_file_path)
        print(f"📊 Planilha Excel salva e formatada: {excel_file_path}")

        # ➕ Ângulos internos a partir do Excel
        angulos_excel = [item["Angulo Interno"] for item in data]
        add_angle_visualization_to_dwg(msp, ordered_points, angulos_excel)

        # ➕ Giro Angular
        try:
            v1 = ordered_points[0]
            v2 = ordered_points[1]
            add_giro_angular_arc_to_dxf(doc_dxf, v1, ponto_az, v2)
            print("Giro horário Az-V1-V2 adicionado com sucesso.")
        except Exception as e:
            print(f"Erro ao adicionar giro angular: {e}")

        # ➕ Camada e rótulo de vértices
        if "Vertices" not in msp.doc.layers:
            msp.doc.layers.add("Vertices", dxfattribs={"color": 1})

        for i, vertex in enumerate(ordered_points):
            msp.add_circle(center=vertex, radius=0.5, dxfattribs={"layer": "Vertices"})
            label_pos = (vertex[0] + 0.3, vertex[1] + 0.3)
            msp.add_text(f"V{i + 1}", dxfattribs={
                "height": 0.3,
                "layer": "Vertices",
                "insert": label_pos
            })

        # ➕ Adicionar arco e rótulo do Azimute
        try:
            azimute = calculate_azimuth(ponto_az, v1)
            add_azimuth_arc_to_dxf(msp, ponto_az, v1, azimute)
            print("Arco do Azimute Az-V1 adicionado com sucesso.")
        except Exception as e:
            print(f"Erro ao adicionar arco do azimute: {e}")

        # ➕ Adicionar distância Az–V1
        try:
            distancia_az_v1 = calculate_distance(ponto_az, v1)
            add_label_and_distance(msp, ponto_az, v1, "", distancia_az_v1)
            print(f"Distância Az-V1 ({distancia_az_v1:.2f} m) adicionada com sucesso.")
        except Exception as e:
            print(f"Erro ao adicionar distância entre Az e V1: {e}")

        # ➕ Salvar DXF
        doc_dxf.saveas(dxf_output_path)
        print(f"📁 Arquivo DXF final salvo em: {dxf_output_path}")

    except Exception as e:
        print(f"❌ Erro ao gerar o memorial descritivo: {e}")
        return None

    return excel_file_path








def generate_initial_text(proprietario, matricula, descricao, area, perimeter, rua, cidade, ponto_az, azimute, distancia):
    """
    Gera o texto inicial do memorial descritivo.
    """
    initial_text = (
        f"MEMORIAL DESCRITIVO\n"
        f"NOME PROPRIETÁRIO / OCUPANTE: {proprietario}\n"
        f"DESCRIÇÃO: {descricao}\n"
        f"DOCUMENTAÇÃO: MATRÍCULA {matricula}\n"
        f"ÁREA DO IMÓVEL: {area:.2f} metros quadrados\n"
        f"PERÍMETRO: {perimeter:.4f} metros\n"
        f"Área localizada na rua {rua}, município de {cidade}, com a seguinte descrição:\n"
        f"O Ponto Az está localizado nas coordenadas E {ponto_az[0]:.3f}, N {ponto_az[1]:.3f}.\n"
        f"Daí, com Azimute de {azimute} e distância de {distancia:.2f} metros, chega-se ao Vértice V1, "
        f"origem da área descrição, alinhado com a rua {rua}."
    )
    return initial_text


def generate_angular_text(az_v1, v1_v2, distancia, angulo, rua, confrontante):
    """
    Gera o texto do ângulo em V1 entre Az-V1 e V1-V2.
    """
    angular_text = (
        f"Daí, visando o Ponto “Az”, com giro angular horário de {angulo} e diste {distancia:.2f} metros, "
        f"chega-se ao Vértice V2, também alinhado com a rua {rua} e limítrofe com {confrontante}."
    )
    return angular_text


def generate_recurring_text(df, rua, confrontantes):
    """
    Gera o texto recorrente enquanto percorre os vértices da poligonal.
    """
    recurring_texts = []
    num_vertices = len(df)
    
    for i in range(1, num_vertices - 1):  # De V2 até o penúltimo vértice
        current = df.iloc[i]
        next_vertex = df.iloc[i + 1]
        confrontante = confrontantes[i % len(confrontantes)]
        
        recurring_text = (
            f"Daí, visando o Vértice {current['V']}, com giro angular horário de {current['Angulo Interno']} "
            f"e diste {current['Distancia(m)']} metros, chega-se ao Vértice {next_vertex['V']}, "
            f"também alinhado à rua {rua} e limítrofe com {confrontante}."
        )
        recurring_texts.append(recurring_text)
    
    return recurring_texts


def generate_final_text(df, rua, confrontantes):
    """
    Gera o texto final ao retornar ao V1.
    """
    last_vertex = df.iloc[-1]
    first_vertex = df.iloc[0]
    confrontante = confrontantes[-1 % len(confrontantes)]  # Confrontante do último trecho
    
    final_text = (
        f"Daí, visando o vértice {last_vertex['V']}, com giro angular de {last_vertex['Angulo Interno']} "
        f"e diste {last_vertex['Distancia(m)']} metros, chega-se ao vértice {first_vertex['V']}, "
        f"origem da presente descrição, no alinhamento da rua {rua} e próximo aos lotes de {confrontante}."
    )
    return final_text


def create_memorial_document(
    uuid_str, proprietario, matricula, descricao, excel_file_path, template_path, 
    output_path, perimeter_dxf, area_dxf, desc_ponto_Az, Coorde_E_ponto_Az, Coorde_N_ponto_Az,
    azimuth, distance, uso_solo, area_imovel, cidade, rua, comarca, RI, caminho_salvar, tipo
):

    try:
        df = pd.read_excel(excel_file_path, engine='openpyxl', dtype=str)

        # Carrega confrontantes diretamente da coluna do Excel
        confrontantes = df['Confrontante'].dropna().tolist()

        # Processa colunas numéricas
        df['Distancia(m)'] = df['Distancia(m)'].str.replace(',', '.').astype(float)
        df['E'] = df['E'].str.replace(',', '').astype(float)
        df['N'] = df['N'].str.replace(',', '').astype(float)

        # Calcular perímetro e área
        perimeter = df['Distancia(m)'].sum()
        x = df['E'].values
        y = df['N'].values
        area = abs(sum(x[i] * y[(i + 1) % len(x)] - x[(i + 1) % len(x)] * y[i] for i in range(len(x))) / 2)

        doc_word = Document(caminho_template)

        # 🔴 Remover linhas vazias ou parágrafos indesejados no topo
        while doc_word.paragraphs and not doc_word.paragraphs[0].text.strip():
            doc_word.paragraphs[0]._element.getparent().remove(doc_word.paragraphs[0]._element)

        for para in doc_word.paragraphs:
            if "copilot" in para.text.lower():
                para._element.getparent().remove(para._element)

        set_default_font(doc_word)

        # Título
        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        p.add_run("MEMORIAL DESCRITIVO").bold = True

        doc_word.add_paragraph()

        # Parágrafos descritivos
        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p.add_run("Objetivo: ").bold = True
        p.add_run(f"Área destinada à servidão de passagem para execução de coletor de fundo pertencente à rede coletora de esgoto de {cidade}/RS.")

        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Matrícula Número: ").bold = True
        p.add_run(f"{matricula_texto} - {rgi}")

        area_total_formatada = str(area_total).replace(".", ",")
        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Área Total do Terreno: ").bold = True
        p.add_run(area_total_formatada)

        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Proprietário: ").bold = True
        p.add_run(f"{proprietario} - CPF/CNPJ: {cpf}")

        p = doc_word.add_paragraph(style='Normal')
        p.add_run("Área de Servidão de Passagem: ").bold = True
        p.add_run(f"{area_dxf:.2f}".replace(".", ",") + " m")
        sup = p.add_run("2")
        sup.font.superscript = True
        sup.font.size = Pt(12)

        doc_word.add_paragraph()

        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        #p.add_run("Descrição: ").bold = True
        p.add_run("Área com ").font.name = 'Arial'

        run1 = p.add_run(f"{area_dxf:.2f}".replace(".", ",")+" m")
        run1.font.name = 'Arial'
        run1.font.size = Pt(12)

        run2 = p.add_run("2")
        run2.font.name = 'Arial'
        run2.font.size = Pt(12)
        run2.font.superscript = True

        p.add_run(f" localizada na {rua}, município de {cidade},com a finalidade de servidão de passagem com a seguinte descrição e confrontações, onde os ângulos foram medidos no sentido horário.").font.name = 'Arial'

        doc_word.add_paragraph()
        doc_word.add_paragraph("Pontos definidos pelas Coordenadas Planas no Sistema U.T.M. – SIRGAS 2000.", style='Normal')
        doc_word.add_paragraph()

        # Coordenadas do ponto Az
        ponto_az_1 = f"{ponto_amarracao[0]:.2f}".replace(".", ",")
        ponto_az_2 = f"{ponto_amarracao[1]:.2f}".replace(".", ",")


        azimute_dms = convert_to_dms(azimute)
        distancia_str = f"{distancia_amarracao_v1:.2f}".replace(".", ",")

        # Linha: ponto de amarração
        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

        p.add_run("O ponto ")
        p.add_run("Az").bold = True
        p.add_run(f", ponto de amarração, está localizado na {desc_ponto_amarracao} nas coordenadas E(X) {ponto_az_1} e N(Y) {ponto_az_2}.")

        p.paragraph_format.space_after = Pt(12)  # ⬅️ Força um espaçamento abaixo do parágrafo

        
        # Linha: Azimute e distância até V1
        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p.add_run("Daí, com Azimute de ").bold = False
        p.add_run(azimute_dms).bold = False
        p.add_run(f" e distância de {distancia_str} metros, chega-se ao vértice ")
        p.add_run("V1").bold = True
        p.add_run(", origem da área descrição, alinhado com a rua " + rua + ".")
        p.paragraph_format.space_after = Pt(12)  # ⬅️ FORÇA espaçamento após esse parágrafo



        # ➤ Percorrer vértices
        for i in range(len(df)):
            current = df.iloc[i]
            next_vertex = df.iloc[(i + 1) % len(df)]
            distancia = f"{current['Distancia(m)']:.2f}".replace(".", ",")
            confrontante = current['Confrontante']
            giro_angular = current['Angulo Interno']

            p = doc_word.add_paragraph(style='Normal')
            p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            p.add_run("Do vértice ").bold = False
            p.add_run(current['V']).bold = True

            if i == 0:
                p.add_run(
                    f", com giro angular horário de {giro_angular_v1_dms} e distância de {distancia} metros, "
                    f"confrontando com área pertencente à {confrontante}, chega-se ao vértice "
                )
            elif next_vertex['V'] == "V1" and i == len(df) - 1:
                p.add_run(
                    f", com giro angular horário de {giro_angular} e distância de {distancia} metros, "
                    f"confrontando com área pertencente à {confrontante}, chega-se ao vértice "
                )
                p.add_run(next_vertex['V']).bold = True
                p.add_run(", origem da presente descrição.")
                doc_word.add_paragraph()
                break
            else:
                p.add_run(
                    f", com giro angular horário de {giro_angular} e distância de {distancia} metros, "
                    f"confrontando com área pertencente à {confrontante}, chega-se ao vértice "
                )

            p.add_run(next_vertex['V']).bold = True
            p.add_run(";")
            doc_word.add_paragraph()

        # Parágrafos descritivos
        p = doc_word.add_paragraph(style='Normal')
        p.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        p.add_run(f"Os angulos foram medidos no sentido horário.")

        data_atual = datetime.now().strftime("%d de %B de %Y")
        doc_word.add_paragraph(f"\nPorto Alegre, RS, {data_atual}.", style='Normal')
        doc_word.add_paragraph("\n\n")
        doc_word.save(output_path)

        print(f"Memorial descritivo salvo em: {output_path}")

    except Exception as e:
        print(f"Erro ao criar o documento memorial: {e}")




def find_excel_file(directory, keywords):
    """
    Busca um arquivo Excel no diretório contendo todas as palavras-chave no nome.
    Se não encontrar, exibe a lista de arquivos disponíveis.
    """
    if not directory or not os.path.exists(directory):
        print(f"Erro: O diretório '{directory}' não existe ou não foi especificado corretamente.")
        return None

    excel_files = [file for file in os.listdir(directory) if file.endswith(".xlsx")]

    if not excel_files:
        print(f"Nenhum arquivo Excel encontrado no diretório: {directory}")
        return None

    for file in excel_files:
        if all(keyword.lower() in file.lower() for keyword in keywords):
            return os.path.join(directory, file)

    # Se nenhum arquivo correspondente foi encontrado, listar os arquivos disponíveis
    print(f"Nenhum arquivo Excel contendo {keywords} foi encontrado em '{directory}'.")
    print("Arquivos disponíveis no diretório:")
    for f in excel_files:
        print(f"  - {f}")

    return None



#função não pode ser usada para LINUX       
# def convert_docx_to_pdf(output_path, pdf_file_path):
#     """
#     Converte um arquivo DOCX para PDF usando a biblioteca comtypes.
#     """
#     try:
#         # Verificar se o arquivo DOCX existe antes de abrir
#         if not os.path.exists(output_path):
#             raise FileNotFoundError(f"Arquivo DOCX não encontrado: {output_path}")
        
#         print(f"Tentando converter o arquivo DOCX: {output_path} para PDF: {pdf_file_path}")

#         word = comtypes.client.CreateObject("Word.Application")
#         word.Visible = False  # Ocultar a interface do Word
#         doc = word.Documents.Open(output_path)
        
#         # Aguardar alguns segundos antes de salvar como PDF
#         import time
#         time.sleep(2)

#         #doc.SaveAs(pdf_file_path, FileFormat=17)  # 17 corresponde ao formato PDF
#         doc.Close()
#         word.Quit()
#         #print(f"Arquivo PDF salvo com sucesso em: {pdf_file_path}")
#     except FileNotFoundError as fnf_error:
#         print(f"Erro: {fnf_error}")
#     except Exception as e:
#         print(f"Erro ao converter DOCX para PDF: {e}")
#     finally:
#         try:
#             word.Quit()
#         except:
#             pass  # Garantir que o Word seja fechado



def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '', filename)

        
def main_poligonal_fechada(uuid_str, excel_path, dxf_path, diretorio_preparado, diretorio_concluido, caminho_template):

    # 🔹 Leitura dos dados do Excel
    df_excel = pd.read_excel(excel_path, sheet_name='Dados_do_Imóvel', header=None, engine='openpyxl')
    dados_imovel = dict(zip(df_excel.iloc[:, 0], df_excel.iloc[:, 1]))

    # 🔹 Extração dos campos
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

    caminho_salvar = diretorio_concluido
    os.makedirs(caminho_salvar, exist_ok=True)

    # 🔍 Determina tipo do memorial a partir do nome do arquivo DXF
    # 🔍 Determina tipo do memorial a partir do nome do arquivo DXF
    dxf_filename = os.path.basename(dxf_path).upper()
    if "ETE" in dxf_filename:
        tipo = "ETE"
        sheet_name = "ETE"
    elif "REM" in dxf_filename:
        tipo = "REM"
        sheet_name = "Confrontantes_Remanescente"
    elif "SER" in dxf_filename:
        tipo = "SER"
        sheet_name = "Confrontantes_Servidao"
    elif "ACE" in dxf_filename:
        tipo = "ACE"
        sheet_name = "Confrontantes_Acesso"
    else:
        logger.error("❌ Tipo de memorial (ETE, REM, SER ou ACE) não identificado no nome do DXF.")
        return

    # 🔍 Busca automática de confrontantes (padrão AZIMUTE_AZ)
    padrao_busca = os.path.join(diretorio_preparado, f"{uuid_str}_FECHADA_{tipo}.xlsx")
    arquivos_encontrados = glob.glob(padrao_busca)
    if not arquivos_encontrados:
        logger.error(f"❌ Nenhum arquivo de confrontantes encontrado com o padrão: {padrao_busca}")
        return None

    excel_confrontantes = arquivos_encontrados[0]

    # Agora carrega exatamente a aba correta (conforme o tipo)
    confrontantes_df = pd.read_excel(excel_confrontantes, sheet_name=sheet_name)

    if confrontantes_df.empty:
        logger.error("❌ Planilha de confrontantes está vazia.")
        return None

    
    # 🔹 Limpa DXF
    dxf_limpo_path = os.path.join(caminho_salvar, f"{uuid_str}_DXF_LIMPO_{matricula}.dxf")
    dxf_file_path = limpar_dxf(dxf_path, dxf_limpo_path)

    # 🔍 Extrai geometria do DXF
    doc, lines, perimeter_dxf, area_dxf, ponto_az, _ = get_document_info_from_dxf(dxf_file_path)

    if not doc or not ponto_az:
        logger.error("❌ Erro ao extrair geometria ou encontrar o ponto Az.")
        return

    try:
        doc_dxf = ezdxf.readfile(dxf_limpo_path)
        msp = doc_dxf.modelspace()
    except Exception as e:
        logger.error(f"❌ Erro ao abrir o DXF com ezdxf: {e}")
        return

    if doc and lines:
        logger.info(f"📐 Área da poligonal: {area_dxf:.6f} m²")

        v1 = lines[0][0]
        v2 = lines[1][0]
        azimute = calculate_azimuth(ponto_az, v1)
        distancia_az_v1 = calculate_distance(ponto_az, v1)
        giro_angular_v1 = calculate_angular_turn(ponto_az, v1, v2)
        giro_angular_v1_dms = convert_to_dms(360 - giro_angular_v1)

        excel_file_path = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_Memorial_{matricula}.xlsx")


        # ✅ Geração do Excel e atualização do DXF
        excel_resultado = create_memorial_descritivo(
            uuid_str=uuid_str,
            doc=doc,
            msp=msp,
            lines=lines,
            proprietario=proprietario,
            matricula=matricula,
            caminho_salvar=caminho_salvar,
            excel_file_path=arquivos_encontrados[0],
            ponto_az=ponto_az,
            distance_az_v1=distancia_az_v1,
            azimute_az_v1=azimute,
            tipo=tipo,
            diretorio_concluido=caminho_salvar
        )


        # ✅ Geração do DOCX
        if excel_resultado:
            output_path_docx = os.path.join(caminho_salvar, f"{uuid_str}_FECHADA_{tipo}_Memorial_{matricula}.docx")
            #assinatura_path = r"C:\Users\Paulo\Documents\CASSINHA\MEMORIAIS DESCRITIVOS\Assinatura.jpg"

            create_memorial_document(
                uuid_str, proprietario, matricula, descricao, excel_file_path, template_path, 
                output_path_docx, perimeter_dxf, area_dxf, desc_ponto_Az, Coorde_E_ponto_Az, Coorde_N_ponto_Az,
                azimuth, distance, uso_solo, area_imovel, cidade, rua, comarca, RI, caminho_salvar, tipo
            )




            # ✅ Geração do PDF
            # time.sleep(2)
            # if os.path.exists(output_path_docx):
            #     pdf_file_path = os.path.join(caminho_salvar, f"FECHADA_{tipo}_Memorial_{matricula}.pdf")
            #     convert_docx_to_pdf(output_path_docx, pdf_file_path)
            #     logger.info(f"✅ PDF salvo em: {pdf_file_path}")
            # else:
            #     logger.info("❌ Arquivo DOCX não gerado para conversão.")

        else:
            logger.error("❌ Planilha Excel não gerada.")

    else:
        logger.error("❌ Não foi possível processar a geometria.")





