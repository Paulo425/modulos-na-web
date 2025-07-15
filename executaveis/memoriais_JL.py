import os
import math
import traceback
from datetime import datetime
from decimal import Decimal, getcontext
import locale
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document
import ezdxf
from shapely.geometry import Polygon
try:
    from ezdxf.math import Vec3 as Vector
except ImportError:
    from ezdxf.math import Vector

getcontext().prec = 28  # Define a precisão para 28 casas decimais

# Força o locale para português em sistemas Windows ou Linux

try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')  # Linux/Mac padrão
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')  # Windows
    except locale.Error as e:
        print(f"⚠️ Locale não pôde ser definido: {e}")
        locale.setlocale(locale.LC_TIME, '')  # fallback para padrão

# Agora sim a data será em português
data_atual = datetime.now().strftime("%d de %B de %Y")


def obter_data_em_portugues():
    meses_pt = {
        "January": "janeiro", "February": "fevereiro", "March": "março",
        "April": "abril", "May": "maio", "June": "junho",
        "July": "julho", "August": "agosto", "September": "setembro",
        "October": "outubro", "November": "novembro", "December": "dezembro"
    }
    data = datetime.now()
    mes_en = data.strftime("%B")
    mes_pt = meses_pt.get(mes_en, mes_en)
    return f"{data.day:02d} de {mes_pt} de {data.year}"


import ezdxf
import math

def limpar_dxf_preservando_original(dxf_original, dxf_saida, log=None):
    import ezdxf

    doc = ezdxf.readfile(dxf_original)
    msp = doc.modelspace()

    novo_doc = ezdxf.new(dxfversion='R2010')
    novo_msp = novo_doc.modelspace()

    ponto_inicial_real = None

    for entity in msp.query('LWPOLYLINE'):
        if entity.closed:
            pontos_bulge = [(p[0], p[1], p[4]) for p in entity.get_points('xyseb')]

            # Captura o ponto inicial real da polilinha
            ponto_inicial_real = (pontos_bulge[0][0], pontos_bulge[0][1])

            novo_msp.add_lwpolyline(
                pontos_bulge, 
                format='xyb', 
                close=True, 
                dxfattribs={'layer': 'LAYOUT_MEMORIAL'}
            )

            if log:
                log.write("✅ Polilinha original preservada com bulges.\n")
                log.write(f"✅ Ponto inicial capturado: {ponto_inicial_real}\n")
            break

    novo_doc.saveas(dxf_saida)

    if log:
        log.write(f"✅ DXF salvo: {dxf_saida}\n")

    # Aqui está o retorno correto com 3 valores:
    return dxf_saida, ponto_inicial_real, ponto_inicial_real




def calculate_signed_area(points):
    area = 0
    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        area += (x1 * y2) - (x2 * y1)
    return area / 2


def calculate_area_with_arcs(points, arcs):
    """
    Calcula a área de uma poligonal que inclui segmentos de linha reta e arcos.

    :param points: Lista de tuplas (x, y) representando os vértices dos segmentos de linha reta.
    :param arcs: Lista de dicionários, cada um contendo informações sobre um arco:
                 {'start_point': (x1, y1), 'end_point': (x2, y2), 'center': (xc, yc),
                  'radius': r, 'start_angle': a1, 'end_angle': a2, 'length': l}
    :return: Área total da poligonal.
    """
    # Calcular a área usando a fórmula do polígono (Shoelace) para os segmentos de linha reta
    num_points = len(points)
    area_linear = 0.0
    for i in range(num_points):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % num_points]
        area_linear += x1 * y2 - x2 * y1
    area_linear = abs(area_linear) / 2.0

    # Calcular a área dos arcos
    area_arcs = 0.0
    for arc in arcs:
#         start_point = Vec2(arc['start_point'])
#         end_point = Vec2(arc['end_point'])
#         center = Vec2(arc['center'])
       
        start_point = (float(arc['start_point'][0]), float(arc['start_point'][1]))
        end_point = (float(arc['end_point'][0]), float(arc['end_point'][1]))
        center = (float(arc['center'][0]), float(arc['center'][1]))

        radius = arc['radius']
        start_angle = math.radians(arc['start_angle'])
        end_angle = math.radians(arc['end_angle'])

        # Determinar a variação do ângulo
        delta_angle = end_angle - start_angle
        if delta_angle <= 0:
            delta_angle += 2 * math.pi

        # Área do setor circular
        sector_area = 0.5 * radius**2 * delta_angle

        # Área do triângulo formado pelo centro e os pontos inicial e final do arco
#         triangle_area = 0.5 * abs((start_point.x - center.x) * (end_point.y - center.y) -
#                                   (end_point.x - center.x) * (start_point.y - center.y))
        triangle_area = 0.5 * abs((start_point[0] - center[0]) * (end_point[1] - center[1]) -
                                  (end_point[0] - center[0]) * (start_point[1] - center[1]))

        # Área do segmento circular
        segment_area = sector_area - triangle_area

        # Determinar a orientação do arco para adicionar ou subtrair a área
        # Aqui assumimos que a orientação positiva indica um arco no sentido anti-horário
        # e negativa para o sentido horário. Isso pode precisar de ajustes conforme a definição dos dados.
        if delta_angle > math.pi:
            area_arcs -= segment_area  # Arco no sentido horário
        else:
            area_arcs += segment_area  # Arco no sentido anti-horário

    # Área total
    area_total = area_linear + area_arcs
    return area_total

def bulge_to_arc_length(start_point, end_point, bulge):
    #chord_length = (end_point - start_point).magnitude
    dx = end_point[0] - start_point[0]
    dy = end_point[1] - start_point[1]
    chord_length = math.hypot(dx, dy)

    sagitta = (bulge * chord_length) / 2
    radius = ((chord_length / 2) ** 2 + sagitta ** 2) / (2 * abs(sagitta))
    angle = 4 * math.atan(abs(bulge))
    arc_length = radius * angle
    return arc_length, radius, angle

import math
from shapely.geometry import Polygon

def get_document_info_from_dxf(dxf_file_path, log=None):
    try:
        if log is None:
            class DummyLog:
                def write(self, msg): pass
            log = DummyLog()
            
        doc = ezdxf.readfile(dxf_file_path)
        msp = doc.modelspace()

        lines = []
        arcs = []
        perimeter_dxf = 0
        ponto_az = None

        for entity in msp.query('LWPOLYLINE'):
            if entity.is_closed:
                polyline_points = entity.get_points('xyseb')
                num_points = len(polyline_points)

                boundary_points = []

                for i in range(num_points):
                    x_start, y_start, _, _, bulge = polyline_points[i]
                    x_end, y_end, _, _, _ = polyline_points[(i + 1) % num_points]

                    #start_point = Vec2(float(x_start), float(y_start))
                    start_point = (float(x_start), float(y_start))

                    #end_point = Vec2(float(x_end), float(y_end))
                    end_point = (float(x_end), float(y_end))

                    if bulge != 0:
                        # Trata-se de arco
                       # chord_length = (end_point - start_point).magnitude
                        dx = end_point[0] - start_point[0]
                        dy = end_point[1] - start_point[1]
                        chord_length = math.hypot(dx, dy)
                        sagitta = (bulge * chord_length) / 2
                        radius = ((chord_length / 2)**2 + sagitta**2) / (2 * abs(sagitta))
                        angle_span_rad = 4 * math.atan(abs(bulge))
                        arc_length = radius * angle_span_rad

                        #chord_midpoint = (start_point + end_point) / 2
                        mid_x = (start_point[0] + end_point[0]) / 2
                        mid_y = (start_point[1] + end_point[1]) / 2
                        chord_midpoint = (mid_x, mid_y)

                        offset_dist = math.sqrt(radius**2 - (chord_length / 2)**2)
#                         dx = end_point[0] - start_point[0]
#                         dy = end_point[1] - start_point[1]
                        dx = float(end_point[0]) - float(start_point[0])
                        dy = float(end_point[1]) - float(start_point[1])

                        #perp_vector = Vec2(-dy, dx).normalize()
                        length = math.hypot(dx, dy)
                        perp_vector = (-dy / length, dx / length)

                        if bulge < 0:
                            #perp_vector = -perp_vector
                            perp_vector = (-perp_vector[0], -perp_vector[1])


                        #center = chord_midpoint + perp_vector * offset_dist
                        center_x = chord_midpoint[0] + perp_vector[0] * offset_dist
                        center_y = chord_midpoint[1] + perp_vector[1] * offset_dist
                        center = (center_x, center_y)


                        #start_angle = math.atan2(start_point.y - center.y, start_point[0] - center[0])
                        start_angle = math.atan2(start_point[1] - center[1], start_point[0] - center[0])

                        end_angle = start_angle + (angle_span_rad if bulge > 0 else -angle_span_rad)

                        arcs.append({
                            'start_point': (start_point[0], start_point[1]),
                            'end_point': (end_point[0], end_point[1]),
                            'start_angle': math.degrees(start_angle),
                            'end_angle': math.degrees(end_angle),
                            'length': arc_length,
                            'bulge': bulge
                        })

                        # Pontos intermediários no arco (para precisão da área)
                        num_arc_points = 100  # mais pontos para maior precisão
                        for t in range(num_arc_points):
                            angle = start_angle + (end_angle - start_angle) * t / num_arc_points
                            arc_x = center[0] + radius * math.cos(angle)
                            arc_y = center[1] + radius * math.sin(angle)
                            boundary_points.append((arc_x, arc_y))

                        segment_length = arc_length
                        perimeter_dxf += segment_length
                    else:
                        # Linha reta
                        lines.append((start_point, end_point))
                        boundary_points.append((start_point[0], start_point[1]))
                       # segment_length = (end_point - start_point).magnitude
                        dx = end_point[0] - start_point[0]
                        dy = end_point[1] - start_point[1]
                        segment_length = math.hypot(dx, dy)

                        perimeter_dxf += segment_length

                # Após loop, calcular a área com Shapely
                polygon = Polygon(boundary_points)
                area_dxf = polygon.area  # área exata do desenho

                break

        if not lines and not arcs:
            print("Nenhuma polilinha fechada encontrada no arquivo DXF.")
            return None, [], [], 0, 0, None

        print(f"Linhas processadas: {len(lines)}")
        print(f"Arcos processados: {len(arcs)}")
        print(f"Perímetro do DXF: {perimeter_dxf:.2f} metros")
        print(f"Área do DXF: {area_dxf:.2f} metros quadrados")
#         print(f"Ponto Az: {ponto_az}")

        return doc, lines, arcs, perimeter_dxf, area_dxf, boundary_points

    except Exception as e:
        print(f"Erro ao obter informações do documento: {e}")
        traceback.print_exc()
        return None, [], [], 0, 0, None

# 🔹 Função para definir a fonte padrão
def set_default_font(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)

def add_arc_labels(doc, msp, start_point, end_point, radius, length, label, log=None):

    if log is None:
        class DummyLog:
            def write(self, msg): pass
        log = DummyLog()
    try:
        #mid_point = Vec2((start_point[0] + end_point[0])/2, (start_point[1] + end_point[1])/2)
        mid_point = ((float(start_point[0]) + float(end_point[0]))/2, (float(start_point[1]) + float(end_point[1]))/2)

        label_radius = f"R={radius:.2f}".replace('.', ',')
        label_length = f"C={length:.2f}".replace('.', ',')

        msp.add_text(
            label_radius,
            dxfattribs={
                "height": 1.0,
                "layer": "LAYOUT_DISTANCIAS",
                "insert": (mid_point[0], mid_point[1])
            }
        )

        msp.add_text(
            label_length,
            dxfattribs={
                "height": 1.0,
                "layer": "LAYOUT_DISTANCIAS",
                "insert": (mid_point[0], mid_point[1] - 2)
            }
        )

        print(f"✅ Rótulos {label_radius} e {label_length} adicionados corretamente no DXF.")
        if log:
            log.write(f"✅ Rótulos {label_radius} e {label_length} adicionados corretamente no DXF.\n")

    except Exception as e:
        print(f"❌ Erro ao adicionar rótulos dos arcos: {e}")
        if log:
            log.write(f"❌ Erro ao adicionar rótulos dos arcos: {e}\n")


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


def add_azimuth_arc(doc, msp, ponto_az, v1, azimuth, log=None):
    """
    Adiciona o arco do azimute no ModelSpace.
    """
    if log is None:
        class DummyLog:
            def write(self, msg): pass
        log = DummyLog()

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
        if log:
            log.write(f"Rótulo do azimute ({azimuth_label}) adicionado com sucesso em {label_position}\n")

    except Exception as e:
        print(f"Erro ao adicionar arco do azimute: {e}")
        if log:
            log.write(f"Erro ao adicionar arco do azimute: {e}\n")


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
    return area / 2.0


def add_label_and_distance(doc, msp, start_point, end_point, label, distance,log=None):
    if log is None:
        class DummyLog:
            def write(self, msg): pass
        log = DummyLog()

    try:
        # Garantir pontos como tuplas de float
        start_point = (float(start_point[0]), float(start_point[1]))
        end_point = (float(end_point[0]), float(end_point[1]))

        # Criar camadas se ainda não existirem
        layers = [
            ("LAYOUT_VERTICES", 2),   # Amarelo
            ("LAYOUT_DISTANCIAS", 4)  # Azul
        ]
        for layer_name, color in layers:
            if layer_name not in doc.layers:
                doc.layers.new(name=layer_name, dxfattribs={"color": color})

        # Adicionar círculo no ponto inicial
        msp.add_circle(center=start_point, radius=0.5, dxfattribs={'layer': 'LAYOUT_VERTICES'})

        # Adicionar rótulo do vértice (ex: V1, V2...)
        ofset_x = 0.3
        ofset_y = 0.3
        msp.add_text(
            label,
            dxfattribs={
                'height': 0.5,
                'layer': 'LAYOUT_VERTICES',
                'insert': (start_point[0] + ofset_x, start_point[1] + ofset_y)
            }
        )

        # Calcular ponto médio e vetor do segmento
        mid_x = (start_point[0] + end_point[0]) / 2
        mid_y = (start_point[1] + end_point[1]) / 2
        dx = end_point[0] - start_point[0]
        dy = end_point[1] - start_point[1]
        length = (dx ** 2 + dy ** 2) ** 0.5

        if length == 0:
            return  # evita divisão por zero

        # Cálculo do ângulo em graus
        angle = math.degrees(math.atan2(dy, dx))

        # Corrige para manter a leitura em pé
        if angle < -90 or angle > 90:
            angle += 180

        # Deslocar o rótulo perpendicularmente ao segmento
        offset = 0.5  # pode ajustar para mais ou menos afastamento

        angle_rad = math.atan2(dy, dx)
        angle_deg = math.degrees(angle_rad)

        # Define o lado correto baseado no ângulo do segmento
        if -90 <= angle_deg <= 90:
            side_factor = 1  # acima da linha
        else:
            side_factor = -1  # abaixo da linha

        perp_x = -dy / length * offset
        perp_y = dx / length * offset
        mid_point_displaced = (mid_x + perp_x, mid_y + perp_y)

        # Formatar distância
        distancia_formatada = f"{distance:.2f} m".replace('.', ',')

        # Inserir texto da distância rotacionado
        msp.add_text(
            distancia_formatada,
            dxfattribs={
                'height': 0.5,
                'layer': 'LAYOUT_DISTANCIAS',
                'rotation': angle,
                'insert': mid_point_displaced
            }
        )

        print(f"✅ DEBUG: '{label}' e distância '{distancia_formatada}' inseridos em {start_point} e {mid_point_displaced} com ângulo {angle:.2f}°")
        if log:
            log.write(f"✅ DEBUG: '{label}' e distância '{distancia_formatada}' inseridos em {start_point} e {mid_point_displaced} com ângulo {angle:.2f}°\n")

    except Exception as e:
        print(f"❌ ERRO GRAVE ao adicionar rótulo '{label}' e distância: {e}")
        print(f"❌ ERRO GRAVE ao adicionar rótulo '{label}' e distância: {e}")











def sanitize_filename(filename):
    # Substitui os caracteres inválidos por um caractere válido (ex: espaço ou underline)
    sanitized_filename = re.sub(r'[\\/*?:"<>|]', "_", filename)  # Substitui caracteres inválidos por "_"
    return sanitized_filename
        
        


# Função para criar memorial descritivo
def create_memorial_descritivo(doc, msp, lines, proprietario, matricula, caminho_salvar, arcs=None,
                               excel_file_path=None, ponto_az=None, distance_az_v1=None,
                               azimute_az_v1=None, ponto_inicial_real=None,  # ✅ Adicionado aqui
                               encoding='ISO-8859-1', boundary_points=None, log=None):

    """
    Cria o memorial descritivo diretamente no arquivo DXF e salva os dados em uma planilha Excel.
    """
    if log is None:
        class DummyLog:
            def write(self, msg): pass
        log = DummyLog()

    if excel_file_path:
        try:
            confrontantes_df = pd.read_excel(excel_file_path)
            confrontantes_dict = dict(zip(confrontantes_df['Código'], confrontantes_df['Confrontante']))
        except Exception as e:
            print(f"Erro ao carregar arquivo de confrontantes: {e}")
            if log:
                log.write(f"Erro ao carregar arquivo de confrontantes: {e}\n")
            confrontantes_dict = {}
    else:
        confrontantes_dict = {}

    if not lines:
        print("Nenhuma linha disponível para criar o memorial descritivo.")
        if log:
            log.write("Nenhuma linha disponível para criar o memorial descritivo.\n")
        return None

    

    
    # Função adicional para verificar sentido horário do arco:
    def is_arc_clockwise(start_pt, end_pt, center):
        start_angle = math.atan2(start_pt[1] - center[1], start_pt[0] - center[0])
        end_angle = math.atan2(end_pt[1] - center[1], end_pt[0] - center[0])
        angle_diff = (end_angle - start_angle) % (2 * math.pi)
        return angle_diff > math.pi
      
    # Sequenciar corretamente os segmentos mantendo coerência geométrica
    # Ao criar elementos, armazene também o bulge original:
    elementos = []

    for line in lines:
        elementos.append(('line', {
            'start_point': line[0],
            'end_point': line[1],
            'bulge': 0
        }))

    if arcs:
        for arc in arcs:
            start_point = arc['start_point']
            end_point = arc['end_point']
            bulge = arc['bulge']

            dx, dy = end_point[0] - start_point[0], end_point[1] - start_point[1]
            chord_length = math.hypot(dx, dy)
            sagitta = (bulge * chord_length) / 2
            radius = ((chord_length / 2)**2 + sagitta**2) / (2 * abs(sagitta))
            angle_span_rad = 4 * math.atan(abs(bulge))
            arc_length = radius * angle_span_rad

            elementos.append(('arc', {
                'start_point': start_point,
                'end_point': end_point,
                'bulge': bulge,          # ✅ Usado APENAS para o DXF
                'radius': radius,        # ✅ Usado APENAS para o Excel
                'length': arc_length     # ✅ Usado APENAS para o Excel
            }))





    # Ajuste definitivo para ordenar corretamente elementos pelo ponto inicial real:
    if ponto_inicial_real:
        index_inicio = None
        for i, elemento in enumerate(elementos):
            pt = elemento[1]['start_point']  # ✅ acesso correto!
            if math.hypot(pt[0] - ponto_inicial_real[0], pt[1] - ponto_inicial_real[1]) < 1e-6:
                index_inicio = i
                break

        if index_inicio is not None:
            elementos = elementos[index_inicio:] + elementos[:index_inicio]

    boundary_points_com_bulge = []  # <-- ✅ ADICIONAR ESTA LINHA AQUI
    # Sequenciar corretamente os segmentos mantendo coerência geométrica
    sequencia_completa = []
    ponto_atual = elementos[0][1]['start_point']  # ✅ Corrigido aqui!

    while elementos:
        for i, elemento in enumerate(elementos):
            tipo, dados = elemento
            start_point, end_point = dados['start_point'], dados['end_point']  # ✅ Corrigido aqui!

            if ponto_atual == start_point:
                sequencia_completa.append(elemento)
                ponto_atual = end_point
                elementos.pop(i)
                break
            elif ponto_atual == end_point:
                # Inverter direção e ajustar bulge corretamente
                if tipo == 'line':
                    elementos[i] = ('line', {
                        'start_point': end_point,
                        'end_point': start_point
                    })
                else:  # tipo == 'arc'
                    elementos[i] = ('arc', {
                        'start_point': end_point,
                        'end_point': start_point,
                        'center': dados['center'],
                        'radius': dados['radius'],
                        'length': dados['length'],
                        'bulge': -dados['bulge']  # Inverte o sinal do bulge corretamente
                    })
                sequencia_completa.append(elementos[i])
                ponto_atual = start_point
                elementos.pop(i)
                break

        else:
            if elementos:
                ponto_atual = elementos[0][1]['start_point']

    # ✅ ADICIONAR ESTE BLOCO COMPLETO LOGO APÓS O LOOP ACIMA TERMINAR
    for idx, (tipo, dados) in enumerate(sequencia_completa):
        bulge = dados['bulge'] if tipo == 'arc' else 0
        boundary_points_com_bulge.append((dados['start_point'][0], dados['start_point'][1], bulge))

    ultimo_ponto = sequencia_completa[-1][1]['end_point']
    boundary_points_com_bulge.append((ultimo_ponto[0], ultimo_ponto[1], 0))
    

    # Calcula a área da poligonal para verificar se precisa inverter
    pontos_para_area = [seg[1]['start_point'] for seg in sequencia_completa]
    pontos_para_area.append(sequencia_completa[-1][1]['end_point'])  # Fecha o polígono corretamente

    simple_ordered_points = [(float(pt[0]), float(pt[1])) for pt in pontos_para_area]
    area = calculate_signed_area(simple_ordered_points)

    # Agora inverter o sentido APENAS se necessário
    if area > 0:
        sequencia_completa.reverse()
        sequencia_corrigida = []

        for tipo, dados in sequencia_completa:
            start, end = dados['start_point'], dados['end_point']  # ✅ Correção aqui!

            if tipo == 'arc':
                sequencia_corrigida.append(('arc', {
                    'start_point': end,
                    'end_point': start,
                    'center': dados['center'],
                    'radius': dados['radius'],
                    'length': dados['length'],
                    'bulge': -dados['bulge']  # ✅ Correção no sinal do bulge
                }))
            else:  # tipo == 'line'
                sequencia_corrigida.append(('line', {
                    'start_point': end,
                    'end_point': start
                }))

        sequencia_completa = sequencia_corrigida
        area = abs(area)


    print(f"Área da poligonal ajustada: {area:.4f} m²")
    if log:
        log.write(f"Área da poligonal ajustada: {area:.4f} m²\n")

       
    # # Fecha corretamente adicionando o último ponto sem bulge
    # ultimo_ponto = sequencia_completa[-1][1]['end_point']
    # boundary_points_com_bulge.append((ultimo_ponto[0], ultimo_ponto[1], 0))


    # ✅ BLOCO DIRETO E CORRIGIDO PARA SUBSTITUIÇÃO

    # Exclui polilinhas antigas no DXF antes de inserir a nova
    for entity in msp.query('LWPOLYLINE[layer=="LAYOUT_MEMORIAL"]'):
        msp.delete_entity(entity)

    # Cria dados para Excel usando sequencia_completa (correto e definitivo)
    dados = []
    num_vertices = len(sequencia_completa)

    boundary_points_com_bulge = []

    for idx, (tipo, dados_seg) in enumerate(sequencia_completa):
        start_point = dados_seg['start_point']
        end_point = dados_seg['end_point']

        if tipo == "line":
            azimuth, distance = calculate_azimuth_and_distance(start_point, end_point)
            azimute_excel = convert_to_dms(azimuth)
            distancia_excel = f"{distance:.2f}".replace(".", ",")
            bulge = 0

        elif tipo == "arc":
            radius = dados_seg['radius']
            distance = dados_seg['length']
            azimute_excel = f"R={radius:.3f}".replace(".", ",")
            distancia_excel = f"C={distance:.3f}".replace(".", ",")
            bulge = dados_seg['bulge']

        label = f"P{idx + 1}"
        divisa = f"P{idx + 1}_P{idx + 2}" if idx + 1 < num_vertices else f"P{idx + 1}_P1"
        confrontante = confrontantes_dict.get(f"V{idx + 1}", "Desconhecido")

        dados.append({
            "V": label,
            "E": f"{start_point[0]:.3f}".replace('.', ','),
            "N": f"{start_point[1]:.3f}".replace('.', ','),
            "Z": "0,000",
            "Divisa": divisa,
            "Azimute": azimute_excel,
            "Distancia(m)": distancia_excel,
            "Confrontante": confrontante,
        })

        boundary_points_com_bulge.append((float(start_point[0]), float(start_point[1]), float(bulge)))

    # Adiciona o último ponto para fechar corretamente a poligonal
    ultimo_ponto = sequencia_completa[-1][1]['end_point']
    boundary_points_com_bulge.append((float(ultimo_ponto[0]), float(ultimo_ponto[1]), 0))

    # Inserção definitiva da polilinha com arcos no DXF
    msp.add_lwpolyline(
        boundary_points_com_bulge,
        format='xyb',
        close=True,
        dxfattribs={"layer": "LAYOUT_MEMORIAL"}
    )

    # Inserção dos rótulos corretamente no DXF
    for idx, (tipo, dados_seg) in enumerate(sequencia_completa):
        start_point = dados_seg['start_point']
        end_point = dados_seg['end_point']
        distance = dados_seg['length'] if tipo == 'arc' else calculate_azimuth_and_distance(start_point, end_point)[1]
        label = f"P{idx + 1}"

        add_label_and_distance(doc, msp, start_point, end_point, label, distance, log=log)

    # Agora salva Excel normalmente
    df = pd.DataFrame(dados, dtype=str)
    excel_output_path = os.path.join(caminho_salvar, f"Memorial_{matricula}.xlsx")
    df.to_excel(excel_output_path, index=False)

    wb = openpyxl.load_workbook(excel_output_path)
    ws = wb.active

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    column_widths = {
        "A": 8, "B": 15, "C": 15, "D": 10, "E": 20, "F": 15,
        "G": 15, "H": 30, "I": 20, "J": 20, "K": 15, "L": 15
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(excel_output_path)

    print(f"Arquivo Excel salvo e formatado em: {excel_output_path}")
    if log:
        log.write(f"Arquivo Excel salvo e formatado em: {excel_output_path}\n")


    # Agora salva o arquivo DXF final com os dados atualizados
    try:
        dxf_output_path = os.path.join(caminho_salvar, f"Memorial_{matricula}.dxf")
        if log:
            log.write("Boundary points com bulge (final para DXF):\n")
            for idx, pt in enumerate(boundary_points_com_bulge, 1):
                log.write(f"{idx}: Coordenadas={pt[:2]}, Bulge={pt[2]}\n")
        doc.saveas(dxf_output_path)
        print(f"Arquivo DXF salvo em: {dxf_output_path}")
        if log:
            log.write(f"Arquivo DXF salvo em: {dxf_output_path}\n")
    except Exception as e:
        erro_msg = f"Erro ao salvar DXF: {e}"
        print(erro_msg)
        if log:
            log.write(erro_msg + "\n")

    return excel_output_path







def create_memorial_document(
    proprietario, matricula, descricao, excel_file_path=None, template_path=None, output_path=None,
    perimeter_dxf=None, area_dxf=None, desc_ponto_Az=None, Coorde_E_ponto_Az=None, Coorde_N_ponto_Az=None,
    azimuth=None, distance=None, log=None
):
    if log is None:
        class DummyLog:
            def write(self, msg): pass
        log = DummyLog()

    try:
        # 🔍 Verificação do template
        print(f"🔎 Caminho do template: {template_path}")
        if log:
            log.write(f"🔎 Template path: {template_path}\n")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"❌ Template não encontrado: {template_path}")

        # 🔍 Verificação do diretório de saída
        output_dir = os.path.dirname(output_path)
        print(f"📁 Caminho de saída do DOCX: {output_path}")
        if log:
            log.write(f"📁 Output DOCX path: {output_path}\n")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        # Função para pular linhas
        def pular_linhas(doc, n_linhas):
            for _ in range(n_linhas):
                doc.add_paragraph("")

        

        # Ler o arquivo Excel se for informado
        if excel_file_path:
            df = pd.read_excel(excel_file_path)
        else:
            df = None

        # Criar o documento Word carregando o template
        doc_word = Document(template_path)
        set_default_font(doc_word)  # Fonte Arial 12
        # Criar o documento Word carregando o template
        doc_word = Document(template_path)
        set_default_font(doc_word)  # Fonte Arial 12

        # Adiciona o preâmbulo centralizado com a variável "descricao"
        p1 = doc_word.add_paragraph(style='Normal')
        p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run1 = p1.add_run("MEMORIAL DESCRITIVO INDIVIDUAL")
        run1.bold = True

        p2 = doc_word.add_paragraph(f"({descricao})", style='Normal')
        p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        doc_word.add_paragraph("")  # Espaço em branco

        # Dados da obra (conteúdo fixo, alinhado normalmente à esquerda)
        dados_obra = [
            "DADOS DA OBRA",
            "Rodovia: BR-277/PR",
            "Trecho: Ponte s/ Rio Emboguaçu – Ponte Internacional Brasil / Paraguai (2ª Ponte)",
            "Sub-trecho: Entr. BR-277(km 722,6) (Acesso a 2ª Ponte Rio Paraná) – Inicio Ponte Internacional Brasil / Paraguai (2ª Ponte)",
            "Segmento: km 00,00 a km 15,00",
            "Lote: Único",
            ""
        ]

        for linha in dados_obra:
            if ":" in linha:
                negrito, restante = linha.split(":", 1)
                par = doc_word.add_paragraph(style='Normal')
                run_bold = par.add_run(negrito + ":")
                run_bold.bold = True
                par.add_run(restante)  # texto normal
            else:
                par = doc_word.add_paragraph(style='Normal')
                run = par.add_run(linha)
                if linha.strip() == "DADOS DA OBRA":
                    run.bold = True

        # Adicionar dados do proprietário
        par1 = doc_word.add_paragraph(style='Normal')
        run1 = par1.add_run("NOME PROPRIETÁRIO / OCUPANTE:")
        run1.bold = True
        par1.add_run(f" {proprietario}")

        par2 = doc_word.add_paragraph(style='Normal')
        run2 = par2.add_run("DOCUMENTAÇÃO:")
        run2.bold = True
        par2.add_run(f" MATRÍCULA {matricula}")

        # ÁREA DO IMÓVEL (m²)
        par3 = doc_word.add_paragraph(style='Normal')
        run3a = par3.add_run("ÁREA DO IMÓVEL (m")
        run3a.bold = True
        run3b = par3.add_run("2")
        run3b.bold = True
        run3b.font.superscript = True
        run3c = par3.add_run("):")
        run3c.bold = True
        area_formatada = f"{area_dxf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        par3.add_run(f" {area_formatada}")

        doc_word.add_paragraph("")  # Uma linha em branco para separar


        # Descrição do perímetro, somente se o arquivo Excel foi fornecido
        if df is not None:
            initial = df.iloc[0]
            par_intro = doc_word.add_paragraph(style='Normal')
            par_intro.add_run("Pontos definidos pelas Coordenadas Planas no Sistema U.T.M. – SIRGAS 2000.\n\n")
            par_intro.add_run("Inicia-se a descrição deste perímetro no vértice ")
            
            run_v = par_intro.add_run(str(initial['V']))
            run_v.bold = True
            
            par_intro.add_run(f", de coordenadas N(Y) {initial['N']} e E(X) {initial['E']}, situado no limite com {initial['Confrontante']}.")



            num_points = len(df)
            for i in range(num_points):
                current = df.iloc[i]
                next_index = (i + 1) % num_points
                next_point = df.iloc[next_index]

                azimute = current['Azimute']
                distancia = current['Distancia(m)']
                confrontante = current['Confrontante']
                destino = next_point['V']
                coord_n = next_point['N']
                coord_e = next_point['E']

                if azimute.startswith("R=") and distancia.startswith("C="):
                    par = doc_word.add_paragraph(style='Normal')
                    par.add_run(f"Deste, segue com raio de {azimute[2:]}m e distância de {distancia[2:]}m, ")
                    par.add_run(f"confrontando neste trecho com {confrontante}, até o vértice ")
                    run_destino = par.add_run(str(destino))
                    run_destino.bold = True
                    par.add_run(f", de coordenadas N(Y) {coord_n} e E(X) {coord_e};")

                else:
                    par = doc_word.add_paragraph(style='Normal')
                    par.add_run(f"Deste, segue com azimute de {azimute} e distância de {distancia} m, ")
                    par.add_run(f"confrontando neste trecho com {confrontante}, até o vértice ")
                    run_destino = par.add_run(str(destino))
                    run_destino.bold = True
                    par.add_run(f", de coordenadas N(Y) {coord_n} e E(X) {coord_e};")

        else:
            # Caso não haja Excel, pode deixar espaço para preenchimento manual
            doc_word.add_paragraph("Descrição do perímetro não incluída neste memorial.", style='Normal')
            pular_linhas(doc_word, 8)

        doc_word.add_paragraph("")  # Uma linha em branco para separar
   
        # Adicionar o fechamento do perímetro e área
        # Formata perímetro e área com separador de milhar (ponto) e decimal (vírgula)
        perimetro_formatado = f"{perimeter_dxf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        area_formatada = f"{area_dxf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        paragrafo_fechamento = doc_word.add_paragraph(
            f"Fechando-se assim o perímetro com {perimetro_formatado} m "
            f"e a área com {area_formatada} m².",
            style='Normal'
        )
        paragrafo_fechamento.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        doc_word.add_paragraph("")
        doc_word.add_paragraph("")
        # Adicionar data
        
        data_atual = obter_data_em_portugues()
        
       # Centralizar data
        paragrafo_data = doc_word.add_paragraph(f"Paraná, {data_atual}.", style='Normal')
        paragrafo_data.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        doc_word.add_paragraph("")  # Uma linha em branco para separar
        doc_word.add_paragraph("")  # Uma linha em branco para separar
        doc_word.add_paragraph("")  # Uma linha em branco para separar

        # Adicionar a imagem da assinatura centralizada
        assinatura = doc_word.add_paragraph()
        assinatura.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        #run = assinatura.add_run()
        #run.add_picture(r"C:\Users\Paulo\OneDrive\Documentos\JL_ADICIONAIS\TEMPLATE_MEMORIAL\assinatura_engenheiro.jpg", width=Inches(2.0))

        # Adicionar informações do engenheiro centralizadas
        infos_engenheiro = [
            "____________________",
            "Rodrigo Luis Schmitz",
            "Técnico em Agrimensura",
            "CFT: 045.300.139-44"
        ]

        for info in infos_engenheiro:
            paragrafo = doc_word.add_paragraph(info, style='Normal')
            paragrafo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


        # Salvar o documento
        doc_word.save(output_path)
        print(f"Memorial descritivo salvo em: {output_path}")
        if log:
            log.write(f"Memorial descritivo salvo em: {output_path}\n")

    except Exception as e:
        print(f"Erro ao criar o documento memorial: {e}")
        if log:
            log.write(f"Erro ao criar o documento memorial: {e}\n")


        
def convert_docx_to_pdf(output_path, pdf_file_path):
    """
    Converte um arquivo DOCX para PDF usando a biblioteca comtypes.
    """
    try:
        # Verificar se o arquivo DOCX existe antes de abrir
        if not os.path.exists(output_path):
            raise FileNotFoundError(f"Arquivo DOCX não encontrado: {output_path}")
           
        
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # Ocultar a interface do Word
        doc = word.Documents.Open(output_path)
        doc.SaveAs(pdf_file_path, FileFormat=17)  # 17 corresponde ao formato PDF
        doc.Close()
        word.Quit()
        print(f"Arquivo PDF salvo em: {pdf_file_path}")
    except FileNotFoundError as fnf_error:
        print(f"Erro: {fnf_error}")
    except Exception as e:
        print(f"Erro ao converter DOCX para PDF: {e}")
    finally:
        try:
            word.Quit()
        except:
            pass  # Garantir que o Word seja fechado


