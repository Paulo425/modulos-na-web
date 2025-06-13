import os
import shutil
import tempfile
import logging
from preparar_planilhas import preparar_planilhas  # certifica-se que essa função exista no mesmo módulo

# Configura o logger para funcionar no ambiente Render
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

def preparar_arquivos(cidade, caminho_excel, caminho_dxf, base_dir):
    try:
        cidade_formatada = cidade.replace(" ", "_")
        cidade_dir = os.path.join(base_dir, cidade_formatada)

        # Criar diretórios temporários
        diretorio_temp = tempfile.mkdtemp()
        diretorio_base = os.path.join(diretorio_temp, cidade_formatada)
        diretorio_preparado = os.path.join(diretorio_base, "PREPARADO")
        diretorio_concluido = os.path.join(diretorio_base, "CONCLUIDO")
        os.makedirs(diretorio_preparado, exist_ok=True)
        os.makedirs(diretorio_concluido, exist_ok=True)

        # Copiar arquivos
        arquivo_excel_recebido = os.path.join(diretorio_base, os.path.basename(caminho_excel))
        arquivo_dxf_recebido = os.path.join(diretorio_base, os.path.basename(caminho_dxf))
        shutil.copy(caminho_excel, arquivo_excel_recebido)
        shutil.copy(caminho_dxf, arquivo_dxf_recebido)

        print(f"✅ Arquivo Excel copiado para: {arquivo_excel_recebido}")
        print(f"✅ Arquivo DXF copiado para: {arquivo_dxf_recebido}")

        # Chamada obrigatória para gerar FECHADA_*_*.xlsx
        preparar_planilhas(arquivo_excel_recebido, diretorio_preparado)

        return {
            "arquivo_excel_recebido": arquivo_excel_recebido,
            "arquivo_dxf_recebido": arquivo_dxf_recebido,
            "diretorio_base": diretorio_base,
            "diretorio_preparado": diretorio_preparado,
            "diretorio_concluido": diretorio_concluido
        }

    except Exception as e:
        logger.error(f"Erro ao preparar os arquivos: {e}")
        print(f"❌ Erro ao preparar os arquivos: {e}")
        return {}
