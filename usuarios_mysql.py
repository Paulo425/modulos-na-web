from sqlalchemy import create_engine, text

# Função de conexão com SQLAlchemy
def create_connection():
    try:
        connection_url = (
            "mysql+mysqlconnector://admin_python:" 
            "jLZW96brdjRqde7LS7ge" 
            "@phoenixappraisal.com.br/admin_phoenix_rio"

        )
        engine = create_engine(connection_url)
        return engine
    except Exception as e:
        print(f"❌ Erro de conexão: {e}")
        return None


def salvar_usuario_mysql(usuario, senha_hash, nivel='tecnico', aprovado=False):
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("""
                INSERT INTO usuarios_memoriais (usuario, senha_hash, nivel, aprovado)
                VALUES (:usuario, :senha_hash, :nivel, :aprovado)
            """)
            conn.execute(sql, {
                "usuario": usuario,
                "senha_hash": senha_hash,
                "nivel": nivel,
                "aprovado": aprovado
            })
            conn.commit()


def buscar_usuario_mysql(usuario):
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("SELECT * FROM usuarios_memoriais WHERE usuario = :usuario")
            result = conn.execute(sql, {"usuario": usuario}).mappings().fetchone()
            return dict(result) if result else None


def aprovar_usuario_mysql(usuario):
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("UPDATE usuarios_memoriais SET aprovado = TRUE WHERE usuario = :usuario")
            conn.execute(sql, {"usuario": usuario})
            conn.commit()


def excluir_usuario_mysql(usuario):
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("DELETE FROM usuarios_memoriais WHERE usuario = :usuario")
            conn.execute(sql, {"usuario": usuario})
            conn.commit()


def listar_pendentes_mysql():
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("SELECT usuario FROM usuarios_memoriais WHERE aprovado = FALSE")
            result = conn.execute(sql).fetchall()
            return [row[0] for row in result]
    return []


def listar_usuarios_mysql():
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("SELECT usuario FROM usuarios_memoriais ORDER BY usuario")
            result = conn.execute(sql).fetchall()
            return [row[0] for row in result]
    return []


def atualizar_senha_mysql(usuario, nova_senha_hash):
    engine = create_connection()
    if engine:
        with engine.connect() as conn:
            sql = text("UPDATE usuarios_memoriais SET senha_hash = :senha WHERE usuario = :usuario")
            conn.execute(sql, {"senha": nova_senha_hash, "usuario": usuario})
            conn.commit()
