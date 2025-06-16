from sqlalchemy import create_engine, text
from werkzeug.security import generate_password_hash


# Configurar a conexão com o seu banco MySQL
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
def recriar_admin():
    engine = create_connection()
    senha = "1234"
    hash_seguro = generate_password_hash(senha, method='pbkdf2:sha256')

    with engine.connect() as conn:
        sql = text("""
            UPDATE usuarios_memoriais
            SET senha_hash = :hash, aprovado = TRUE
            WHERE usuario = 'admin'
        """)
        conn.execute(sql, {"hash": hash_seguro})
        conn.commit()
        print(f"✅ Usuário admin atualizado com nova senha (1234) e hash seguro.")

if __name__ == "__main__":
    recriar_admin()
