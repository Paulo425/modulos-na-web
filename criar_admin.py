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
def criar_admin():
    engine = create_connection()
    usuario = "admin"
    senha = "1234"
    senha_hash = generate_password_hash(senha)

    with engine.connect() as conn:
        # Verifica se o admin já existe
        check_sql = text("SELECT 1 FROM usuarios_memoriais WHERE usuario = :usuario")
        resultado = conn.execute(check_sql, {"usuario": usuario}).fetchone()

        if resultado:
            print("⚠️ O usuário 'admin' já existe no banco.")
        else:
            insert_sql = text("""
                INSERT INTO usuarios_memoriais (usuario, senha_hash, nivel, aprovado)
                VALUES (:usuario, :senha_hash, :nivel, :aprovado)
            """)
            conn.execute(insert_sql, {
                "usuario": usuario,
                "senha_hash": senha_hash,
                "nivel": "admin",
                "aprovado": True
            })
            conn.commit()
            print("✅ Usuário 'admin' criado com sucesso!")

if __name__ == "__main__":
    criar_admin()
