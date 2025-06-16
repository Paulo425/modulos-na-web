from sqlalchemy import create_engine, text

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

def criar_tabela():
    engine = create_connection()
    with engine.connect() as conn:
        sql = text("""
        CREATE TABLE IF NOT EXISTS usuarios_memoriais (
            id INT AUTO_INCREMENT PRIMARY KEY,
            usuario VARCHAR(100) NOT NULL UNIQUE,
            senha_hash TEXT NOT NULL,
            nivel VARCHAR(50) DEFAULT 'tecnico',
            aprovado BOOLEAN DEFAULT TRUE
        );
        """)
        conn.execute(sql)
        conn.commit()
        print("✅ Tabela 'usuarios_memoriais' criada ou já existente.")

if __name__ == "__main__":
    criar_tabela()
