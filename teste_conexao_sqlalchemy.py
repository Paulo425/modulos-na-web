from sqlalchemy import create_engine
from sqlalchemy import text

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


def testar_conexao():
    try:
        engine = create_connection()
        if engine:
            with engine.connect() as conn:
                versao = conn.execute(text("SELECT VERSION()")).scalar()
                print(f"✅ Conexão estabelecida com sucesso! Versão do MySQL: {versao}")
        else:
            print("❌ Engine não foi criada.")
    except Exception as e:
        print(f"❌ Erro ao testar conexão: {type(e).__name__}: {e}")


if __name__ == "__main__":
    testar_conexao()
