import os
import sys
import subprocess
import shutil

def checar_pyinstaller():
    """Verifica se o PyInstaller está instalado e o instala se necessário."""
    try:
        import PyInstaller
        print("PyInstaller já está instalado.")
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller instalado com sucesso.")

def instalar_dependencias():
    """Instala as dependências do projeto."""
    print("Instalando dependências...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    print("Dependências instaladas com sucesso.")

def compilar_backend():
    """Compila o backend em um executável."""
    print("Compilando o backend...")
    
    # Diretório de saída para o executável
    output_dir = os.path.join("..", "..", "dist")
    os.makedirs(output_dir, exist_ok=True)
    
    # Comando do PyInstaller - usando o Python do ambiente virtual
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=OrganizadorPlanilhas",
        "--onefile",
        "--windowed",
        "--add-data=organizador.py;.",
        "--clean",
        "api.py"
    ]
    
    # Executar o comando
    subprocess.check_call(cmd)
    
    # Mover o executável para o diretório de saída
    shutil.copy(
        os.path.join("dist", "OrganizadorPlanilhas.exe"),
        os.path.join(output_dir, "OrganizadorPlanilhas.exe")
    )
    
    print(f"Backend compilado com sucesso! Executável disponível em: {os.path.abspath(output_dir)}")

if __name__ == "__main__":
    # Mudar para o diretório do script
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # Verificar PyInstaller
    checar_pyinstaller()
    
    # Instalar dependências
    instalar_dependencias()
    
    # Compilar o backend
    compilar_backend()
    
    print("Processo de compilação concluído!") 