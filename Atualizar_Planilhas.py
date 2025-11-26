import logging
import win32com.client as win32
from pathlib import Path
from typing import List, Optional
from contextlib import contextmanager

# Configuração de Logs Profissional
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger("ExcelBot")

class ExcelRefresher:
    """
    Classe responsável pela orquestração segura do Microsoft Excel.
    Utiliza o padrão RAII (Resource Acquisition Is Initialization) via Context Managers.
    """

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.app = None

    def __enter__(self):
        """Inicializa a instância do Excel de forma isolada."""
        try:
            # DispatchEx força uma nova instância, evitando conflito com planilhas abertas pelo usuário
            self.app = win32.client.DispatchEx("Excel.Application")
            self.app.Visible = self.visible
            self.app.DisplayAlerts = False
            logger.info("Instância do Excel iniciada com sucesso.")
            return self
        except Exception as e:
            logger.critical(f"Falha ao iniciar o Excel: {e}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Garante o encerramento do processo do Excel, prevenindo 'zumbis' na memória."""
        if self.app:
            self.app.Quit()
            # Força a liberação do objeto COM
            del self.app
            logger.info("Instância do Excel encerrada e memória liberada.")

    def refresh_workbook(self, file_path: Path) -> bool:
        """
        Abre, atualiza conexões de dados e salva uma planilha.
        
        Args:
            file_path (Path): Caminho do arquivo Excel.
            
        Returns:
            bool: True se sucesso, False se falha.
        """
        path_obj = Path(file_path)
        
        if not path_obj.exists():
            logger.error(f"Arquivo não encontrado: {path_obj}")
            return False

        wb = None
        try:
            logger.info(f"Abrindo: {path_obj.name}")
            wb = self.app.Workbooks.Open(str(path_obj.resolve()))

            # Desabilita atualização de tela para performance
            self.app.ScreenUpdating = False

            logger.info("Disparando RefreshAll...")
            wb.RefreshAll()
            
            # Método nativo: Aguarda o término das queries assíncronas (Power Query/Pivot)
            # Isso elimina a necessidade de time.sleep() arbitrários
            self.app.CalculateUntilAsyncQueriesDone()
            
            wb.Save()
            logger.info(f"✅ Sucesso: {path_obj.name} atualizada e salva.")
            return True

        except Exception as e:
            logger.error(f"❌ Erro ao processar {path_obj.name}: {e}")
            return False
        
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            self.app.ScreenUpdating = True

# --- Execução Principal ---
if __name__ == "__main__":
    
    # Lista de arquivos usando raw strings ou Path objects
    # DICA: Mantenha caminhos relativos ou absolutos claros
    LISTA_PLANILHAS = [
        r"C:\Relatorios\Vendas_2023.xlsx",
        r"C:\Relatorios\Estoque_Consolidado.xlsx"
    ]

    logger.info(">>> Iniciando Pipeline de Atualização Excel <<<")

    # O uso do 'with' garante que o Excel feche no final, sem exceção.
    try:
        with ExcelRefresher(visible=False) as bot:
            for arquivo in LISTA_PLANILHAS:
                bot.refresh_workbook(arquivo)
    except Exception as e:
        logger.fatal(f"Erro fatal na execução do robô: {e}")

    logger.info(">>> Processo Finalizado <<<")
