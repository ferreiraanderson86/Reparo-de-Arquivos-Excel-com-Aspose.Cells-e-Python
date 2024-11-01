from __future__ import absolute_import
import os
from jpype import *

# Adiciona o caminho dos JARs do Aspose.Cells
__cells_jar_dir__ = os.path.dirname(__file__)
addClassPath(os.path.join(__cells_jar_dir__, "lib", "aspose-cells-24.10.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "lib", "bcprov-jdk15on-1.68.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "lib", "bcpkix-jdk15on-1.68.jar"))
addClassPath(os.path.join(__cells_jar_dir__, "lib", "JavaClassBridge.jar"))

# Inicia a JVM (Java Virtual Machine)
startJVM()

class Workbook:
    def __init__(self):
        # Importa a classe Workbook da biblioteca Aspose.Cells
        self.WorkbookClass = JClass("com.aspose.cells.Workbook")
        self.workbook = None

    def open(self, file_path):
        # Abre o arquivo Excel
        self.workbook = self.WorkbookClass(file_path)

    def save(self, output_path):
        # Salva o arquivo Excel
        self.workbook.save(output_path)

# Exemplo de uso da classe Workbook
if __name__ == "__main__":
    # Caminho do arquivo corrompido
    file_path = 'C:/Users/FR76/Downloads/TABELAO_20241027220001.xlsx'
    # Caminho para salvar o arquivo reparado
    output_path = 'C:/Users/FR76/Downloads/TABELAO_REPARADO.xlsx'

    try:
        workbook = Workbook()
        workbook.open(file_path)  # Abre o arquivo
        workbook.save(output_path)  # Salva o arquivo reparado
        print(f'Arquivo reparado e salvo com sucesso: {output_path}')
    except Exception as e:
        print(f"Ocorreu um erro ao tentar reparar o arquivo: {e}")
    finally:
        # Para a JVM após o uso
        shutdownJVM()