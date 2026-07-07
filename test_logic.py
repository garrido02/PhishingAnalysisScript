import os
import pandas as pd
import pytest

# Importamos os teus módulos (garante que os nomes dos ficheiros são process.py e analyze.py)
import process as prc
import analyze as anl
from analyze import extract_campaign_info

# --- 1. TESTES UNITÁRIOS ---


def test_extract_campaign_info_LC():
    """Testa se extrai corretamente localização e tipo LC (Credentials)."""
    loc, c_type = extract_campaign_info("2025C2_PTAOMZ_D_LCGemini2.csv")
    assert loc == "PTAOMZ"
    assert c_type == "Credentials"


def test_extract_campaign_info_LCA():
    """Testa se a prioridade do LCA funciona e não é intercetada pelo LC."""
    loc, c_type = extract_campaign_info("2025C2_MADRID_D_LCAGemini.csv")
    assert loc == "MADRID"
    assert c_type == "Credentials & Anex"


def test_extract_campaign_info_A():
    """Testa a extração do tipo Anex."""
    loc, c_type = extract_campaign_info("2024_LISBOA_D_A_Test.csv")
    assert loc == "LISBOA"
    assert c_type == "Anex"


def test_extract_campaign_info_unknown():
    """Testa o comportamento com um ficheiro fora do formato esperado."""
    # Agora sim, um nome de ficheiro com menos de 4 partes
    loc, c_type = extract_campaign_info("campanha_invalida.csv")
    assert loc == "Unknown"
    assert c_type == "Unknown"


# --- 2. TESTES DE INTEGRAÇÃO (Ficheiros) ---


def test_process_file_creates_excel(tmp_path):
    """
    Testa se o process_file lê um CSV falso e cria o Excel master corretamente.
    O tmp_path é gerido pelo pytest para não sujar o teu disco.
    """
    # 1. Preparar caminhos temporários
    input_csv = tmp_path / "2025C2_PORTO_D_LCTest.csv"
    output_excel = tmp_path / "master_results.xlsx"

    # 2. Criar um CSV falso com cabeçalhos que o teu config.py reconheça
    # Assumimos que "Email" e "Opened At" estão no teu config.PHISHING_HEADERS
    mock_data = pd.DataFrame(
        {
            "Email": ["user1@empresa.com", "user2@empresa.com"],
            "Opened At": ["2023-01-01 10:00:00", None],
        }
    )
    mock_data.to_csv(input_csv, index=False)

    # 3. Executar a função
    prc.process_file(str(input_csv), str(output_excel))

    # 4. Validar os resultados
    assert os.path.exists(output_excel), "O ficheiro Excel não foi criado."

    # Ler o Excel criado e verificar se os dados estão lá
    df_result = pd.read_excel(output_excel, sheet_name="Raw Data")

    assert (
        "Source File" in df_result.columns
    ), "A coluna 'Source File' não foi injetada."
    assert df_result.iloc[0]["Source File"] == "2025C2_PORTO_D_LCTest.csv"
    assert len(df_result) == 2, "Nem todas as linhas foram processadas."


def test_generate_report_creates_sheets(tmp_path):
    """
    Testa se o generate_report consegue ler a 'Raw Data' e criar as tabelas de Analytics.
    """
    master_excel = tmp_path / "master_results.xlsx"

    # 1. Criar dados base (Raw Data) para o teste do relatório
    mock_raw_data = pd.DataFrame(
        {
            "Source File": ["2025C2_TEST_D_LCTest.csv", "2025C2_TEST_D_LCTest.csv"],
            "Email": ["user1@teste.com", "user2@teste.com"],
            "Opened At": ["2025-01-01 10:00", "2025-01-01 11:00"],
            "Clicked At": ["2025-01-01 10:05", None],
            "Data Entered At": [None, None],
            "Attachment Opened At": [None, None],
        }
    )

    # Guardar num Excel temporário
    mock_raw_data.to_excel(master_excel, index=False, sheet_name="Raw Data")

    # 2. Executar a geração do relatório
    anl.generate_report(str(master_excel))

    # 3. Validar se as sheets de output foram criadas sem apagar a Raw Data
    xls = pd.ExcelFile(master_excel)

    assert "Analytics" in xls.sheet_names, "A sheet 'Analytics' não foi gerada."
    assert (
        "Analysis - Users" in xls.sheet_names
    ), "A sheet 'Analysis - Users' não foi gerada."

    # Garantir que a lógica não apagou os dados originais
    assert "Raw Data" in xls.sheet_names, "A sheet original 'Raw Data' desapareceu!"
