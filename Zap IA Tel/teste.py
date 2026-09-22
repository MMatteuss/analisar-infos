import uiautomator2 as u2
import pandas as pd
from datetime import datetime
import time
import requests
import json
from ollama import chat
import logging
import sys
import os

# ============================================================
# Configuração
# ============================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

d = u2.connect('127.0.0.1')
d.set_fastinput_ime(True)

# Coordenadas
COORD_ADD_CONTATO   = (1005, 2106)   # botão "+ nova conversa"
COORD_LUPA_PESQUISA = (950, 167)     # campo de pesquisa
COORD_PRIMEIRO_CONT = (912, 380)     # primeiro contato no resultado
COORD_VOLTAR        = (56, 174)      # botão voltar (usado como fallback)

# Relatório
relatorioName = f"relatorio {datetime.now().strftime('%d-%m-%Y %H-%M-%S')}.txt"
relatorio = relatorioName.replace(":", "-").replace("/", "-")

# Contadores
numDisparo = 0
mensagensEnviadas = 0

# Caminho do Excel
CAMINHO_EXCEL = "/data/data/com.termux/files/home/dadosAL.xlsx"

try:
    df = pd.read_excel(CAMINHO_EXCEL)
except FileNotFoundError:
    logging.error(f"Arquivo Excel não encontrado em: {CAMINHO_EXCEL}")
    sys.exit(1)

if 'telefone' not in df.columns or 'cnpjCompleto' not in df.columns:
    logging.error("O Excel deve conter as colunas 'telefone' e 'cnpjCompleto'.")
    sys.exit(1)


# ============================================================
# Funções auxiliares (substituindo pyautogui)
# ============================================================
def digitar_texto(texto: str):
    d.clipboard = texto
    time.sleep(0.3)
    # Cola via tecla de atalho do Android (funciona no WhatsApp)
    d.press("paste")

def clicar(x: int, y: int, delay: float = 0.5):
    d.click(x, y)
    time.sleep(delay)


def telefone_nao_encontrado() -> bool:
    return (
        d(textContains="Nenhum resultado").exists(timeout=2)
        or d(textContains="No results found").exists(timeout=1)
        or d(text="Nenhum resultado encontrado").exists
    )


# ============================================================
# Consulta CNPJ (BrasilAPI)
# ============================================================

def consulta_cnpj(cnpj: str):
    cnpj = ''.join(filter(str.isdigit, cnpj))
    if len(cnpj) != 14:
        raise ValueError("CNPJ deve ter 14 dígitos")
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    return resp.json()


def formatar_dados_cnpj_filtrado(dados: dict) -> str:
    linhas = []
    for chave, valor in dados.items():
        if valor is None or (isinstance(valor, str) and valor.strip() == '') or (isinstance(valor, (list, dict)) and not valor):
            continue
        if isinstance(valor, list):
            if valor and isinstance(valor[0], dict):
                valor_str = json.dumps(valor, indent=2, ensure_ascii=False)
            else:
                valor_str = ', '.join(str(v) for v in valor)
        elif isinstance(valor, dict):
            valor_str = json.dumps(valor, indent=2, ensure_ascii=False)
        else:
            valor_str = str(valor)
        linhas.append(f"{chave}: {valor_str}")
    return '\n'.join(linhas)

# ============================================================
# Geração de mensagem com IA (Ollama)
# ============================================================
#region IA
def gerar_mensagem_ia(dados_empresa: str) -> str:
    MODELO = "qwen3:4b"

    prompt = f"""
Você é um assistente comercial da Yale Solution especializado em prospecção B2B.

Sua função é analisar UMA empresa por execução, identificar oportunidades comerciais plausíveis e criar uma única mensagem de prospecção personalizada.

==================================================
REGRA ABSOLUTA: CADA EMPRESA É UMA ANÁLISE ISOLADA
==================================================

A variável abaixo contém os dados da EMPRESA ATUAL:

{dados_empresa}

Esses são os ÚNICOS dados disponíveis sobre a empresa.
NUNCA utilize informações de empresas analisadas anteriormente.
Considere esta execução como uma análise completamente nova e independente.

==================================================
SOBRE A YALE SOLUTION
=====================
A Yale Solution oferece: Desenvolvimento Web, Aplicativos Mobile, Sistemas Completos (ERPs, gestão, logística), Bots & Automação (chatbots, WhatsApp, email, IA), Leads da Receita Federal, Manutenção & Vistoria, Infraestrutura de Rede, Consultoria Técnica.

==================================================
OBJETIVO
========
Analise SOMENTE a empresa atual. Identifique atividade econômica, segmento, características explícitas e determine quais serviços da Yale Solution têm relevância comercial plausível.

Pode escolher 1, 2, 3 ou mais serviços — desde que tenham relação lógica real. NÃO inclua serviços apenas para aumentar a quantidade.

==================================================
REGRAS CRÍTICAS
===============
- NÃO invente informações, problemas, sistemas, sites, apps, funcionários, clientes.
- Se não está nos dados, é DESCONHECIDO.
- Use linguagem de possibilidade ("pode ajudar", "pode ser útil").
- NUNCA diga "vocês precisam" ou "identificamos um problema".
- NÃO mencione CNPJ, CNAE, capital social, natureza jurídica.
- NÃO faça promessas de resultado.
- Mensagem: 350-600 caracteres, profissional, natural, para WhatsApp.
- Estruture: saudação → personalização real → serviços relevantes → como ajudam → pergunta final.

==================================================
FORMATO DE SAÍDA
================
RETORNE SOMENTE A MENSAGEM FINAL. SEM títulos, sem JSON, sem Markdown, sem explicações, sem colchetes.
"""
    try:
        resposta = chat(
            model=MODELO,
            messages=[{"role": "user", "content": prompt}],
            options={"temperature": 0.7}
        )
        return resposta["message"]["content"].strip()
    except Exception as e:
        logging.error(f"Erro ao chamar Ollama: {e}")
        return "Olá! Somos a Yale Solution e oferecemos soluções em tecnologia para empresas. Podemos conversar?"
#endregion

# ============================================================
# Navegação WhatsApp
# ============================================================
def abrir_nova_conversa():
    """Aperta no botão de adicionar contato/nova conversa."""
    clicar(*COORD_ADD_CONTATO, delay=1.0)


def pesquisar_telefone(telefone: str):
    """Aperta na lupa e digita o telefone."""
    clicar(*COORD_LUPA_PESQUISA, delay=0.8)
    # Limpa o campo caso tenha texto anterior
    try:
        d(focused=True).clear_text()
    except Exception:
        pass
    time.sleep(0.3)
    d.send_keys(str(telefone), clear=True)
    time.sleep(1.2)


def clicar_primeiro_contato():
    """Clica no primeiro resultado da busca."""
    clicar(*COORD_PRIMEIRO_CONT, delay=1.2)


def voltar_tela(vezes: int = 1):
    """Usa o botão físico/gesto de voltar do Android (mais confiável que coordenada)."""
    for _ in range(vezes):
        d.press("back")
        time.sleep(0.4)


# ============================================================
# Função principal
# ============================================================
def main():
    global numDisparo, mensagensEnviadas

    with open(relatorio, 'w', encoding='utf-8') as arquivo:
        arquivo.write("Relatório de Telefones\n")
        arquivo.write(f"Gerado em; {datetime.now().strftime('%d/%m/%Y')};{datetime.now().strftime('%H:%M:%S')};\n\n")

        for index, row in df.iterrows():
            telefone = row['telefone']
            cnpj = row['cnpjCompleto']

            numDisparo += 1
            hora_atual = datetime.now().strftime('%H:%M:%S')
            logging.info(f"Processando telefone: {telefone} - {hora_atual}")

            try:
                # 1 Abrir nova conversa
                abrir_nova_conversa()
                time.sleep(1)

                # 2 Pesquisar telefone
                pesquisar_telefone(telefone)

                # 3 Verificar se encontrou
                if telefone_nao_encontrado():
                    logging.warning(f"Telefone {telefone} não encontrado.")
                    arquivo.write(
                        f"{telefone};Telefone não encontrado;"
                        f"{datetime.now().strftime('%d/%m/%Y')};"
                        f"{datetime.now().strftime('%H:%M:%S')};\n"
                    )
                    # Fecha a busca e volta
                    voltar_tela(2)
                    continue

                # 4 Clicar no primeiro contato
                clicar_primeiro_contato()
                time.sleep(1.5)

                # 5 Consultar CNPJ
                try:
                    dados_json = consulta_cnpj(str(cnpj))
                    dados_empresa = formatar_dados_cnpj_filtrado(dados_json)
                except Exception as e:
                    logging.error(f"Erro na consulta CNPJ para {cnpj}: {e}")
                    dados_empresa = f"CNPJ: {cnpj}\n(Informações não disponíveis)"

                arquivo.write(
                    f"{telefone};encontrado;"
                    f"{datetime.now().strftime('%d/%m/%Y')};"
                    f"{datetime.now().strftime('%H:%M:%S')};\n"
                )

                # 6 Gerar mensagem com IA
                mensagem = gerar_mensagem_ia(dados_empresa)
                logging.info(f"Mensagem gerada: {mensagem[:50]}...")

                # 7 Colar e enviar a mensagem
                digitar_texto(mensagem)
                time.sleep(1)

                d.press("enter")
                time.sleep(1.2)

                arquivo.write(f"Mensagem enviada: {mensagem}\n")

                # 8 Voltar para tela inicial do WhatsApp
                voltar_tela(3)

                mensagensEnviadas += 1

                # Pausa a cada 3 mensagens
                if mensagensEnviadas % 3 == 0:
                    logging.info("Pausa de 30 segundos após 3 envios...")
                    time.sleep(30)

            except Exception as e:
                logging.error(f"Erro inesperado no telefone {telefone}: {e}")
                try:
                    voltar_tela(3)
                except Exception:
                    pass
                continue

            if numDisparo >= 100:
                logging.info("Limite de 100 disparos atingido. Encerrando.")
                break

    logging.info(f"Processo finalizado. Total de disparos: {numDisparo}")

if __name__ == "__main__":
    main()