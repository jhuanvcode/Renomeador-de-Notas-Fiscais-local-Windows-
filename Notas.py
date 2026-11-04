"""
RENOMEADOR DE NOTAS FISCAIS

REGRAS DE SEGURANÇA
  1. Só usa o nome de quem EMITIU (prestador/fornecedor/emitente).
  2. A própria empresa (quem RECEBE a nota) é BLOQUEADA: nota emitida por ela não é renomeada.
  3. Endereço, bairro, CEP, e-mail, telefone e rótulos NUNCA viram nome.
  4. Se não achar o nome com segurança -> o arquivo mantém o nome original.

RECURSOS DE DESEMPENHO E HISTÓRICO
  - HISTÓRICO PERMANENTE: Salva em disco ('historico_renomeacao.json') todos os
    arquivos já processados e renomeados. Em execuções futuras (.exe), o script
    ignora instantaneamente as notas já tratadas e só analisa as novas!
"""

import os
import re
import csv
import sys
import time
import json
import logging
import tempfile
import subprocess
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Optional, List


try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None


# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Configure via variáveis de ambiente (veja .env.example). NUNCA coloque caminhos reais,
# nomes de usuário ou CNPJ diretamente no código.
PASTA_BASE = os.environ.get("NOTAS_PASTA_BASE", "")

# True  = só simula  |   False = renomeia de verdade
DRY_RUN = os.environ.get("DRY_RUN", "false").lower() in ("true", "1")

EXTENSOES_VALIDAS = (".pdf", ".xml", ".csv")

# Limitar Tamanho máximo do nome no arquivo
NOME_MAX_CARACTERES = 20

# "" = sem espaço (ABCDANEURO...)   |   "_" = ABC_DA_NEURO
SEPARADOR_PALAVRAS = ""

# None = todas as páginas
MAX_PAGINAS_PDF = None

# None = processa tudo, se quiser uma quantidade especifica colocar o numero desejado.
MAX_ARQUIVOS_TESTE = None

# Espera o OneDrive terminar de baixar/sincronizar o arquivo
VERIFICAR_SINCRONIZACAO = True

# Arquivos que eu quero que ignore
# Nomes de arquivo (separados por vírgula, em minúsculo) a ignorar. Ex: "comprovante.pdf,manual.pdf"
ARQUIVOS_SEM_AVISO = {
    x.strip().lower()
    for x in os.environ.get("ARQUIVOS_IGNORADOS", "").split(",")
    if x.strip()
}

# ------------------------------------------------------------
# HISTÓRICO PERMANENTE 
# ------------------------------------------------------------
USAR_HISTORICO = True

# True = ignora o histórico e força a reanálise de todos os arquivos
FORCAR_REPROCESSAMENTO = False

# ------------------------------------------------------------
# TRAVA DA PRÓPRIA EMPRESA
# ------------------------------------------------------------
# CNPJ da empresa que RECEBE as notas (tomadora). Obrigatório.
EMPRESA_CNPJ = os.environ.get("EMPRESA_CNPJ", "")
BLOQUEAR_TODAS_AS_FILIAIS = True     # bloqueia qualquer CNPJ com a mesma raiz (8 primeiros dígitos)

# Termos do nome da própria empresa (bloqueio quando não há CNPJ legível).
# Grupos separados por "|"; termos dentro do grupo separados por "+" (todos devem aparecer).
# Sem acentos e em minúsculo. Ex: "minhaempresa+matriz|minhaempresa+filial"
EMPRESA_TERMOS_BLOQUEIO = os.environ.get("EMPRESA_TERMOS_BLOQUEIO", "")

SALVAR_DEBUG = os.environ.get("SALVAR_DEBUG", "false").lower() in ("true", "1")
PASTA_DEBUG = os.path.join(PASTA_BASE, "debug_extracao")

if getattr(sys, "frozen", False):
    _DIR_SCRIPT = os.path.dirname(sys.executable)
else:
    _DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_LOG = os.path.join(_DIR_SCRIPT, "log_renomeacao_nome.txt")
ARQUIVO_RELATORIO = os.path.join(_DIR_SCRIPT, "relatorio_renomeacao_nome.csv")
ARQUIVO_OCR_PS1 = os.path.join(_DIR_SCRIPT, "_ocr_engine.ps1")
ARQUIVO_HISTORICO = os.path.join(_DIR_SCRIPT, "historico_renomeacao.json")


# ============================================================
# LOG
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, mode="w", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ============================================================
# TEXTO
# ============================================================

def remover_acentos(texto) -> str:
    if not texto:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def normalizar_texto(texto) -> str:
    texto = remover_acentos(texto).lower()
    return re.sub(r"\s+", " ", texto).strip()


# ============================================================
# CNPJ/CPF 
# ============================================================

_PESOS_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
_PESOS_2 = [6] + _PESOS_1


def _dv(base: str, pesos) -> int:
    soma = sum((ord(c) - 48) * p for c, p in zip(base, pesos))
    r = soma % 11
    return 0 if r < 2 else 11 - r


def cnpj_valido(c: str) -> bool:
    c = c.upper()
    if not re.fullmatch(r"[0-9A-Z]{12}\d{2}", c) or len(set(c)) == 1:
        return False
    d1 = _dv(c[:12], _PESOS_1)
    d2 = _dv(c[:12] + str(d1), _PESOS_2)
    return c[12:] == f"{d1}{d2}"


def cpf_valido(c: str) -> bool:
    if not re.fullmatch(r"\d{11}", c) or len(set(c)) == 1:
        return False
    for n in (9, 10):
        soma = sum(int(c[i]) * (n + 1 - i) for i in range(n))
        if ((soma * 10) % 11) % 10 != int(c[n]):
            return False
    return True


def limpar_documento(bruto) -> Optional[str]:
    if not bruto:
        return None
    s = re.sub(r"[^0-9A-Za-z]", "", str(bruto)).upper()
    if len(s) == 14 and cnpj_valido(s):
        return s
    if len(s) == 11 and cpf_valido(s):
        return s
    return None


EMPRESA_DOC = limpar_documento(EMPRESA_CNPJ)
if EMPRESA_DOC is None:
    # Não exibe o valor informado, só orienta.
    raise SystemExit("EMPRESA_CNPJ ausente ou inválido. Defina a variável de ambiente EMPRESA_CNPJ.")


def doc_bloqueado(doc: Optional[str]) -> bool:
    if not doc:
        return False
    if doc == EMPRESA_DOC:
        return True
    return BLOQUEAR_TODAS_AS_FILIAIS and len(doc) == 14 and doc[:8] == EMPRESA_DOC[:8]


_S = r"\s*"
RE_CNPJ_FMT = re.compile(
    r"(?<![0-9A-Za-z])[0-9A-Z]{2}" + _S + r"\." + _S + r"[0-9A-Z]{3}" + _S + r"\." + _S
    + r"[0-9A-Z]{3}" + _S + r"/" + _S + r"[0-9A-Z]{4}" + _S + r"-" + _S + r"\d{2}(?![0-9A-Za-z])"
)
RE_CPF_FMT = re.compile(
    r"(?<!\d)\d{3}" + _S + r"\." + _S + r"\d{3}" + _S + r"\." + _S + r"\d{3}" + _S
    + r"-" + _S + r"\d{2}(?!\d)"
)
RE_BRUTO = re.compile(r"(?<!\d)(?:\d{14}|\d{11})(?!\d)")


def documentos_no_texto(texto: str) -> List[str]:
    """Todos os CNPJ/CPF válidos do texto, na ordem em que aparecem."""
    def achar(regexes):
        achados = []
        for rx in regexes:
            for m in rx.finditer(texto or ""):
                d = limpar_documento(m.group())
                if d:
                    achados.append((m.start(), d))
        achados.sort()
        saida = []
        for _, d in achados:
            if d not in saida:
                saida.append(d)
        return saida
    return achar([RE_CNPJ_FMT, RE_CPF_FMT]) or achar([RE_BRUTO])


# ============================================================
# HISTÓRICO PERMANENTE
# ============================================================

_historico_dados = {}
_historico_modificado = False


def _chave(p: str) -> str:
    return os.path.normcase(os.path.abspath(p))


def carregar_historico() -> dict:
    """Carrega o histórico de arquivos já renomeados/processados."""
    if not USAR_HISTORICO or FORCAR_REPROCESSAMENTO:
        return {}
    if not os.path.isfile(ARQUIVO_HISTORICO):
        return {}
    try:
        with open(ARQUIVO_HISTORICO, "r", encoding="utf-8") as f:
            dados = json.load(f)
            if isinstance(dados, dict):
                return dados
    except Exception as e:
        log.warning(f"Não foi possível ler o histórico ({e}). Iniciando novo.")
    return {}


def salvar_historico():
    """Salva o histórico permanentemente em disco."""
    global _historico_modificado
    if not USAR_HISTORICO or DRY_RUN or not _historico_modificado:
        return
    try:
        pasta = os.path.dirname(ARQUIVO_HISTORICO)
        if pasta:
            os.makedirs(pasta, exist_ok=True)
        temp_hist = ARQUIVO_HISTORICO + ".tmp"
        with open(temp_hist, "w", encoding="utf-8") as f:
            json.dump(_historico_dados, f, ensure_ascii=False, indent=2)
        if os.path.exists(ARQUIVO_HISTORICO):
            os.replace(temp_hist, ARQUIVO_HISTORICO)
        else:
            os.rename(temp_hist, ARQUIVO_HISTORICO)
        _historico_modificado = False
    except Exception as e:
        log.error(f"Erro ao salvar histórico permanente: {e}")


def arquivo_ja_processado(caminho: str) -> bool:
    """Verifica se o arquivo já foi processado anteriormente e permanece inalterado."""
    if not USAR_HISTORICO or FORCAR_REPROCESSAMENTO:
        return False
    chave = _chave(caminho)
    registro = _historico_dados.get(chave)
    if not registro:
        return False
    try:
        stat = os.stat(caminho)
        if registro.get("tamanho") == stat.st_size and abs(registro.get("mtime", 0) - stat.st_mtime) < 1.0:
            return True
    except OSError:
        pass
    return False


def registrar_no_historico(caminho: str, nome_original: str, status: str, nome_fornecedor: str = ""):
    """Adiciona um arquivo ao histórico permanente (efetivado em modo real)."""
    global _historico_modificado
    if not USAR_HISTORICO or DRY_RUN:
        return
    try:
        stat = os.stat(caminho)
        chave = _chave(caminho)
        _historico_dados[chave] = {
            "caminho": caminho,
            "nome_original": nome_original,
            "nome_arquivo": os.path.basename(caminho),
            "nome_fornecedor": nome_fornecedor,
            "status": status,
            "tamanho": stat.st_size,
            "mtime": stat.st_mtime,
            "data_registro": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        _historico_modificado = True
    except OSError:
        pass


# ============================================================
# NOME 
# ============================================================

class PrestadorEhEmpresaPropria(Exception):
    """O 'prestador' encontrado é a própria empresa (nota emitida por ela)."""


PRIMEIRA_PALAVRA_INVALIDA = {
    "rua", "r", "avenida", "av", "travessa", "tv", "alameda", "praca", "estrada",
    "rodovia", "largo", "beco", "viela", "logradouro", "bairro", "cidade",
    "estado", "municipio", "cep", "endereco", "numero", "complemento", "quadra",
    "telefone", "fone", "email", "e-mail", "inscricao", "indicador", "cnpj",
    "cpf", "nif", "nome", "razao", "prestador", "fornecedor", "tomador",
    "emitente", "destinatario", "adquirente", "empresa", "optante", "simples",
    "regime", "codigo", "sigla", "uf", "data", "valor", "local",
}
PALAVRA_INVALIDA_EM_QUALQUER_LUGAR = {"bairro", "cep", "complemento", "logradouro", "endereco"}
RE_TELEFONE = re.compile(r"\(?\d{2}\)?\s*9?\d{4}-?\d{4}")


def _eh_empresa_propria_bruto(s: str) -> bool:
    """Bloqueio por nome, configurado via EMPRESA_TERMOS_BLOQUEIO (sem termos fixos no código)."""
    n = normalizar_texto(s).replace(" ", "")
    for grupo in EMPRESA_TERMOS_BLOQUEIO.split("|"):
        termos = [t.strip().lower() for t in grupo.split("+") if t.strip()]
        if termos and all(t in n for t in termos):
            return True
    return False


def limpar_nome(raw, confiavel=False) -> Optional[str]:
    if not raw:
        return None
    s = re.sub(r"\s+", " ", remover_acentos(raw)).strip(" :;-|")
    if not s or "@" in s:
        return None

    s = re.sub(r"^\d{2}[\.,]?\d{3}[\.,]?\d{3}\s*", "", s).strip()
    s = re.sub(r"^\d{8}\s*", "", s).strip()
    s = re.sub(r"\s*\d{11}\s*$", "", s).strip()

    if _eh_empresa_propria_bruto(s):
        raise PrestadorEhEmpresaPropria(s)

    palavras = re.findall(r"[A-Za-z0-9\-\.]+", s.lower())
    if not palavras:
        return None

    if not confiavel:
        primeira = palavras[0].strip(".-")
        if primeira in PRIMEIRA_PALAVRA_INVALIDA:
            return None
        if any(p.strip(".,-") in PALAVRA_INVALIDA_EM_QUALQUER_LUGAR for p in palavras):
            return None

    if re.search(r",\s*\d+", s):              
        return None
    if documentos_no_texto(s):                
        return None

    limpo = re.sub(r'[\\/:*?"<>|]', "", s)
    limpo = re.sub(r"[^A-Za-z0-9 ]", " ", limpo)
    limpo = re.sub(r"\s+", " ", limpo).strip()
    compacto = limpo.replace(" ", "")
    if len(compacto) < 3:
        return None
    if sum(c.isdigit() for c in compacto) > len(compacto) * 0.30:
        return None

    nome = limpo.upper().replace(" ", SEPARADOR_PALAVRAS)
    return nome[:NOME_MAX_CARACTERES].strip("_")


# ============================================================
# OCR HELPER 
# ============================================================

def _garantir_script_ocr(caminho_ps1: str):
    codigo_ps = r'''param (
    [Parameter(Mandatory=$true)]
    [string]$ImagePath
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }

function Await-WinRT($asyncOp, $type) {
    $asTask = $asTaskGeneric.MakeGenericMethod($type)
    $netTask = $asTask.Invoke($null, @($asyncOp))
    return $netTask.GetAwaiter().GetResult()
}

[Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime] | Out-Null

$fullPath = (Resolve-Path $ImagePath).Path
$opFile = [Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)
$file = Await-WinRT $opFile ([Windows.Storage.StorageFile])

$opStream = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
$stream = Await-WinRT $opStream ([Windows.Storage.Streams.IRandomAccessStream])

$opDecoder = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
$decoder = Await-WinRT $opDecoder ([Windows.Graphics.Imaging.BitmapDecoder])

$opBitmap = $decoder.GetSoftwareBitmapAsync()
$bitmap = Await-WinRT $opBitmap ([Windows.Graphics.Imaging.SoftwareBitmap])

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) {
    $lang = [Windows.Globalization.Language]::new("pt-BR")
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
}

$opResult = $engine.RecognizeAsync($bitmap)
$result = Await-WinRT $opResult ([Windows.Media.Ocr.OcrResult])

$words = @()
foreach ($line in $result.Lines) {
    foreach ($word in $line.Words) {
        $words += [PSCustomObject]@{
            Text = $word.Text
            X = [math]::Round($word.BoundingRect.X, 1)
            Y = [math]::Round($word.BoundingRect.Y, 1)
            W = [math]::Round($word.BoundingRect.Width, 1)
            H = [math]::Round($word.BoundingRect.Height, 1)
        }
    }
}

$words | ConvertTo-Json -Compress
'''
    try:
        with open(caminho_ps1, "w", encoding="utf-8") as f:
            f.write(codigo_ps)
    except Exception as e:
        log.debug(f"Falha gravando script OCR: {e}")


def executar_ocr_pagina(pdf_path: str, num_pagina: int = 0) -> List[dict]:
    if pdfium is None:
        return []
    tmp_path = None
    try:
        _garantir_script_ocr(ARQUIVO_OCR_PS1)
        pdf = pdfium.PdfDocument(pdf_path)
        if num_pagina >= len(pdf):
            return []
        page = pdf[num_pagina]
        img = page.render(scale=3.0).to_pil()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        img.save(tmp_path)

        cmd = [
            "powershell",
            "-ExecutionPolicy", "Bypass",
            "-File", ARQUIVO_OCR_PS1,
            "-ImagePath", tmp_path,
        ]
        res = subprocess.run(cmd, capture_output=True)
        if res.returncode != 0:
            return []

        out = res.stdout.decode("utf-8", errors="replace").strip()
        if not out:
            return []
        ocr_words = json.loads(out)
        return [
            {
                "text": w["Text"],
                "x0": w["X"],
                "x1": w["X"] + w["W"],
                "top": w["Y"],
                "bottom": w["Y"] + w["H"],
            }
            for w in ocr_words
        ]
    except Exception as erro:
        log.debug(f"Erro no OCR de {pdf_path}: {erro}")
        return []
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


# ============================================================
# PDF 
# ============================================================

def _titulo_prestador(linha: str) -> bool:
    return bool(re.search(r"\b(prestador|fornecedor|emitente)\b", normalizar_texto(linha)))


def _titulo_tomador(linha: str) -> bool:
    t = normalizar_texto(linha)
    return bool(re.search(r"\b(tomador|adquirente|destinatario)\b", t)) or "destinatario da operacao" in t


def secoes_prestador(textos: List[str]):
    secoes = []
    for i, t in enumerate(textos):
        if _titulo_prestador(t):
            fim = len(textos)
            for j in range(i + 1, len(textos)):
                if _titulo_tomador(textos[j]):
                    fim = j
                    break
            secoes.append((i, fim))
    return secoes


def secoes_com_plano_b(textos: List[str]):
    secoes = secoes_prestador(textos)
    if secoes:
        return secoes
    fim_pre = next((j for j, t in enumerate(textos) if _titulo_tomador(t)), None)
    if fim_pre:
        return [(0, fim_pre)]
    return []


def agrupar_linhas(words, tol=5):
    linhas = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if linhas and abs(w["top"] - linhas[-1]["top"]) <= tol:
            linhas[-1]["words"].append(w)
        else:
            linhas.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
    for l in linhas:
        l["words"].sort(key=lambda w: w["x0"])
        l["bottom"] = max(w["bottom"] for w in l["words"])
        l["texto"] = " ".join(w["text"] for w in l["words"])
    return linhas


def celulas(linha, gap=12):
    cs = []
    for w in linha["words"]:
        if cs and w["x0"] - cs[-1]["x1"] <= gap:
            cs[-1]["text"] += " " + w["text"]
            cs[-1]["x1"] = w["x1"]
        else:
            cs.append({"text": w["text"], "x0": w["x0"], "x1": w["x1"]})
    return cs

RE_ROTULO_NOME = re.compile(
    r"^(?:n\s*ome\s*[/|1iI\\-]?\s*nome\s*empresarial|"
    r"n\s*ome\s*[/|1iI\\-]?\s*razao\s*social|"
    r"nome\s*empresarial|"
    r"razao\s*social|"
    r"nome\s*do\s*prestador|"
    r"nome\s*do\s*fornecedor|"
    r"nome\s*do\s*emitente|"
    r"n\s*ome)"
    r"\s*(?::\s*(.*))?$",
    re.IGNORECASE,
)

RE_NOME_MESMA_LINHA = re.compile(
    r"^\s*n\s*ome\s*[/|1iI\\-]?\s*(?:razao\s*social|nome\s*empresarial)\s*:\s*(.+)$",
    re.IGNORECASE,
)

RE_ROTULO_SOZINHO = re.compile(r"^(?:emitente|prestador|fornecedor)\s*:\s*$", re.IGNORECASE)
RE_PREFIXO_DOC = re.compile(
    r"^\s*(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})\s*[-–]\s*"
)


def _sem_ruido(linha_texto: str) -> str:
    """Remove e-mail e telefone de uma linha."""
    toks = [t for t in linha_texto.split() if "@" not in t]
    return RE_TELEFONE.sub(" ", " ".join(toks))


def nome_na_secao(linhas, ini, fim):
    for i in range(ini, fim):

        m_linha = RE_NOME_MESMA_LINHA.match(remover_acentos(linhas[i]["texto"]))
        if m_linha:
            n = limpar_nome(m_linha.group(1), confiavel=True)
            if n:
                return n, "mesma linha do rótulo (modelo prefeitura)"

        if RE_ROTULO_SOZINHO.match(remover_acentos(linhas[i]["texto"]).strip()):
            valor = " ".join(linhas[k]["texto"] for k in range(i + 1, min(i + 3, fim)))
            valor = RE_PREFIXO_DOC.sub("", valor).strip()
            n = limpar_nome(valor, confiavel=True)
            if n:
                return n, "valor abaixo de 'Emitente:' (portal)"

        if "nota fiscal de servicos eletronica" in normalizar_texto(linhas[i]["texto"]):
            for k in range(i + 1, min(i + 3, fim)):
                n = limpar_nome(_sem_ruido(linhas[k]["texto"]))
                if n:
                    return n, "linha abaixo do título da NFS-e (sem rótulo)"

        cs = celulas(linhas[i])
        for c, cel in enumerate(cs):
            m = RE_ROTULO_NOME.match(remover_acentos(cel["text"]).strip())
            if not m:
                continue

            if m.group(1):
                n = limpar_nome(m.group(1), confiavel=True)
                if n:
                    return n, "mesma célula do rótulo"
                continue

            xl = cel["x0"] - 15
            xr = cs[c + 1]["x0"] - 5 if c + 1 < len(cs) else float("inf")
            for k in range(i + 1, min(i + 4, fim)):
                if linhas[k]["top"] - linhas[i]["bottom"] > 60:
                    break
                raw = " ".join(w["text"] for w in linhas[k]["words"] if xl <= w["x0"] < xr)
                n = limpar_nome(raw, confiavel=True)
                if n:
                    return n, "abaixo do rótulo (mesma coluna)"

            for k in range(i + 1, min(i + 4, fim)):
                n = limpar_nome(_sem_ruido(linhas[k]["texto"]), confiavel=True)
                if n:
                    return n, "linha abaixo do rótulo (plano B)"

    for i in range(ini, fim):
        n = limpar_nome(_sem_ruido(linhas[i]["texto"]), confiavel=False)
        if n:
            return n, "linha da seção prestador (fallback)"

    return None, ""


@dataclass
class Resultado:
    nome: Optional[str] = None
    metodo: str = ""
    motivo: str = ""
    texto_debug: str = ""


def extrair_nome_pdf(caminho: str) -> Resultado:
    res = Resultado()
    debug = []

    if pdfplumber is not None:
        try:
            with pdfplumber.open(caminho) as pdf:
                paginas = pdf.pages if MAX_PAGINAS_PDF is None else pdf.pages[:MAX_PAGINAS_PDF]
                for numero, pg in enumerate(paginas, 1):
                    try:
                        words = pg.extract_words(x_tolerance=2, y_tolerance=2)
                    except Exception as e:
                        log.debug(f"Falha lendo página {numero}: {e}")
                        continue
                    if not words:
                        continue
                    linhas = agrupar_linhas(words, tol=3)
                    textos = [l["texto"] for l in linhas]
                    debug.append(f"--- página {numero} ---\n" + "\n".join(textos))

                    for ini, fim in secoes_com_plano_b(textos):
                        docs = documentos_no_texto("\n".join(textos[ini:fim]))
                        if docs and doc_bloqueado(docs[0]):
                            res.motivo = "nota emitida pela própria empresa (prestador = própria empresa)"
                            res.texto_debug = "\n".join(debug)
                            return res
                        try:
                            nome, como = nome_na_secao(linhas, ini, fim)
                        except PrestadorEhEmpresaPropria:
                            res.motivo = "nota emitida pela própria empresa (nome do prestador = própria empresa)"
                            res.texto_debug = "\n".join(debug)
                            return res
                        if nome:
                            res.nome, res.metodo = nome, f"{como} | pág {numero}"
                            return res
        except Exception as e:
            log.debug(f"Erro pdfplumber em {caminho}: {e}")

    if not res.nome:
        log.info(f"Tentando OCR nativo para: {os.path.basename(caminho)}")
        try:
            total_pags = 1
            if pdfium is not None:
                try:
                    pdf = pdfium.PdfDocument(caminho)
                    total_pags = len(pdf) if MAX_PAGINAS_PDF is None else min(len(pdf), MAX_PAGINAS_PDF)
                except Exception:
                    total_pags = 1

            for num_pg in range(total_pags):
                ocr_words = executar_ocr_pagina(caminho, num_pagina=num_pg)
                if not ocr_words:
                    continue
                linhas_ocr = agrupar_linhas(ocr_words, tol=8)
                textos_ocr = [l["texto"] for l in linhas_ocr]
                debug.append(f"--- página {num_pg + 1} (OCR) ---\n" + "\n".join(textos_ocr))

                for ini, fim in secoes_com_plano_b(textos_ocr):
                    docs = documentos_no_texto("\n".join(textos_ocr[ini:fim]))
                    if docs and doc_bloqueado(docs[0]):
                        res.motivo = "nota emitida pela própria empresa (prestador = própria empresa)"
                        res.texto_debug = "\n".join(debug)
                        return res
                    try:
                        nome, como = nome_na_secao(linhas_ocr, ini, fim)
                    except PrestadorEhEmpresaPropria:
                        res.motivo = "nota emitida pela própria empresa (nome do prestador = própria empresa)"
                        res.texto_debug = "\n".join(debug)
                        return res
                    if nome:
                        res.nome, res.metodo = nome, f"OCR | {como} | pág {num_pg + 1}"
                        return res
        except Exception as erro_ocr:
            log.debug(f"Falha OCR em {caminho}: {erro_ocr}")

    res.texto_debug = "\n".join(debug)
    if not debug:
        res.motivo = "PDF sem texto e OCR não retornou caracteres"
    else:
        res.motivo = "nome do prestador não encontrado com segurança"
    return res


# ============================================================
# XML
# ============================================================

def tag_local(tag) -> str:
    return (tag or "").split("}")[-1].lower()


BLOCOS_PRESTADOR = [
    "emit", "emitente", "emissor",
    "prestadorservico", "prestadorservicos", "identificacaoprestador",
    "prest", "prestador", "fornecedor",
    "dadosprestador", "dados_prestador", "infprestador",
    "dadosemitente", "dados_emitente", "infemitente",
]
TAGS_NOME = [
    "xnome", "razaosocial", "nomeempresarial", "nomerazaosocial",
    "nome", "xfant", "nomefantasia",
]
TAGS_DOC = {"cnpj", "cpf", "cnpjcpf", "cpfcnpj", "documento"}
TAGS_TOMADOR = {
    "toma", "tomador", "tomadorservico", "dest", "destinatario",
    "adquirente", "interm", "intermediario", "intermediarioservico",
    "dadostomador", "dados_tomador", "inftomador",
    "dadosdestinatario", "dados_destinatario", "infdestinatario",
}


def iter_sem_tomador(el):
    yield el
    for filho in el:
        if tag_local(filho.tag) in TAGS_TOMADOR:
            continue
        yield from iter_sem_tomador(filho)


def _nome_no_elemento(el):
    achados = {}
    for x in iter_sem_tomador(el):
        t = tag_local(x.tag)
        if t in TAGS_NOME and x.text and x.text.strip() and t not in achados:
            achados[t] = x.text
    for t in TAGS_NOME:                      
        if t in achados:
            n = limpar_nome(achados[t], confiavel=True)
            if n:
                return n
    return None


def extrair_nome_xml(caminho: str) -> Resultado:
    res = Resultado()
    try:
        root = ET.parse(caminho).getroot()
    except Exception as e:
        res.motivo = f"XML inválido: {e}"
        return res

    pai = {filho: p for p in root.iter() for filho in p}

    try:
        for tag_bloco in BLOCOS_PRESTADOR:
            for bloco in root.iter():
                if tag_local(bloco.tag) != tag_bloco:
                    continue

                for x in iter_sem_tomador(bloco):
                    if tag_local(x.tag) in TAGS_DOC and x.text:
                        d = limpar_documento(x.text)
                        if d and doc_bloqueado(d):
                            res.motivo = "nota emitida pela própria empresa (emitente = própria empresa)"
                            return res

                nome = _nome_no_elemento(bloco)
                como = f"bloco <{tag_bloco}>"
                if not nome and bloco in pai:
                    nome = _nome_no_elemento(pai[bloco])
                    como = f"vizinho de <{tag_bloco}>"
                if nome:
                    res.nome, res.metodo = nome, como
                    return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa (nome do emitente = própria empresa)"
        return res

    try:
        for table in root.iter():
            if tag_local(table.tag) == "table":
                rows = [r for r in table if tag_local(r.tag) == "row"]
                for r_idx, r in enumerate(rows):
                    cells = [c.text.strip() for c in r if c.text and c.text.strip()]
                    if any(re.search(r"\b(tomador|adquirente|destinatario)\b", normalizar_texto(c)) for c in cells):
                        break
                    for c_idx, text in enumerate(cells):
                        if re.match(r"^n\s*ome\s*[/|1iI\\-]?\s*(?:nome\s*empresarial|razao\s*social)", normalizar_texto(text)):
                            if r_idx + 1 < len(rows):
                                next_cells = [c.text.strip() for c in rows[r_idx + 1] if c.text and c.text.strip()]
                                if c_idx < len(next_cells):
                                    n = limpar_nome(next_cells[c_idx], confiavel=True)
                                    if n:
                                        res.nome, res.metodo = n, "tabela XML (coluna de Nome Empresarial)"
                                        return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa (tabela XML)"
        return res

    try:
        textos = []
        for el in root.iter():
            t = (el.text or "").strip()
            tag = tag_local(el.tag)
            if t and tag not in ("tag", "brush", "size", "point", "font"):
                textos.append(t)

        full_str = " ".join(textos)
        idx_tomador = -1
        for m in re.finditer(r"\b(tomador|adquirente|destinatario)\b", full_str, re.IGNORECASE):
            idx_tomador = m.start()
            break

        secao_prest = full_str[:idx_tomador] if idx_tomador != -1 else full_str
        m_rotulo = re.search(
            r"nome\s*[/|1iI\\-]?\s*(?:nome\s*empresarial|razao\s*social)\s*(?:[0-9\.\,\-\s]{8,20})?\s*([A-Za-z0-9\s]{3,60}?)(?=\s*(?:e\s*-\s*mail|endereco|municipio|simples|cpf|cnpj|telefone|$))",
            secao_prest,
            re.IGNORECASE,
        )
        if m_rotulo:
            n = limpar_nome(m_rotulo.group(1), confiavel=True)
            if n:
                res.nome, res.metodo = n, "texto de layout XML (<Text>)"
                return res
    except PrestadorEhEmpresaPropria:
        res.motivo = "nota emitida pela própria empresa (layout XML)"
        return res

    res.motivo = "nome do emitente/prestador não encontrado no XML"
    return res


# ============================================================
# CSV
# ============================================================

def _abrir_csv(caminho: str):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(caminho, "r", encoding=enc, newline="") as f:
                conteudo = f.read()
            amostra = conteudo[:5000]
            try:
                delim = csv.Sniffer().sniff(amostra, delimiters=";,\t").delimiter
            except csv.Error:
                delim = ";" if ";" in amostra else ("," if "," in amostra else "\t")
            return list(csv.reader(conteudo.splitlines(), delimiter=delim))
        except UnicodeDecodeError:
            continue
    return []


def _coluna_nome_prestador(cabecalho: str) -> bool:
    n = normalizar_texto(cabecalho)
    if any(x in n for x in ("tomador", "adquirente", "destinatario", "intermediario")):
        return False
    if "cnpj" in n or "cpf" in n:
        return False
    if n in ("prestador", "fornecedor", "emitente"):
        return True
    return any(x in n for x in ("prestador", "fornecedor", "emitente")) and (
        "razao" in n or "nome" in n
    )


def extrair_nome_csv(caminho: str) -> Resultado:
    res = Resultado()
    linhas = _abrir_csv(caminho)
    if len(linhas) < 2:
        res.motivo = "CSV vazio ou ilegível"
        return res

    if len(linhas) == 2 and linhas[1] and str(linhas[1][0]).strip().lower() == "total":
        res.motivo = "relatório de NFTS sem movimentação (total 0)"
        return res

    for i, linha in enumerate(linhas[:30]):
        indices = [k for k, c in enumerate(linha) if _coluna_nome_prestador(c)]
        if not indices:
            continue
        nomes = []
        try:
            for dados in linhas[i + 1:]:
                for k in indices:
                    if k < len(dados):
                        n = limpar_nome(dados[k])
                        if n and n not in nomes:
                            nomes.append(n)
        except PrestadorEhEmpresaPropria:
            res.motivo = "nota emitida pela própria empresa (prestador = própria empresa)"
            return res
        if len(nomes) == 1:
            res.nome, res.metodo = nomes[0], "coluna do prestador no CSV"
        elif len(nomes) > 1:
            res.motivo = f"CSV com {len(nomes)} prestadores diferentes: ambíguo"
        else:
            res.motivo = "coluna do prestador sem nome válido"
        return res

    res.motivo = "coluna de nome do prestador não encontrada"
    return res


# ============================================================
# DEBUG
# ============================================================

def salvar_debug(nome_arquivo: str, res: Resultado):
    if not SALVAR_DEBUG or not res.texto_debug:
        return
    try:
        os.makedirs(PASTA_DEBUG, exist_ok=True)
        seguro = re.sub(r'[\\/:*?"<>|]', "_", nome_arquivo)
        with open(os.path.join(PASTA_DEBUG, f"debug_{seguro}.txt"), "w", encoding="utf-8") as f:
            f.write(f"ARQUIVO: {nome_arquivo}\nMOTIVO: {res.motivo}\n\n{'=' * 80}\n\n")
            f.write(res.texto_debug[:30000])
    except Exception:
        pass


# ============================================================
# RENOMEAÇÃO
# ============================================================

_reservados = set()
_relatorio = []


def arquivo_pronto(caminho: str) -> bool:
    """Evita ler arquivo que o OneDrive ainda está baixando."""
    if not VERIFICAR_SINCRONIZACAO:
        return True
    try:
        t1 = os.path.getsize(caminho)
        time.sleep(0.3)
        t2 = os.path.getsize(caminho)
        if t1 != t2 or t2 == 0:
            return False
        with open(caminho, "rb") as f:
            f.read(1)
        return True
    except Exception:
        return False


def escolher_destino(pasta: str, base: str, ext: str, original: str) -> Optional[str]:
    n = 1
    while True:
        nome = base if n == 1 else f"{base}_{n}"
        cand = os.path.join(pasta, nome + ext)
        if _chave(cand) == _chave(original):
            return None
        if not os.path.exists(cand) and _chave(cand) not in _reservados:
            _reservados.add(_chave(cand))
            return cand
        n += 1


def registrar(original, novo, status, res: Optional[Resultado], caminho):
    _relatorio.append({
        "pasta": os.path.dirname(caminho),
        "arquivo_original": original,
        "arquivo_novo": novo,
        "status": status,
        "nome_encontrado": res.nome if res and res.nome else "",
        "como_encontrou": res.metodo if res else "",
        "motivo": res.motivo if res else "",
    })


def processar_arquivo(caminho: str):
    nome = os.path.basename(caminho)
    ext = os.path.splitext(nome)[1]
    if ext.lower() not in EXTENSOES_VALIDAS or nome.startswith("~$"):
        return

    if arquivo_ja_processado(caminho):
        log.debug(f"PULADO (já no histórico) | {nome}")
        registrar(nome, nome, "PULADO (já no histórico)", None, caminho)
        return

    if not arquivo_pronto(caminho):
        log.warning(f"MANTIDO (arquivo ainda sincronizando/ilegível) | {nome}")
        registrar(nome, nome, "MANTIDO", Resultado(motivo="arquivo não está pronto"), caminho)
        return

    if nome.lower() in ARQUIVOS_SEM_AVISO:
        log.info(f"IGNORADO (não é nota fiscal) | {nome}")
        registrar(nome, nome, "NÃO É NOTA FISCAL (ignorado)", Resultado(motivo="documento interno / comprovante"), caminho)
        registrar_no_historico(caminho, nome_original=nome, status="NÃO É NOTA FISCAL (ignorado)")
        return

    try:
        e = ext.lower()
        if e == ".pdf":
            res = extrair_nome_pdf(caminho)
        elif e == ".xml":
            res = extrair_nome_xml(caminho)
        else:
            res = extrair_nome_csv(caminho)
    except Exception as erro:
        log.error(f"Erro processando {nome}: {erro}")
        registrar(nome, nome, "ERRO (mantido)", Resultado(motivo=str(erro)), caminho)
        return

    if not res.nome:
        if "relatório de nfts" in res.motivo.lower():
            log.info(f"MANTIDO (relatório sem lançamentos) | {nome}")
            registrar(nome, nome, "RELATÓRIO VAZIO (sem notas)", res, caminho)
            registrar_no_historico(caminho, nome_original=nome, status="RELATÓRIO VAZIO (sem notas)")
        else:
            log.warning(f"MANTIDO (nome não encontrado) | {nome} | {res.motivo}")
            salvar_debug(nome, res)
            registrar(nome, nome, "MANTIDO", res, caminho)
        return

    destino = escolher_destino(os.path.dirname(caminho), res.nome, ext, caminho)
    if destino is None:
        registrar(nome, nome, "JÁ NO PADRÃO", res, caminho)
        registrar_no_historico(caminho, nome_original=nome, status="JÁ NO PADRÃO", nome_fornecedor=res.nome)
        return
    novo_nome = os.path.basename(destino)

    if DRY_RUN:
        log.info(f"[SIMULAÇÃO] {nome} -> {novo_nome}  ({res.metodo})")
        registrar(nome, novo_nome, "SIMULADO", res, caminho)
        return

    try:
        os.rename(caminho, destino)
        log.info(f"RENOMEADO | {nome} -> {novo_nome}  ({res.metodo})")
        registrar(nome, novo_nome, "RENOMEADO", res, caminho)
        _historico_dados.pop(_chave(caminho), None)
        registrar_no_historico(destino, nome_original=nome, status="RENOMEADO", nome_fornecedor=res.nome)
    except Exception as erro:
        log.error(f"Erro renomeando {nome}: {erro}")
        registrar(nome, nome, "ERRO (mantido)", Resultado(motivo=str(erro)), caminho)


def processar_pasta(pasta_base: str):
    global _historico_dados
    if not os.path.isdir(pasta_base):
        log.error("Pasta não encontrada. Verifique a variável de ambiente NOTAS_PASTA_BASE.")
        return

    _historico_dados = carregar_historico()
    if USAR_HISTORICO and not FORCAR_REPROCESSAMENTO:
        log.info(f"Histórico permanente carregado: {len(_historico_dados)} arquivo(s) registrado(s).")

    debug_abs = _chave(PASTA_DEBUG)
    total = 0
    parar = False
    arquivos_desde_salvamento = 0

    for raiz, subpastas, arquivos in os.walk(pasta_base):
        subpastas[:] = [p for p in subpastas if _chave(os.path.join(raiz, p)) != debug_abs]
        for arq in sorted(arquivos):
            if not arq.lower().endswith(EXTENSOES_VALIDAS):
                continue
            if MAX_ARQUIVOS_TESTE is not None and total >= MAX_ARQUIVOS_TESTE:
                log.info(f"LIMITE DE TESTE ATINGIDO ({MAX_ARQUIVOS_TESTE} arquivos).")
                parar = True
                break
            total += 1
            processar_arquivo(os.path.join(raiz, arq))

            arquivos_desde_salvamento += 1
            if arquivos_desde_salvamento >= 50:
                salvar_historico()
                arquivos_desde_salvamento = 0

        if parar:
            break

    salvar_historico()

    try:
        with open(ARQUIVO_RELATORIO, "w", encoding="utf-8-sig", newline="") as f:
            campos = ["pasta", "arquivo_original", "arquivo_novo", "status",
                      "nome_encontrado", "como_encontrou", "motivo"]
            w = csv.DictWriter(f, fieldnames=campos, delimiter=";")
            w.writeheader()
            w.writerows(_relatorio)
    except Exception as e:
        log.error(f"Não consegui salvar o relatório: {e}")

    contagem = {}
    for r in _relatorio:
        contagem[r["status"]] = contagem.get(r["status"], 0) + 1
    log.info("=" * 80)
    log.info(f"CONCLUÍDO | {total} arquivo(s) analisado(s)")
    for status, qtd in sorted(contagem.items()):
        log.info(f"  {status}: {qtd}")
    log.info(f"Relatório: {ARQUIVO_RELATORIO}")
    if USAR_HISTORICO:
        log.info(f"Histórico permanente: {ARQUIVO_HISTORICO} ({len(_historico_dados)} registros)")
    log.info("MODO SIMULAÇÃO: nada foi alterado." if DRY_RUN else "MODO REAL: arquivos renomeados.")


if __name__ == "__main__":
    log.info("=" * 80)
    log.info("RENOMEAÇÃO PELO NOME DO FORNECEDOR (sem CNPJ)")
    log.info("Pasta: definida via NOTAS_PASTA_BASE")
    log.info("Trava da própria empresa: ATIVA (CNPJ e termos de nome configurados por variável de ambiente)")
    log.info(f"Nome máximo: {NOME_MAX_CARACTERES} caracteres")
    log.info(f"Histórico permanente: {'ATIVADO' if USAR_HISTORICO else 'DESATIVADO'}")
    log.info(f"Modo: {'SIMULAÇÃO (DRY_RUN)' if DRY_RUN else 'RENOMEAÇÃO REAL'}")
    log.info("=" * 80)
    processar_pasta(PASTA_BASE)
    try:
        input("\nPressione ENTER para fechar...")
    except EOFError:
        pass