# <option> define itens dentro de uma lista, é uma tag de HTML

from __future__ import annotations

import re
from html import unescape
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


BASE_URL = "https://areacliente.mssistemas.com.br" # Url de onde vem os dados
CLASSES_URL = f"{BASE_URL}/Conveniados/lista/convenios/" # Busca os convenios dos sites
SUBCLASSES_URL = f"{BASE_URL}/Conveniados/getSubclasse/" # Busca as subclasses
OUTPUT_XLSX = Path(__file__).with_name("gabetta_explorer.xlsx") # Carrega o arquivo já existente 

HEADERS = { # Faz a pesquisa como se fosse pelo navegador pra não dar B.O de request
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

OPTION_RE = re.compile(  #<option value="12">Alimentação</option> extrai o valor 12 || usamos para ler classes e subclasses
    r"<option\b[^>]*value=(['\"]?)(.*?)\1[^>]*>(.*?)</option>",
    re.IGNORECASE | re.DOTALL,
)
CARD_RE = re.compile( # Usamos para ler o nome e cidade
    r'<div class="conveniados-lista-item">(.*?)</div>\s*</a>',
    re.IGNORECASE | re.DOTALL,
)

KNOWN_CONVENIOS = { # Codigo dos convenios que achamos dentro da baseurl
    851: {
        "convenio_nome": "Santa Luzia Funeraria",
        "evidencia": "Logo extraido do login: 'Santa Luzia funeraria'.",
    },
    852: {
        "convenio_nome": "Grupo Nacional Pax",
        "evidencia": "Logo e redes sociais no Fale Conosco: instagram/gruponacionalpax.",
    },
    860: {
        "convenio_nome": "Interplan Assistencia Funeral Familiar",
        "evidencia": "Logo extraido do login: 'INTERPLAN assistencia funeral familiar'.",
    },
    861: {
        "convenio_nome": "Plan Minas / Viva Feliz",
        "evidencia": "Assets do login mostram 'PLAN MINAS' e fundo com marca 'viva feliz'.",
    },
    868: {
        "convenio_nome": "Fenix PAFF",
        "evidencia": "css_login.css aponta para logo externo em fenixpaff.com.br.",
    },
    872: {
        "convenio_nome": "BHMINAS Plano Assistencial Funerario",
        "evidencia": "Logo extraido do login: 'BHMINAS Plano Assistencial Funerario'.",
    },
    874: {
        "convenio_nome": "Grupo Gabetta",
        "evidencia": "Logo extraido do login e codigo exposto no site publico do Grupo Gabetta.",
    },
    880: {
        "convenio_nome": "Flor de Lotus Adamantina",
        "evidencia": "Links sociais no Fale Conosco: facebook/instagram flordelotusadamantina.",
    },
    890: {
        "convenio_nome": "A Viagem Plano de Protecao Familiar",
        "evidencia": "Logo extraido do login e redes sociais 'planoaviagem'.",
    },
    894: {
        "convenio_nome": "Grupo Sefex Assistencia Familiar",
        "evidencia": "Logo extraido do site: 'GRUPO SEFEX Assistencia Familiar'.",
    },
    898: {
        "convenio_nome": "TerraBrasil Organizacao Funeraria",
        "evidencia": "Logo extraido do login: 'terrabrasil Organizacao Funeraria'.",
    },
    899: {
        "convenio_nome": "TerraBrasil Organizacao Funeraria",
        "evidencia": "Mesmo padrao visual, telefone e logo do codigo 898.",
    },
} # OBS: Não é comprovado que o nome de cada convenio, supomos por conta do logo


def fetch_html(url: str, params: dict | None = None) -> str: # Faz a requisição http usando os headres montando em modelo url
    if params:
        url = f"{url}?{urlencode(params)}"

    request = Request(url, headers=HEADERS)
    with urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8", errors="ignore")


def clean_text(text: str) -> str: #  Limpa textos vindos do HTML
    text = unescape(text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_title(html: str) -> str: # Extrai o titulo do HTML
    match = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    return clean_text(match.group(1)) if match else ""


def extract_brand(html: str) -> str: # Ajuda a ver se a página tem algum nome de marca no cabeçalho
    match = re.search(
        r'<a class="navbar-brand"[^>]*>\s*(.*?)\s*</a>',
        html,
        re.IGNORECASE | re.DOTALL,
    )
    return clean_text(match.group(1)) if match else ""


def count_options(html: str) -> int: # Conta quantas tags no caso os <option> tem no html
    return len(OPTION_RE.findall(html))


def parse_options(html: str, ignored_labels: set[str] | None = None) -> list[dict]: # Pega os <option> e transforma em lista de dicionarios
    ignored = {label.casefold() for label in (ignored_labels or set())}
    options: list[dict] = []

    for _, value, label in OPTION_RE.findall(html):
        value = clean_text(value)
        label = clean_text(label)

        if not value or not label:
            continue
        if label.casefold() in ignored:
            continue

        options.append({"id": value, "nome": label})

    return options


def extract_classes(codigo: int) -> list[dict]: # Busca todas as classes
    html = fetch_html(CLASSES_URL, params={"codigo": codigo})
    match = re.search(
        r'<select[^>]*name=["\']classe["\'][^>]*>(.*?)</select>',
        html,
        re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return []

    return parse_options(match.group(1), ignored_labels={"Selecione"})


def extract_subclasses(codigo: int, classe_id: str) -> list[dict]: # Busca as subclasses de uma classe específica
    html = fetch_html(
        SUBCLASSES_URL,
        params={"classe": classe_id, "codigo": codigo, "subclasse": ""},
    )
    return parse_options(
        html,
        ignored_labels={"Selecione a Classe", "Todas", "Selecione"},
    )


def extract_name(card_html: str) -> str: # Pega, de dentro do card, o nome exibido do conveniado
    match = re.search(
        r'<div class="conveniados-lista-nome">\s*(.*?)\s*<div class="span float-right">',
        card_html,
        re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return ""
    return clean_text(match.group(1)) # DOCES DIVERSOS - CAFÉ COM DELICIAS vira CAFÉ COM DELICIAS


def extract_city(card_html: str) -> str: # Extrai a cidade
    match = re.search(
        r'<span class="conveniados-lista-cidade">\s*(.*?)\s*</span>',
        card_html,
        re.IGNORECASE | re.DOTALL,
    )
    if not match:
        return ""
    return clean_text(match.group(1))


def normalize_place_name(nome: str, subclasse: str) -> str: # Ela remove o nome antes do prefixo
    nome_limpo = nome.strip(" -")
    subclasse_limpa = subclasse.strip()
    prefixos = [
        f"{subclasse_limpa} - ",
        f"{subclasse_limpa}-",
        f"{subclasse_limpa}  - ",
    ]

    nome_upper = nome_limpo.upper()
    for prefixo in prefixos:
        if nome_upper.startswith(prefixo.upper()):
            return nome_limpo[len(prefixo):].strip(" -")

    partes = re.split(r"\s+-\s+", nome_limpo, maxsplit=1)
    if len(partes) == 2:
        _, restante = partes
        if restante:
            return restante.strip(" -")

    return nome_limpo


def extract_conveniados(codigo: int, classe: dict, subclasse: dict) -> list[dict]: # Organiza tudo
    html = fetch_html(
        CLASSES_URL, # Dando o request la nos conveios do site
        params={ # Trás os parametros que pedi dos clientes
            "codigo": codigo,
            "classe": classe["id"],
            "subclasse": subclasse["id"],
        },
    )

    registros: list[dict] = [] # Passando como lista
    for card_html in CARD_RE.findall(html): # Extraindo nome e cidade
        nome = extract_name(card_html)
        cidade = extract_city(card_html)
        if not nome or not cidade:
            continue

        registros.append( # Aqui junta tudo
            {
                "Nome do lugar": normalize_place_name(nome, subclasse["nome"]),
                "classe": classe["nome"],
                "subclasse": subclasse["nome"],
                "cidade": cidade,
            }
        )

    return registros


def montar_df_conveniados_codigo(codigo: int) -> pd.DataFrame: # Monta o df
    registros: list[dict] = [] # Transforma em lista

    for classe in extract_classes(codigo): # Busca as classes
        for subclasse in extract_subclasses(codigo, classe["id"]): # Para cada classe as subclasses
            registros.extend(extract_conveniados(codigo, classe, subclasse)) # Para cada subclasse os conveniados delas

    df = pd.DataFrame(registros) # Junta tudo em uma lista e converte pra df
    if df.empty:
        return pd.DataFrame(columns=["Nome do lugar", "classe", "subclasse", "cidade"])

    return df.drop_duplicates().sort_values( # Remove duplicados e ordena
        ["classe", "subclasse", "cidade", "Nome do lugar"],
        ignore_index=True,
    )


def build_sheet_name(codigo: int, convenio_nome: str) -> str: # Monta o noem da aba do excel
    base = f"{codigo} - {convenio_nome}" if convenio_nome else str(codigo)
    base = re.sub(r'[:\\/?*\[\]]', "-", base)
    return base[:31]


def probe_codigo(codigo: int) -> dict: # Ele analisa a estrutura basica do codigo para ver se é valido ou se é uma cópia de outro codigo
    html = fetch_html(CLASSES_URL, params={"codigo": codigo})
    has_classe_select = 'name="classe"' in html or "name='classe'" in html
    preview = clean_text(html[:250])
    custom_match = re.search(r"/assets/custom/(\d+)/", html)

    return {
        "codigo": codigo,
        "tem_select_classe": has_classe_select,
        "qtd_options": count_options(html),
        "qtd_cards": html.count("conveniados-lista-item"),
        "custom_asset_codigo": custom_match.group(1) if custom_match else "",
        "title": extract_title(html),
        "brand": extract_brand(html),
        "preview": preview,
    }


def scan_codigo_range(start: int = 850, end: int = 900) -> pd.DataFrame: # Para cada codigo ele chama o probe_codigo || Serve praticamente para uma varredura
    registros: list[dict] = []

    for codigo in range(start, end + 1):
        try:
            registros.append(probe_codigo(codigo))
        except Exception as exc:
            registros.append(
                {
                    "codigo": codigo,
                    "tem_select_classe": False,
                    "qtd_options": 0,
                    "qtd_cards": 0,
                    "custom_asset_codigo": "",
                    "title": "",
                    "brand": "",
                    "preview": f"erro: {exc}",
                }
            )

    df = pd.DataFrame(registros)
    return df.sort_values(["codigo"], ignore_index=True)


def mapear_convenios(df_codigos: pd.DataFrame) -> pd.DataFrame: # Pega o resultado da varredura e se o codigo estiver no KNOWN_CONVENIOS trás o nome e evidencia, se não, deixa ele vazaio
    registros: list[dict] = []

    for _, row in df_codigos.iterrows():
        codigo = int(row["codigo"])
        known = KNOWN_CONVENIOS.get(codigo)

        if known:
            registros.append(
                {
                    "codigo": codigo,
                    "convenio_nome": known["convenio_nome"],
                    "status": "identificado",
                    "evidencia": known["evidencia"],
                }
            )
        else:
            preview = str(row.get("preview", ""))
            status = "empresa_nao_encontrada" if "Empresa não encontrada" in preview else "nao_identificado"
            registros.append(
                {
                    "codigo": codigo,
                    "convenio_nome": "",
                    "status": status,
                    "evidencia": preview[:160],
                }
            )

    return pd.DataFrame(registros).sort_values(["codigo"], ignore_index=True)


def build_duplicate_signature(row: pd.Series) -> str: # Aqui serve para acharmos as duplicadas dos codigos
    return "|".join(
        [
            str(bool(row.get("tem_select_classe", False))),
            str(int(row.get("qtd_options", 0))),
            str(int(row.get("qtd_cards", 0))),
        ]
    )


def build_initial_sheet(df_codigos: pd.DataFrame, df_convenios: pd.DataFrame) -> pd.DataFrame: # Monta a aba inicial
    base = df_codigos.merge(df_convenios, on="codigo", how="left")
    base = base.loc[base["tem_select_classe"]].copy() # Filtra só códigos que têm estrutura de classe
    base["assinatura"] = base.apply(build_duplicate_signature, axis=1) # Cria a assinatura de duplicidade

    registros: list[dict] = []

    for _, grupo in base.groupby("assinatura", sort=True): # Agrupa os códigos parecidos
        grupo = grupo.sort_values(["codigo"], ignore_index=True)
        identificados = grupo.loc[grupo["convenio_nome"].fillna("").str.strip() != ""]

        if not identificados.empty:
            oficial = identificados.sort_values(["codigo"], ignore_index=True).iloc[0]
        else:
            oficial = grupo.iloc[0]

        codigo_oficial = int(oficial["codigo"]) # Escolhe um codigo oficial para o grupo
        convenio_nome = str(oficial.get("convenio_nome", "")).strip()
        codigos_adjacentes = [ # Coloca os outros codigos adjascentes
            str(int(codigo))
            for codigo in grupo["codigo"].tolist()
            if int(codigo) != codigo_oficial
        ]

        registros.append( # Resultado final
            {
                "codigo": codigo_oficial,
                "convenio_nome": convenio_nome,
                "status": str(oficial.get("status", "")),
                "codigos adjascentes": ", ".join(codigos_adjacentes),
                "qtd_options": int(oficial.get("qtd_options", 0)),
                "qtd_cards": int(oficial.get("qtd_cards", 0)),
                "evidencia": str(oficial.get("evidencia", "")),
            }
        )

    df_inicial = pd.DataFrame(registros)
    if df_inicial.empty:
        return df_inicial

    return df_inicial.sort_values(["codigo"], ignore_index=True)


if __name__ == "__main__": # Aqui é onde unimos tudo no fim das contas
    df_codigos = scan_codigo_range(850, 900) # Varre os codigos
    df_convenios = mapear_convenios(df_codigos) # Mapeia os convenios
    df_inicial = build_initial_sheet(df_codigos, df_convenios) # Monta a tabela inicial
    codigos_para_exportar = df_inicial["codigo"].astype(int).tolist()

    with pd.ExcelWriter(OUTPUT_XLSX) as writer: # Ajuda a salvar o arquivo excel sem que de b.o
        df_inicial.to_excel(writer, sheet_name="Inicial", index=False)
        df_codigos.to_excel(writer, sheet_name="scan_850_900", index=False)
        df_convenios.to_excel(writer, sheet_name="convenios_nomeados", index=False)
        used_sheet_names: set[str] = set()

        for codigo in codigos_para_exportar:
            convenio_nome = ""
            row = df_inicial.loc[df_inicial["codigo"] == codigo, "convenio_nome"]
            if not row.empty:
                convenio_nome = str(row.iloc[0]).strip()

            sheet_name = build_sheet_name(codigo, convenio_nome)
            original_sheet_name = sheet_name
            suffix = 1
            while sheet_name in used_sheet_names:
                suffix_str = f"_{suffix}"
                sheet_name = f"{original_sheet_name[:31 - len(suffix_str)]}{suffix_str}"
                suffix += 1
            used_sheet_names.add(sheet_name)

            df_sheet = montar_df_conveniados_codigo(codigo)
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
    # Prints da vida
    print("Convenios identificados na faixa 850-900:")
    print(df_convenios.loc[df_convenios["status"] == "identificado", ["codigo", "convenio_nome"]])
    print()
    print("Codigos oficiais na aba Inicial:")
    print(df_inicial[["codigo", "convenio_nome", "codigos adjascentes"]])
    print()
    print("Abas de convenios geradas:")
    print([build_sheet_name(codigo, str(df_inicial.loc[df_inicial['codigo'] == codigo, 'convenio_nome'].iloc[0]).strip()) for codigo in codigos_para_exportar])
    print()
    print(f"Arquivo salvo em: {OUTPUT_XLSX}")
