# -*- coding: utf-8 -*-

"""
AUTOMAÇÃO COMPLETA — VAGAS UNISYS / WORKDAY

Não requer:
    - Playwright
    - Selenium
    - requests
    - pandas
    - openpyxl
    - navegador aberto
    - HAR
    - intervenção manual

Usa somente bibliotecas nativas do Python.

O programa:
    1. Acessa o portal público da Unisys.
    2. Obtém cookies da sessão.
    3. Obtém o token CSRF, quando fornecido.
    4. Consulta a API CXS do Workday.
    5. Descobre quantas vagas existem.
    6. Percorre todos os offsets automaticamente.
    7. Extrai os dados das vagas.
    8. Remove duplicidades.
    9. Cria um arquivo XLSX sem depender de nenhuma biblioteca externa.

Endpoint descoberto no HAR:

POST
https://unisys.wd5.myworkdayjobs.com/wday/cxs/unisys/External/jobs

Payload:

{
    "appliedFacets": {},
    "limit": 20,
    "offset": 0,
    "searchText": ""
}
"""

import json
import re
import sys
import time
import html
import zipfile
from io import BytesIO
from datetime import datetime
from urllib.request import (
    Request,
    build_opener,
    HTTPBasicAuthHandler,
    HTTPPasswordMgrWithDefaultRealm,
)
from urllib.error import HTTPError, URLError
from http.cookiejar import CookieJar
from xml.sax.saxutils import escape


# ============================================================
# CONFIGURAÇÃO
# ============================================================

PORTAL_URL = (
    "https://unisys.wd5.myworkdayjobs.com/"
    "pt-BR/External"
)

API_URL = (
    "https://unisys.wd5.myworkdayjobs.com/"
    "wday/cxs/unisys/External/jobs"
)

LIMIT = 20

OUTPUT_XLSX = "Unisys_vagas_completa.xlsx"

# Pequena pausa entre páginas.
# Pode ser aumentada para 1.0 ou 2.0 se necessário.
DELAY = 0.5

TIMEOUT = 60

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


# ============================================================
# SESSÃO HTTP
# ============================================================

cookies = CookieJar()

opener = build_opener(
    HTTPBasicAuthHandler(),
)

# Substitui o opener por um que também mantenha cookies.
from urllib.request import HTTPCookieProcessor

opener = build_opener(
    HTTPCookieProcessor(cookies)
)


# ============================================================
# FUNÇÕES HTTP
# ============================================================

def http_get(url):
    """
    Faz GET usando a sessão persistente.
    """

    request = Request(
        url,
        method="GET",
        headers={
            "User-Agent": USER_AGENT,
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
            "Cache-Control": "no-cache",
        },
    )

    with opener.open(request, timeout=TIMEOUT) as response:
        return response.read(), response.headers


def http_post(url, payload, csrf=None):
    """
    POST para o endpoint CXS do Workday.
    """

    body = json.dumps(
        payload,
        ensure_ascii=False,
        separators=(",", ":")
    ).encode("utf-8")

    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
        "Origin": "https://unisys.wd5.myworkdayjobs.com",
        "Referer": PORTAL_URL,
        "X-Requested-With": "XMLHttpRequest",
    }

    if csrf:
        headers["X-CALYPSO-CSRF-TOKEN"] = csrf

    request = Request(
        url,
        data=body,
        method="POST",
        headers=headers,
    )

    with opener.open(request, timeout=TIMEOUT) as response:
        raw = response.read()

        return json.loads(
            raw.decode("utf-8-sig")
        )


# ============================================================
# CSRF
# ============================================================

def get_csrf_token():
    """
    Procura o token CSRF nos cookies.

    Workday/Calypso normalmente utiliza:
        CALYPSO_CSRF_TOKEN

    O nome pode variar; por isso procuramos também
    nomes contendo CSRF.
    """

    for cookie in cookies:
        name = cookie.name.upper()

        if "CSRF" in name:
            return cookie.value

    return None


# ============================================================
# LIMPEZA DE HTML
# ============================================================

def html_to_text(value):
    """
    Converte a descrição HTML da vaga para texto.
    """

    if not value:
        return ""

    text = str(value)

    # Quebras de bloco
    text = re.sub(
        r"<\s*(br|p|div|li|ul|ol|h[1-6]|tr)[^>]*>",
        "\n",
        text,
        flags=re.I,
    )

    text = re.sub(
        r"<\s*/\s*(p|div|li|ul|ol|h[1-6]|tr)\s*>",
        "\n",
        text,
        flags=re.I,
    )

    # Remove tags
    text = re.sub(
        r"<[^>]+>",
        "",
        text,
        flags=re.S,
    )

    # Entidades HTML
    text = html.unescape(text)

    # Normalização
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


# ============================================================
# CÓDIGO REQ
# ============================================================

def get_req(job):
    """
    Procura REQxxxxx no objeto da vaga.
    """

    campos = [
        job.get("externalPath", ""),
        job.get("title", ""),
    ]

    bullet_fields = job.get("bulletFields") or []

    if isinstance(bullet_fields, list):
        campos.extend(
            str(x) for x in bullet_fields
        )

    for campo in campos:

        match = re.search(
            r"\bREQ\d+\b",
            str(campo),
            flags=re.I,
        )

        if match:
            return match.group(0).upper()

    return ""


# ============================================================
# URL DA VAGA
# ============================================================

def get_job_url(external_path):

    if not external_path:
        return ""

    if external_path.startswith("http://"):
        return external_path

    if external_path.startswith("https://"):
        return external_path

    return (
        "https://unisys.wd5.myworkdayjobs.com"
        + external_path
    )


# ============================================================
# EXTRAÇÃO DE UMA PÁGINA
# ============================================================

def extract_page(offset):

    payload = {
        "appliedFacets": {},
        "limit": LIMIT,
        "offset": offset,
        "searchText": "",
    }

    for tentativa in range(1, 5):

        try:

            csrf = get_csrf_token()

            data = http_post(
                API_URL,
                payload,
                csrf,
            )

            if not isinstance(data, dict):
                raise RuntimeError(
                    "Resposta não é JSON objeto."
                )

            if "jobPostings" not in data:
                raise RuntimeError(
                    "Resposta não contém jobPostings."
                )

            return data

        except Exception as e:

            print(
                f"   tentativa {tentativa}/4 "
                f"falhou: {e}"
            )

            if tentativa < 4:
                time.sleep(
                    tentativa * 2
                )
            else:
                raise


# ============================================================
# NORMALIZAÇÃO DA VAGA
# ============================================================

def normalize_job(job, offset):

    return {
        "Descrição da vaga":
            job.get("title", "").strip(),

        "Código da vaga":
            get_req(job),

        "Data da publicação":
            job.get("postedOn", "").strip(),

        "Localização":
            job.get("locationsText", "").strip(),

        "Descrição completa":
            html_to_text(
                job.get(
                    "jobDescription",
                    ""
                )
            ),

        "URL da vaga":
            get_job_url(
                job.get(
                    "externalPath",
                    ""
                )
            ),

        "Offset":
            offset,
    }


# ============================================================
# DEDUPLICAÇÃO
# ============================================================

def deduplicate(rows):

    result = []

    seen_req = set()
    seen_url = set()

    for row in rows:

        req = row["Código da vaga"]
        url = row["URL da vaga"]

        if req:

            if req in seen_req:
                continue

            seen_req.add(req)

        elif url:

            if url in seen_url:
                continue

            seen_url.add(url)

        result.append(row)

    return result


# ============================================================
# XML / XLSX
# ============================================================

def column_name(number):

    result = ""

    while number:

        number, remainder = divmod(
            number - 1,
            26
        )

        result = chr(
            65 + remainder
        ) + result

    return result


def cell_xml(
    reference,
    value,
    style=None
):

    if value is None:
        value = ""

    value = str(value)

    style_xml = (
        f' s="{style}"'
        if style is not None
        else ""
    )

    return (
        f'<c r="{reference}" '
        f't="inlineStr"{style_xml}>'
        f'<is><t xml:space="preserve">'
        f'{escape(value)}'
        f'</t></is></c>'
    )


def worksheet_xml(headers, rows):

    xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns='
        '"http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetViews>',
        '<sheetView workbookViewId="0">',
        '</sheetView>',
        '</sheetViews>',
        '<sheetData>',
    ]

    # Cabeçalho
    xml.append('<row r="1">')

    for index, header in enumerate(
        headers,
        start=1
    ):

        ref = (
            column_name(index)
            + "1"
        )

        xml.append(
            cell_xml(
                ref,
                header,
                style=1
            )
        )

    xml.append("</row>")

    # Dados
    for row_number, row in enumerate(
        rows,
        start=2
    ):

        xml.append(
            f'<row r="{row_number}">'
        )

        for col_number, header in enumerate(
            headers,
            start=1
        ):

            ref = (
                column_name(col_number)
                + str(row_number)
            )

            xml.append(
                cell_xml(
                    ref,
                    row.get(header, "")
                )
            )

        xml.append("</row>")

    xml.extend([
        "</sheetData>",
        "</worksheet>",
    ])

    return "".join(xml)


def create_xlsx(
    filename,
    sheets
):

    workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets>
        <sheet name="Vagas" sheetId="1" r:id="rId1"/>
        <sheet name="Controle da extração" sheetId="2" r:id="rId2"/>
        <sheet name="Paginação" sheetId="3" r:id="rId3"/>
      </sheets>
    </workbook>
    """

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
       Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
       Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2"
       Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
       Target="worksheets/sheet2.xml"/>
      <Relationship Id="rId3"
       Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
       Target="worksheets/sheet3.xml"/>
    </Relationships>
    """

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
       Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
       Target="xl/workbook.xml"/>
    </Relationships>
    """

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

      <fonts count="2">
        <font>
          <sz val="11"/>
          <name val="Calibri"/>
        </font>
        <font>
          <b/>
          <sz val="11"/>
          <name val="Calibri"/>
        </font>
      </fonts>

      <fills count="2">
        <fill>
          <patternFill patternType="none"/>
        </fill>
        <fill>
          <patternFill patternType="gray125"/>
        </fill>
      </fills>

      <borders count="1">
        <border>
          <left/>
          <right/>
          <top/>
          <bottom/>
          <diagonal/>
        </border>
      </borders>

      <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
      </cellStyleXfs>

      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
      </cellXfs>

    </styleSheet>
    """

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

      <Default Extension="rels"
       ContentType="application/vnd.openxmlformats-package.relationships+xml"/>

      <Default Extension="xml"
       ContentType="application/xml"/>

      <Override PartName="/xl/workbook.xml"
       ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>

      <Override PartName="/xl/styles.xml"
       ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>

      <Override PartName="/xl/worksheets/sheet1.xml"
       ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>

      <Override PartName="/xl/worksheets/sheet2.xml"
       ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>

      <Override PartName="/xl/worksheets/sheet3.xml"
       ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>

    </Types>
    """

    with zipfile.ZipFile(
        filename,
        "w",
        compression=zipfile.ZIP_DEFLATED
    ) as z:

        z.writestr(
            "[Content_Types].xml",
            content_types
        )

        z.writestr(
            "_rels/.rels",
            root_rels
        )

        z.writestr(
            "xl/workbook.xml",
            workbook_xml
        )

        z.writestr(
            "xl/_rels/workbook.xml.rels",
            rels_xml
        )

        z.writestr(
            "xl/styles.xml",
            styles_xml
        )

        for index, sheet in enumerate(
            sheets,
            start=1
        ):

            xml = worksheet_xml(
                sheet["headers"],
                sheet["rows"]
            )

            z.writestr(
                f"xl/worksheets/sheet{index}.xml",
                xml
            )


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

def main():

    inicio = datetime.now()

    print()
    print("=" * 72)
    print("UNISYS / WORKDAY — EXTRAÇÃO AUTOMÁTICA")
    print("=" * 72)
    print()

    print("1. Abrindo sessão no portal...")

    try:

        _, headers = http_get(
            PORTAL_URL
        )

    except HTTPError as e:

        print(
            f"Erro HTTP ao abrir portal: "
            f"{e.code} {e.reason}"
        )

        sys.exit(1)

    except URLError as e:

        print(
            "Erro de conexão:"
        )

        print(e)

        sys.exit(1)

    csrf = get_csrf_token()

    if csrf:

        print(
            "   Token CSRF obtido."
        )

    else:

        print(
            "   Nenhum token CSRF encontrado; "
            "tentando mesmo assim."
        )

    print()
    print("2. Consultando offset=0...")

    try:

        first = extract_page(0)

    except Exception as e:

        print()
        print(
            "Não foi possível consultar a API:"
        )
        print(e)

        print()
        print(
            "O Workday pode ter alterado a proteção "
            "da API ou estar exigindo uma sessão "
            "mais específica."
        )

        sys.exit(1)

    total = int(
        first.get(
            "total",
            0
        )
    )

    first_jobs = (
        first.get(
            "jobPostings",
            []
        )
        or []
    )

    if total <= 0:

        print(
            "A API não retornou um total válido."
        )

        sys.exit(1)

    print(
        f"   Total informado pelo Workday: {total}"
    )

    print(
        f"   Primeira página: "
        f"{len(first_jobs)} vagas"
    )

    all_rows = []

    pagination = []

    # --------------------------------------------------------
    # PRIMEIRA PÁGINA
    # --------------------------------------------------------

    for job in first_jobs:

        all_rows.append(
            normalize_job(
                job,
                0
            )
        )

    pagination.append({
        "Página": 1,
        "Offset": 0,
        "Quantidade": len(first_jobs),
        "Status": "OK",
    })

    # --------------------------------------------------------
    # DEMAIS PÁGINAS
    # --------------------------------------------------------

    pagina = 2

    for offset in range(
        LIMIT,
        total,
        LIMIT
    ):

        print(
            f"[{pagina:02d}] "
            f"offset={offset:03d} ...",
            end="",
            flush=True
        )

        try:

            data = extract_page(
                offset
            )

            jobs = (
                data.get(
                    "jobPostings",
                    []
                )
                or []
            )

            for job in jobs:

                all_rows.append(
                    normalize_job(
                        job,
                        offset
                    )
                )

            pagination.append({
                "Página": pagina,
                "Offset": offset,
                "Quantidade": len(jobs),
                "Status": "OK",
            })

            print(
                f" {len(jobs)} vagas"
            )

            # Segurança: se o Workday devolver
            # menos registros, chegamos ao final.
            if len(jobs) < LIMIT:

                break

        except Exception as e:

            pagination.append({
                "Página": pagina,
                "Offset": offset,
                "Quantidade": 0,
                "Status": f"ERRO: {e}",
            })

            print(
                f" ERRO: {e}"
            )

            raise

        pagina += 1

        time.sleep(
            DELAY
        )

    # --------------------------------------------------------
    # DEDUPLICAÇÃO
    # --------------------------------------------------------

    antes = len(all_rows)

    all_rows = deduplicate(
        all_rows
    )

    depois = len(all_rows)

    print()
    print(
        f"Registros recebidos: {antes}"
    )

    print(
        f"Registros únicos:    {depois}"
    )

    # --------------------------------------------------------
    # CONTROLE
    # --------------------------------------------------------

    fim = datetime.now()

    controle = [
        {
            "Campo": "Fonte",
            "Valor": PORTAL_URL,
        },
        {
            "Campo": "Endpoint",
            "Valor": API_URL,
        },
        {
            "Campo": "Total informado pelo Workday",
            "Valor": total,
        },
        {
            "Campo": "Registros recebidos",
            "Valor": antes,
        },
        {
            "Campo": "Registros únicos",
            "Valor": depois,
        },
        {
            "Campo": "Limit",
            "Valor": LIMIT,
        },
        {
            "Campo": "Início",
            "Valor": inicio.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        },
        {
            "Campo": "Fim",
            "Valor": fim.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
        },
        {
            "Campo": "Status",
            "Valor": (
                "OK"
                if depois == total
                else
                "ATENÇÃO — quantidade diferente do total"
            ),
        },
        {
            "Campo": "Observação",
            "Valor": (
                "Data de publicação preservada "
                "exatamente como retornada pelo Workday."
            ),
        },
    ]

    # --------------------------------------------------------
    # XLSX
    # --------------------------------------------------------

    headers_vagas = [
        "Descrição da vaga",
        "Código da vaga",
        "Data da publicação",
        "Localização",
        "Descrição completa",
        "URL da vaga",
        "Offset",
    ]

    headers_controle = [
        "Campo",
        "Valor",
    ]

    headers_pag = [
        "Página",
        "Offset",
        "Quantidade",
        "Status",
    ]

    create_xlsx(
        OUTPUT_XLSX,
        [
            {
                "headers":
                    headers_vagas,
                "rows":
                    all_rows,
            },
            {
                "headers":
                    headers_controle,
                "rows":
                    controle,
            },
            {
                "headers":
                    headers_pag,
                "rows":
                    pagination,
            },
        ],
    )

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    print()
    print("=" * 72)
    print("EXTRAÇÃO CONCLUÍDA")
    print("=" * 72)

    print(
        f"Total Workday:    {total}"
    )

    print(
        f"Registros obtidos: {antes}"
    )

    print(
        f"Registros únicos:  {depois}"
    )

    print(
        f"Arquivo gerado:    {OUTPUT_XLSX}"
    )

    print("=" * 72)
    print()


if __name__ == "__main__":

    main()