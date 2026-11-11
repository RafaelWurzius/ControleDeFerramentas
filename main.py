from flask import Flask, render_template, request, redirect, url_for, flash
import psycopg2, os, uuid, re, unicodedata

from openpyxl.drawing.image import Image as ExcelImage

from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from copy import copy

from io import BytesIO
from flask import send_file
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify


# ============================================================
# CONFIGURAÇÃO DO FLASK
# ============================================================

app = Flask(__name__)

app.secret_key = "uma-chave-secreta-aqui"


# ============================================================
# DIRETÓRIO BASE DO PROJETO
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# ARQUIVOS EXCEL
# ============================================================

CAMINHO_FERRAMENTAS = os.path.join(
    BASE_DIR,
    "ferramentas.xlsx"
)

CAMINHO_DEFEITO = os.path.join(
    BASE_DIR,
    "defeito.xlsx"
)


# ============================================================
# CONFIGURAÇÃO DO BANCO DE DADOS
# ============================================================

DB_CONFIG = {
    "host": "localhost",
    "database": "controle_de_ferramentas",
    "user": "usuario",
    "password": "sua_senha",
    "port": "5432"
}


# ============================================================
# CONFIGURAÇÃO DE UPLOAD
# ============================================================

UPLOAD_FOLDER = os.path.join(
    BASE_DIR,
    "static",
    "uploads"
)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

ALLOWED_EXTENSIONS = {
    "png",
    "jpg",
    "jpeg",
    "webp"
}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ============================================================
# CARREGAR LISTA DE FERRAMENTAS
# ============================================================

def carregar_ferramentas():

    try:

        workbook = load_workbook(
            CAMINHO_FERRAMENTAS,
            read_only=True,
            data_only=True
        )

        sheet = workbook.active

        # Obtém os valores da primeira linha
        cabecalhos = [
            celula.value
            for celula in sheet[1]
        ]

        # Procura a coluna chamada "Nome"
        if "Nome" not in cabecalhos:

            workbook.close()

            print(
                "ERRO: A coluna 'Nome' não foi encontrada "
                "em ferramentas.xlsx."
            )

            return []

        coluna_nome = cabecalhos.index("Nome")

        ferramentas = []

        # Percorre as linhas a partir da segunda
        for linha in sheet.iter_rows(
            min_row=2,
            values_only=True
        ):

            if (
                coluna_nome < len(linha)
                and linha[coluna_nome] is not None
            ):

                ferramentas.append(
                    str(linha[coluna_nome])
                )

        workbook.close()

        return ferramentas

    except FileNotFoundError:

        print(
            "ERRO: O arquivo ferramentas.xlsx "
            "não foi encontrado."
        )

        return []

    except Exception as erro:

        print(
            f"ERRO AO LER ferramentas.xlsx: {erro}"
        )

        return []


# ============================================================
# CONEXÃO COM POSTGRESQL
# ============================================================

def conectar_banco():

    return psycopg2.connect(
        **DB_CONFIG
    )


# ============================================================
# VERIFICAR EXTENSÃO DA IMAGEM
# ============================================================

def extensao_permitida(nome_arquivo):

    if "." not in nome_arquivo:
        return False

    extensao = nome_arquivo.rsplit(
        ".",
        1
    )[1].lower()

    return extensao in ALLOWED_EXTENSIONS

# ============================================================
# INSERIR IMAGEM NO EXECEL
# ============================================================
def inserir_imagem_excel(
    sheet,
    caminho_foto,
    celula,
    buffers_imagens,
    largura_max=120,
    altura_max=80
):
    """
    Insere uma imagem na posição de uma célula do Excel.
    Mantém os buffers vivos até o salvamento do arquivo.
    """

    if not caminho_foto or not os.path.isfile(caminho_foto):
        return False

    try:
        buffer = BytesIO()

        with PILImage.open(caminho_foto) as imagem:
            imagem.thumbnail((largura_max, altura_max))

            # Converte para um formato compatível com o Excel.
            if imagem.mode not in ("RGB", "RGBA"):
                imagem = imagem.convert("RGB")

            imagem.save(buffer, format="PNG")

        buffer.seek(0)

        imagem_excel = XLImage(buffer)
        imagem_excel.width = imagem.width
        imagem_excel.height = imagem.height
        imagem_excel.anchor = celula

        sheet.add_image(imagem_excel)

        # Evita que o buffer seja liberado antes de salvar.
        buffers_imagens.append(buffer)

        return True

    except Exception as erro:
        print(f"Erro ao inserir imagem no Excel: {erro}")
        return False

# ============================================================
# ADICIONAR REGISTRO NO EXCEL
# ============================================================

def adicionar_no_excel(
    ferramenta,
    descricao,
    nome_foto,
    ciclo,
    data
):

    # --------------------------------------------------------
    # Verifica se o arquivo existe
    # --------------------------------------------------------

    if os.path.exists(CAMINHO_DEFEITO):

        workbook = load_workbook(
            CAMINHO_DEFEITO
        )

        sheet = workbook.active

    else:

        # Caso o arquivo ainda não exista,
        # cria um novo arquivo.

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "defeito"

        # Cabeçalho
        sheet.append([
            "ferramenta",
            "descricao",
            "evidencia_foto",
            "ciclo",
            "data"
        ])


    # --------------------------------------------------------
    # Descobre a próxima linha
    # --------------------------------------------------------

    nova_linha = sheet.max_row + 1


    # --------------------------------------------------------
    # Insere os dados
    # --------------------------------------------------------

    sheet.cell(
        row=nova_linha,
        column=1
    ).value = ferramenta

    sheet.cell(
        row=nova_linha,
        column=2
    ).value = descricao

    # A coluna de evidência fica visualmente vazia.
    # A imagem será ancorada nessa posição.
    sheet.cell(
        row=nova_linha,
        column=3
    ).value = None

    sheet.cell(
        row=nova_linha,
        column=4
    ).value = ciclo

    sheet.cell(
        row=nova_linha,
        column=5
    ).value = data


    # --------------------------------------------------------
    # COPIA A FORMATAÇÃO DA LINHA ANTERIOR
    # --------------------------------------------------------

    # Só copia se existir uma linha anterior
    # além do cabeçalho.

    if nova_linha > 2:

        linha_anterior = nova_linha - 1

        # Percorre todas as colunas existentes
        for coluna in range(
            1,
            sheet.max_column + 1
        ):

            origem = sheet.cell(
                row=linha_anterior,
                column=coluna
            )

            destino = sheet.cell(
                row=nova_linha,
                column=coluna
            )


            # ------------------------------------------------
            # Copiar estilo
            # ------------------------------------------------

            if origem.has_style:

                destino._style = copy(
                    origem._style
                )


            # ------------------------------------------------
            # Copiar fonte
            # ------------------------------------------------

            destino.font = copy(
                origem.font
            )


            # ------------------------------------------------
            # Copiar preenchimento
            # ------------------------------------------------

            destino.fill = copy(
                origem.fill
            )


            # ------------------------------------------------
            # Copiar bordas
            # ------------------------------------------------

            destino.border = copy(
                origem.border
            )


            # ------------------------------------------------
            # Copiar alinhamento
            # ------------------------------------------------

            destino.alignment = copy(
                origem.alignment
            )


            # ------------------------------------------------
            # Copiar proteção
            # ------------------------------------------------

            destino.protection = copy(
                origem.protection
            )


            # ------------------------------------------------
            # Copiar formato numérico
            # ------------------------------------------------

            destino.number_format = (
                origem.number_format
            )


    # --------------------------------------------------------
    # Copiar altura da linha anterior
    # --------------------------------------------------------

    if nova_linha > 2:

        linha_anterior = nova_linha - 1

        altura = sheet.row_dimensions[
            linha_anterior
        ].height

        if altura is not None:

            sheet.row_dimensions[
                nova_linha
            ].height = altura

# --------------------------------------------------------
# INSERIR A FOTO NA COLUNA DE EVIDÊNCIA
# --------------------------------------------------------

    buffers_imagens = []

    sheet.column_dimensions["C"].width = 20

    caminho_foto = (
        os.path.join(app.config["UPLOAD_FOLDER"], nome_foto)
        if nome_foto
        else None
    )

    if caminho_foto and os.path.isfile(caminho_foto):
        inseriu = inserir_imagem_excel(
            sheet=sheet,
            caminho_foto=caminho_foto,
            celula=f"C{nova_linha}",
            buffers_imagens=buffers_imagens,
            largura_max=120,
            altura_max=80
        )

        if inseriu:
            # Altura em pontos; ajuste se desejar fotos maiores.
            sheet.row_dimensions[nova_linha].height = 65
    # --------------------------------------------------------
    # Salvar arquivo
    # --------------------------------------------------------

    workbook.save(
        CAMINHO_DEFEITO
    )

    workbook.close()


# ============================================================
# ROTA PRINCIPAL
# ============================================================

@app.route("/")
def index():

    ferramentas = carregar_ferramentas()

    return render_template(
        "index.html",
        ferramentas=ferramentas
    )

# ============================================================
# FUNÇÕES AUXILIARES PARA FILTRAR REGISTROS
# ============================================================

def obter_filtros_registros():
    return {
        "nome": request.args.get("nome", "").strip(),
        "descricao": request.args.get("descricao", "").strip(),
        "data": request.args.get("data", "").strip()
    }


def consultar_registros(filtros):
    conexao = None
    cursor = None

    try:
        conexao = conectar_banco()
        cursor = conexao.cursor()

        consulta = """
            SELECT
                id,
                ferramenta,
                descricao,
                evidencia_foto,
                ciclo,
                data
            FROM defeito
            WHERE 1 = 1
        """

        parametros = []

        # Pesquisa por qualquer trecho do nome
        if filtros["nome"]:
            consulta += " AND ferramenta ILIKE %s"
            parametros.append(f"%{filtros['nome']}%")

        # Pesquisa por qualquer trecho da descrição
        if filtros["descricao"]:
            consulta += " AND descricao ILIKE %s"
            parametros.append(f"%{filtros['descricao']}%")

        # Pesquisa flexível por data
        valor_data = filtros["data"].strip()

        if valor_data:
            # Data completa: 24/10/2026 ou 2026-10-24
            data_completa = re.fullmatch(
                r"(\d{2})/(\d{2})/(\d{4})",
                valor_data
            )

            data_iso = re.fullmatch(
                r"(\d{4})-(\d{2})-(\d{2})",
                valor_data
            )

            ano_mes = re.fullmatch(
                r"(\d{4})-(\d{1,2})",
                valor_data
            )

            ano = re.fullmatch(r"\d{4}", valor_data)
            numero = re.fullmatch(r"\d{1,2}", valor_data)

            if data_completa:
                dia, mes, ano_valor = data_completa.groups()
                valor_iso = f"{ano_valor}-{mes}-{dia}"

                consulta += " AND data::date = %s::date"
                parametros.append(valor_iso)

            elif data_iso:
                consulta += " AND data::date = %s::date"
                parametros.append(valor_data)

            elif ano_mes:
                ano_valor, mes_valor = ano_mes.groups()
                mes_valor = int(mes_valor)

                if not 1 <= mes_valor <= 12:
                    consulta += " AND 1 = 0"
                else:
                    consulta += """
                        AND EXTRACT(YEAR FROM data::date) = %s
                        AND EXTRACT(MONTH FROM data::date) = %s
                    """
                    parametros.extend([int(ano_valor), mes_valor])

            elif ano:
                consulta += " AND EXTRACT(YEAR FROM data::date) = %s"
                parametros.append(int(valor_data))

            elif numero:
                numero_valor = int(valor_data)

                if 1 <= numero_valor <= 12:
                    # Ex.: 08 encontra o dia 8 ou o mês de agosto
                    consulta += """
                        AND (
                            EXTRACT(DAY FROM data::date) = %s
                            OR EXTRACT(MONTH FROM data::date) = %s
                        )
                    """
                    parametros.extend([numero_valor, numero_valor])

                elif 13 <= numero_valor <= 31:
                    consulta += " AND EXTRACT(DAY FROM data::date) = %s"
                    parametros.append(numero_valor)

                else:
                    consulta += " AND 1 = 0"

            else:
                # Permite digitar o nome do mês, como "outubro"
                texto = unicodedata.normalize(
                    "NFKD",
                    valor_data
                ).encode("ASCII", "ignore").decode("ASCII").lower()

                meses = {
                    "janeiro": 1, "jan": 1,
                    "fevereiro": 2, "fev": 2,
                    "marco": 3, "mar": 3,
                    "abril": 4, "abr": 4,
                    "maio": 5, "mai": 5,
                    "junho": 6, "jun": 6,
                    "julho": 7, "jul": 7,
                    "agosto": 8, "ago": 8,
                    "setembro": 9, "set": 9,
                    "outubro": 10, "out": 10,
                    "novembro": 11, "nov": 11,
                    "dezembro": 12, "dez": 12
                }

                mes_encontrado = next(
                    (
                        mes
                        for nome_mes, mes in meses.items()
                        if nome_mes.startswith(texto)
                    ),
                    None
                )

                if mes_encontrado:
                    consulta += """
                        AND EXTRACT(MONTH FROM data::date) = %s
                    """
                    parametros.append(mes_encontrado)
                else:
                    # Entrada inválida: não retorna resultados por data
                    consulta += " AND 1 = 0"

        consulta += " ORDER BY id DESC"

        cursor.execute(consulta, tuple(parametros))
        return cursor.fetchall()

    finally:
        if cursor:
            cursor.close()

        if conexao:
            conexao.close()

# ============================================================
# ROTA PARA LISTAR OS REGISTROS COM FILTROS
# ============================================================

@app.route("/registros")
def registros():
    filtros = obter_filtros_registros()

    try:
        lista_registros = consultar_registros(filtros)

        return render_template(
            "registros.html",
            registros=lista_registros,
            nome=filtros["nome"],
            descricao=filtros["descricao"],
            data=filtros["data"]
        )

    except Exception as erro:
        print(f"ERRO AO CARREGAR REGISTROS: {erro}")

        flash(
            "Erro ao carregar os registros.",
            "error"
        )

        return redirect(url_for("index"))


# ============================================================
# BAIXAR EM EXCEL OS REGISTROS FILTRADOS
# ============================================================

@app.route("/baixar-registros")
def baixar_registros():
    workbook = None

    try:
        filtros = obter_filtros_registros()
        lista_registros = consultar_registros(filtros)

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Registros"

        sheet.append([
            "ID",
            "Ferramenta",
            "Descrição",
            "Evidência",
            "Ciclo",
            "Data"
        ])

        # Larguras das colunas
        larguras = {
            "A": 10,
            "B": 25,
            "C": 50,
            "D": 22,
            "E": 15,
            "F": 18
        }

        for coluna, largura in larguras.items():
            sheet.column_dimensions[coluna].width = largura

        from openpyxl.styles import Font, PatternFill, Alignment

        for celula in sheet[1]:
            celula.font = Font(bold=True)
            celula.fill = PatternFill(
                fill_type="solid",
                fgColor="D9EAF7"
            )
            celula.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # Preencher linhas e inserir imagens
        for registro in lista_registros:
            id_registro = registro[0]
            ferramenta = registro[1]
            descricao = registro[2]
            nome_foto = registro[3]
            ciclo = registro[4]
            data = registro[5]

            sheet.append([
                id_registro,
                ferramenta,
                descricao,
                None,
                ciclo,
                data
            ])

            numero_linha = sheet.max_row

            for coluna in range(1, 7):
                sheet.cell(
                    row=numero_linha,
                    column=coluna
                ).alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

            # Altura inicial para a imagem
            sheet.row_dimensions[numero_linha].height = 65

            if nome_foto:
                caminho_foto = os.path.join(
                    UPLOAD_FOLDER,
                    os.path.basename(nome_foto)
                )

                if os.path.isfile(caminho_foto):
                    try:
                        imagem = ExcelImage(caminho_foto)

                        # Redimensionar sem distorcer a imagem
                        largura_original = imagem.width
                        altura_original = imagem.height

                        limite_largura = 120
                        limite_altura = 80

                        escala = min(
                            limite_largura / largura_original,
                            limite_altura / altura_original
                        )

                        imagem.width = int(largura_original * escala)
                        imagem.height = int(altura_original * escala)

                        # A imagem fica ancorada na coluna D
                        sheet.add_image(
                            imagem,
                            f"D{numero_linha}"
                        )

                        sheet.row_dimensions[
                            numero_linha
                        ].height = max(
                            65,
                            imagem.height * 0.75
                        )

                    except Exception as erro_imagem:
                        print(
                            f"Erro ao inserir imagem do registro "
                            f"{id_registro}: {erro_imagem}"
                        )

                        sheet.cell(
                            row=numero_linha,
                            column=4
                        ).value = "Erro ao inserir imagem"

                else:
                    sheet.cell(
                        row=numero_linha,
                        column=4
                    ).value = "Arquivo não encontrado"

            else:
                sheet.cell(
                    row=numero_linha,
                    column=4
                ).value = "Sem imagem"

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        arquivo = BytesIO()
        workbook.save(arquivo)
        arquivo.seek(0)

        return send_file(
            arquivo,
            as_attachment=True,
            download_name="registros_defeitos.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as erro:
        print(f"ERRO AO GERAR EXCEL: {erro}")

        flash(
            "Erro ao gerar o arquivo Excel. Verifique o terminal.",
            "error"
        )

        return redirect(url_for("registros"))

    finally:
        if workbook:
            workbook.close()
# ============================================================
# ROTA PARA ADICIONAR DEFEITO
# ============================================================

@app.route(
    "/adicionar",
    methods=["POST"]
)
def adicionar():

    # --------------------------------------------------------
    # Receber dados do formulário
    # --------------------------------------------------------

    ferramenta = request.form.get(
        "ferramenta"
    )

    descricao = request.form.get(
        "descricao"
    )

    ciclo = request.form.get(
        "ciclo"
    )

    data = request.form.get(
        "data"
    )

    foto_camera = request.files.get("evidencia_camera")
    foto_arquivos = request.files.get("evidencia")

    # Usa a foto da câmera quando ela foi selecionada.
    # Caso contrário, usa a imagem dos arquivos.
    foto = (
        foto_camera
        if foto_camera and foto_camera.filename
        else foto_arquivos
    )

    em_manutencao = (
        request.form.get("em_manutencao", "false") == "true"
    )

    # --------------------------------------------------------
    # Verificar campos obrigatórios
    # --------------------------------------------------------

    if not ferramenta or not descricao:

        flash(
            "Preencha a ferramenta e a descrição.",
            "error"
        )

        return redirect(
            url_for("index")
        )


    # --------------------------------------------------------
    # SALVAR FOTO
    # --------------------------------------------------------

    nome_foto = None

    if foto and foto.filename:

        if not extensao_permitida(
            foto.filename
        ):

            flash(
                "Formato de imagem não permitido.",
                "error"
            )

            return redirect(
                url_for("index")
            )


        # Nome seguro do arquivo
        nome_original = secure_filename(
            foto.filename
        )


        # Obtém extensão
        extensao = nome_original.rsplit(
            ".",
            1
        )[1].lower()


        # Gera nome único
        nome_foto = (
            f"{uuid.uuid4().hex}.{extensao}"
        )


        # Caminho completo
        caminho_foto = os.path.join(
            app.config["UPLOAD_FOLDER"],
            nome_foto
        )


        # Salvar imagem
        foto.save(
            caminho_foto
        )


    # ========================================================
    # INSERIR NO POSTGRESQL
    # ========================================================

    conexao = None

    try:

        conexao = conectar_banco()

        cursor = conexao.cursor()


        cursor.execute(
            """
            INSERT INTO defeito (
                ferramenta,
                descricao,
                evidencia_foto,
                ciclo,
                data,
                em_manutencao
            )
            VALUES (%s, %s, %s, %s, %s, %s)
            """,
            (
                ferramenta,
                descricao,
                nome_foto,
                ciclo,
                data,
                em_manutencao
            )
        )


        conexao.commit()

        cursor.close()


        # ====================================================
        # INSERIR NO EXCEL
        # ====================================================

        try:

            adicionar_no_excel(
                ferramenta,
                descricao,
                nome_foto,
                ciclo,
                data
            )

        except PermissionError:

            flash(
                "O registro foi salvo no PostgreSQL, "
                "mas não foi possível atualizar "
                "defeito.xlsx. Feche a planilha no Excel "
                "e tente novamente.",
                "error"
            )

            return redirect(
                url_for("index")
            )


        except Exception as erro_excel:

            print(
                f"ERRO AO ATUALIZAR EXCEL: {erro_excel}"
            )

            flash(
                "O registro foi salvo no PostgreSQL, "
                "mas ocorreu um erro ao atualizar "
                "defeito.xlsx.",
                "error"
            )

            return redirect(
                url_for("index")
            )


        # ====================================================
        # SUCESSO
        # ====================================================

        flash(
            "Defeito cadastrado com sucesso!",
            "success"
        )


    except Exception as erro:

        print(
            f"ERRO NO POSTGRESQL: {erro}"
        )


        if conexao:

            conexao.rollback()


        flash(
            "Erro ao cadastrar o defeito.",
            "error"
        )


    finally:

        if conexao:

            conexao.close()


    return redirect(
        url_for("index")
    )


# ============================================================
# LISTAR SOMENTE AS FERRAMENTAS EM MANUTENÇÃO
# ============================================================

@app.route("/manutencao")
def manutencao():
    conexao = None

    try:
        conexao = conectar_banco()
        cursor = conexao.cursor()

        cursor.execute(
            """
            SELECT
                id,
                ferramenta,
                descricao,
                evidencia_foto,
                ciclo,
                data,
                em_manutencao
            FROM defeito
            WHERE em_manutencao = TRUE
            ORDER BY id DESC
            """
        )

        registros = cursor.fetchall()
        cursor.close()

        return render_template(
            "manutencao.html",
            registros=registros
        )

    except Exception as erro:
        print(f"ERRO AO CARREGAR MANUTENÇÕES: {erro}")
        flash("Não foi possível carregar os registros de manutenção.", "error")
        return redirect(url_for("index"))

    finally:
        if conexao:
            conexao.close()


# ============================================================
# ALTERAR O STATUS DE MANUTENÇÃO DE UM REGISTRO
# ============================================================

@app.route("/alterar-manutencao/<int:id>", methods=["POST"])
def alterar_manutencao(id):
    conexao = None

    try:
        novo_status = request.form.get("em_manutencao") == "true"

        conexao = conectar_banco()
        cursor = conexao.cursor()

        cursor.execute(
            """
            UPDATE defeito
            SET em_manutencao = %s
            WHERE id = %s
            """,
            (novo_status, id)
        )

        if cursor.rowcount == 0:
            conexao.rollback()
            flash("Registro não encontrado.", "error")
        else:
            conexao.commit()
            flash("Status de manutenção atualizado.", "success")

        cursor.close()

    except Exception as erro:
        if conexao:
            conexao.rollback()

        print(f"ERRO AO ALTERAR MANUTENÇÃO: {erro}")
        flash("Não foi possível atualizar o status.", "error")

    finally:
        if conexao:
            conexao.close()

    return redirect(request.referrer or url_for("manutencao"))

# ============================================================
# ROTA DE ULTIMO REGISTRO
# ============================================================

@app.route("/ultimo-registro/<path:nome_ferramenta>")
def ultimo_registro_ferramenta(nome_ferramenta):
    conexao = None
    cursor = None

    try:
        conexao = conectar_banco()
        cursor = conexao.cursor()

        cursor.execute(
            """
            SELECT
                id,
                ferramenta,
                descricao,
                evidencia_foto,
                ciclo,
                data,
                em_manutencao
            FROM defeito
            WHERE LOWER(TRIM(ferramenta)) = LOWER(TRIM(%s))
            ORDER BY id DESC
            LIMIT 1
            """,
            (nome_ferramenta,)
        )

        registro = cursor.fetchone()

        if not registro:
            return jsonify({
                "encontrado": False
            })

        return jsonify({
            "encontrado": True,
            "id": registro[0],
            "ferramenta": registro[1],
            "descricao": registro[2] or "",
            "foto": registro[3],
            "ciclo": registro[4],
            "data": str(registro[5]) if registro[5] else "",
            "em_manutencao": bool(registro[6])
        })

    except Exception as erro:
        print(f"Erro ao buscar último registro: {erro}")
        return jsonify({
            "erro": "Não foi possível consultar o último registro."
        }), 500

    finally:
        if cursor:
            cursor.close()
        if conexao:
            conexao.close()
            
# ============================================================
# EXECUTAR APLICAÇÃO
# ============================================================

if __name__ == "__main__":

    app.run(
        debug=True
    )