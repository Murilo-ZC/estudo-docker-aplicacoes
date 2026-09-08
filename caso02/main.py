from flask import Flask, jsonify, request

app = Flask(__name__)


class DadosInvalidos(Exception):
    """Erro de entrada: o cliente mandou algo que nao da para calcular."""


def extrair_operandos():
    """Le o corpo JSON da requisicao e devolve os operandos como float.

    Levanta DadosInvalidos quando o corpo esta ausente, nao e um objeto
    JSON, nao tem os campos obrigatorios ou tem valores nao numericos.
    """
    dados = request.get_json(silent=True)

    if not isinstance(dados, dict):
        raise DadosInvalidos("envie um corpo JSON com os campos 'a' e 'b'")

    ausentes = [campo for campo in ("a", "b") if campo not in dados]
    if ausentes:
        raise DadosInvalidos(f"campos ausentes: {', '.join(ausentes)}")

    try:
        return float(dados["a"]), float(dados["b"])
    except (TypeError, ValueError):
        raise DadosInvalidos("'a' e 'b' precisam ser numeros")


def primeira_linha(texto):
    """Primeira linha nao vazia de uma docstring (ou string vazia)."""
    for linha in (texto or "").strip().splitlines():
        if linha.strip():
            return linha.strip()
    return ""


def resposta(operacao, a, b, resultado):
    """Formato unico de resposta de sucesso, igual para as quatro rotas."""
    return jsonify(operacao=operacao, a=a, b=b, resultado=resultado)


@app.errorhandler(DadosInvalidos)
def tratar_dados_invalidos(erro):
    """Transforma DadosInvalidos em uma resposta 400 com mensagem legivel."""
    return jsonify(erro=str(erro)), 400


@app.post("/soma")
def soma():
    """Soma dois numeros: a + b."""
    a, b = extrair_operandos()
    return resposta("soma", a, b, a + b)


@app.post("/subtracao")
def subtracao():
    """Subtrai o segundo numero do primeiro: a - b."""
    a, b = extrair_operandos()
    return resposta("subtracao", a, b, a - b)


@app.post("/multiplicacao")
def multiplicacao():
    """Multiplica dois numeros: a * b."""
    a, b = extrair_operandos()
    return resposta("multiplicacao", a, b, a * b)


@app.post("/divisao")
def divisao():
    """Divide o primeiro numero pelo segundo: a / b (b diferente de zero)."""
    a, b = extrair_operandos()
    if b == 0:
        return jsonify(erro="divisao por zero nao e permitida"), 400
    return resposta("divisao", a, b, a / b)


@app.get("/rotas")
def rotas():
    """Lista as rotas disponiveis lendo o proprio mapa de URLs do Flask.

    Nao ha lista escrita a mao: quem adicionar uma rota nova aparece aqui
    automaticamente.
    """
    disponiveis = []
    for regra in app.url_map.iter_rules():
        if regra.endpoint == "static":
            continue
        metodos = sorted(regra.methods - {"HEAD", "OPTIONS"})
        disponiveis.append(
            {
                "rota": str(regra.rule),
                "metodos": metodos,
                "descricao": primeira_linha(
                    app.view_functions[regra.endpoint].__doc__
                ),
            }
        )

    disponiveis.sort(key=lambda item: item["rota"])
    return jsonify(total=len(disponiveis), rotas=disponiveis)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)