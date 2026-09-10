import streamlit as st
import io
import math
import shutil
import subprocess
import os
import tempfile
import hashlib
import requests
import base64
from datetime import datetime
import chess
import chess.pgn
import chess.engine
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit.components.v1 as components

st.set_page_config(page_title="E4Chess", page_icon="♟️", layout="wide")

STOCKFISH_PATH = shutil.which("stockfish") or "/usr/games/stockfish"
HEADERS_HTTP = {"User-Agent": "E4Chess/1.0"}


# ═══════════════════════════════════════════════════════════════════
#  APIs
# ═══════════════════════════════════════════════════════════════════
@st.cache_data(ttl=86400, show_spinner=False)
def consultar_mestres(fen):
    try:
        url = "https://explorer.lichess.ovh/masters?fen=" + requests.utils.quote(fen) + "&moves=1"
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return {}


@st.cache_data(ttl=1800, show_spinner=False)
def buscar_noticias(rss_url, limite=5):
    try:
        import xml.etree.ElementTree as ET
        r = requests.get(rss_url, headers=HEADERS_HTTP, timeout=8)
        if r.status_code != 200:
            return []
        root = ET.fromstring(r.content)
        itens = []
        for item in root.iter("item"):
            titulo = item.findtext("title", "")
            link = item.findtext("link", "")
            data = item.findtext("pubDate", "")
            if titulo and link:
                itens.append({"titulo": titulo, "link": link, "data": data})
            if len(itens) >= limite:
                break
        return itens
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def buscar_ticker_noticias():
    fontes = [
        "http://www.theweekinchess.com/twic-rss-feed",
        "http://feeds.feedburner.com/chessbase",
        "http://www.chessdom.com/rss",
    ]
    manchetes = []
    for url in fontes:
        try:
            manchetes.extend(buscar_noticias(url, 3))
            if len(manchetes) >= 8:
                break
        except Exception:
            continue
    return manchetes[:10]


@st.cache_data(ttl=7200, show_spinner=False)
def buscar_ranking_fide():
    try:
        r = requests.get(
            "https://fide-players.fly.dev/players?limit=10&sort=rating&order=desc",
            timeout=8,
        )
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return [
        {"rank": 1, "name": "Magnus Carlsen", "country": "NOR", "rating": 2830},
        {"rank": 2, "name": "Hikaru Nakamura", "country": "USA", "rating": 2802},
        {"rank": 3, "name": "Fabiano Caruana", "country": "USA", "rating": 2796},
        {"rank": 4, "name": "Arjun Erigaisi", "country": "IND", "rating": 2791},
        {"rank": 5, "name": "Gukesh D", "country": "IND", "rating": 2783},
        {"rank": 6, "name": "Nodirbek Abdusattorov", "country": "UZB", "rating": 2777},
        {"rank": 7, "name": "Ian Nepomniachtchi", "country": "FID", "rating": 2770},
        {"rank": 8, "name": "Alireza Firouzja", "country": "FRA", "rating": 2765},
        {"rank": 9, "name": "Wesley So", "country": "USA", "rating": 2758},
        {"rank": 10, "name": "Leinier Dominguez", "country": "USA", "rating": 2752},
    ]


# ═══════════════════════════════════════════════════════════════════
#  LIVRO DE ABERTURAS
# ═══════════════════════════════════════════════════════════════════
LIVRO = {
    ("e4", "e5", "Nf3", "Nc6", "Bc4"): "Abertura Italiana",
    ("e4", "e5", "Nf3", "Nc6", "Bb5"): "Abertura Espanhola",
    ("e4", "c5"): "Defesa Siciliana",
    ("e4", "e6"): "Defesa Francesa",
    ("e4", "c6"): "Defesa Caro-Kann",
    ("e4", "d5"): "Defesa Escandinava",
    ("e4", "Nf6"): "Defesa Alekhine",
    ("e4", "d6"): "Defesa Pirc",
    ("e4", "g6"): "Defesa Moderna",
    ("e4", "e5", "Nf3", "Nf6"): "Defesa Petroff",
    ("e4", "e5", "Nc3"): "Abertura Vienense",
    ("e4", "e5", "Bc4"): "Abertura do Bispo",
    ("e4", "e5", "f4"): "Gambito do Rei",
    ("e4", "e5"): "Abertura Aberta",
    ("d4", "d5", "c4"): "Gambito da Dama",
    ("d4", "Nf6", "c4", "g6"): "India do Rei",
    ("d4", "Nf6", "c4", "e6"): "India da Dama",
    ("d4", "d5", "Nf3", "Nf6", "Bf4"): "Sistema London",
    ("d4", "f5"): "Defesa Holandesa",
    ("d4", "e6"): "Peao da Dama",
    ("d4", "d5"): "Peao da Dama",
    ("c4",): "Abertura Inglesa",
    ("Nf3",): "Abertura Reti",
    ("d4",): "Peao da Dama",
    ("e4",): "Peao do Rei",
}


def identificar_abertura(sans):
    melhor = None
    tam = 0
    for seq, nome in LIVRO.items():
        if len(seq) <= len(sans):
            if tuple(sans[:len(seq)]) == seq and len(seq) > tam:
                melhor = nome
                tam = len(seq)
    return melhor


# ═══════════════════════════════════════════════════════════════════
#  FÓRMULAS
# ═══════════════════════════════════════════════════════════════════
def cp_para_winpercent(cp):
    cp = max(-1500, min(1500, cp))
    return 50 + 50 * (2 / (1 + math.exp(-0.00368208 * cp)) - 1)


def precisao_lichess(wp_a, wp_d):
    queda = max(0, wp_a - wp_d)
    if queda <= 0:
        return 100.0
    if queda >= 50:
        return 0.0
    valor = 103.1668 * math.exp(-0.04354 * queda) - 3.1669
    if valor < 0:
        return 0.0
    if valor > 100:
        return 100.0
    return valor


def classificar(loss, is_best, san, mate_a, mate_d, em_livro, em_mestre):
    if em_mestre and loss <= 30:
        return "Lance de Mestre"
    if em_livro:
        return "Livro"
    if "#" in san:
        return "Melhor"
    if is_best:
        return "Melhor"
    if mate_a is not None and mate_a > 0 and mate_d is None:
        return "Gafe"
    if mate_a is not None and mate_d is not None:
        if abs(mate_d) > abs(mate_a) + 2:
            return "Gafe"
    if loss <= 10:
        return "Excelente"
    if loss <= 40:
        return "Bom"
    if loss <= 90:
        return "Imprecisao"
    if loss <= 200:
        return "Erro"
    if loss <= 500:
        return "Erro Grave"
    return "Gafe"


def detectar_plataforma(h):
    s = (h.get("Site", "") or "").lower()
    e = (h.get("Event", "") or "").lower()
    if "chess.com" in s or "chess.com" in e:
        return "Chess.com"
    if "lichess" in s or "lichess" in e:
        return "Lichess"
    return "Desconhecida"


def validar_pgn(texto):
    if not texto or not texto.strip():
        raise ValueError("Cole um PGN.")
    game = chess.pgn.read_game(io.StringIO(texto))
    if game is None:
        raise ValueError("PGN invalido.")
    lances = list(game.mainline_moves())
    if not lances:
        raise ValueError("PGN sem lances.")
    board = game.board()
    for i, m in enumerate(lances, 1):
        if not board.is_legal(m):
            raise ValueError("Lance ilegal numero " + str(i))
        board.push(m)
    return game


# ═══════════════════════════════════════════════════════════════════
#  E4 RATING
# ═══════════════════════════════════════════════════════════════════
def calcular_e4_rating(prec, mestres, gafes, erros_graves, rating_oponente, rating_jogador):
    if prec >= 90:
        bonus = 300 + (prec - 90) * 20
    elif prec >= 80:
        bonus = 150 + (prec - 80) * 15
    elif prec >= 70:
        bonus = 50 + (prec - 70) * 10
    else:
        bonus = max(0, prec * 0.5)
    bonus += mestres * 50
    bonus += 150 if (gafes == 0 and erros_graves == 0) else 0
    bonus -= gafes * 100
    bonus -= erros_graves * 40

    try:
        if rating_oponente in ["—", "*", "", None]:
            ro = 1200
        else:
            ro = int(rating_oponente)
    except Exception:
        ro = 1200

    try:
        if rating_jogador in ["—", "*", "", None]:
            rj = 1200
        else:
            rj = int(rating_jogador)
    except Exception:
        rj = 1200

    dif = ro - rj
    if dif >= 400:
        fator = 1.5
    elif dif >= 100:
        fator = 1.2
    else:
        fator = 1.0

    e4 = (ro + bonus) * fator
    if e4 < 400:
        return 400
    if e4 > 2800:
        return 2800
    return int(e4)


# ═══════════════════════════════════════════════════════════════════
#  MOTOR STOCKFISH
# ═══════════════════════════════════════════════════════════════════
def analisar_stockfish(game, prof, progress_cb=None):
    try:
        eng = chess.engine.SimpleEngine.popen_uci(STOCKFISH_PATH)
    except Exception as e:
        st.warning("Stockfish: " + str(e))
        return None

    try:
        eng.configure({"Threads": 1, "Hash": 128})
    except Exception:
        pass

    board = game.board()
    lances = list(game.mainline_moves())
    n = len(lances)
    aval = []
    melhores = []

    for i in range(n):
        info = eng.analyse(board, chess.engine.Limit(depth=prof))
        sc = info["score"].pov(chess.WHITE)
        cp = sc.score(mate_score=100000)
        if cp is None:
            cp = 0
        mate = None
        if sc.is_mate():
            mate = sc.mate()
        pv = info.get("pv") or []
        melhor = pv[0] if pv else None
        aval.append((cp, mate))
        melhores.append(melhor)
        if progress_cb:
            try:
                progress_cb(i + 1, n + 1)
            except Exception:
                pass
        board.push(lances[i])

    info = eng.analyse(board, chess.engine.Limit(depth=prof))
    sc = info["score"].pov(chess.WHITE)
    cp_final = sc.score(mate_score=100000)
    if cp_final is None:
        cp_final = 0
    mate_final = None
    if sc.is_mate():
        mate_final = sc.mate()
    aval.append((cp_final, mate_final))

    if progress_cb:
        try:
            progress_cb(n + 1, n + 1)
        except Exception:
            pass

    eng.quit()

    resultados = []
    for i, move in enumerate(lances):
        cor = "brancas" if i % 2 == 0 else "negras"
        cp_i, mate_i = aval[i]
        cp_n, mate_n = aval[i + 1]
        best = melhores[i]

        if cor == "brancas":
            loss = max(0, cp_i - cp_n)
            cp_played = cp_n
        else:
            loss = max(0, cp_n - cp_i)
            cp_played = -cp_n

        is_best = (best is not None and move == best)
        resultados.append({
            "cp_best": cp_i if cor == "brancas" else -cp_i,
            "cp_played": cp_played,
            "loss": loss,
            "is_best": is_best,
            "mate_antes": mate_i,
            "mate_depois": mate_n,
            "cp_white_before": cp_i,
            "cp_white_after": cp_n,
        })

    return resultados


def analisar(game, prof=15, progress_cb=None):
    headers = dict(game.headers)
    lances = list(game.mainline_moves())
    cab = {
        "brancas": headers.get("White", "Brancas"),
        "negras": headers.get("Black", "Pretas"),
        "rating_brancas": headers.get("WhiteElo", "—"),
        "rating_negras": headers.get("BlackElo", "—"),
        "resultado": headers.get("Result", "*"),
        "evento": headers.get("Event", "—"),
        "data": headers.get("Date", "—"),
        "plataforma": detectar_plataforma(headers),
    }

    eng_res = analisar_stockfish(game, prof=prof, progress_cb=progress_cb)

    dados = []
    stats = {
        "brancas": {"classes": {}, "precs": [], "mestres": 0},
        "negras": {"classes": {}, "precs": [], "mestres": 0},
    }
    curva = [0]
    marcadores = []
    sans = []
    abertura = None
    board = game.board()

    for i, move in enumerate(lances):
        cor = "brancas" if i % 2 == 0 else "negras"
        san = board.san(move)
        fen = board.fen()
        board.push(move)
        sans.append(san)

        nome = identificar_abertura(sans)
        em_livro = tuple(sans) in LIVRO
        if em_livro and abertura is None:
            abertura = nome

        dm = consultar_mestres(fen)
        em_mestre = False
        if dm and dm.get("moves") and i >= 16:
            total_jogos = 0
            for m in dm["moves"]:
                total_jogos += (m.get("white", 0) or 0) + (m.get("black", 0) or 0)
            if total_jogos >= 20:
                for mv in dm["moves"]:
                    if mv.get("san") == san:
                        freq = ((mv.get("white", 0) or 0) + (mv.get("black", 0) or 0)) / total_jogos
                        if 0.03 <= freq <= 0.25:
                            em_mestre = True
                        break

        if eng_res and i < len(eng_res):
            info = eng_res[i]
            wp_a = cp_para_winpercent(info["cp_best"])
            wp_d = cp_para_winpercent(info["cp_played"])
            prec = precisao_lichess(wp_a, wp_d)
            if em_mestre:
                prec = min(99.9, prec + 8)
            if em_livro:
                prec = min(99.9, prec + 3)
            if "#" in san:
                prec = 100.0
            cls = classificar(
                info["loss"], info["is_best"], san,
                info["mate_antes"], info["mate_depois"],
                em_livro, em_mestre
            )
            loss = info["loss"]
            ev = info["cp_played"]
            cp_w = info["cp_white_after"]
        else:
            prec = 100.0 if (em_livro or em_mestre) else 50.0
            if em_mestre:
                cls = "Lance de Mestre"
            elif em_livro:
                cls = "Livro"
            else:
                cls = "Bom"
            loss = 0
            ev = 0
            cp_w = 0

        cp_vis = max(-1000, min(1000, cp_w))
        curva.append(cp_vis / 100)

        if cls in ("Erro", "Erro Grave", "Gafe", "Imprecisao"):
            marcadores.append({"ply": i + 1, "cp": cp_vis / 100, "classe": cls})

        dados.append({
            "n": i + 1, "cor": cor, "san": san, "classe": cls,
            "precisao": prec, "loss": loss, "eval": ev, "mestre": em_mestre,
        })

        stats[cor]["classes"][cls] = stats[cor]["classes"].get(cls, 0) + 1
        stats[cor]["precs"].append(prec)
        if em_mestre:
            stats[cor]["mestres"] += 1

    pb = sum(stats["brancas"]["precs"]) / max(1, len(stats["brancas"]["precs"]))
    pn = sum(stats["negras"]["precs"]) / max(1, len(stats["negras"]["precs"]))

    def nota(p):
        if p >= 95:
            return "A+"
        if p >= 90:
            return "A"
        if p >= 85:
            return "B+"
        if p >= 80:
            return "B"
        if p >= 75:
            return "C+"
        if p >= 70:
            return "C"
        if p >= 60:
            return "D"
        return "F"

    r_b = calcular_e4_rating(
        pb, stats["brancas"]["mestres"],
        stats["brancas"]["classes"].get("Gafe", 0),
        stats["brancas"]["classes"].get("Erro Grave", 0),
        cab["rating_negras"], cab["rating_brancas"]
    )
    r_n = calcular_e4_rating(
        pn, stats["negras"]["mestres"],
        stats["negras"]["classes"].get("Gafe", 0),
        stats["negras"]["classes"].get("Erro Grave", 0),
        cab["rating_brancas"], cab["rating_negras"]
    )

    return {
        "cabecalho": cab,
        "lances": dados,
        "estatisticas": stats,
        "precisao_brancas": pb,
        "precisao_negras": pn,
        "nota_brancas": nota(pb),
        "nota_negras": nota(pn),
        "rating_e4_brancas": r_b,
        "rating_e4_negras": r_n,
        "curva": curva,
        "marcadores": marcadores,
        "abertura": abertura or "Nao identificada",
    }


# ═══════════════════════════════════════════════════════════════════
#  CONQUISTAS E SUPERAÇÃO
# ═══════════════════════════════════════════════════════════════════
def calcular_superacao(d, cor):
    cab = d["cabecalho"]
    chave = "rating_" + cor
    try:
        if cab[chave] in ["—", "*", "", None]:
            rating_real = 0
        else:
            rating_real = int(cab[chave])
    except Exception:
        rating_real = 0

    performance = d["rating_e4_" + cor]
    if rating_real == 0:
        return None

    diferenca = performance - rating_real

    if diferenca >= 600:
        nivel = "Mestre"
        emoji = "👑"
    elif diferenca >= 400:
        nivel = "Expert"
        emoji = "🌟"
    elif diferenca >= 250:
        nivel = "Avançado"
        emoji = "⭐"
    elif diferenca >= 100:
        nivel = "Intermediário"
        emoji = "✨"
    elif diferenca >= 0:
        nivel = "Acima do seu nível"
        emoji = "🎯"
    else:
        return None

    return {
        "diferenca": diferenca,
        "nivel": nivel,
        "emoji": emoji,
        "rating_real": rating_real,
        "performance": performance,
    }


def calcular_conquistas(d, cor):
    stats = d["estatisticas"][cor]
    classes = stats["classes"]
    prec = d["precisao_" + cor]
    c = []

    if stats["mestres"] >= 1:
        c.append(("🏅", "Lance de Mestre", str(stats["mestres"]) + " lance(s) reproduzido(s)"))
    if stats["mestres"] >= 5:
        c.append(("👑", "Reprodutor de Mestres", str(stats["mestres"]) + " lances de alto nivel"))
    if prec >= 90:
        c.append(("🎯", "Precisão de Campeão", str(round(prec, 1)) + "% de precisao"))
    if classes.get("Gafe", 0) == 0 and classes.get("Erro Grave", 0) == 0:
        c.append(("🛡️", "Defensor", "Sem erros graves nem gafes"))
    if classes.get("Livro", 0) >= 6:
        c.append(("📖", "Mestre das Aberturas", str(classes.get("Livro", 0)) + " lances de livro"))
    if classes.get("Melhor", 0) >= 5:
        c.append(("⭐", "Caçador de Melhores", str(classes["Melhor"]) + " melhores lances"))
    if classes.get("Gafe", 0) >= 3:
        c.append(("⚠️", "Precisa Estudar", str(classes["Gafe"]) + " gafes cometidas"))

    sup = calcular_superacao(d, cor)
    if sup and sup["diferenca"] >= 250:
        c.append((
            sup["emoji"],
            "Performance " + sup["nivel"],
            "Jogou como " + str(sup["performance"]) + " (+" + str(sup["diferenca"]) + " acima)"
        ))

    return c


# ═══════════════════════════════════════════════════════════════════
#  CACHE
# ═══════════════════════════════════════════════════════════════════
def obter_analise(pgn, prof):
    chave = "v3_" + hashlib.md5(pgn.encode()).hexdigest() + "_" + str(prof)

    if "cache" not in st.session_state:
        st.session_state.cache = {}

    if chave in st.session_state.cache:
        st.info("⚡ Análise recuperada do cache")
        return st.session_state.cache[chave]

    game = validar_pgn(pgn)
    barra = st.progress(0, text="Analisando com Stockfish...")

    def cb(i, total):
        pct = min(1.0, i / total)
        barra.progress(pct, text="Posição " + str(i) + "/" + str(total))

    d = analisar(game, prof=prof, progress_cb=cb)
    barra.empty()

    st.session_state.cache[chave] = d
    return d


# ═══════════════════════════════════════════════════════════════════
#  GRÁFICO
# ═══════════════════════════════════════════════════════════════════
def gerar_grafico(d):
    c = d["curva"]
    xs = list(range(len(c)))

    fig, ax = plt.subplots(figsize=(9, 4), dpi=100, facecolor="#0d1117")
    ax.set_facecolor("#0d1117")

    pos = [max(0, y) for y in c]
    neg = [min(0, y) for y in c]
    ax.fill_between(xs, pos, 0, color="#f0f6fc", alpha=0.92)
    ax.fill_between(xs, neg, 0, color="#0d1117", alpha=0.95)
    ax.plot(xs, c, color="#58a6ff", linewidth=1.4)
    ax.axhline(0, color="#30363d", linewidth=1)

    cores = {
        "Imprecisao": "#d29922",
        "Erro": "#f0883e",
        "Erro Grave": "#f85149",
        "Gafe": "#ff2222",
    }

    for cl, cor in cores.items():
        pts = [m for m in d["marcadores"] if m["classe"] == cl]
        if pts:
            xs2 = [p["ply"] for p in pts]
            ys2 = [p["cp"] for p in pts]
            ax.scatter(xs2, ys2, color=cor, s=65, edgecolors="#0d1117", linewidths=1.2, label=cl)

    ax.set_xlim(0, max(1, len(c) - 1))
    ax.set_ylim(-8, 8)
    ax.tick_params(colors="#8b949e", labelsize=8)
    ax.grid(True, color="#21262d", linestyle=":", alpha=0.5)

    for sp in ax.spines.values():
        sp.set_color("#30363d")

    leg = ax.legend(loc="upper right", fontsize=7, facecolor="#161b22", edgecolor="#30363d")
    if leg:
        for t in leg.get_texts():
            t.set_color("#c9d1d9")

    fig.tight_layout()
    return fig


# ═══════════════════════════════════════════════════════════════════
#  IMAGEM PNG
# ═══════════════════════════════════════════════════════════════════
def gerar_imagem(d):
    c = d["cabecalho"]
    pb = d["precisao_brancas"]
    pn = d["precisao_negras"]
    nb = d["nota_brancas"]
    nn = d["nota_negras"]
    rb = d["rating_e4_brancas"]
    rn = d["rating_e4_negras"]

    def cor_nota(n):
        tabela = {
            "A+": "#3fb950", "A": "#3fb950",
            "B+": "#58a6ff", "B": "#58a6ff",
            "C+": "#d29922", "C": "#d29922",
            "D": "#f0883e", "F": "#f85149",
        }
        return tabela.get(n, "#c9d1d9")

    fig = plt.figure(figsize=(10, 12), dpi=100, facecolor="#0d1117")
    fig.subplots_adjust(left=0.06, right=0.94, top=0.96, bottom=0.04)

    fig.text(0.5, 0.970, "E4Chess", ha="center", fontsize=28, color="#58a6ff", weight="bold")
    fig.text(0.5, 0.947, "Analise Profissional de Xadrez", ha="center",
             fontsize=11, color="#8b949e", style="italic")

    y_cards = 0.82

    cards = [
        (0.08, c["brancas"], c["rating_brancas"], pb, nb, cor_nota(nb), rb),
        (0.54, c["negras"], c["rating_negras"], pn, nn, cor_nota(nn), rn),
    ]

    for x0, nome, rating, prec, nota, cor_n, r4 in cards:
        box = FancyBboxPatch(
            (x0, y_cards), 0.38, 0.11,
            boxstyle="round,pad=0.01",
            facecolor="#161b22", edgecolor="#30363d", linewidth=1.5
        )
        fig.add_artist(box)
        fig.text(x0 + 0.19, y_cards + 0.093, nome, ha="center", fontsize=12,
                 color="#c9d1d9", weight="bold")
        fig.text(x0 + 0.19, y_cards + 0.070, "Rating " + str(rating),
                 ha="center", fontsize=9, color="#8b949e")
        fig.text(x0 + 0.19, y_cards + 0.030, str(round(prec, 1)) + "%",
                 ha="center", fontsize=24, color=cor_n, weight="bold")
        fig.text(x0 + 0.19, y_cards - 0.002, "Nota " + nota,
                 ha="center", fontsize=10, color=cor_n, weight="bold")
        fig.text(x0 + 0.19, y_cards - 0.025, "E4 Rating: " + str(r4),
                 ha="center", fontsize=9, color="#58a6ff")

    fig.text(0.5, 0.775, "Resultado: " + str(c["resultado"]),
             ha="center", fontsize=15, color="#58a6ff", weight="bold")
    fig.text(0.5, 0.754, "Abertura: " + str(d["abertura"]),
             ha="center", fontsize=10, color="#c9d1d9")
    fig.text(0.5, 0.736, str(c["plataforma"]) + " · " + str(c["data"]),
             ha="center", fontsize=9, color="#8b949e")

    ax_g = fig.add_axes([0.08, 0.44, 0.84, 0.24])
    ax_g.set_facecolor("#0d1117")
    xs = list(range(len(d["curva"])))
    cc = d["curva"]
    pos = [max(0, y) for y in cc]
    neg = [min(0, y) for y in cc]
    ax_g.fill_between(xs, pos, 0, color="#f0f6fc", alpha=0.92)
    ax_g.fill_between(xs, neg, 0, color="#0d1117", alpha=0.95)
    ax_g.plot(xs, cc, color="#58a6ff", linewidth=1.2)
    ax_g.axhline(0, color="#30363d", linewidth=0.8)

    cores = {
        "Imprecisao": "#d29922",
        "Erro": "#f0883e",
        "Erro Grave": "#f85149",
        "Gafe": "#ff2222",
    }

    for cl, cor_er in cores.items():
        pts = [m for m in d["marcadores"] if m["classe"] == cl]
        if pts:
            xs2 = [p["ply"] for p in pts]
            ys2 = [p["cp"] for p in pts]
            ax_g.scatter(xs2, ys2, color=cor_er, s=40, edgecolors="#0d1117", linewidths=0.8)

    ax_g.set_xlim(0, max(1, len(cc) - 1))
    ax_g.set_ylim(-8, 8)
    ax_g.tick_params(colors="#8b949e", labelsize=7)
    ax_g.grid(True, color="#21262d", linestyle=":", alpha=0.5)

    for sp in ax_g.spines.values():
        sp.set_color("#30363d")

    ax_g.set_title("Evolucao da Partida", color="#c9d1d9", fontsize=10)

    ax_s = fig.add_axes([0.08, 0.06, 0.84, 0.34])
    ax_s.axis("off")
    ax_s.set_facecolor("#0d1117")

    cats = ["Lance de Mestre", "Livro", "Melhor", "Excelente", "Bom",
            "Imprecisao", "Erro", "Erro Grave", "Gafe"]

    ax_s.text(0.15, 0.96, c["brancas"][:12], transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#58a6ff", weight="bold")
    ax_s.text(0.5, 0.96, "Categoria", transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#c9d1d9", weight="bold")
    ax_s.text(0.85, 0.96, c["negras"][:12], transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#f0883e", weight="bold")

    for i, cat in enumerate(cats):
        y = 0.84 - i * 0.087
        cb_ = d["estatisticas"]["brancas"]["classes"].get(cat, 0)
        cn_ = d["estatisticas"]["negras"]["classes"].get(cat, 0)
        ax_s.text(0.15, y, str(cb_), transform=ax_s.transAxes,
                  ha="center", fontsize=9, color="#c9d1d9")
        ax_s.text(0.5, y, cat, transform=ax_s.transAxes,
                  ha="center", fontsize=9, color="#c9d1d9")
        ax_s.text(0.85, y, str(cn_), transform=ax_s.transAxes,
                  ha="center", fontsize=9, color="#c9d1d9")

    fig.text(0.5, 0.015, "E4Chess - e4chess.streamlit.app",
             ha="center", fontsize=8, color="#8b949e", style="italic")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    fig.savefig(tmp.name, facecolor="#0d1117")
    plt.close(fig)
    return tmp.name


# ═══════════════════════════════════════════════════════════════════
#  LEGENDA INSTAGRAM
# ═══════════════════════════════════════════════════════════════════
def legenda_instagram(d):
    c = d["cabecalho"]
    linhas = []
    linhas.append("♟️ Analisei minha partida no E4Chess!")
    linhas.append("")
    linhas.append("⚔️ " + str(c["brancas"]) + " vs " + str(c["negras"]))
    linhas.append("🏆 Resultado: " + str(c["resultado"]))
    linhas.append("🎯 Precisao: " + str(round(d["precisao_brancas"], 1)) + "% (Nota " + d["nota_brancas"] + ")")
    linhas.append("📈 E4 Rating: " + str(d["rating_e4_brancas"]))
    linhas.append("📖 Abertura: " + str(d["abertura"]))
    linhas.append("🏅 Lances de Mestre: " + str(d["estatisticas"]["brancas"]["mestres"]))
    linhas.append("")
    linhas.append("Analise a sua tambem em e4chess.streamlit.app")
    linhas.append("#xadrez #chess #e4chess #analise #xadrezbrasil")
    return "\n".join(linhas)


# ═══════════════════════════════════════════════════════════════════
#  WORD
# ═══════════════════════════════════════════════════════════════════
def gerar_word(d):
    doc = Document()
    t = doc.add_heading("E4Chess - Relatorio de Analise", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    c = d["cabecalho"]
    p = doc.add_paragraph("Gerado em " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True

    doc.add_heading("1. Metadados", 1)
    linhas_meta = [
        ("Brancas", c["brancas"] + " (" + str(c["rating_brancas"]) + ") - " + str(round(d["precisao_brancas"], 1)) + "% (Nota " + d["nota_brancas"] + ") - E4 Rating " + str(d["rating_e4_brancas"])),
        ("Negras", c["negras"] + " (" + str(c["rating_negras"]) + ") - " + str(round(d["precisao_negras"], 1)) + "% (Nota " + d["nota_negras"] + ") - E4 Rating " + str(d["rating_e4_negras"])),
        ("Resultado", str(c["resultado"])),
        ("Abertura", str(d["abertura"])),
        ("Plataforma", str(c["plataforma"])),
        ("Data", str(c["data"])),
    ]
    for label, valor in linhas_meta:
        doc.add_paragraph("• " + label + ": " + valor)

    doc.add_heading("2. Conquistas", 1)
    for cor in ["brancas", "negras"]:
        par = doc.add_paragraph("- " + cor.capitalize() + " -")
        par.runs[0].bold = True
        for icon, nome, desc in calcular_conquistas(d, cor):
            doc.add_paragraph(icon + " " + nome + ": " + desc)

    doc.add_heading("3. Estatisticas por Categoria", 1)
    tab2 = doc.add_table(rows=1, cols=3)
    tab2.style = "Light Grid Accent 1"
    h = tab2.rows[0].cells
    h[0].text = "Categoria"
    h[1].text = c["brancas"]
    h[2].text = c["negras"]
    cats = ["Lance de Mestre", "Livro", "Melhor", "Excelente", "Bom",
            "Imprecisao", "Erro", "Erro Grave", "Gafe"]
    for cat in cats:
        r = tab2.add_row().cells
        r[0].text = cat
        r[1].text = str(d["estatisticas"]["brancas"]["classes"].get(cat, 0))
        r[2].text = str(d["estatisticas"]["negras"]["classes"].get(cat, 0))

    doc.add_heading("4. Historico Completo", 1)
    tab = doc.add_table(rows=1, cols=5)
    tab.style = "Light Grid Accent 1"
    h = tab.rows[0].cells
    h[0].text = "#"
    h[1].text = "Lance"
    h[2].text = "Classe"
    h[3].text = "Perda"
    h[4].text = "Precisao"
    for l in d["lances"]:
        r = tab.add_row().cells
        r[0].text = str(l["n"])
        r[1].text = l["san"]
        r[2].text = l["classe"]
        r[3].text = str(l["loss"]) + "cp"
        r[4].text = str(round(l["precisao"])) + "%"

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


# ═══════════════════════════════════════════════════════════════════
#  INTERFACE PRINCIPAL
# ═══════════════════════════════════════════════════════════════════
st.markdown("# ♟️ E4Chess")
st.markdown("**Analise profissional de partidas de xadrez**")

try:
    ticker = buscar_ticker_noticias()
    if ticker:
        itens_lista = []
        for t in ticker[:6]:
            itens_lista.append("📰 " + t["titulo"])
        itens = " · ".join(itens_lista)

        st.markdown(
            "<style>"
            ".ticker-wrap { background: linear-gradient(90deg,#0d1117,#161b22,#0d1117);"
            " border-top:1px solid #30363d; border-bottom:1px solid #30363d;"
            " overflow:hidden; padding:8px 0; margin-bottom:12px; border-radius:8px; }"
            ".ticker { display:inline-block; white-space:nowrap;"
            " animation:rolar 60s linear infinite; color:#58a6ff; font-size:13px;"
            " font-weight:500; padding-left:100%; }"
            "@keyframes rolar { 0% { transform:translateX(0); } 100% { transform:translateX(-200%); } }"
            "</style>"
            "<div class='ticker-wrap'><div class='ticker'>" + itens + "</div></div>",
            unsafe_allow_html=True
        )
except Exception:
    pass


abas = st.tabs([
    "🎯 Analisar",
    "🏆 Ranking",
    "📰 Notícias",
    "📊 Como funciona o E4 Rating",
    "🧠 Pensadores",
    "🕯️ Nossa História",
])

aba_analise = abas[0]
aba_ranking = abas[1]
aba_noticias = abas[2]
aba_rating = abas[3]
aba_pensadores = abas[4]
aba_historia = abas[5]


# ═══════════════════════════════════════════════════════════════════
#  ABA ANALISAR
# ═══════════════════════════════════════════════════════════════════
with aba_analise:
    pgn_text = st.text_area("Cole o PGN da partida:", height=200,
                            placeholder="Cole aqui o PGN completo...")
    prof = st.select_slider(
        "Profundidade",
        options=[10, 12, 15, 18],
        value=15,
        help="10=rapido | 12=padrao | 15=profundo | 18=maximo"
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        btn_analisar = st.button("ANALISAR", use_container_width=True, type="primary")
    with col2:
        btn_limpar = st.button("LIMPAR", use_container_width=True)
    with col3:
        btn_recarregar = st.button("🔄 RECARREGAR", use_container_width=True)

    if btn_recarregar:
        st.session_state.cache = {}
        st.rerun()

    if btn_analisar and pgn_text.strip():
        try:
            d = obter_analise(pgn_text, prof)
            if d is None:
                st.error("Falha na análise.")
            else:
                c = d["cabecalho"]
                st.success("✅ Analise concluida! Abertura: " + str(d["abertura"]))

                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        "♙ " + c["brancas"],
                        str(round(d["precisao_brancas"], 1)) + "%",
                        "Nota " + d["nota_brancas"] + " · E4 " + str(d["rating_e4_brancas"])
                    )
                with col2:
                    st.metric(
                        "♟ " + c["negras"],
                        str(round(d["precisao_negras"], 1)) + "%",
                        "Nota " + d["nota_negras"] + " · E4 " + str(d["rating_e4_negras"])
                    )

                sup_b = calcular_superacao(d, "brancas")
                sup_n = calcular_superacao(d, "negras")

                if sup_b or sup_n:
                    st.markdown("### ⭐ Destaques de Superação")
                    col1, col2 = st.columns(2)

                    with col1:
                        if sup_b:
                            st.markdown(
                                "<div style='background:linear-gradient(135deg,#1f6feb,#8957e5);"
                                "padding:20px;border-radius:12px;color:#fff;text-align:center;'>"
                                "<div style='font-size:14px;opacity:.85;'>" + c["brancas"] + "</div>"
                                "<div style='font-size:36px;'>" + sup_b["emoji"] + "</div>"
                                "<div style='font-size:22px;font-weight:700;margin:8px 0;'>"
                                "Você jogou como " + str(sup_b["performance"]) + "!</div>"
                                "<div style='font-size:13px;opacity:.9;'>"
                                "Rating real: " + str(sup_b["rating_real"]) + " → "
                                "<b>+" + str(sup_b["diferenca"]) + " acima</b></div>"
                                "</div>",
                                unsafe_allow_html=True
                            )

                    with col2:
                        if sup_n:
                            st.markdown(
                                "<div style='background:linear-gradient(135deg,#1f6feb,#8957e5);"
                                "padding:20px;border-radius:12px;color:#fff;text-align:center;'>"
                                "<div style='font-size:14px;opacity:.85;'>" + c["negras"] + "</div>"
                                "<div style='font-size:36px;'>" + sup_n["emoji"] + "</div>"
                                "<div style='font-size:22px;font-weight:700;margin:8px 0;'>"
                                "Você jogou como " + str(sup_n["performance"]) + "!</div>"
                                "<div style='font-size:13px;opacity:.9;'>"
                                "Rating real: " + str(sup_n["rating_real"]) + " → "
                                "<b>+" + str(sup_n["diferenca"]) + " acima</b></div>"
                                "</div>",
                                unsafe_allow_html=True
                            )

                st.markdown("**Resultado:** " + str(c["resultado"]) + " | **Abertura:** " + str(d["abertura"]))

                st.markdown("### 🏆 Conquistas")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**" + c["brancas"] + "**")
                    for ic, nm, ds in calcular_conquistas(d, "brancas"):
                        st.markdown(ic + " **" + nm + "** — " + ds)
                with col2:
                    st.markdown("**" + c["negras"] + "**")
                    for ic, nm, ds in calcular_conquistas(d, "negras"):
                        st.markdown(ic + " **" + nm + "** — " + ds)

                st.markdown("### 📈 Evolução da Partida")
                fig = gerar_grafico(d)
                st.pyplot(fig)

                st.markdown("### 📊 Estatísticas")
                cats_tab = ["Lance de Mestre", "Livro", "Melhor", "Excelente", "Bom",
                            "Imprecisao", "Erro", "Erro Grave", "Gafe"]
                cat_data = []
                for cat in cats_tab:
                    cat_data.append({
                        "Categoria": cat,
                        c["brancas"]: d["estatisticas"]["brancas"]["classes"].get(cat, 0),
                        c["negras"]: d["estatisticas"]["negras"]["classes"].get(cat, 0),
                    })
                st.dataframe(cat_data, use_container_width=True)

                st.markdown("### 📋 Histórico")
                hist = []
                for l in d["lances"][:80]:
                    hist.append({
                        "#": l["n"],
                        "Lance": l["san"],
                        "Classe": l["classe"],
                        "Perda": str(l["loss"]) + "cp",
                        "Precisão": str(round(l["precisao"])) + "%",
                    })
                st.dataframe(hist, use_container_width=True)

                st.markdown("### 📥 Exportar e Compartilhar")
                col1, col2 = st.columns(2)

                with col1:
                    word_path = gerar_word(d)
                    with open(word_path, "rb") as f:
                        st.download_button(
                            "📄 Word",
                            f,
                            file_name="e4chess_" + c["brancas"] + "_vs_" + c["negras"] + ".docx",
                            use_container_width=True
                        )

                with col2:
                    img_path = gerar_imagem(d)
                    with open(img_path, "rb") as f:
                        st.download_button(
                            "🖼️ Imagem PNG",
                            f,
                            file_name="e4chess_" + c["brancas"] + "_vs_" + c["negras"] + ".png",
                            mime="image/png",
                            use_container_width=True
                        )

                st.markdown("### 📱 Compartilhar")
                legenda = legenda_instagram(d)
                st.text_area("Legenda (copie para Instagram):", value=legenda, height=180)

                with open(img_path, "rb") as f:
                    img_b64 = base64.b64encode(f.read()).decode()

                html_botao = """
                <div style="font-family:sans-serif; padding:10px;">
                  <button onclick="compartilhar()"
                    style="background:#E4405F;color:#fff;border:none;padding:14px 24px;
                           border-radius:10px;font-size:16px;font-weight:600;cursor:pointer;
                           width:100%; margin-bottom:8px;">
                    📤 Compartilhar no Instagram
                  </button>
                  <button onclick="compartilharWhats()"
                    style="background:#25D366;color:#fff;border:none;padding:12px 24px;
                           border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;
                           width:100%;">
                    💬 Compartilhar no WhatsApp
                  </button>
                </div>
                <script>
                const imgData = "data:image/png;base64,IMG_BASE64";
                const legenda = LEGENDA_JS;
                async function compartilhar() {
                  try {
                    const blob = await (await fetch(imgData)).blob();
                    const file = new File([blob], 'e4chess.png', { type: 'image/png' });
                    if (navigator.canShare && navigator.canShare({ files: [file] })) {
                      await navigator.share({ files: [file], title: 'E4Chess', text: legenda });
                    } else {
                      alert('Baixe a imagem e poste no Instagram.');
                    }
                  } catch (e) { console.log(e); }
                }
                function compartilharWhats() {
                  const url = 'https://wa.me/?text=' + encodeURIComponent(legenda);
                  window.open(url, '_blank');
                }
                </script>
                """
                html_botao = html_botao.replace("IMG_BASE64", img_b64)
                html_botao = html_botao.replace("LEGENDA_JS", repr(legenda))

                components.html(html_botao, height=200)

        except Exception as e:
            st.error("❌ " + str(e))

    if btn_limpar:
        st.rerun()


# ═══════════════════════════════════════════════════════════════════
#  ABA RANKING
# ═══════════════════════════════════════════════════════════════════
with aba_ranking:
    st.markdown("## 🏆 Ranking Mundial FIDE")
    st.markdown("*Os melhores jogadores do mundo, segundo a Federação Internacional de Xadrez.*")
    st.markdown("---")

    with st.spinner("Carregando ranking..."):
        ranking = buscar_ranking_fide()

    if ranking:
        for j in ranking:
            try:
                pos = j.get("rank", "—")
                nome = j.get("name", "—")
                pais = j.get("country", "—")
                rating = j.get("rating", "—")
                if pos == 1:
                    medalha = "🥇"
                elif pos == 2:
                    medalha = "🥈"
                elif pos == 3:
                    medalha = "🥉"
                else:
                    medalha = "#" + str(pos)

                st.markdown(
                    "<div style='background:#161b22;border:1px solid #30363d;"
                    "border-radius:10px;padding:14px;margin:8px 0;"
                    "display:flex;justify-content:space-between;align-items:center;'>"
                    "<div style='display:flex;align-items:center;gap:14px;'>"
                    "<div style='font-size:22px;min-width:40px;'>" + medalha + "</div>"
                    "<div>"
                    "<div style='font-weight:700;color:#c9d1d9;font-size:15px;'>" + str(nome) + "</div>"
                    "<div style='color:#8b949e;font-size:12px;'>" + str(pais) + "</div>"
                    "</div></div>"
                    "<div style='font-size:22px;font-weight:700;color:#58a6ff;'>" + str(rating) + "</div>"
                    "</div>",
                    unsafe_allow_html=True
                )
            except Exception:
                continue
    else:
        st.info("Ranking temporariamente indisponível.")

    st.markdown("---")
    st.markdown("### 🎯 Próximos Duelos Importantes")
    duelos = [
        ("Global Chess League 2026", "5 a 13 de setembro · Bengaluru, Índia"),
        ("World Chess Championship 2026", "22 nov a 13 dez · Genebra, Suíça"),
        ("Norway Chess 2026", "Maio/Junho · Oslo, Noruega"),
    ]
    for nome, data in duelos:
        st.markdown("**" + nome + "** — " + data)


# ═══════════════════════════════════════════════════════════════════
#  ABA NOTÍCIAS
# ═══════════════════════════════════════════════════════════════════
with aba_noticias:
    st.markdown("## 📰 Notícias do Xadrez")
    st.markdown("---")

    st.markdown("### 🌍 Internacionais")
    with st.spinner("Carregando..."):
        noticias_int = buscar_noticias("http://www.theweekinchess.com/twic-rss-feed", 5)

    if noticias_int:
        for n in noticias_int:
            st.markdown("**[" + n["titulo"] + "](" + n["link"] + ")**  \n_" + n["data"] + "_")
    else:
        st.info("Não foi possível carregar agora. Tente novamente em instantes.")

    st.markdown("---")
    st.markdown("### 🇧🇷 Nacionais")
    st.info("Em breve: notícias da Confederação Brasileira de Xadrez.")

    st.markdown("---")
    st.markdown("### 🏙️ Paraíba e Nordeste")
    st.info("Em breve: cobertura do xadrez regional.")

    st.markdown("---")
    st.markdown("### 🗓️ Próximos Torneios")
    torneios = [
        ("Global Chess League 2026", "5 a 13 de setembro · Bengaluru, Índia"),
        ("World Chess Championship 2026", "22 nov a 13 dez · Genebra, Suíça"),
        ("Norway Chess 2026", "Maio/Junho · Oslo, Noruega"),
    ]
    for nome, data in torneios:
        st.markdown("**" + nome + "** — " + data)


# ═══════════════════════════════════════════════════════════════════
#  ABA COMO FUNCIONA O E4 RATING
# ═══════════════════════════════════════════════════════════════════
with aba_rating:
    st.markdown("## 📊 Como funciona o E4 Rating")
    st.markdown("""
O **E4 Rating** é a nossa métrica de performance. Ele responde à pergunta:

> *"Qual rating um jogador precisaria ter para jogar exatamente assim contra este oponente?"*

---

### 🎯 Fórmula Transparente

**E4 Rating = (Rating do Oponente + Bônus de Performance) × Fator de Dificuldade**

---

### 1️⃣ Rating do Oponente (Base)
Ponto de partida. Se você venceu um jogador de 1500, sua nota começa em 1500.

### 2️⃣ Bônus de Performance (por precisão)
- Precisão **≥ 90%** → Bônus de **+300 a +500**
- Precisão **80-89%** → Bônus de **+150 a +300**
- Precisão **70-79%** → Bônus de **+50 a +150**
- Precisão **< 70%** → Bônus de **0 a +50**

### 3️⃣ Bônus Adicionais (Diferenciais E4)
- **+50 pontos** por cada Lance de Mestre reproduzido
- **+150 pontos** se não cometeu Gafe nem Erro Grave
- **-100 pontos** por cada Gafe
- **-40 pontos** por cada Erro Grave

### 4️⃣ Fator de Dificuldade
- Oponente muito superior (diferença ≥ 400) → **× 1.5**
- Oponente superior (diferença ≥ 100) → **× 1.2**
- Partida equilibrada → **× 1.0**

---

### 🏆 Por que é Justo?

- **Não é achismo:** cada ponto tem justificativa.
- **Premia o esforço:** jogar bem contra um adversário forte eleva a nota.
- **Útil para todos os níveis:** um iniciante que joga uma partida perfeita contra oponente similar pode alcançar nota alta.

---

### 📚 Fontes de Referência

- **Sistema Elo da FIDE** — fórmula de rating oficial
- **Regra dos 400 pontos** — diferença máxima considerada
- **K-factor** — coeficiente de desenvolvimento
- **Filosofia do Chess.com** — precisão e classificação por lance
- **Filosofia do Lichess** — probabilidade de vitória em cada posição

*O E4 Rating não substitui o rating oficial da FIDE. É uma métrica complementar e transparente.*
    """)


# ═══════════════════════════════════════════════════════════════════
#  ABA PENSADORES
# ═══════════════════════════════════════════════════════════════════
with aba_pensadores:
    st.markdown("## 🧠 Pensadores e o Xadrez")
    st.markdown("*Grandes mentes que encontraram no tabuleiro um espelho da própria inteligência.*")
    st.markdown("---")
    st.markdown("""
> *"O xadrez é a ginástica da inteligência."*
> — **Blaise Pascal**, matemático e filósofo
    """)
    st.markdown("---")

    st.markdown("### 📜 Linha do Tempo do Xadrez")
    linha_tempo = [
        ("Século VI", "Índia", "Chaturanga", "O jogo nasce como 'os quatro elementos de um exército' em sânscrito."),
        ("Século VII", "Pérsia", "Shatranj", "Surge a palavra 'Xeque-Mate' (Shah Mat = o rei está morto)."),
        ("Século IX", "Mundo Árabe", "Difusão", "Os árabes levam o jogo para o norte da África e Península Ibérica."),
        ("Século XV", "Europa", "Renascença", "As regras modernas surgem na Itália. Dama e Bispo ganham mobilidade."),
        ("1886", "EUA", "Steinitz", "Wilhelm Steinitz torna-se o primeiro campeão mundial oficial."),
        ("1972", "Islândia", "Fischer", "Bobby Fischer vence Spassky no 'Match do Século'."),
        ("1985", "Rússia", "Kasparov", "Aos 22 anos, Kasparov é o campeão mundial mais jovem."),
        ("2013", "Noruega", "Carlsen", "Magnus Carlsen inicia sua era de domínio no xadrez mundial."),
    ]
    for periodo, local, nome, desc in linha_tempo:
        st.markdown(
            "<div style='background:#161b22;border-left:4px solid #58a6ff;"
            "border-radius:8px;padding:12px;margin:8px 0;'>"
            "<div style='color:#58a6ff;font-weight:700;font-size:15px;'>" + periodo + " · " + local + "</div>"
            "<div style='color:#c9d1d9;font-weight:600;margin:4px 0;'>" + nome + "</div>"
            "<div style='color:#8b949e;font-size:13px;'>" + desc + "</div>"
            "</div>",
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### 🏆 Campeões Mundiais e Seus Estilos")

    campeoes = [
        ("👑", "Wilhelm Steinitz", "1886-1894", "O Cientista", "Introduziu os princípios posicionais da base estratégica moderna."),
        ("🧮", "Emanuel Lasker", "1894-1921", "O Matemático", "Doutor em matemática, reinou 27 anos com abordagem psicológica."),
        ("🌟", "José Raúl Capablanca", "1921-1927", "O Gênio Natural", "Considerado o maior talento natural da história do xadrez."),
        ("⚙️", "Mikhail Botvinnik", "1948-1963", "O Engenheiro", "Doutor em engenharia, criou o primeiro programa de xadrez soviético."),
        ("🎭", "Mikhail Tal", "1960-1961", "O Mágico", "Conhecido pelo estilo de ataque brilhante e sacrificial."),
        ("🦅", "Bobby Fischer", "1972-1975", "O Prodígio", "Campeão aos 29 anos, revolucionou o xadrez americano."),
        ("👊", "Garry Kasparov", "1985-2000", "O Rei", "Campeão mundial mais jovem da história aos 22 anos."),
        ("🎼", "Magnus Carlsen", "2013-2023", "O Mozart", "Dominou a era moderna com precisão sobre-humana."),
    ]
    for emoji, nome, periodo, titulo, desc in campeoes:
        st.markdown(
            "<div style='background:#161b22;border:1px solid #30363d;border-radius:12px;"
            "padding:16px;margin:10px 0;'>"
            "<div style='display:flex;align-items:center;gap:12px;'>"
            "<div style='font-size:32px;'>" + emoji + "</div>"
            "<div>"
            "<div style='font-weight:700;color:#58a6ff;font-size:16px;'>" + nome + "</div>"
            "<div style='color:#8b949e;font-size:12px;'>" + periodo + " · <b>" + titulo + "</b></div>"
            "</div></div>"
            "<div style='color:#c9d1d9;font-size:13px;margin-top:8px;line-height:1.5;'>" + desc + "</div>"
            "</div>",
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### 📖 Citações Bíblicas e Reflexões")
    st.markdown("""
> *"Pois com a sabedoria se faz a guerra, e a vitória está na multidão dos conselheiros."*
> — **Provérbios 24:6**

> *"Prepare-se o cavalo para o dia da batalha, mas a vitória vem do Senhor."*
> — **Provérbios 21:31**

> *"Se te mostrares fraco no dia da angústia, a tua força é pequena."*
> — **Provérbios 24:10**

A Bíblia valoriza a sabedoria e o planejamento estratégico. O xadrez é uma metáfora para a vida.
    """)

    st.markdown("---")
    st.markdown("### 🧠 Grandes Pensadores e o Xadrez")

    pensadores_chess = [
        ("🔭", "Galileu Galilei", "O astrônomo italiano jogava xadrez e via no tabuleiro um exercício de lógica."),
        ("📜", "Baruch Spinoza", "O filósofo holandês encontrou no xadrez uma metáfora para a busca da verdade."),
        ("➗", "Gottfried Wilhelm Leibniz", "O matemático alemão via no xadrez uma expressão da harmonia do universo."),
        ("⚗️", "Dmitri Mendeleev", "O químico russo, criador da tabela periódica, era enxadrista entusiasta."),
        ("💻", "Alan Turing", "O pai da computação criou o Turochamp, um dos primeiros programas de xadrez (1948)."),
        ("🎯", "John von Neumann", "O matemático húngaro desenvolveu a Teoria dos Jogos."),
        ("📡", "Claude Shannon", "Determinou a árvore de complexidade do xadrez (o 'Número de Shannon')."),
    ]
    for emoji, nome, desc in pensadores_chess:
        st.markdown(
            "<div style='background:#161b22;border:1px solid #30363d;border-radius:10px;"
            "padding:14px;margin:8px 0;'>"
            "<div style='font-weight:700;color:#c9d1d9;font-size:15px;'>" + emoji + " " + nome + "</div>"
            "<div style='color:#8b949e;font-size:13px;margin-top:4px;'>" + desc + "</div>"
            "</div>",
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### 🕯️ E houve também **Ramatis Santos Pessoa de Luna**")
    st.markdown("""
> *"Meu pai era matemático e professor. Não foi campeão mundial, não escreveu livros famosos, mas carregou a mesma essência desses grandes pensadores: uma mente brilhante, um coração generoso e o amor pelo conhecimento.*
> 
> *Foi ele quem me ensinou a jogar xadrez. Cada lance que analiso hoje neste aplicativo carrega um pouco da inteligência que ele me transmitiu.*
> 
> *Esta aba é para ele. E para todos os pais que ensinam seus filhos a pensar."*
> 
> **— Joaldo Farias Pessoa de Luna**, criador do E4Chess
    """)


# ═══════════════════════════════════════════════════════════════════
#  ABA NOSSA HISTÓRIA
# ═══════════════════════════════════════════════════════════════════
with aba_historia:
    st.markdown("## 🕯️ Nossa História")
    st.markdown("### *O Legado de Ramatis Santos Pessoa de Luna*")
    st.markdown("---")
    st.markdown("""
> *"Todo grande projeto nasce de uma grande inspiração."*

O **E4Chess** é uma homenagem de um filho ao seu pai.

**Ramatis Santos Pessoa de Luna** foi matemático e professor. Um homem de mente brilhante, de coração generoso e de sorriso fácil. Daqueles que enxergam beleza nos números e ensinam com paciência.

Foi ele quem ensinou ao filho **Joaldo Farias Pessoa de Luna** os primeiros movimentos no tabuleiro. Foi ele quem plantou a semente deste projeto.

Hoje, cada análise gerada pelo E4Chess carrega um pouco daquela herança. É a forma que encontramos de manter viva a inteligência, o amor e a paixão pelo xadrez de um pai que continua inspirando o filho em cada lance.

**Este projeto é para ele.**

---

*Em memória de Ramatis Santos Pessoa de Luna — o professor que ensinou o primeiro lance.*
    """)
