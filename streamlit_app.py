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


@st.cache_data(ttl=86400, show_spinner=False)
def consultar_mestres(fen):
    try:
        url = "https://explorer.lichess.ovh/masters?fen=" + requests.utils.quote(fen) + "&moves=1&topGames=20"
        r = requests.get(url, timeout=6)
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


LIVRO = {
    ("e4", "e5", "Nf3", "Nc6", "Bc4"): "Abertura Italiana",
    ("e4", "e5", "Nf3", "Nc6", "Bb5"): "Abertura Espanhola",
    ("e4", "e5", "Nf3", "Nc6", "d4"): "Abertura Escocesa",
    ("e4", "e5", "Nf3", "Nc6", "Nc3"): "Abertura dos Quatro Cavalos",
    ("e4", "e5", "Nf3", "Nc6", "Bc4", "Bc5"): "Italiana Clássica",
    ("e4", "e5", "Nf3", "Nc6", "Bc4", "Nf6"): "Defesa dos Dois Cavalos",
    ("e4", "e5", "Nf3", "Nf6"): "Defesa Petroff",
    ("e4", "e5", "Nc3"): "Abertura Vienense",
    ("e4", "e5", "Bc4"): "Abertura do Bispo",
    ("e4", "e5", "f4"): "Gambito do Rei",
    ("e4", "e5", "d4"): "Gambito do Centro",
    ("e4", "e5"): "Abertura Aberta",
    ("e4", "c5"): "Defesa Siciliana",
    ("e4", "c5", "Nf3", "d6"): "Siciliana Najdorf",
    ("e4", "c5", "Nf3", "Nc6"): "Siciliana Clássica",
    ("e4", "c5", "Nf3", "e6"): "Siciliana Kan",
    ("e4", "c5", "Nc3"): "Siciliana Fechada",
    ("e4", "c5", "c3"): "Siciliana Alapin",
    ("e4", "e6"): "Defesa Francesa",
    ("e4", "e6", "d4", "d5"): "Francesa Clássica",
    ("e4", "c6"): "Defesa Caro-Kann",
    ("e4", "c6", "d4", "d5"): "Caro-Kann Clássica",
    ("e4", "d5"): "Defesa Escandinava",
    ("e4", "Nf6"): "Defesa Alekhine",
    ("e4", "d6"): "Defesa Pirc",
    ("e4", "g6"): "Defesa Moderna",
    ("e4", "Nc6"): "Defesa Nimzowitsch",
    ("d4", "d5", "c4"): "Gambito da Dama",
    ("d4", "d5", "c4", "e6"): "Gambito da Dama Recusado",
    ("d4", "d5", "c4", "c6"): "Defesa Eslava",
    ("d4", "d5", "c4", "dxc4"): "Gambito da Dama Aceito",
    ("d4", "Nf6", "c4", "g6"): "Índia do Rei",
    ("d4", "Nf6", "c4", "e6"): "Índia da Dama",
    ("d4", "Nf6", "c4", "e6", "Nc3", "Bb4"): "Nimzo-Índia",
    ("d4", "Nf6", "c4", "e6", "Nf3", "b6"): "Índia da Dama Clássica",
    ("d4", "Nf6", "c4", "c5"): "Defesa Benoni",
    ("d4", "Nf6", "c4", "d6"): "Defesa Índia Antiga",
    ("d4", "d5", "Nf3", "Nf6", "Bf4"): "Sistema London",
    ("d4", "f5"): "Defesa Holandesa",
    ("d4", "e6"): "Peão da Dama",
    ("d4", "d5"): "Peão da Dama",
    ("c4",): "Abertura Inglesa",
    ("c4", "e5"): "Inglesa Simétrica",
    ("c4", "Nf6"): "Inglesa Anglo-Índia",
    ("Nf3",): "Abertura Réti",
    ("Nf3", "d5", "c4"): "Réti Gambito",
    ("d4",): "Peão da Dama",
    ("e4",): "Peão do Rei",
    ("b3",): "Abertura Larsen",
    ("g3",): "Abertura Benko",
    ("f4",): "Abertura Bird",
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


def detectar_sacrificio(board_antes, board_depois, cor):
    valores = {chess.PAWN: 1, chess.KNIGHT: 3, chess.BISHOP: 3,
               chess.ROOK: 5, chess.QUEEN: 9, chess.KING: 0}
    def mat(b, c):
        total = 0
        for sq in chess.SQUARES:
            p = b.piece_at(sq)
            if p:
                v = valores.get(p.piece_type, 0)
                total += v if p.color == c else -v
        return total
    antes = mat(board_antes, cor)
    depois = mat(board_depois, cor)
    return (antes - depois) >= 2


def classificar(loss, is_best, san, mate_a, mate_d, em_livro, em_mestre, is_top3=False, sacrificio=False):
    if is_best and em_mestre:
        return "🌟 Lendário"
    if em_mestre and loss <= 30:
        return "💎 Lance de Mestre"
    if sacrificio and is_best:
        return "🏅 Brilhante"
    if "#" in san:
        return "👑 Xeque-Mate"
    if em_livro:
        return "📖 Livro"
    if is_best:
        return "⭐ Melhor"
    if is_top3 and loss <= 15:
        return "✨ Excelente"
    if mate_a is not None and mate_a > 0 and mate_d is None:
        return "🚨 Erro Grave"
    if mate_a is not None and mate_d is not None:
        if abs(mate_d) > abs(mate_a) + 2:
            return "🚨 Erro Grave"
    if loss <= 25:
        return "✅ Bom"
    if loss <= 90:
        return "🤏 Imprecisão"
    if loss <= 200:
        return "⚠️ Erro"
    if loss <= 500:
        return "🚨 Erro Grave"
    return "💀 Gafe"


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


MESTRES_LENDARIOS = {
    "Carlsen, Magnus": ("Magnus Carlsen", "🇳🇴", "Escola Nórdica", "O Mozart do Xadrez"),
    "Carlsen,Magnus": ("Magnus Carlsen", "🇳🇴", "Escola Nórdica", "O Mozart do Xadrez"),
    "Kasparov, Garry": ("Garry Kasparov", "🇷🇺", "Escola Soviética", "O Rei do Xadrez"),
    "Karpov, Anatoly": ("Anatoly Karpov", "🇷🇺", "Escola Soviética", "A Jiboia"),
    "Tal, Mikhail": ("Mikhail Tal", "🇱🇻", "Escola Soviética", "O Mágico de Riga"),
    "Petrosian, Tigran": ("Tigran Petrosian", "🇦🇲", "Escola Soviética", "A Muralha de Ferro"),
    "Spassky, Boris": ("Boris Spassky", "🇷🇺", "Escola Soviética", "O Cavalheiro"),
    "Botvinnik, Mikhail": ("Mikhail Botvinnik", "🇷🇺", "Escola Soviética", "O Patriarca"),
    "Smyslov, Vasily": ("Vasily Smyslov", "🇷🇺", "Escola Soviética", "O Cantor"),
    "Kramnik, Vladimir": ("Vladimir Kramnik", "🇷🇺", "Escola Soviética", "O Urso"),
    "Nepomniachtchi, Ian": ("Ian Nepomniachtchi", "🇷🇺", "Escola Soviética", "O Desafiante"),
    "Karjakin, Sergey": ("Sergey Karjakin", "🇷🇺", "Escola Soviética", "O Ministro"),
    "Fischer, Robert": ("Bobby Fischer", "🇺🇸", "Escola Americana", "O Prodígio"),
    "Fischer, Robert James": ("Bobby Fischer", "🇺🇸", "Escola Americana", "O Prodígio"),
    "Nakamura, Hikaru": ("Hikaru Nakamura", "🇺🇸", "Escola Americana", "O Rei do Blitz"),
    "Caruana, Fabiano": ("Fabiano Caruana", "🇺🇸", "Escola Americana", "O Estrategista"),
    "So, Wesley": ("Wesley So", "🇺🇸", "Escola Americana", "O Silencioso"),
    "Aronian, Levon": ("Levon Aronian", "🇺🇸", "Escola Armênia", "O Mago Armênio"),
    "Anand, Viswanathan": ("Viswanathan Anand", "🇮🇳", "Escola Indiana", "O Tigre de Madras"),
    "Gukesh, D": ("Gukesh D", "🇮🇳", "Escola Indiana", "O Fenômeno"),
    "Gukesh D": ("Gukesh D", "🇮🇳", "Escola Indiana", "O Fenômeno"),
    "Erigaisi, Arjun": ("Arjun Erigaisi", "🇮🇳", "Escola Indiana", "A Fera"),
    "Praggnanandhaa, R": ("Rameshbabu Praggnanandhaa", "🇮🇳", "Escola Indiana", "A Joia"),
    "Vidit, Santosh Gujrathi": ("Vidit Gujrathi", "🇮🇳", "Escola Indiana", "O Estrategista Indiano"),
    "Harikrishna, Pentala": ("Pentala Harikrishna", "🇮🇳", "Escola Indiana", "O Pilar"),
    "Firouzja, Alireza": ("Alireza Firouzja", "🇫🇷", "Escola Francesa", "O Mozart Persa"),
    "Vachier-Lagrave, Maxime": ("Maxime Vachier-Lagrave", "🇫🇷", "Escola Francesa", "O Francês Voador"),
    "Bacrot, Etienne": ("Étienne Bacrot", "🇫🇷", "Escola Francesa", "O Veterano"),
    "Ding, Liren": ("Ding Liren", "🇨🇳", "Escola Chinesa", "O Campeão Silencioso"),
    "Wei, Yi": ("Wei Yi", "🇨🇳", "Escola Chinesa", "O Estrategista Chinês"),
    "Wang, Hao": ("Wang Hao", "🇨🇳", "Escola Chinesa", "O Mestre Chinês"),
    "Abdusattorov, Nodirbek": ("Nodirbek Abdusattorov", "🇺🇿", "Escola Uzbeque", "O Menino Maravilha"),
    "Kasimdzhanov, Rustam": ("Rustam Kasimdzhanov", "🇺🇿", "Escola Uzbeque", "O Campeão Uzbeque"),
    "Giri, Anish": ("Anish Giri", "🇳🇱", "Escola Holandesa", "O Analista"),
    "Van Foreest, Jorden": ("Jorden Van Foreest", "🇳🇱", "Escola Holandesa", "A Nova Geração"),
    "Radjabov, Teimour": ("Teimour Radjabov", "🇦🇿", "Escola Azeri", "O Estrategista Azeri"),
    "Mamedyarov, Shakhriyar": ("Shakhriyar Mamedyarov", "🇦🇿", "Escola Azeri", "O Selvagem"),
    "Leko, Peter": ("Peter Leko", "🇭🇺", "Escola Húngara", "O Preciso"),
    "Topalov, Veselin": ("Veselin Topalov", "🇧🇬", "Escola Búlgara", "O Touro de Sófia"),
    "Adams, Michael": ("Michael Adams", "🇬🇧", "Escola Britânica", "O Estrategista Inglês"),
    "Howell, David": ("David Howell", "🇬🇧", "Escola Britânica", "O Comentarista"),
    "Nisipeanu, Liviu-Dieter": ("Liviu-Dieter Nisipeanu", "🇩🇪", "Escola Alemã", "O Alemão"),
}


def identificar_mestre_lendario(nome_lichess):
    if not nome_lichess:
        return None
    if nome_lichess in MESTRES_LENDARIOS:
        return MESTRES_LENDARIOS[nome_lichess]
    for chave, dados in MESTRES_LENDARIOS.items():
        if chave.lower() in nome_lichess.lower() or nome_lichess.lower() in chave.lower():
            return dados
    return None


def buscar_partida_historica(fen, san_jogado):
    try:
        url = (
            "https://explorer.lichess.ovh/masters?fen="
            + requests.utils.quote(fen)
            + "&moves=1&topGames=20"
        )
        r = requests.get(url, timeout=6)
        if r.status_code != 200:
            return None
        data = r.json()
        top = data.get("topGames") or []
        for g in top:
            uci = g.get("uci", "")
            if not uci or len(uci) < 4:
                continue
            try:
                board = chess.Board(fen)
                mv = chess.Move.from_uci(uci)
                san_game = board.san(mv)
            except Exception:
                continue
            if san_game != san_jogado:
                continue
            brancas_nome = (g.get("white") or {}).get("name", "")
            negras_nome = (g.get("black") or {}).get("name", "")
            mestre = identificar_mestre_lendario(brancas_nome) or \
                     identificar_mestre_lendario(negras_nome)
            if not mestre:
                continue
            return {
                "mestre": mestre[0],
                "bandeira": mestre[1],
                "escola": mestre[2],
                "apelido": mestre[3],
                "brancas": brancas_nome,
                "negras": negras_nome,
                "ano": g.get("year", "—"),
                "resultado": g.get("winner", "—"),
                "id": g.get("id", ""),
                "uci": uci,
                "san": san_jogado,
            }
    except Exception:
        pass
    return None


def card_mestre_lendario(info):
    if not info:
        return ""
    nome = info["mestre"]
    emoji = info["bandeira"]
    escola = info["escola"]
    apelido = info["apelido"]
    brancas = info["brancas"]
    negras = info["negras"]
    ano = info["ano"]
    game_id = info["id"]
    link = "https://lichess.org/" + str(game_id) if game_id else ""
    return (
        "<div style='background:linear-gradient(135deg,#8a6d1a 0%,#d4a017 50%,#f5c542 100%);"
        "padding:20px;border-radius:14px;color:#0d1117;margin:12px 0;"
        "box-shadow:0 4px 24px rgba(212,160,23,0.35);'>"
        "<div style='display:flex;align-items:center;gap:14px;'>"
        f"<div style='font-size:42px;'>{emoji}✨</div>"
        "<div>"
        "<div style='font-size:12px;font-weight:700;text-transform:uppercase;"
        "letter-spacing:1.5px;opacity:0.75;'>Você jogou como um Mestre</div>"
        f"<div style='font-size:22px;font-weight:800;margin-top:2px;'>{nome}</div>"
        f"<div style='font-size:13px;font-weight:600;opacity:0.85;'>"
        f"{escola} · {apelido}</div>"
        "</div></div>"
        "<div style='margin-top:14px;padding:12px;background:rgba(13,17,23,0.85);"
        "border-radius:10px;color:#f5c542;font-size:13px;line-height:1.5;'>"
        f"📜 Em <b>{ano}</b>, na partida <b>{brancas} vs {negras}</b>, "
        f"o lance <b>{info['san']}</b> foi jogado exatamente assim. "
        "Você está pensando como os grandes mestres!"
        "</div>"
        + (f"<a href='{link}' target='_blank' style='display:inline-block;"
           "margin-top:12px;padding:10px 18px;background:#0d1117;color:#f5c542;"
           "border-radius:8px;text-decoration:none;font-weight:700;font-size:13px;'>"
           "🔗 Ver a partida original</a>" if link else "")
        + "</div>"
    )


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


def analisar_stockfish(game, prof, progress_cb=None):
    try:
        eng = chess.engine.SimpleEngine.popen_uci(STOCKFISH_PATH)
    except Exception as e:
        st.warning("Stockfish: " + str(e))
        return None

    try:
        eng.configure({"Threads": 2, "Hash": 128})
    except Exception:
        pass

    board = game.board()
    lances = list(game.mainline_moves())
    n = len(lances)
    aval = []
    melhores = []

    if n > 100:
        st.warning("⚠️ Partida longa (" + str(n) + " lances). A análise pode demorar.")

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
        best_move = pv[0] if pv else None
        aval.append((cp, mate))
        melhores.append([best_move] if best_move else [])
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

        is_best = (best is not None and len(best) > 0 and move == best[0])
        is_top3 = (isinstance(best, list) and move in best)
        resultados.append({
            "cp_best": cp_i if cor == "brancas" else -cp_i,
            "cp_played": cp_played,
            "loss": loss,
            "is_best": is_best,
            "is_top3": is_top3,
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
        board_antes = board.copy()
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

        feito_lendario = None
        if em_mestre:
            feito_lendario = buscar_partida_historica(fen, san)

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

            sacrifice = detectar_sacrificio(
                board_antes, board,
                chess.WHITE if cor == "brancas" else chess.BLACK
            )

            cls = classificar(
                info["loss"], info["is_best"], san,
                info["mate_antes"], info["mate_depois"],
                em_livro, em_mestre,
                is_top3=info.get("is_top3", False),
                sacrificio=sacrifice
            )
            loss = info["loss"]
            ev = info["cp_played"]
            cp_w = info["cp_white_after"]
        else:
            prec = 100.0 if (em_livro or em_mestre) else 50.0
            if em_mestre:
                cls = "💎 Lance de Mestre"
            elif em_livro:
                cls = "📖 Livro"
            else:
                cls = "✅ Bom"
            loss = 0
            ev = 0
            cp_w = 0

        cp_vis = max(-1000, min(1000, cp_w))
        curva.append(cp_vis / 100)

        if cls in ("⚠️ Erro", "🚨 Erro Grave", "💀 Gafe", "🤏 Imprecisão"):
            marcadores.append({"ply": i + 1, "cp": cp_vis / 100, "classe": cls})

        dados.append({
            "n": i + 1, "cor": cor, "san": san, "classe": cls,
            "precisao": prec, "loss": loss, "eval": ev, "mestre": em_mestre,
            "feito_lendario": feito_lendario,
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
        stats["brancas"]["classes"].get("💀 Gafe", 0),
        stats["brancas"]["classes"].get("🚨 Erro Grave", 0),
        cab["rating_negras"], cab["rating_brancas"]
    )
    r_n = calcular_e4_rating(
        pn, stats["negras"]["mestres"],
        stats["negras"]["classes"].get("💀 Gafe", 0),
        stats["negras"]["classes"].get("🚨 Erro Grave", 0),
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
        "feitos_lendarios": [l for l in dados if l.get("feito_lendario")],
    }

PECAS_SVG = {
    'K': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="none" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22.5 11.63V6M20 8h5" stroke-linejoin="miter"/><path d="M22.5 25s4.5-7.5 3-10.5c0 0-1-2.5-3-2.5s-3 2.5-3 2.5c-1.5 3 3 10.5 3 10.5" fill="#fff" stroke-linecap="butt" stroke-linejoin="miter"/><path d="M11.5 37c5.5 3.5 15.5 3.5 21 0v-7s9-4.5 6-10.5c-4-6.5-13.5-3.5-16 4V27v-3.5c-3.5-7.5-13-10.5-16-4-3 6 5 10 5 10V37z" fill="#fff"/><path d="M11.5 30c5.5-3 15.5-3 21 0m-21 3.5c5.5-3 15.5-3 21 0m-21 3.5c5.5-3 15.5-3 21 0"/></g></svg>',
    'Q': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#fff" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M8 12a2 2 0 1 1-4 0 2 2 0 1 1 4 0zM24.5 7.5a2 2 0 1 1-4 0 2 2 0 1 1 4 0zM41 12a2 2 0 1 1-4 0 2 2 0 1 1 4 0zM16 8.5a2 2 0 1 1-4 0 2 2 0 1 1 4 0zM33 9a2 2 0 1 1-4 0 2 2 0 1 1 4 0z"/><path d="M9 26c8.5-1.5 21-1.5 27 0l2-12-7 11V11l-5.5 13.5-3-15-3 15-5.5-14V25L7 14l2 12z" stroke-linecap="butt"/><path d="M9 26c0 2 1.5 2 2.5 4 1 1.5 1 1 .5 3.5-1.5 1-1.5 2.5-1.5 2.5-1.5 1.5.5 2.5.5 2.5 6.5 1 16.5 1 23 0 0 0 1.5-1 0-2.5 0 0 .5-1.5-1-2.5-.5-2.5-.5-2 .5-3.5 1-2 2.5-2 2.5-4-8.5-1.5-18.5-1.5-27 0z" stroke-linecap="butt"/><path d="M11.5 30c3.5-1 18.5-1 22 0M12 33.5c6-1 15-1 21 0" fill="none"/></g></svg>',
    'R': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#fff" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 39h27v-3H9v3zM12 36v-4h21v4H12zM11 14V9h4v2h5V9h5v2h5V9h4v5" stroke-linecap="butt"/><path d="M34 14l-3 3H14l-3-3"/><path d="M31 17v12.5H14V17" stroke-linecap="butt" stroke-linejoin="miter"/><path d="M31 29.5l1.5 2.5h-20l1.5-2.5"/><path d="M11 14h23" fill="none" stroke-linejoin="miter"/></g></svg>',
    'B': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#fff" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><g fill="#fff" stroke-linecap="butt"><path d="M9 36c3.39-.97 10.11.43 13.5-2 3.39 2.43 10.11 1.03 13.5 2 0 0 1.65.54 3 2-.68.97-1.65.99-3 .5-3.39-.97-10.11.46-13.5-1-3.39 1.46-10.11.03-13.5 1-1.354.49-2.323.47-3-.5 1.354-1.94 3-2 3-2z"/><path d="M15 32c2.5 2.5 12.5 2.5 15 0 .5-1.5 0-2 0-2 0-2.5-2.5-4-2.5-4 5.5-1.5 6-11.5-5-15.5-11 4-10.5 14-5 15.5 0 0-2.5 1.5-2.5 4 0 0-.5.5 0 2z"/><path d="M25 8a2.5 2.5 0 1 1-5 0 2.5 2.5 0 1 1 5 0z"/></g><path d="M17.5 26h10M15 30h15m-7.5-14.5v5M20 18h5" stroke-linejoin="miter"/></g></svg>',
    'N': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="none" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10c10.5 1 16.5 8 16 29H15c0-9 10-6.5 8-21" fill="#fff"/><path d="M24 18c.38 2.91-5.55 7.37-8 9-3 2-2.82 4.34-5 4-1.042-.94 1.41-3.04 0-3-1 0 .19 1.23-1 2-1 0-4.003 1-4-4 0-2 6-12 6-12s1.89-1.9 2-3.5c-.73-.994-.5-2-.5-3 1-1 3 2.5 3 2.5h2s.78-1.992 2.5-3c1 0 1 3 1 3" fill="#fff"/><path d="M9.5 25.5a.5.5 0 1 1-1 0 .5.5 0 1 1 1 0zM14.933 15.75a.5 1.5 30 1 1-.866-.5.5 1.5 30 1 1 .866.5z" fill="#000" stroke="#000"/></g></svg>',
    'P': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><path d="M22.5 9c-2.21 0-4 1.79-4 4 0 .89.29 1.71.78 2.38C17.33 16.5 16 18.59 16 21c0 2.03.94 3.84 2.41 5.03-3 1.06-7.41 5.55-7.41 13.47h23c0-7.92-4.41-12.41-7.41-13.47 1.47-1.19 2.41-3 2.41-5.03 0-2.41-1.33-4.5-3.28-5.62.49-.67.78-1.49.78-2.38 0-2.21-1.79-4-4-4z" fill="#fff" stroke="#000" stroke-width="1.5" stroke-linecap="round"/></svg>',
    'k': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="none" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22.5 11.63V6" stroke-linejoin="miter"/><path d="M22.5 25s4.5-7.5 3-10.5c0 0-1-2.5-3-2.5s-3 2.5-3 2.5c-1.5 3 3 10.5 3 10.5" fill="#000" stroke-linecap="butt" stroke-linejoin="miter"/><path d="M11.5 37c5.5 3.5 15.5 3.5 21 0v-7s9-4.5 6-10.5c-4-6.5-13.5-3.5-16 4V27v-3.5c-3.5-7.5-13-10.5-16-4-3 6 5 10 5 10V37z" fill="#000"/><path d="M20 8h5" stroke-linejoin="miter"/><path d="M32 29.5s8.5-4 6.03-9.65C34.15 14 25 18 22.5 24.5l.01 2.1-.01-2.1C20 18 9.906 14 6.997 19.85c-2.497 5.65 4.853 9 4.853 9"/><path d="M11.5 30c5.5-3 15.5-3 21 0m-21 3.5c5.5-3 15.5-3 21 0m-21 3.5c5.5-3 15.5-3 21 0"/></g></svg>',
    'q': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#000" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><g fill="#000" stroke="none"><circle cx="6" cy="12" r="2.75"/><circle cx="14" cy="9" r="2.75"/><circle cx="22.5" cy="8" r="2.75"/><circle cx="31" cy="9" r="2.75"/><circle cx="39" cy="12" r="2.75"/></g><path d="M9 26c8.5-1.5 21-1.5 27 0l2.5-12.5L31 25l-.3-14.1-5.2 13.6-3-14.5-3 14.5-5.2-13.6L14 25 6.5 13.5 9 26z" stroke-linecap="butt"/><path d="M9 26c0 2 1.5 2 2.5 4 1 1.5 1 1 .5 3.5-1.5 1-1.5 2.5-1.5 2.5-1.5 1.5.5 2.5.5 2.5 6.5 1 16.5 1 23 0 0 0 1.5-1 0-2.5 0 0 .5-1.5-1-2.5-.5-2.5-.5-2 .5-3.5 1-2 2.5-2 2.5-4-8.5-1.5-18.5-1.5-27 0z" stroke-linecap="butt"/><path d="M11 38.5a35 35 1 0 0 23 0" fill="none" stroke-linecap="butt"/><path d="M11 29a35 35 1 0 1 23 0M12.5 31.5h20M11.5 34.5a35 35 1 0 0 22 0M10.5 37.5a35 35 1 0 0 24 0" fill="none" stroke="#fff"/></g></svg>',
    'r': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#000" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 39h27v-3H9v3zM12.5 32l1.5-2.5h17l1.5 2.5h-20zM12 36v-4h21v4H12z" stroke-linecap="butt"/><path d="M14 29.5v-13h17v13H14z" stroke-linecap="butt" stroke-linejoin="miter"/><path d="M14 16.5L11 14h23l-3 2.5H14zM11 14V9h4v2h5V9h5v2h5V9h4v5H11z" stroke-linecap="butt"/><path d="M12 35.5h21M13 31.5h19M14 29.5h17M14 16.5h17M11 14h23" fill="none" stroke="#fff" stroke-width="1" stroke-linejoin="miter"/></g></svg>',
    'b': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="#000" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><g fill="#000" stroke-linecap="butt"><path d="M9 36c3.39-.97 10.11.43 13.5-2 3.39 2.43 10.11 1.03 13.5 2 0 0 1.65.54 3 2-.68.97-1.65.99-3 .5-3.39-.97-10.11.46-13.5-1-3.39 1.46-10.11.03-13.5 1-1.354.49-2.323.47-3-.5 1.354-1.94 3-2 3-2z"/><path d="M15 32c2.5 2.5 12.5 2.5 15 0 .5-1.5 0-2 0-2 0-2.5-2.5-4-2.5-4 5.5-1.5 6-11.5-5-15.5-11 4-10.5 14-5 15.5 0 0-2.5 1.5-2.5 4 0 0-.5.5 0 2z"/><path d="M25 8a2.5 2.5 0 1 1-5 0 2.5 2.5 0 1 1 5 0z"/></g><path d="M17.5 26h10M15 30h15m-7.5-14.5v5M20 18h5" stroke="#fff" stroke-linejoin="miter"/></g></svg>',
    'n': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><g fill="none" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10c10.5 1 16.5 8 16 29H15c0-9 10-6.5 8-21" fill="#000"/><path d="M24 18c.38 2.91-5.55 7.37-8 9-3 2-2.82 4.34-5 4-1.042-.94 1.41-3.04 0-3-1 0 .19 1.23-1 2-1 0-4.003 1-4-4 0-2 6-12 6-12s1.89-1.9 2-3.5c-.73-.994-.5-2-.5-3 1-1 3 2.5 3 2.5h2s.78-1.992 2.5-3c1 0 1 3 1 3" fill="#000"/><path d="M9.5 25.5a.5.5 0 1 1-1 0 .5.5 0 1 1 1 0zM14.933 15.75a.5 1.5 30 1 1-.866-.5.5 1.5 30 1 1 .866.5z" fill="#fff" stroke="#fff"/><path d="M24.55 10.4l-.45 1.45.5.15c3.15 1 5.65 2.49 7.9 6.75S35.75 29.06 35.25 39l-.05.5h2.25l.05-.5c.5-10.06-.88-16.85-3.25-21.34-2.37-4.49-5.79-6.64-9.19-7.16l-.51-.1z" fill="#fff" stroke="none"/></g></svg>',
    'p': '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 45 45"><path d="M22.5 9c-2.21 0-4 1.79-4 4 0 .89.29 1.71.78 2.38C17.33 16.5 16 18.59 16 21c0 2.03.94 3.84 2.41 5.03-3 1.06-7.41 5.55-7.41 13.47h23c0-7.92-4.41-12.41-7.41-13.47 1.47-1.19 2.41-3 2.41-5.03 0-2.41-1.33-4.5-3.28-5.62.49-.67.78-1.49.78-2.38 0-2.21-1.79-4-4-4z" fill="#000" stroke="#000" stroke-width="1.5" stroke-linecap="round"/></svg>',
}

TEMAS_TABULEIRO = {
    "Marrom Clássico": ("#f0d9b5", "#b58863"),
    "Verde Lichess": ("#eeeed2", "#769656"),
    "Azul Oceano": ("#dee3e6", "#8ca2ad"),
    "Cinza Moderno": ("#e8e8e8", "#a0a0a0"),
    "Roxo E4": ("#e8d5f2", "#8e5da8"),
    "Madeira Escura": ("#d9b382", "#7a5230"),
}

CORES_SETAS = ["#003088", "#0d6ec0", "#5ea3d4"]


def render_board_html(fen, setas=None, size=380, tema="Marrom Clássico",
                      estilo="Clássico", show_coords=True,
                      last_move=None, legal_moves=None, rotated=False):
    board = chess.Board(fen)
    light, dark = TEMAS_TABULEIRO.get(tema, TEMAS_TABULEIRO["Marrom Clássico"])

    legal_targets = set()
    if legal_moves:
        for mv in legal_moves:
            try:
                legal_targets.add(chess.square_name(mv.to_square))
            except Exception:
                pass

    last_squares = set()
    if last_move:
        try:
            last_squares.add(chess.square_name(last_move.from_square))
            last_squares.add(chess.square_name(last_move.to_square))
        except Exception:
            pass

    ranks = list(range(8, 0, -1)) if not rotated else list(range(1, 9))
    files = list(range(8)) if not rotated else list(range(7, -1, -1))
    first_file = files[0]
    last_rank = ranks[-1]

    squares_html = []
    for rank in ranks:
        for file in files:
            sq_name = chr(ord('a') + file) + str(rank)
            sq = chess.parse_square(sq_name)
            piece = board.piece_at(sq)
            is_light = (file + rank) % 2 == 0
            bg = light if is_light else dark
            if sq_name in last_squares:
                bg = "#f7ec74"

            svg_piece = ""
            if piece:
                svg_piece = PECAS_SVG.get(piece.symbol(), "")

            coord_color = dark if is_light else light
            rank_label = ""
            file_label = ""
            if show_coords:
                if file == first_file:
                    rank_label = f'<span class="e4rk" style="color:{coord_color};">{rank}</span>'
                if rank == last_rank:
                    file_label = f'<span class="e4fl" style="color:{coord_color};">{chr(ord("a") + file)}</span>'

            if sq_name in legal_targets and not piece:
                inner = '<span class="e4dot"></span>'
            elif sq_name in legal_targets and piece:
                inner = svg_piece + '<span class="e4ring"></span>'
            else:
                inner = svg_piece

            squares_html.append(
                f'<div class="e4sq" style="background:{bg};">{rank_label}{file_label}{inner}</div>'
            )

    arrows = ""
    if setas:
        for i, (frm, to, cor) in enumerate(setas):
            try:
                f1 = ord(frm[0]) - ord('a')
                r1 = int(frm[1]) - 1
                f2 = ord(to[0]) - ord('a')
                r2 = int(to[1]) - 1
                if rotated:
                    x1, y1 = (7 - f1) + 0.5, r1 + 0.5
                    x2, y2 = (7 - f2) + 0.5, r2 + 0.5
                else:
                    x1, y1 = f1 + 0.5, (7 - r1) + 0.5
                    x2, y2 = f2 + 0.5, (7 - r2) + 0.5
                arrows += (
                    f'<defs><marker id="ah{i}" markerWidth="4" markerHeight="4" '
                    f'refX="2" refY="2" orient="auto">'
                    f'<polygon points="0,0 4,2 0,4" fill="{cor}"/></marker></defs>'
                    f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" '
                    f'stroke="{cor}" stroke-width="0.18" opacity="0.85" '
                    f'marker-end="url(#ah{i})"/>'
                )
            except Exception:
                continue

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<style>
html, body {{
    margin: 0; padding: 0;
    background: transparent;
    width: 100%; height: 100%;
    overflow: hidden;
}}
.e4wrap {{
    width: 100%; height: 100%;
    display: flex; align-items: center; justify-content: center;
    padding: 4px; box-sizing: border-box;
}}
.e4board {{
    aspect-ratio: 1 / 1;
    width: min(100%, 100vh);
    height: auto; max-height: 100%;
    display: grid;
    grid-template-columns: repeat(8, 1fr);
    grid-template-rows: repeat(8, 1fr);
    border: 2px solid #30363d;
    border-radius: 6px;
    overflow: hidden;
    position: relative;
    box-shadow: 0 4px 16px rgba(0,0,0,0.5);
    box-sizing: border-box;
}}
.e4sq {{
    display: flex; align-items: center; justify-content: center;
    position: relative;
    user-select: none;
    overflow: hidden;
}}
.e4sq svg {{
    width: 92%; height: 92%;
    display: block;
}}
.e4dot {{
    position: absolute; width: 30%; height: 30%;
    background: rgba(20, 85, 30, 0.5);
    border-radius: 50%;
}}
.e4ring {{
    position: absolute; inset: 6%;
    border: 0.4vmin solid rgba(20, 85, 30, 0.5);
    border-radius: 50%;
    box-sizing: border-box;
}}
.e4board-svg {{
    position: absolute; top: 0; left: 0;
    width: 100%; height: 100%;
    pointer-events: none; z-index: 10;
}}
.e4rk {{
    position: absolute; top: 3%; left: 5%;
    font-size: 1.6vmin; font-weight: 700; line-height: 1;
    font-family: -apple-system, sans-serif; opacity: 0.85;
}}
.e4fl {{
    position: absolute; bottom: 3%; right: 5%;
    font-size: 1.6vmin; font-weight: 700; line-height: 1;
    font-family: -apple-system, sans-serif; opacity: 0.85;
}}
</style>
</head>
<body>
<div class="e4wrap">
<div class="e4board">
{''.join(squares_html)}
<svg class="e4board-svg" viewBox="0 0 8 8" preserveAspectRatio="none">
{arrows}
</svg>
</div>
</div>
</body>
</html>"""
    return html


@st.cache_data(ttl=600, show_spinner=False)
def analisar_top3(fen, prof=12):
    try:
        eng = chess.engine.SimpleEngine.popen_uci(STOCKFISH_PATH)
        eng.configure({"Threads": 1, "Hash": 128, "MultiPV": 3})
        board = chess.Board(fen)
        infos = eng.analyse(board, chess.engine.Limit(depth=prof), multipv=3)
        eng.quit()
        if not isinstance(infos, list):
            infos = [infos]
        setas = []
        for i, info in enumerate(infos[:3]):
            pv = info.get("pv") or []
            if pv:
                mv = pv[0]
                setas.append((
                    chess.square_name(mv.from_square),
                    chess.square_name(mv.to_square),
                    CORES_SETAS[i],
                ))
        return setas
    except Exception:
        return []


def melhor_lance_stockfish(fen, prof):
    try:
        eng = chess.engine.SimpleEngine.popen_uci(STOCKFISH_PATH)
        eng.configure({"Threads": 1, "Hash": 64})
        board = chess.Board(fen)
        info = eng.analyse(board, chess.engine.Limit(depth=prof))
        eng.quit()
        pv = info.get("pv") or []
        return pv[0] if pv else None
    except Exception:
        return None


def init_tab_state():
    defaults = {
        "tb_modo": "📖 Análise",
        "tb_theme": "Marrom Clássico",
        "tb_size": 380,
        "tb_coords": True,
        "tb_auto_rotate": True,
        "tb_show_legal": True,
        "tb_game_obj": None,
        "tb_moves": [],
        "tb_index": 0,
        "tb_initial_fen": chess.STARTING_FEN,
        "tb_players": {"brancas": "Brancas", "negras": "Negras"},
        "tb_stockfish_level": 8,
        "tb_stockfish_color": "negras",
        "tb_correspond_pgn": "",
        "tb_finished": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_tab_game(starting_fen=chess.STARTING_FEN, players=None):
    st.session_state.tb_initial_fen = starting_fen
    st.session_state.tb_game_obj = None
    st.session_state.tb_moves = []
    st.session_state.tb_index = 0
    st.session_state.tb_finished = False
    if players:
        st.session_state.tb_players = players


def get_current_board():
    board = chess.Board(st.session_state.tb_initial_fen)
    for mv in st.session_state.tb_moves[:st.session_state.tb_index]:
        board.push(mv)
    return board


def current_fen():
    return get_current_board().fen()


def pgn_atual():
    game = chess.pgn.Game()
    node = game
    for mv in st.session_state.tb_moves:
        node = node.add_variation(mv)
    return str(game)


def make_move(mv):
    board = get_current_board()
    if mv not in board.legal_moves:
        return False
    st.session_state.tb_moves = st.session_state.tb_moves[:st.session_state.tb_index] + [mv]
    st.session_state.tb_index += 1
    return True

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
        nivel, emoji = "Mestre", "👑"
    elif diferenca >= 400:
        nivel, emoji = "Expert", "🌟"
    elif diferenca >= 250:
        nivel, emoji = "Avançado", "⭐"
    elif diferenca >= 100:
        nivel, emoji = "Intermediário", "✨"
    elif diferenca >= 0:
        nivel, emoji = "Acima do seu nível", "🎯"
    else:
        return None

    return {
        "diferenca": diferenca, "nivel": nivel, "emoji": emoji,
        "rating_real": rating_real, "performance": performance,
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
    if classes.get("💀 Gafe", 0) == 0 and classes.get("🚨 Erro Grave", 0) == 0:
        c.append(("🛡️", "Defensor", "Sem erros graves nem gafes"))
    if classes.get("📖 Livro", 0) >= 6:
        c.append(("📖", "Mestre das Aberturas", str(classes.get("📖 Livro", 0)) + " lances de livro"))
    if classes.get("⭐ Melhor", 0) >= 5:
        c.append(("⭐", "Caçador de Melhores", str(classes["⭐ Melhor"]) + " melhores lances"))
    if classes.get("🌟 Lendário", 0) >= 1:
        c.append(("🌟", "Toque de Gênio", str(classes["🌟 Lendário"]) + " lance(s) Lendário(s)"))
    if classes.get("🏅 Brilhante", 0) >= 1:
        c.append(("🏅", "Brilho E4", str(classes["🏅 Brilhante"]) + " lance(s) Brilhante(s)"))
    if classes.get("💀 Gafe", 0) >= 3:
        c.append(("⚠️", "Precisa Estudar", str(classes["💀 Gafe"]) + " gafes cometidas"))

    sup = calcular_superacao(d, cor)
    if sup and sup["diferenca"] >= 250:
        c.append((
            sup["emoji"],
            "Performance " + sup["nivel"],
            "Jogou como " + str(sup["performance"]) + " (+" + str(sup["diferenca"]) + " acima)"
        ))
    return c


@st.cache_data(ttl=7200, show_spinner=False)
def obter_analise_cache(pgn, prof):
    game = validar_pgn(pgn)
    return analisar(game, prof=prof)


def obter_analise(pgn, prof):
    chave = "v5_" + hashlib.md5(pgn.encode()).hexdigest() + "_" + str(prof)
    if "cache" not in st.session_state:
        st.session_state.cache = {}
    if chave in st.session_state.cache:
        st.info("⚡ Análise recuperada do cache")
        return st.session_state.cache[chave]

    barra = st.progress(0, text="Analisando com Stockfish...")
    d = obter_analise_cache(pgn, prof)
    barra.empty()
    st.session_state.cache[chave] = d
    return d


def cores_grafico():
    return {"🤏 Imprecisão": "#d29922", "⚠️ Erro": "#f0883e",
            "🚨 Erro Grave": "#f85149", "💀 Gafe": "#ff2222"}


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

    for cl, cor in cores_grafico().items():
        pts = [m for m in d["marcadores"] if m["classe"] == cl]
        if pts:
            xs2 = [p["ply"] for p in pts]
            ys2 = [p["cp"] for p in pts]
            ax.scatter(xs2, ys2, color=cor, s=65, edgecolors="#0d1117",
                       linewidths=1.2, label=cl)

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


def gerar_imagem(d):
    c = d["cabecalho"]
    pb, pn = d["precisao_brancas"], d["precisao_negras"]
    nb, nn = d["nota_brancas"], d["nota_negras"]
    rb, rn = d["rating_e4_brancas"], d["rating_e4_negras"]

    def cor_nota(n):
        return {"A+": "#3fb950", "A": "#3fb950", "B+": "#58a6ff", "B": "#58a6ff",
                "C+": "#d29922", "C": "#d29922", "D": "#f0883e", "F": "#f85149"}.get(n, "#c9d1d9")

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
        box = FancyBboxPatch((x0, y_cards), 0.38, 0.11,
                             boxstyle="round,pad=0.01",
                             facecolor="#161b22", edgecolor="#30363d", linewidth=1.5)
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
    ax_g.fill_between(xs, [max(0, y) for y in cc], 0, color="#f0f6fc", alpha=0.92)
    ax_g.fill_between(xs, [min(0, y) for y in cc], 0, color="#0d1117", alpha=0.95)
    ax_g.plot(xs, cc, color="#58a6ff", linewidth=1.2)
    ax_g.axhline(0, color="#30363d", linewidth=0.8)
    for cl, cor_er in cores_grafico().items():
        pts = [m for m in d["marcadores"] if m["classe"] == cl]
        if pts:
            ax_g.scatter([p["ply"] for p in pts], [p["cp"] for p in pts],
                         color=cor_er, s=40, edgecolors="#0d1117", linewidths=0.8)
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
    cats = ["🌟 Lendário", "💎 Lance de Mestre", "🏅 Brilhante",
            "📖 Livro", "⭐ Melhor", "✨ Excelente", "✅ Bom",
            "🤏 Imprecisão", "⚠️ Erro", "🚨 Erro Grave", "💀 Gafe"]

    ax_s.text(0.15, 0.96, c["brancas"][:12], transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#58a6ff", weight="bold")
    ax_s.text(0.5, 0.96, "Categoria", transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#c9d1d9", weight="bold")
    ax_s.text(0.85, 0.96, c["negras"][:12], transform=ax_s.transAxes,
              ha="center", fontsize=10, color="#f0883e", weight="bold")
    for i, cat in enumerate(cats):
        y = 0.84 - i * 0.072
        cb_ = d["estatisticas"]["brancas"]["classes"].get(cat, 0)
        cn_ = d["estatisticas"]["negras"]["classes"].get(cat, 0)
        ax_s.text(0.15, y, str(cb_), transform=ax_s.transAxes,
                  ha="center", fontsize=8, color="#c9d1d9")
        ax_s.text(0.5, y, cat, transform=ax_s.transAxes,
                  ha="center", fontsize=8, color="#c9d1d9")
        ax_s.text(0.85, y, str(cn_), transform=ax_s.transAxes,
                  ha="center", fontsize=8, color="#c9d1d9")

    fig.text(0.5, 0.015, "E4Chess - e4chess.streamlit.app",
             ha="center", fontsize=8, color="#8b949e", style="italic")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    fig.savefig(tmp.name, facecolor="#0d1117")
    plt.close(fig)
    return tmp.name


def legenda_instagram(d):
    c = d["cabecalho"]
    classes_b = d["estatisticas"]["brancas"]["classes"]
    lendarios = classes_b.get("🌟 Lendário", 0)
    mestres = classes_b.get("💎 Lance de Mestre", 0)
    brilhantes = classes_b.get("🏅 Brilhante", 0)

    linhas = ["♟️ Analisei minha partida no E4Chess!", ""]
    linhas.append("⚔️ " + str(c["brancas"]) + " vs " + str(c["negras"]))
    linhas.append("🏆 Resultado: " + str(c["resultado"]))
    linhas.append("🎯 Precisão: " + str(round(d["precisao_brancas"], 1)) + "% (Nota " + d["nota_brancas"] + ")")
    linhas.append("📈 E4 Rating: " + str(d["rating_e4_brancas"]))
    linhas.append("📖 Abertura: " + str(d["abertura"]))
    if lendarios > 0:
        linhas.append("🌟 Lances Lendários: " + str(lendarios))
    if mestres > 0:
        linhas.append("💎 Lances de Mestre: " + str(mestres))
    if brilhantes > 0:
        linhas.append("🏅 Lances Brilhantes: " + str(brilhantes))
    linhas.append("")
    linhas.append("Análise exclusiva com E4 Rating — só no E4Chess!")
    linhas.append("👉 e4chess.streamlit.app")
    linhas.append("#xadrez #chess #e4chess #analise #xadrezbrasil #xeque")
    return "\n".join(linhas)


def gerar_word(d):
    doc = Document()
    t = doc.add_heading("E4Chess - Relatorio de Analise", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    c = d["cabecalho"]
    p = doc.add_paragraph("Gerado em " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True

    doc.add_heading("1. Metadados", 1)
    metas = [
        ("Brancas", c["brancas"] + " (" + str(c["rating_brancas"]) + ") - " + str(round(d["precisao_brancas"], 1)) + "% (Nota " + d["nota_brancas"] + ") - E4 Rating " + str(d["rating_e4_brancas"])),
        ("Negras", c["negras"] + " (" + str(c["rating_negras"]) + ") - " + str(round(d["precisao_negras"], 1)) + "% (Nota " + d["nota_negras"] + ") - E4 Rating " + str(d["rating_e4_negras"])),
        ("Resultado", str(c["resultado"])),
        ("Abertura", str(d["abertura"])),
        ("Plataforma", str(c["plataforma"])),
        ("Data", str(c["data"])),
    ]
    for label, valor in metas:
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
    cats = ["🌟 Lendário", "💎 Lance de Mestre", "🏅 Brilhante", "📖 Livro",
            "⭐ Melhor", "✨ Excelente", "✅ Bom", "🤏 Imprecisão",
            "⚠️ Erro", "🚨 Erro Grave", "💀 Gafe"]
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
    "♟️ Tabuleiro",
    "🏆 Ranking",
    "📰 Notícias",
    "📊 Como funciona o E4 Rating",
    "🧠 Pensadores",
    "🕯️ Nossa História",
])

aba_analise = abas[0]
aba_tabuleiro = abas[1]
aba_ranking = abas[2]
aba_noticias = abas[3]
aba_rating = abas[4]
aba_pensadores = abas[5]
aba_historia = abas[6]


with aba_analise:
    pgn_text = st.text_area("Cole o PGN da partida:", height=200,
                            placeholder="Cole aqui o PGN completo...")
    prof = st.select_slider(
        "Profundidade",
        options=[6, 8, 10, 12, 15],
        value=8,
        help="6=instantaneo | 8=rapido | 10=padrao | 12=profundo | 15=maximo"
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

                feitos = d.get("feitos_lendarios", [])
                if feitos:
                    st.markdown("### 🌟 Você jogou como os Grandes Mestres!")
                    vistos = set()
                    for f in feitos:
                        if not f.get("feito_lendario"):
                            continue
                        info_lend = f["feito_lendario"]
                        chave = (info_lend["mestre"], info_lend["ano"])
                        if chave in vistos:
                            continue
                        vistos.add(chave)
                        st.markdown(card_mestre_lendario(info_lend), unsafe_allow_html=True)

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
                cats_tab = ["🌟 Lendário", "💎 Lance de Mestre", "🏅 Brilhante",
                            "📖 Livro", "⭐ Melhor", "✨ Excelente", "✅ Bom",
                            "🤏 Imprecisão", "⚠️ Erro", "🚨 Erro Grave", "💀 Gafe"]
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
                st.text_area("Legenda (copie para Instagram):", value=legenda, height=200)

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


with aba_tabuleiro:
    init_tab_state()
    st.markdown("## ♟️ Tabuleiro Interativo")
    st.markdown("*4 modos: Análise · Presencial · Stockfish · Correspondência*")
    st.markdown("---")

    with st.expander("⚙️ Personalizar Tabuleiro", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            idx_tema = list(TEMAS_TABULEIRO.keys()).index(st.session_state.tb_theme)
            st.session_state.tb_theme = st.selectbox("🎨 Tema", list(TEMAS_TABULEIRO.keys()), index=idx_tema)
        with col2:
            st.session_state.tb_size = st.select_slider("📏 Tamanho", options=[300, 340, 380, 420, 460], value=st.session_state.tb_size)
        with col3:
            st.session_state.tb_coords = st.checkbox("🔢 Coordenadas", value=st.session_state.tb_coords)
        col4, col5 = st.columns(2)
        with col4:
            st.session_state.tb_show_legal = st.checkbox("🟢 Mostrar lances legais", value=st.session_state.tb_show_legal)
        with col5:
            st.session_state.tb_auto_rotate = st.checkbox("🔄 Rotação automática", value=st.session_state.tb_auto_rotate)

    modos = ["📖 Análise", "👥 Presencial", "🤖 Stockfish", "📨 Correspondência"]
    idx_modo = modos.index(st.session_state.tb_modo) if st.session_state.tb_modo in modos else 0
    st.session_state.tb_modo = st.radio("Escolha o modo:", modos, index=idx_modo, horizontal=True, key="tb_modo_radio")
    st.markdown("---")
    modo = st.session_state.tb_modo

    if modo == "📖 Análise":
        col_pgn, col_btn = st.columns([3, 1])
        with col_pgn:
            pgn_tab = st.text_area("Cole o PGN:", height=100, key="tab_board_pgn",
                                   placeholder="Cole aqui o PGN para carregar no tabuleiro...")
        with col_btn:
            st.markdown("&nbsp;")
            if st.button("📥 Carregar", use_container_width=True, type="primary"):
                try:
                    g = validar_pgn(pgn_tab)
                    st.session_state.tb_game_obj = g
                    st.session_state.tb_initial_fen = g.board().fen()
                    st.session_state.tb_moves = list(g.mainline_moves())
                    st.session_state.tb_index = 0
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

        if st.session_state.tb_moves:
            prof_tab = st.select_slider("Profundidade das sugestões", options=[8, 10, 12, 15], value=12, key="tb_prof_slider")
            board = get_current_board()
            fen = board.fen()
            with st.spinner("Calculando as 3 melhores jogadas..."):
                setas = analisar_top3(fen, prof_tab)
            last_mv = None
            if st.session_state.tb_index > 0:
                last_mv = st.session_state.tb_moves[st.session_state.tb_index - 1]
            html_board = render_board_html(
                fen, setas=setas, size=st.session_state.tb_size,
                tema=st.session_state.tb_theme,
                show_coords=st.session_state.tb_coords, last_move=last_mv,
            )
            components.html(html_board, height=440)

            total = len(st.session_state.tb_moves)
            idx = st.session_state.tb_index

            st.markdown(
                "<style>"
                ".nav-lich { display:flex; gap:6px; justify-content:center; "
                "margin-top:8px; margin-bottom:6px; flex-wrap:nowrap; }"
                ".nav-lich .stButton { flex:1; }"
                ".nav-lich .stButton button { "
                "padding:8px 4px !important; font-size:15px !important; "
                "min-height:0 !important; height:44px !important; "
                "background:#21262d !important; color:#c9d1d9 !important; "
                "border:1px solid #30363d !important; border-radius:8px !important; }"
                ".nav-lich .stButton button:hover { "
                "background:#30363d !important; color:#58a6ff !important; }"
                "</style>",
                unsafe_allow_html=True
            )

            c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
            with c1:
                if st.button("⏮", use_container_width=True, help="Início", key="nav_ini"):
                    st.session_state.tb_index = 0
                    st.rerun()
            with c2:
                if st.button("◀", use_container_width=True, help="Anterior", key="nav_ant"):
                    st.session_state.tb_index = max(0, idx - 1)
                    st.rerun()
            with c3:
                if st.button("▶", use_container_width=True, help="Próximo", key="nav_prox"):
                    st.session_state.tb_index = min(total, idx + 1)
                    st.rerun()
            with c4:
                if st.button("⏭", use_container_width=True, help="Fim", key="nav_fim"):
                    st.session_state.tb_index = total
                    st.rerun()

            turno = "Brancas" if board.turn == chess.WHITE else "Negras"
            st.markdown(
                f"<div style='text-align:center;color:#8b949e;font-size:13px;"
                f"margin-top:4px;'>Lance <b style='color:#58a6ff;'>{idx}</b> / "
                f"{total} — Vez das <b style='color:#c9d1d9;'>{turno}</b></div>",
                unsafe_allow_html=True
            )

            if setas:
                nomes = ["🥇 Melhor", "🥈 2ª melhor", "🥉 3ª melhor"]
                st.markdown("### 🤖 Sugestões do Stockfish")
                for i, (frm, to, cor) in enumerate(setas):
                    st.markdown(
                        f"<div style='display:flex;align-items:center;gap:10px;"
                        f"padding:6px 10px;margin:4px 0;background:#161b22;"
                        f"border-left:4px solid {cor};border-radius:6px;'>"
                        f"<span style='color:#c9d1d9;font-weight:600;'>{nomes[i]}:</span>"
                        f"<span style='color:#58a6ff;font-weight:700;'>{frm} → {to}</span>"
                        f"</div>", unsafe_allow_html=True,
                    )

            st.markdown("### 🎮 Jogar lance livre")
            legal = sorted([board.san(m) for m in board.legal_moves])
            if legal:
                col_mv, col_btn_mv = st.columns([3, 1])
                with col_mv:
                    chosen = st.selectbox("Escolha um lance:", ["—"] + legal, key="tb_move_analise")
                with col_btn_mv:
                    st.markdown("&nbsp;")
                    if st.button("Jogar ▶", use_container_width=True) and chosen != "—":
                        try:
                            mv = board.parse_san(chosen)
                            make_move(mv)
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))
            else:
                st.info("Fim de jogo — sem lances legais.")

            if st.button("🔄 Reiniciar"):
                reset_tab_game()
                st.rerun()

    elif modo == "👥 Presencial":
        st.markdown("### 👥 Jogar Presencialmente")
        st.markdown("*Jogue com um amigo no mesmo celular.*")
        col_a, col_b = st.columns(2)
        with col_a:
            nome_b = st.text_input("Nome Brancas", value=st.session_state.tb_players["brancas"], key="pres_b")
        with col_b:
            nome_n = st.text_input("Nome Negras", value=st.session_state.tb_players["negras"], key="pres_n")
        st.session_state.tb_players = {"brancas": nome_b, "negras": nome_n}

        if st.button("🚀 Começar Nova Partida", type="primary", use_container_width=True):
            reset_tab_game()
            st.rerun()

        board = get_current_board()
        fen = board.fen()
        turno = board.turn
        rotated = (turno == chess.BLACK) if st.session_state.tb_auto_rotate else False
        last_mv = None
        if st.session_state.tb_index > 0:
            last_mv = st.session_state.tb_moves[st.session_state.tb_index - 1]
        legal = list(board.legal_moves)
        html_board = render_board_html(
            fen, size=st.session_state.tb_size,
            tema=st.session_state.tb_theme,
            show_coords=st.session_state.tb_coords, last_move=last_mv,
            legal_moves=legal if st.session_state.tb_show_legal else None,
            rotated=rotated,
        )
        components.html(html_board, height=440)

        vez_nome = st.session_state.tb_players["brancas"] if turno == chess.WHITE else st.session_state.tb_players["negras"]
        st.markdown(f"### ♟️ Vez de: **{vez_nome}**")

        st.markdown("**Escolha o lance:**")
        legal_san = sorted([board.san(m) for m in legal])
        if legal_san:
            col_mv, col_btn_mv, col_undo = st.columns([3, 1, 1])
            with col_mv:
                chosen = st.selectbox("Lance:", ["—"] + legal_san, key="pres_move")
            with col_btn_mv:
                st.markdown("&nbsp;")
                if st.button("▶ Jogar", use_container_width=True, type="primary") and chosen != "—":
                    try:
                        mv = board.parse_san(chosen)
                        make_move(mv)
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))
            with col_undo:
                st.markdown("&nbsp;")
                if st.button("↩️ Desfazer", use_container_width=True):
                    if st.session_state.tb_index > 0:
                        st.session_state.tb_index -= 1
                        st.session_state.tb_moves = st.session_state.tb_moves[:st.session_state.tb_index]
                        st.rerun()
        else:
            if board.is_checkmate():
                st.success("👑 Xeque-Mate!")
            elif board.is_stalemate():
                st.info("🤝 Empate por Rei Afogado.")

        if st.button("🔄 Reiniciar Partida"):
            reset_tab_game()
            st.rerun()

    elif modo == "🤖 Stockfish":
        st.markdown("### 🤖 Jogar contra Stockfish")
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            niveis = {"🟢 Fácil": 3, "🟡 Médio": 8, "🔴 Difícil": 12, "👑 Mestre": 18}
            nivel = st.selectbox("Nível:", list(niveis.keys()), index=1, key="sf_nivel")
            prof_sf = niveis[nivel]
        with col_n2:
            cor_ia = st.selectbox("IA joga com:", ["Negras", "Brancas"], key="sf_cor")

        if st.button("🚀 Começar Nova Partida", type="primary", use_container_width=True):
            reset_tab_game()
            st.session_state.tb_stockfish_level = prof_sf
            st.session_state.tb_stockfish_color = "negras" if cor_ia == "Negras" else "brancas"
            st.rerun()

        board = get_current_board()
        fen = board.fen()
        turno = board.turn
        ia_cor = chess.BLACK if st.session_state.tb_stockfish_color == "negras" else chess.WHITE
        rotated = (turno == chess.BLACK)
        last_mv = None
        if st.session_state.tb_index > 0:
            last_mv = st.session_state.tb_moves[st.session_state.tb_index - 1]
        legal = list(board.legal_moves)
        html_board = render_board_html(
            fen, size=st.session_state.tb_size,
            tema=st.session_state.tb_theme,
            show_coords=st.session_state.tb_coords, last_move=last_mv,
            legal_moves=legal if st.session_state.tb_show_legal else None,
            rotated=rotated,
        )
        components.html(html_board, height=440)

        if turno == ia_cor and not board.is_game_over():
            with st.spinner("🤖 Stockfish pensando..."):
                mv_ia = melhor_lance_stockfish(fen, st.session_state.tb_stockfish_level)
                if mv_ia:
                    make_move(mv_ia)
                    st.rerun()

        vez_txt = "Você" if turno != ia_cor else "Stockfish"
        st.markdown(f"### ♟️ Vez: **{vez_txt}**")

        if turno != ia_cor and not board.is_game_over():
            st.markdown("**Escolha seu lance:**")
            legal_san = sorted([board.san(m) for m in legal])
            if legal_san:
                col_mv, col_btn_mv, col_undo = st.columns([3, 1, 1])
                with col_mv:
                    chosen = st.selectbox("Lance:", ["—"] + legal_san, key="sf_move")
                with col_btn_mv:
                    st.markdown("&nbsp;")
                    if st.button("▶ Jogar", use_container_width=True, type="primary") and chosen != "—":
                        try:
                            mv = board.parse_san(chosen)
                            make_move(mv)
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))
                with col_undo:
                    st.markdown("&nbsp;")
                    if st.button("↩️ Desfazer", use_container_width=True):
                        if st.session_state.tb_index >= 2:
                            st.session_state.tb_index -= 2
                            st.session_state.tb_moves = st.session_state.tb_moves[:st.session_state.tb_index]
                            st.rerun()
        elif board.is_game_over():
            if board.is_checkmate():
                vencedor = "Stockfish" if turno != ia_cor else "Você"
                st.success(f"👑 Xeque-Mate! Vencedor: {vencedor}")
            else:
                st.info("🤝 Empate.")

        if st.button("🔄 Reiniciar Partida"):
            reset_tab_game()
            st.rerun()

    elif modo == "📨 Correspondência":
        st.markdown("### 📨 Jogar por Correspondência")
        st.info("💡 Você joga → copia o PGN → envia para seu amigo → ele cola no E4Chess dele → devolve → você cola aqui.")

        pgn_recebido = st.text_area("Cole o PGN que seu amigo enviou:",
                                     value=st.session_state.get("tb_correspond_pgn", ""),
                                     height=120, key="correspond_pgn_input")
        col_car, col_novo = st.columns(2)
        with col_car:
            if st.button("📥 Carregar PGN", use_container_width=True, type="primary"):
                try:
                    g = validar_pgn(pgn_recebido)
                    st.session_state.tb_initial_fen = chess.STARTING_FEN
                    st.session_state.tb_moves = list(g.mainline_moves())
                    st.session_state.tb_index = len(st.session_state.tb_moves)
                    st.session_state.tb_correspond_pgn = pgn_recebido
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
        with col_novo:
            if st.button("🆕 Nova Partida", use_container_width=True):
                reset_tab_game()
                st.session_state.tb_correspond_pgn = ""
                st.rerun()

        board = get_current_board()
        fen = board.fen()
        turno = "Brancas" if board.turn == chess.WHITE else "Negras"
        last_mv = None
        if st.session_state.tb_index > 0:
            last_mv = st.session_state.tb_moves[st.session_state.tb_index - 1]
        legal = list(board.legal_moves)
        html_board = render_board_html(
            fen, size=st.session_state.tb_size,
            tema=st.session_state.tb_theme,
            show_coords=st.session_state.tb_coords, last_move=last_mv,
            legal_moves=legal if st.session_state.tb_show_legal else None,
        )
        components.html(html_board, height=440)

        st.markdown(f"### ♟️ Vez das **{turno}**")
        if not board.is_game_over():
            legal_san = sorted([board.san(m) for m in legal])
            if legal_san:
                chosen = st.selectbox("Lance:", ["—"] + legal_san, key="corr_move")
                if st.button("▶ Jogar", use_container_width=True, type="primary") and chosen != "—":
                    try:
                        mv = board.parse_san(chosen)
                        make_move(mv)
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))
        else:
            st.success("Partida encerrada!")

        st.markdown("#### 📤 Enviar para seu amigo")
        pgn_gerado = pgn_atual()
        st.code(pgn_gerado, language="text")
        st.markdown(
            f"<a href='https://wa.me/?text={requests.utils.quote(pgn_gerado)}' "
            f"target='_blank' style='display:inline-block;padding:12px 24px;"
            f"background:#25D366;color:#fff;border-radius:10px;"
            f"text-decoration:none;font-weight:700;'>"
            f"💬 Enviar PGN pelo WhatsApp</a>",
            unsafe_allow_html=True,
        )


with aba_ranking:
    st.markdown("## 🏆 Ranking Mundial FIDE")
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
                medalha = "🥇" if pos == 1 else "🥈" if pos == 2 else "🥉" if pos == 3 else "#" + str(pos)
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
                    "</div>", unsafe_allow_html=True
                )
            except Exception:
                continue
    else:
        st.info("Ranking temporariamente indisponível.")


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
        st.info("Não foi possível carregar agora.")
    st.markdown("---")
    st.markdown("### 🇧🇷 Nacionais")
    st.info("Em breve: notícias da Confederação Brasileira de Xadrez.")
    st.markdown("---")
    st.markdown("### 🗓️ Próximos Torneios")
    torneios = [
        ("Global Chess League 2026", "5 a 13 de setembro · Bengaluru, Índia"),
        ("World Chess Championship 2026", "22 nov a 13 dez · Genebra, Suíça"),
        ("Norway Chess 2026", "Maio/Junho · Oslo, Noruega"),
    ]
    for nome, data in torneios:
        st.markdown("**" + nome + "** — " + data)


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
- **Premia o esforço:** jogar bem contra adversário forte eleva a nota.
- **Útil para todos os níveis.**

*O E4 Rating não substitui o rating oficial da FIDE. É uma métrica complementar e transparente.*
    """)


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
            "</div>", unsafe_allow_html=True
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
