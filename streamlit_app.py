import streamlit as st
import io, math, shutil, subprocess, os, tempfile
from datetime import datetime
import chess, chess.pgn, chess.engine
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="E4Chess", page_icon="♟️", layout="wide")

if not shutil.which("stockfish"):
    try:
        subprocess.run(["apt-get","update","-qq"], check=False)
        subprocess.run(["apt-get","install","-y","-qq","stockfish"], check=False)
    except Exception:
        pass
STOCKFISH_PATH = shutil.which("stockfish") or "/usr/games/stockfish"

VALORES_PECAS = {chess.PAWN:1, chess.KNIGHT:3, chess.BISHOP:3, chess.ROOK:5, chess.QUEEN:9, chess.KING:0}

def detectar_plataforma(h):
    s = (h.get("Site","") or "").lower(); e = (h.get("Event","") or "").lower()
    if "chess.com" in s or "chess.com" in e: return "Chess.com"
    if "lichess" in s or "lichess" in e: return "Lichess"
    return "Desconhecida"

def validar_pgn(texto):
    if not texto or not texto.strip(): raise ValueError("Cole um PGN.")
    game = chess.pgn.read_game(io.StringIO(texto))
    if game is None: raise ValueError("PGN invalido.")
    lances = list(game.mainline_moves())
    if not lances: raise ValueError("PGN sem lances.")
    return game

def cp_prob(cp):
    cp = max(-1500, min(1500, cp))
    return 1.0/(1.0+math.pow(10,-cp/400.0))

def prec_lance(cp_a, cp_d, cor):
    if cor == "brancas": a, d = cp_a, cp_d
    else: a, d = -cp_a, -cp_d
    perda = max(0, cp_prob(a) - cp_prob(d))
    return max(0, min(100, 100*(1-perda*2.2)))

def classificar(p, san, melhor):
    if "#" in san: return "Melhor", "Xeque-mate"
    if san == melhor: return "Melhor", "Melhor lance"
    if p >= 95: return "Excelente", "Alto nivel"
    if p >= 85: return "Bom", "Solido"
    if p >= 70: return "Imprecisao", "Poderia ser melhor"
    if p >= 50: return "Erro", "Erro tatico"
    return "Erro grave", "Erro grave"

def analisar_stockfish(game, prof=10):
    try:
        eng = chess.engine.SimpleEngine.popen_uci(STOCKFISH_PATH)
        try: eng.configure({"Threads":2,"Hash":128})
        except: pass
        board = game.board(); res = []
        for move in list(game.mainline_moves()):
            info = eng.analyse(board, chess.engine.Limit(depth=prof))
            sc = info["score"].pov(chess.WHITE)
            cp = sc.score(mate_score=100000) or 0
            pv = info.get("pv") or []
            melhor = board.san(pv[0]) if pv else ""
            san = board.san(move)
            board.push(move)
            res.append({"cp":cp,"melhor_san":melhor,"san":san})
        eng.quit(); return res
    except Exception as e:
        st.warning(f"Stockfish: {e}"); return None

def analisar(game):
    headers = dict(game.headers); lances = list(game.mainline_moves()); board = game.board()
    cab = {"brancas": headers.get("White","Brancas"),"negras": headers.get("Black","Pretas"),
           "rating_brancas": headers.get("WhiteElo","-"),"rating_negras": headers.get("BlackElo","-"),
           "resultado": headers.get("Result","*"),"evento": headers.get("Event","-"),
           "data": headers.get("Date","-"),"plataforma": detectar_plataforma(headers)}
    eng_res = analisar_stockfish(game)
    dados = []; stats = {"brancas":{"classes":{},"precs":[]},"negras":{"classes":{},"precs":[]}}
    curva = [0]; marcadores = []
    for i, move in enumerate(lances):
        cor = "brancas" if i%2==0 else "negras"
        san = board.san(move)
        board.push(move)
        prec = 100.0; melhor = san; ev = 0
        if eng_res and i < len(eng_res):
            info = eng_res[i]; ev = info["cp"]; melhor = info["melhor_san"]
            cp_a = eng_res[i-1]["cp"] if i>0 else 20
            prec = prec_lance(cp_a, ev, cor)
        cls, motivo = classificar(prec, san, melhor)
        cp_vis = max(-1000, min(1000, ev))
        curva.append(cp_vis/100)
        if cls in ("Erro","Erro grave","Imprecisao"):
            marcadores.append({"ply":i+1,"cp":cp_vis/100,"classe":cls})
        dados.append({"n":i+1,"cor":cor,"san":san,"classe":cls,"motivo":motivo,
                      "precisao":prec,"eval":ev,"melhor":melhor})
        stats[cor]["classes"][cls] = stats[cor]["classes"].get(cls,0)+1
        stats[cor]["precs"].append(prec)
    pb = sum(stats["brancas"]["precs"])/max(1,len(stats["brancas"]["precs"]))
    pn = sum(stats["negras"]["precs"])/max(1,len(stats["negras"]["precs"]))
    def nota(p):
        return "A+" if p>=95 else "A" if p>=90 else "B+" if p>=85 else "B" if p>=80 else "C+" if p>=75 else "C" if p>=70 else "D" if p>=60 else "F"
    return {"cabecalho":cab,"lances":dados,"estatisticas":stats,
            "precisao_brancas":pb,"precisao_negras":pn,
            "nota_brancas":nota(pb),"nota_negras":nota(pn),
            "curva":curva,"marcadores":marcadores}

def gerar_grafico(d):
    c = d["curva"]; xs = list(range(len(c)))
    fig, ax = plt.subplots(figsize=(9,4), dpi=100, facecolor="#0d1117")
    ax.set_facecolor("#0d1117")
    ax.fill_between(xs,[max(0,y) for y in c],0,color="#f0f6fc",alpha=0.92)
    ax.fill_between(xs,[min(0,y) for y in c],0,color="#0d1117",alpha=0.95)
    ax.plot(xs,c,color="#58a6ff",linewidth=1.4)
    ax.axhline(0,color="#30363d",linewidth=1)
    for cl,cor in {"Imprecisao":"#d29922","Erro":"#f0883e","Erro grave":"#f85149"}.items():
        pts = [m for m in d["marcadores"] if m["classe"]==cl]
        if pts:
            ax.scatter([p["ply"] for p in pts],[p["cp"] for p in pts],
                       color=cor,s=65,edgecolors="#0d1117",linewidths=1.2,label=cl)
    ax.set_xlim(0,max(1,len(c)-1)); ax.set_ylim(-8,8)
    ax.tick_params(colors="#8b949e",labelsize=8)
    ax.grid(True,color="#21262d",linestyle=":",alpha=0.5)
    for sp in ax.spines.values(): sp.set_color("#30363d")
    leg = ax.legend(loc="upper right",fontsize=7,facecolor="#161b22",edgecolor="#30363d")
    if leg:
        for t in leg.get_texts(): t.set_color("#c9d1d9")
    fig.tight_layout(); return fig

def gerar_word(d):
    doc = Document()
    t = doc.add_heading("E4Chess - Relatorio", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = d["cabecalho"]
    doc.add_paragraph(f"Brancas: {c['brancas']} ({c['rating_brancas']}) - {d['precisao_brancas']:.1f}% Nota {d['nota_brancas']}")
    doc.add_paragraph(f"Negras: {c['negras']} ({c['rating_negras']}) - {d['precisao_negras']:.1f}% Nota {d['nota_negras']}")
    doc.add_paragraph(f"Resultado: {c['resultado']}")
    doc.add_heading("Historico", 1)
    tab = doc.add_table(rows=1, cols=4); tab.style = "Light Grid Accent 1"
    h = tab.rows[0].cells
    h[0].text="#"; h[1].text="Lance"; h[2].text="Classe"; h[3].text="Precisao"
    for l in d["lances"]:
        r = tab.add_row().cells
        r[0].text=str(l["n"]); r[1].text=l["san"]; r[2].text=l["classe"]; r[3].text=f"{l['precisao']:.0f}%"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name); return tmp.name

st.markdown("# ♟️ E4Chess")
st.markdown("**Analise profissional de partidas de xadrez**")

pgn_text = st.text_area("Cole o PGN da partida:", height=200,
                        placeholder="Cole aqui o PGN completo...")

col1, col2 = st.columns(2)
with col1:
    analisar_btn = st.button("ANALISAR", use_container_width=True, type="primary")
with col2:
    limpar_btn = st.button("LIMPAR", use_container_width=True)

if analisar_btn and pgn_text.strip():
    with st.spinner("Analisando com Stockfish... aguarde ~30s"):
        try:
            game = validar_pgn(pgn_text)
            d = analisar(game)
            c = d["cabecalho"]
            st.success("Analise concluida!")
            col1, col2 = st.columns(2)
            with col1:
                st.metric(f"{c['brancas']}", f"{d['precisao_brancas']:.1f}%", f"Nota {d['nota_brancas']}")
            with col2:
                st.metric(f"{c['negras']}", f"{d['precisao_negras']:.1f}%", f"Nota {d['nota_negras']}")
            st.markdown(f"**Resultado:** {c['resultado']}")
            st.markdown("### Evolucao da Partida")
            fig = gerar_grafico(d)
            st.pyplot(fig)
            st.markdown("### Estatisticas")
            cat_data = []
            for cat in ["Melhor","Excelente","Bom","Imprecisao","Erro","Erro grave"]:
                cat_data.append({
                    "Categoria": cat,
                    c["brancas"]: d["estatisticas"]["brancas"]["classes"].get(cat,0),
                    c["negras"]: d["estatisticas"]["negras"]["classes"].get(cat,0)
                })
            st.dataframe(cat_data, use_container_width=True)
            st.markdown("### Historico")
            hist_data = [{"#": l["n"], "Lance": l["san"], "Classe": l["classe"], "Precisao": f"{l['precisao']:.0f}%"} for l in d["lances"][:60]]
            st.dataframe(hist_data, use_container_width=True)
            word_path = gerar_word(d)
            with open(word_path, "rb") as f:
                st.download_button("Baixar Relatorio Word", f,
                                   file_name="e4chess_relatorio.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Erro: {e}")

if limpar_btn:
    st.rerun()
