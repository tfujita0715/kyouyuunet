import pandas as pd
import spacy
from collections import Counter, defaultdict
from itertools import combinations
import networkx as nx
import matplotlib.pyplot as plt
import japanize_matplotlib
import os

# --- 設定項目 ---
# CSVファイルが保存されているディレクトリ
CSV_DIRECTORY = '.'
# 分析対象のテキストが格納されている列番号 (ヘッダーがないため0番目の列を指定)
COLUMN_NAME = 0
# 共起ネットワーク図の出力先ディレクトリ
OUTPUT_DIR = 'output_networks'
# 抽出する品詞 (NOUN:名詞, PROPN:固有名詞, VERB:動詞, ADJ:形容詞)
TARGET_POS = ['NOUN', 'PROPN', 'VERB', 'ADJ']
# 共起の最小頻度 (この回数以上共起する単語ペアのみ描画)
MIN_COOCCURRENCE = 3
# ----------------

def analyze_and_create_network(filepath, column_name):
    """
    CSVファイルを読み込み、テキストデータから共起ネットワークを生成・保存する
    """
    print(f"--- 分析開始: {filepath} ---")
    
    # --- 1. データの読み込み (Excel/CSV両対応) ---
    df = None
    # try:
    #     # まずExcelファイルとして読み込みを試す
    #     df = pd.read_excel(filepath, engine='openpyxl')
    #     print(f"[情報] Excelファイルとして読み込みました: {os.path.basename(filepath)}")
    #except Exception:
        # Excelで失敗した場合、CSVとしてエンコーディングを試す
        #print(f"[情報] Excel形式での読み込みに失敗。CSVとして再試行します...")
    for enc in ['utf-8', 'cp932', 'shift_jis']:
        try:
            # ヘッダーがないことを想定し、列番号で読み込む
            df = pd.read_csv(filepath, encoding=enc, header=None)
            print(f"[情報] エンコーディング '{enc}' でCSVファイルを読み込みました。")
            break
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue

    if df is None:
        print(f"[エラー] サポートされている形式(Excel/CSV)でファイル '{os.path.basename(filepath)}' を読み込めませんでした。処理をスキップします。")
        return

    try:
        if column_name not in df.columns:
            print(f"[エラー] 列名 '{column_name}' が見つかりません。処理をスキップします。")
            return
        # 欠損値を除外
        texts = df[column_name].dropna().tolist()
        if not texts:
            print("[情報] 分析対象のテキストが見つかりません。処理をスキップします。")
            return
    except Exception as e:
        print(f"[エラー] データ処理中にエラーが発生しました: {e}")
        return

    # --- 2. 形態素解析 ---
    nlp = spacy.load('ja_ginza')
    words_per_sentence = []
    for text in texts:
        doc = nlp(str(text))
        for sent in doc.sents:
            # TARGET_POSで指定された品詞の単語の原型を抽出
            words = [token.lemma_ for token in sent if token.pos_ in TARGET_POS]
            if len(words) > 1:
                words_per_sentence.append(words)

    # --- 3. 共起ペアのカウント ---
    co_occurrence = Counter()
    for words in words_per_sentence:
        # 1文中の単語の組み合わせをカウント
        for pair in combinations(sorted(set(words)), 2):
            co_occurrence[pair] += 1
    
    if not co_occurrence:
        print("[情報] 共起ペアが見つかりませんでした。")
        return

    # --- 4. ネットワークグラフの作成 (NetworkX) ---
    G = nx.Graph()
    
    # 頻度の高いペアのみをエッジとして追加
    for pair, weight in co_occurrence.items():
        if weight >= MIN_COOCCURRENCE:
            G.add_edge(pair[0], pair[1], weight=weight)

    if G.number_of_nodes() == 0:
        print(f"[情報] 最小共起回数({MIN_COOCCURRENCE}回)以上のペアがなかったため、グラフは生成されませんでした。")
        return

    # --- 5. グラフの描画・保存 (Matplotlib) ---
    plt.figure(figsize=(12, 12))
    
    # ノードの大きさを接続しているエッジの重み(頻度)の合計に比例させる
    node_size = [d * 100 for n, d in G.degree(weight='weight')]
    
    # エッジの太さを重み(頻度)に比例させる
    edge_width = [d['weight'] * 0.5 for u, v, d in G.edges(data=True)]

    pos = nx.spring_layout(G, k=0.8, seed=42) # kでノード間の反発力を調整

    nx.draw_networkx_nodes(G, pos, node_size=node_size, node_color='skyblue', alpha=0.8)
    nx.draw_networkx_edges(G, pos, width=edge_width, edge_color='gray', alpha=0.6)
    nx.draw_networkx_labels(G, pos, font_size=10, font_family='Yu Gothic')
    
    plt.title(f'共起ネットワーク: {os.path.basename(filepath)}', fontsize=16)
    plt.axis('off')
    
    # ファイルに保存
    output_filename = os.path.join(OUTPUT_DIR, f"network_{os.path.splitext(os.path.basename(filepath))[0]}.png")
    plt.savefig(output_filename, bbox_inches='tight')
    plt.close()
    
    print(f"✓ グラフを保存しました: {output_filename}")


if __name__ == '__main__':
    # 出力用ディレクトリがなければ作成
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    # カレントディレクトリにあるCSVファイルを取得
    csv_files = [f for f in os.listdir(CSV_DIRECTORY) if f.endswith('.csv')]
    
    if not csv_files:
        print(f"[エラー] ディレクトリ '{CSV_DIRECTORY}' にCSVファイルが見つかりません。")
    else:
        for filename in csv_files:
            filepath = os.path.join(CSV_DIRECTORY, filename)
            analyze_and_create_network(filepath, COLUMN_NAME)
            print("- すべての処理が完了しました ---")
#このコードはAIのアシストの元作成されました
