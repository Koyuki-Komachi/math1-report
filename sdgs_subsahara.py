
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.font_manager as fm

def process_sdg_data():
    # ファイルパス
    input_file = 'Sub_Saharan_African_Countries_SDG.xlsx'
    output_file = 'sdgs_subsahara_graph.xlsx'
    
    try:
        # データの読み込み
        df = pd.read_excel(input_file)
        
        # カラム名を小文字に変換し、余分なスペースを削除
        df.columns = [col.strip().lower() for col in df.columns]
        
        # 必要なカラムの確認
        required_columns = ['year'] + [f'goal{i}' for i in range(1, 18)]
        for col in required_columns:
            if col not in df.columns:
                print(f"エラー: '{col}' カラムが存在しません。")
                return
        
        # 年次ごとの平均を格納するリスト
        years = list(range(2000, 2024))
        goals = [f'goal{i}' for i in range(1, 18)]
        averages_dict = { 'year': [] }
        for goal in goals:
            averages_dict[goal] = []
        
        for year in years:
            yearly_data = df[df['year'] == year]
            if yearly_data.empty:
                print(f"警告: {year} 年のデータが存在しません。")
                continue
            averages_dict['year'].append(year)
            for goal in goals:
                # 空セルや0を除外
                valid_data = yearly_data[goal].replace(0, pd.NA).dropna()
                average = valid_data.mean()
                averages_dict[goal].append(average)
        
        # 平均値のデータフレーム作成
        averages_df = pd.DataFrame(averages_dict)
        
        # 表をExcelに保存
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            averages_df.to_excel(writer, sheet_name='平均値', index=False)
        
        # フォントを設定
        font_prop = fm.FontProperties(fname='/mnt/data/file-ngwyeoEN29l1M3O1QpdxCwkj')
        
        # グラフの作成
        plt.figure(figsize=(14, 8))
        colors = plt.cm.viridis(range(len(goals)))  # Use a colormap to generate distinct colors
        for i, goal in enumerate(goals):
            plt.plot(averages_df['year'], averages_df[goal], marker='o', label=goal.capitalize(), color=colors[i])
        
        plt.title('Sub-Saharan Africaの各ゴールの年次平均値 (2000-2023)', fontsize=16, fontproperties=font_prop)
        plt.xlabel('年', fontsize=14, fontproperties=font_prop)
        plt.ylabel('平均スコア', fontsize=14, fontproperties=font_prop)
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # グラフを画像として保存
        graph_image = 'sdg_averages_graph.png'
        plt.savefig(graph_image, dpi=300)
        plt.close()
        
        # Excelファイルにグラフを挿入
        wb = load_workbook(output_file)
        ws = wb.create_sheet(title='グラフ')
        
        img = OpenpyxlImage(graph_image)
        img.anchor = 'A1'
        ws.add_image(img)
        
        wb.save(output_file)
        print(f"処理が完了しました。結果は '{output_file}' に保存されました。")
    
    except FileNotFoundError:
        print(f"エラー: ファイル '{input_file}' が見つかりません。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    process_sdg_data()
