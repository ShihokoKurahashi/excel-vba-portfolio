import json
import datetime

def save_task():
    """
    VBAでの「ファイル入出力」や「データ整形」の知識をPythonへ応用する学習サンプルです。
    """
    print("--- Task Logger ---")
    task_name = input("記録する内容を入力してください: ")
    
    # 現在時刻の取得（VBAのNow関数に相当）
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # データの辞書作成（VBAの構造体に相当）
    log_data = {
        "timestamp": now_str,
        "content": task_name
    }
    
    # ファイルへの追記保存
    try:
        with open("tasks.json", "a", encoding="utf-8") as f:
            json.dump(log_data, f, ensure_ascii=False)
            f.write("\n")
        print(f"『{task_name}』を保存しました。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    save_task()
