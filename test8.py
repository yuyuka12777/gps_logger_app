import serial
import pynmea2
import keyboard
import pyperclip
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import serial.tools.list_ports

class GPSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GPS座標取得ツール")
        self.root.geometry("600x500")
        
        self.ser = None
        self.running = False
        self.coordinates_history = []
        self.selected_key = 'F9'
        self.key_listener = None  # キーリスナーを追跡
        
        self.setup_ui()
        self.scan_ports()
        
    def setup_ui(self):
        # COMポート選択フレーム
        port_frame = ttk.LabelFrame(self.root, text="COMポート設定", padding=10)
        port_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(port_frame, text="COMポート:").grid(row=0, column=0, sticky="w")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(port_frame, textvariable=self.port_var, width=20)
        self.port_combo.grid(row=0, column=1, padx=5)
        
        ttk.Button(port_frame, text="再スキャン", command=self.scan_ports).grid(row=0, column=2, padx=5)
        ttk.Button(port_frame, text="接続", command=self.connect_serial).grid(row=0, column=3, padx=5)
        ttk.Button(port_frame, text="切断", command=self.disconnect_serial).grid(row=0, column=4, padx=5)
        
        # キー選択フレーム
        key_frame = ttk.LabelFrame(self.root, text="トリガーキー設定", padding=10)
        key_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(key_frame, text="座標取得キー:").grid(row=0, column=0, sticky="w")
        self.key_var = tk.StringVar(value=self.selected_key)
        key_combo = ttk.Combobox(key_frame, textvariable=self.key_var, values=['F9', 'F10', 'F11', 'F12', 'Space', 'Enter'], width=15)
        key_combo.grid(row=0, column=1, padx=5)
        key_combo.bind('<<ComboboxSelected>>', self.on_key_changed)
        
        # 制御ボタンフレーム
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        self.start_button = ttk.Button(control_frame, text="監視開始", command=self.start_monitoring)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="監視停止", command=self.stop_monitoring, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        ttk.Button(control_frame, text="履歴クリア", command=self.clear_history).pack(side="left", padx=5)
        
        # ステータス表示
        self.status_var = tk.StringVar(value="待機中")
        status_label = ttk.Label(self.root, textvariable=self.status_var, font=("Arial", 10, "bold"))
        status_label.pack(pady=5)
        
        # 座標履歴表示
        history_frame = ttk.LabelFrame(self.root, text="座標履歴", padding=10)
        history_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.history_text = scrolledtext.ScrolledText(history_frame, height=15, width=70)
        self.history_text.pack(fill="both", expand=True)
        
    def scan_ports(self):
        """利用可能なCOMポートをスキャン"""
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            self.port_var.set(ports[0])
            
    def connect_serial(self):
        """シリアルポートに接続"""
        try:
            port = self.port_var.get()
            if not port:
                messagebox.showerror("エラー", "COMポートを選択してください")
                return
                
            self.ser = serial.Serial(port, baudrate=4800, timeout=1)
            self.status_var.set(f"接続中: {port}")
            messagebox.showinfo("成功", f"{port}に接続しました")
        except Exception as e:
            messagebox.showerror("エラー", f"接続に失敗しました: {e}")
            
    def disconnect_serial(self):
        """シリアルポートから切断"""
        if self.ser:
            self.ser.close()
            self.ser = None
            self.status_var.set("切断済み")
            
    def on_key_changed(self, event):
        """選択キーが変更された時の処理"""
        self.selected_key = self.key_var.get()
        
    def get_latest_gps(self, timeout=1.0):
        """最新のGPS座標を取得"""
        if not self.ser:
            return None
            
        start_time = time.time()
        latest_coords = None
        while time.time() - start_time < timeout:
            try:
                line = self.ser.readline().decode('ascii', errors='replace')
                if line.startswith('$GNGGA'):
                    msg = pynmea2.parse(line)
                    if msg.gps_qual > 0:  # 有効なFixか確認
                        lat = msg.latitude
                        lon = msg.longitude
                        latest_coords = f"{lat},{lon}"
            except pynmea2.ParseError:
                continue
            except Exception:
                break
        return latest_coords
        
    def on_key_event(self, event):
        """キーイベントハンドラー（イベントベース）"""
        key_mapping = {
            'F9': 'f9',
            'F10': 'f10', 
            'F11': 'f11',
            'F12': 'f12',
            'Space': 'space',
            'Enter': 'enter'
        }
        
        target_key = key_mapping.get(self.selected_key, 'f9')
        
        # キーが押された瞬間のみ処理（key downイベント）
        if event.event_type == keyboard.KEY_DOWN and event.name == target_key and self.running:
            coords = self.get_latest_gps()
            if coords:
                # 座標をクリップボードにコピー
                pyperclip.copy(coords)
                
                # 履歴に追加
                timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                history_entry = f"[{timestamp}] {coords}"
                self.coordinates_history.append(history_entry)
                
                # GUIを更新（メインスレッドで実行）
                self.root.after(0, self.update_history_display, history_entry)
                self.root.after(0, lambda: self.status_var.set(f"座標コピー完了: {coords}"))
                
                # キーボード操作
                keyboard.send('ctrl+v')
                keyboard.send('enter')
            else:
                self.root.after(0, lambda: self.status_var.set("有効なGPSデータが取得できませんでした"))

    def monitor_key_press(self):
        """キー入力を監視（ポーリング方式をバックアップとして保持）"""
        key_mapping = {
            'F9': 'f9',
            'F10': 'f10', 
            'F11': 'f11',
            'F12': 'f12',
            'Space': 'space',
            'Enter': 'enter'
        }
        
        key_code = key_mapping.get(self.selected_key, 'f9')
        key_pressed = False  # キーの状態を追跡
        
        while self.running:
            try:
                current_key_state = keyboard.is_pressed(key_code)
                
                # キーが押された瞬間を検出（前回は押されていなくて、今回は押されている）
                if current_key_state and not key_pressed:
                    coords = self.get_latest_gps()
                    if coords:
                        # 座標をクリップボードにコピー
                        pyperclip.copy(coords)
                        
                        # 履歴に追加
                        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                        history_entry = f"[{timestamp}] {coords}"
                        self.coordinates_history.append(history_entry)
                        
                        # GUIを更新（メインスレッドで実行）
                        self.root.after(0, self.update_history_display, history_entry)
                        self.root.after(0, lambda: self.status_var.set(f"座標コピー完了: {coords}"))
                        
                        # キーボード操作
                        keyboard.send('ctrl+v')
                        keyboard.send('enter')
                    else:
                        self.root.after(0, lambda: self.status_var.set("有効なGPSデータが取得できませんでした"))
                
                # キーの状態を更新
                key_pressed = current_key_state
                        
                time.sleep(0.05)  # チェック間隔を短くして反応性を向上
            except Exception as e:
                self.root.after(0, lambda: self.status_var.set(f"エラー: {e}"))
                break
                
    def update_history_display(self, entry):
        """履歴表示を更新"""
        self.history_text.insert(tk.END, entry + "\n")
        self.history_text.see(tk.END)
        
    def start_monitoring(self):
        """監視を開始"""
        if not self.ser:
            messagebox.showerror("エラー", "COMポートに接続してください")
            return
            
        self.running = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_var.set(f"監視中 (キー: {self.selected_key})")
        
        # イベントベースのキーリスナーを開始（プログラマブルキーボード対応）
        try:
            self.key_listener = keyboard.hook(self.on_key_event)
        except Exception as e:
            print(f"イベントベースリスナー開始失敗: {e}")
        
        # 別スレッドでポーリング方式のキー監視も開始（バックアップ）
        self.monitor_thread = threading.Thread(target=self.monitor_key_press, daemon=True)
        self.monitor_thread.start()
        
    def stop_monitoring(self):
        """監視を停止"""
        self.running = False
        
        # イベントベースリスナーを停止
        if self.key_listener:
            try:
                keyboard.unhook(self.key_listener)
                self.key_listener = None
            except Exception as e:
                print(f"キーリスナー停止エラー: {e}")
        
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.status_var.set("監視停止")
        
    def clear_history(self):
        """履歴をクリア"""
        self.coordinates_history.clear()
        self.history_text.delete(1.0, tk.END)
        self.status_var.set("履歴をクリアしました")
        
    def on_closing(self):
        """アプリケーション終了時の処理"""
        self.stop_monitoring()
        
        # イベントリスナーを確実に停止
        if self.key_listener:
            try:
                keyboard.unhook(self.key_listener)
            except:
                pass
                
        if self.ser:
            self.ser.close()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = GPSApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
