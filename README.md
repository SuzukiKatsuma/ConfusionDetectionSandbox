# 困惑検知システム

## 機能

### 1. 操作ログ記録
* 対象アプリケーションを選択して操作ログを記録する
  * 操作ログは押下 (down) イベントのみを記録し、解放 (up) イベントは記録していない
* ログは `C:\Users\<user-name>\Documents\ConfusionDetectionLogs` に CSV 形式で保存される

#### ログの詳細
```csv
timestamp, target_pid, target_app_name, operation, x, y, delta, virtual_key
2026-01-17T16:29:28.797,35280,P5R,mouse_l,511,881,,
2026-01-17T16:29:32.649,35280,P5R,keyboard,,,,13
2026-01-17T16:29:36.515,35280,P5R,keyboard,,,,32
```

1. **timestamp**  
`YYYY-MM-DDTHH:mm:ss.fff` 形式 (UTC)

2. **target_pid**  
アプリケーションのプロセス ID (一意にアプリを特定)

3. **target_app_name**  
`MainWindowTitle (ProcessName)` 形式の対象アプリケーションの名前

4. **operation**  
操作種別

| 操作 ID | 対応操作 |
| :-: | :-- |
| keyboard | キーボード |
| mouse_l | マウスの左クリック |
| mouse_r | マウスの右クリック |
| mouse_m | マウスの中央ボタン押下 |
| wheel_v | マウスホイールの上下回転 |
| wheel_h | マウスホイールの左右回転 |

5. **x**  
マウス操作時のクライアント X（取れないときは空）

6. **y**  
マウス操作時のクライアント Y（取れないときは空）

7. **delta**  
ホイール量（クリックやキー操作の場合は空）

8. **virtual_key**  
論理的にどのキーが入力されたかを特定する。  
Windows Virtual-Key Code（例：A=65, Enter=13）。マウス操作時は空。
