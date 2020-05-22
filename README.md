# 概要
「アラームをかけてもなかなか起きられない！」そんな学生さんのために「**大学生が絶対に起きられる目覚まし時計**」を作りました。

指定した時刻になるとアラーム音が鳴るだけでなく、あらかじめ人質として指定しておいたWordファイルが開かれ、文字が一文字ずつ🍟の絵文字に変換されていきます。

文書の全てが🍟になる前にESCキーを押せば止まります。そして「保存しないで閉じる」を行えば文書は無傷です。

もし文書の最後まで🍟になってしまったら「時すでに遅し」です。**勝手に上書き保存**されて**Wordも閉じられてしまいます**のでご注意ください。

これでレポートを守るために、きちんと起きることができますね。


# 使い方
- Python3環境が前提です。
- アラームが起動するためにはPCが電源につながっている必要があります。
- 必要なモジュールをインストールします。
    - docx
    - keyboard
    - pyautogui
    - pywin32
    - playsound
- template.xmlのコメントになっている部分を実行環境に合わせて編集してください。
- make_schedule.pyのUSERNAMEを実行環境のユーザーネームに書き換えてください。
- コマンドラインから下記のようにコマンドを実行します 。

```
>python PerfectAlarm.py 07:10
```
- 人質にしたいWordファイルを選びます。
- 実行ユーザーパスワードを求められるので入力します
- PCはスリープ状態であっても動作します。

# 注意
- スクリプトの実行は全て自己責任でお願いします。
- 実行環境によってはうまく動作しない可能性があります。申し訳ございません。