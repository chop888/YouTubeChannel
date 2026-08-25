# キャラクター表情差分リスト

YMM4で動く立ち絵として使用するために必要な表情パーツの一覧。
一般的な立ち絵制作では基本9表情前後をベースに作ることが多いが、
このチャンネル（比較型・おすすめ〇選型・雑学系など）で実際に使う場面を想定し、優先度をつけている。

## 必須（これがないと会話が成立しない）

| 表情 | 用途 | えりか（進行役） | ここな（リアクション役） |
|---|---|---|---|
| 通常 | ベース、説明時 | ◎ | ◎ |
| 笑顔 | 良い情報を伝える時 | ◎ | ◎ |
| 驚き | 意外な情報・リアクション | ○ | ◎（特に重要） |

## 優先度高（比較・〇選・雑学系で頻出）

| 表情 | 用途 |
|---|---|
| 困り顔 | 「うーん」「悩ましい」という場面 |
| 納得顔（頷き） | 「なるほど」という場面。目を細める程度でも表現可 |
| ジト目／半目 | ツッコミ、あきれ気味のリアクション（特にここな向き） |

## あると表現の幅が広がる（任意）

| 表情 | 用途 |
|---|---|
| 怒り（軽め） | 「え〜それはひどい」など軽いツッコミ |
| 照れ／赤面 | 商品を褒めすぎた時など |
| 泣き顔／半泣き | コスパの悪さを嘆く時などコミカルな場面 |
| キラキラ目 | 「これ良い！」と強く推す時 |

## 補助パーツ（表情に組み合わせて使う、漫符）

必要な表情の分だけ用意すれば十分で、無理にすべて揃える必要はない。

- 汗（1滴・複数滴の2段階）：焦り・気まずさの表現
- 「！」マーク：驚きの強調
- 青ざめ線（縦線）：ドン引き・動揺

## パーツ分けの基本（表情の組み合わせ方）

表情は「目」「眉」「口」を別レイヤーで用意し、組み合わせて作ると効率が良い。
例：「困り眉」＋「笑顔の口」＝「困り笑い」のように、パーツの掛け合わせで表情の種類を増やせる。

## 作成の優先順位（おすすめの進め方）

1. まず「通常・笑顔・驚き」の3表情＋口パク・まばたきで1本動画を作ってみる
2. 実際に台本を当てはめてみて、足りないと感じた表情だけ追加していく
3. 全部を最初から揃えようとせず、「困ったら都度追加する」方式で進める

## 制作ツール・手法

- Photopea（無料・ブラウザ・インストール不要）を使用
- 新規に描く必要はなく、既存パーツの「切り取り」「拡大縮小」「回転」「線1本の描き足し」で対応可能
- 詳細な作成手順は制作メモ（別途チャットで相談した手順）を参照

# AI生成用 表情差分リスト・プロンプト集

口パク・まばたき用パーツは別途Photopeaで手動加工が必要（本リストの対象外）。
ここでは「顔全体の表情差分（通常・驚き・笑顔など）」をAI画像生成で作る場合のリストとプロンプトをまとめる。

## 表情リストと優先度

| 表情 | えりか | ここな | 優先度 |
|---|---|---|---|
| 通常 | 完成済み | 完成済み | - |
| 驚き | 必要 | 必要（特に重要） | 高 |
| 笑顔 | 必要 | 必要 | 高 |
| 納得顔（頷き/細目） | 必要 | 必要 | 中 |
| ジト目／半目 | 任意 | 必要 | 中 |
| 困り顔 | 必要 | 任意 | 中 |
| 照れ／赤面 | 任意 | 任意 | 低 |
| 怒り（軽め） | 任意 | 必要 | 低 |
| キラキラ目 | 任意 | 必要 | 低 |
| 泣き顔／半泣き | 任意 | 任意 | 低 |

## 基本テンプレート

### 日本語

```
添付した画像のキャラクターと完全に同じ絵柄・同じ顔立ち・同じ髪型・同じ髪色を維持したまま、
表情だけを変更してください。他の要素（輪郭、髪型、髪色、瞳の色、線の太さ、塗りのスタイル）は
一切変えないでください。

表情の変更内容：{ここに表情ごとの指示を挿入}

背景は白の単色のまま、顔のみのクローズアップ構図も維持してください。
```

### English

```
Using the attached image as reference, keep the exact same character design —
same face shape, same hairstyle, same hair color, same eye color, same line
thickness, and same coloring style. Only change the facial expression.
Do not alter any other feature.

Expression change: {insert expression-specific instruction here}

Keep the plain white background and the same face-only close-up composition.
```

## 表情ごとの差し替えフレーズ

| 表情 | 日本語（挿入用） | English（挿入用） |
|---|---|---|
| 驚き | 目を大きく見開き、眉を上げて驚いた表情にする | Open the eyes wide and raise the eyebrows for a surprised expression |
| 笑顔 | 目を少し細めて、口角を上げた明るい笑顔にする | Slightly narrow the eyes and raise the corners of the mouth into a bright smile |
| 納得顔（頷き） | 目を柔らかく細め、小さくうなずくような落ち着いた表情にする | Softly narrow the eyes into a calm, nodding/understanding expression |
| ジト目／半目 | 目を半分閉じ気味にして、あきれ気味・冷静な半目にする | Half-close the eyes for a deadpan, unimpressed half-lidded look |
| 困り顔 | 眉を八の字にして、少し困ったような表情にする | Angle the eyebrows into a worried slant for a troubled expression |
| 照れ／赤面 | 頬の赤みを強くし、目を少し伏し目がちにする | Increase the blush on the cheeks and slightly lower the gaze for a shy expression |
| 怒り（軽め） | 眉を吊り上げ気味にして、少しむっとした表情にする | Slightly raise and angle the eyebrows for a mildly annoyed expression |
| キラキラ目 | 瞳の中にきらめきを増やし、目を輝かせた期待感のある表情にする | Add extra sparkle highlights in the eyes for an excited, shining-eyed expression |
| 泣き顔／半泣き | 目に涙を浮かべ、眉を下げて半泣きの表情にする | Add tears welling up in the eyes and lower the eyebrows for a half-crying expression |

## 使い方の手順

1. 「通常」の画像をツールにアップロードする
2. 基本テンプレートの `{ここに表情ごとの指示を挿入}` に、作りたい表情のフレーズを当てはめる
3. 生成結果を確認し、髪型・輪郭・線の太さがズレていないかチェックする
4. ズレていたら「一切変えないでください」の部分を強調する、または「添付画像に忠実に」という言葉を追加して再生成する

## 注意点

- 生成のたびに顔が微妙に変わることがあるため、複数回試して一番近いものを採用する
- 口パク・まばたき用のパーツ分けは、この表情差分とは別工程（Photopeaでの手動加工、またはフリー素材の流用）で対応する
