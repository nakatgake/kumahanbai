from PIL import Image, ImageDraw, ImageFont
import os

def create_hanko():
    # サイズと色
    size = 400
    padding = 20
    border_width = 15
    red_color = (180, 0, 0, 255)  # 深みのある朱色
    
    # 画像作成 (RGBA - 透明背景)
    img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    
    # 外枠を描画
    draw.rectangle([padding, padding, size-padding, size-padding], outline=red_color, width=border_width)
    
    # テキスト設定
    # 株式会社熊ノ護化研 (9文字) -> 3x3の配置
    # 右から左、上から下に読む配置:
    # [3] [2] [1]
    #  研  熊  株
    #  化  ノ  会
    #  護  護  社  <-- いや、10文字にするために「印」を足すのが一般的だが、社名のみでも可
    
    text = "株式会社熊ノ護化研"
    chars = [
        ["社", "会", "株"], # 右端
        ["護", "ノ", "熊"], # 真ん中
        ["印", "研", "化"]  # 左端 (10文字目として「印」を補うのが一般的)
    ]
    
    font_path = "static/fonts/NotoSansJP-Regular.otf"
    if not os.path.exists(font_path):
        font_path = "C:\\Windows\\Fonts\\msgothic.ttc"
        
    font_size = 85
    try:
        font = ImageFont.truetype(font_path, font_size)
    except:
        font = ImageFont.load_default()

    # 文字を描画
    col_width = (size - 2*padding) // 3
    row_height = (size - 2*padding) // 3
    
    for c_idx, col_chars in enumerate(chars):
        x = size - padding - (c_idx + 1) * col_width + (col_width - font_size) // 2
        for r_idx, char in enumerate(col_chars):
            y = padding + r_idx * row_height + (row_height - font_size) // 2
            draw.text((x, y), char, font=font, fill=red_color)
            
    # 保存
    os.makedirs('static/images', exist_ok=True)
    img.save('static/images/seal.png')
    print("True transparent seal generated at static/images/seal.png")

if __name__ == "__main__":
    create_hanko()
