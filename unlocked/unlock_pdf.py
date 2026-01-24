import pikepdf
from pathlib import Path

# 设置目录（当前目录）
input_dir = Path('.')
output_dir = Path('./unlocked')
output_dir.mkdir(exist_ok=True)

for pdf_file in input_dir.glob('*.pdf'):
    try:
        pdf = pikepdf.open(pdf_file)
        output_path = output_dir / pdf_file.name
        pdf.save(output_path)
        pdf.close()
        print(f"✅ {pdf_file.name}")
    except Exception as e:
        print(f"❌ {pdf_file.name}: {e}")

print(f"\n完成！解锁后的文件在 {output_dir.absolute()}")