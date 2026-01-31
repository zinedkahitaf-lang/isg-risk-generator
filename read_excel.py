import pandas as pd
import os

path = r'C:\Users\zined\.gemini\antigravity\scratch\isg_risk_generator'
files = os.listdir(path)
xlsx_file = [f for f in files if f.endswith('.xlsx')][0]
file_path = os.path.join(path, xlsx_file)

print(f"Dosya: {xlsx_file}")

# Sayfaları listele
xl = pd.ExcelFile(file_path)
print(f"Sayfalar: {xl.sheet_names}")

# TOSYALI sayfasını oku
df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0], header=None)

# Dosyaya kaydet
with open(os.path.join(path, 'excel_output.txt'), 'w', encoding='utf-8') as f:
    f.write(f"Dosya: {xlsx_file}\n")
    f.write(f"Sayfa: {xl.sheet_names[0]}\n")
    f.write(f"Boyut: {df.shape[0]} satır x {df.shape[1]} sütun\n\n")
    f.write("="*100 + "\n")
    f.write("İlk 50 satır:\n")
    f.write("="*100 + "\n\n")
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    pd.set_option('display.max_colwidth', 40)
    f.write(df.head(50).to_string())

print(f"Dosya kaydedildi!")
print(f"Toplam: {df.shape[0]} satır, {df.shape[1]} sütun")
