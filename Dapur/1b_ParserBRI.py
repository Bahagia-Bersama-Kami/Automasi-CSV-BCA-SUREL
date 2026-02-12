import os
import configparser
from datetime import datetime

def proses_rename_pdf():
    config = configparser.ConfigParser()
    config_file = 'config.conf'
    
    if not os.path.exists(config_file):
        print(f"Error: File {config_file} tidak ditemukan!")
        return

    config.read(config_file)
    
    if 'BANK_MAPPING' not in config:
        print("Error: Header [BANK_MAPPING] tidak ditemukan di config.conf")
        return
        
    bank_mapping = config['BANK_MAPPING']
    
    folder_path = os.getcwd()
    files = os.listdir(folder_path)

    for filename in files:
        if filename.startswith("IBIZ_") and filename.endswith(".pdf"):
            try:
                parts = filename.split('_')
                kode_bank = parts[1]
                tanggal_raw = parts[2] 
                
                if kode_bank in bank_mapping:
                    nama_cabang = bank_mapping[kode_bank]
                else:
                    print(f"Lewati: Kode bank {kode_bank} tidak ada di config.")
                    continue
                    
                date_obj = datetime.strptime(tanggal_raw, "%Y%m%d")
                tanggal_baru = date_obj.strftime("%d %b").upper()
                
                nama_baru = f"BRI {nama_cabang} {tanggal_baru}.pdf"
                
                old_file_path = os.path.join(folder_path, filename)
                new_file_path = os.path.join(folder_path, nama_baru)
                
                if not os.path.exists(new_file_path):
                    os.rename(old_file_path, new_file_path)
                    print(f"Sukses: '{filename}' -> '{nama_baru}'")
                else:
                    print(f"Info: File '{nama_baru}' sudah ada, rename dibatalkan.")

            except Exception as e:
                print(f"Error memproses file {filename}: {e}")

if __name__ == "__main__":
    proses_rename_pdf()