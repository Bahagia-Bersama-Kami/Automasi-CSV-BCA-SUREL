import os
import glob
import shutil
import sys
import configparser

def main():
    if not os.path.exists('config.conf'):
        print("--> Error: File config.conf tidak ditemukan.")
        sys.exit()

    config = configparser.ConfigParser()
    
    try:
        config.read('config.conf')
        
        if 'DIRECTORY' not in config:
            print("--> Error: Header [DIRECTORY] tidak ditemukan dalam config.conf.")
            sys.exit()

        dest_pusat = config['DIRECTORY'].get('pusat', '').strip()
        dest_depo = config['DIRECTORY'].get('depo', '').strip()
        dest_bri = config['DIRECTORY'].get('bri', '').strip()

    except Exception as e:
        print(f"--> Gagal memproses config.conf: {e}")
        sys.exit()

    if not dest_pusat or not dest_depo or not dest_bri:
        print("--> Error: Path untuk 'pusat', 'depo', atau 'bri' kosong atau tidak ditemukan di bawah [DIRECTORY].")
        sys.exit()

    files_bca = glob.glob('BCA *.xlsx')
    files_bri = glob.glob('BRI *.pdf')
    all_files = files_bca + files_bri

    if not all_files:
        print("--> Tidak ditemukan file Excel BCA *.xlsx atau PDF BRI *.pdf. Proses dilewati.")
        return

    print(f"--> Ditemukan {len(all_files)} file. Memulai pemindahan...")

    jumlah_berhasil = 0

    for file_path in all_files:
        filename = os.path.basename(file_path)
        target_folder = None

        if filename.startswith("BCA 3777"):
            target_folder = dest_pusat
        elif filename.startswith("BCA "):
            target_folder = dest_depo
        elif filename.startswith("BRI "):
            target_folder = dest_bri
        
        if target_folder:
            if os.path.exists(target_folder):
                try:
                    destination_path = os.path.join(target_folder, filename)
                    shutil.move(file_path, destination_path)
                    print(f"--> Sukses dipindahkan: {filename} ke {target_folder}")
                    jumlah_berhasil += 1
                except Exception as e:
                    print(f"--> Gagal memindahkan {filename}: {e}")
            else:
                print(f"--> Folder tujuan tidak ditemukan: {target_folder}")
        else:
            print(f"--> File {filename} tidak masuk kriteria.")

    print(f"--> Proses Selesai. Total file dipindahkan: {jumlah_berhasil}")

if __name__ == "__main__":
    main()