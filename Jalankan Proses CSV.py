import os
import sys
import shutil
import glob
import subprocess

def clean_folder(folder_path, extensions):
    for ext in extensions:
        files = glob.glob(os.path.join(folder_path, ext))
        for f in files:
            try:
                os.remove(f)
            except Exception as e:
                print(f"--> Gagal menghapus {f}: {e}")

def main():
    folder_dapur = "Dapur"
    folder_input = "Input"
    folder_output = "Output"

    if not os.path.exists(folder_dapur):
        print(f"--> Folder {folder_dapur} tidak ditemukan.")
        input("--> Tekan enter untuk keluar")
        sys.exit()

    if not os.path.exists(folder_input):
        print(f"--> Folder {folder_input} tidak ditemukan.")
        input("--> Tekan enter untuk keluar")
        sys.exit()

    if not os.path.exists(folder_output):
        print(f"--> Folder {folder_output} tidak ditemukan.")
        input("--> Tekan enter untuk keluar")
        sys.exit()

    required_dapur = ['__init__.py', '1_bcacsv2excel.py', '2_SentEmail.py', '3_MovingDoc.py', 'config.conf']
    missing_dapur = [f for f in required_dapur if not os.path.exists(os.path.join(folder_dapur, f))]

    if missing_dapur:
        print(f"--> File berikut tidak ditemukan di {folder_dapur}: {missing_dapur}")
        input("--> Tekan enter untuk keluar")
        sys.exit()

    input_csvs = glob.glob(os.path.join(folder_input, "*.csv"))
    if not input_csvs:
        print(f"--> Tidak ditemukan file CSV di folder {folder_input}.")
        input("--> Tekan enter untuk keluar")
        sys.exit()

    print("--> Membersihkan folder Dapur dan Output sebelum memulai...")
    clean_folder(folder_dapur, ["*.csv", "*.xlsx"])
    clean_folder(folder_output, ["*.csv", "*.xlsx"])

    print(f"--> Menyalin {len(input_csvs)} file CSV dari {folder_input} ke {folder_dapur}...")
    for csv_file in input_csvs:
        shutil.copy(csv_file, folder_dapur)

    print("--> Pilih Proses:")
    print("--> 1. Hanya konversi CSV ke Excel")
    print("--> 2. Konversi ke Excel dan Kirim Surel")
    print("--> 3. Konversi, Kirim Surel, dan Pindahkan File")
    pilihan = input("--> Masukkan pilihan (1/2/3): ").strip().upper()

    if pilihan not in ['1', '2', '3']:
        print("--> Pilihan tidak valid.")
        clean_folder(folder_dapur, ["*.csv", "*.xlsx"])
        input("--> Tekan enter untuk keluar")
        sys.exit()

    try:
        print("--> Menjalankan 1_bcacsv2excel.py...")
        subprocess.run([sys.executable, '1_bcacsv2excel.py'], cwd=folder_dapur, check=True)

        if pilihan == '2' or pilihan == '3':
            print("--> Menjalankan 2_SentEmail.py...")
            subprocess.run([sys.executable, '2_SentEmail.py'], cwd=folder_dapur, check=True)

        if pilihan == '3':
            print("--> Menjalankan 3_MovingDoc.py...")
            subprocess.run([sys.executable, '3_MovingDoc.py'], cwd=folder_dapur, check=True)
        
        if pilihan == '1' or pilihan == '2':
            excel_files = glob.glob(os.path.join(folder_dapur, "*.xlsx"))
            print(f"--> Memindahkan {len(excel_files)} file Excel ke {folder_output}...")
            for f in excel_files:
                filename = os.path.basename(f)
                shutil.move(f, os.path.join(folder_output, filename))
        
        print("--> Proses sukses. Menghapus file CSV asli di folder Input...")
        clean_folder(folder_input, ["*.csv"])

    except subprocess.CalledProcessError as e:
        print(f"--> Terjadi kesalahan saat menjalankan script: {e}")
    except Exception as e:
        print(f"--> Terjadi error: {e}")

    print("--> Membersihkan file sampah di Dapur...")
    clean_folder(folder_dapur, ["*.csv", "*.xlsx"])

    input("--> Tekan enter untuk keluar")

if __name__ == "__main__":
    main()