import os
import glob
import smtplib
import configparser
import sys
from email.message import EmailMessage

def kirim_email(pengirim, password, penerima, subject, body, path_file, nama_file):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = pengirim
    msg['To'] = penerima
    msg.set_content(body)

    with open(path_file, 'rb') as f:
        file_data = f.read()
    
    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=nama_file)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(pengirim, password)
            smtp.send_message(msg)
        print(f"--> Sukses mengirim email ke {penerima} untuk file {nama_file}")
    except Exception as e:
        print(f"--> Gagal mengirim ke {penerima}. Error: {e}")

def main():
    if not os.path.exists('config.conf'):
        print("--> Error: File config.conf tidak ditemukan.")
        sys.exit()

    config = configparser.ConfigParser()
    try:
        config.read('config.conf')
        
        email_pengirim = config['EMAIL']['PENGIRIM']
        email_password = config['EMAIL']['PASSWORD']
        
        subject_template = config['ISI EMAIL']['SUBJECT']
        content_template = config['ISI EMAIL']['CONTENT'].replace('\\n', '\n')
        
        daftar_penerima = config['PENERIMA']
    except Exception as e:
        print(f"--> Gagal membaca konfigurasi: {e}")
        sys.exit()

    excel_files = glob.glob('BCA *.xlsx')

    if not excel_files:
        print("--> Tidak ditemukan file Excel BCA *.xlsx.")
        sys.exit()

    print(f"--> Ditemukan {len(excel_files)} file Excel. Memulai proses pengiriman email...")

    for path_file in excel_files:
        nama_file = os.path.basename(path_file)
        
        try:
            parts = nama_file.split()
            if len(parts) < 2:
                print(f"--> Format nama file tidak valid: {nama_file}")
                continue

            kode_unik = parts[1]

            if kode_unik in daftar_penerima:
                email_tujuan = daftar_penerima[kode_unik]
                subject_final = subject_template.replace('{nama_file}', nama_file)
                
                print(f"--> Memproses {nama_file} (Kode: {kode_unik}) tujuan {email_tujuan}")
                kirim_email(email_pengirim, email_password, email_tujuan, subject_final, content_template, path_file, nama_file)
            else:
                print(f"--> Kode {kode_unik} pada file {nama_file} tidak terdaftar di config.")

        except Exception as e:
            print(f"--> Error memproses file {nama_file}: {e}")

    print("--> Proses selesai.")

if __name__ == "__main__":
    main()