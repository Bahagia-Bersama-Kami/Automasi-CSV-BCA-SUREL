import os
import glob
import smtplib
import configparser
import sys
import json
from email.message import EmailMessage
from email.utils import make_msgid

HISTORY_FILE = 'email_history.json'

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_history(history):
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=4)

def kirim_email(pengirim, password, penerima, subject, body, path_file, nama_file, reply_to_id=None, thread_ids=None):
    msg = EmailMessage()
    
    final_subject = subject
    if reply_to_id:
        if not subject.lower().startswith("re:"):
            final_subject = f"Re: {subject}"
        
        msg['In-Reply-To'] = reply_to_id
        if thread_ids:
            msg['References'] = " ".join(thread_ids)
        else:
            msg['References'] = reply_to_id

    msg['Subject'] = final_subject
    msg['From'] = pengirim
    msg['To'] = penerima
    msg.set_content(body)
    
    msg_id = make_msgid()
    msg['Message-ID'] = msg_id

    with open(path_file, 'rb') as f:
        file_data = f.read()
    
    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=nama_file)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(pengirim, password)
            smtp.send_message(msg)
        print(f"--> Sukses mengirim email ke {penerima} untuk file {nama_file}")
        return msg_id
    except Exception as e:
        print(f"--> Gagal mengirim ke {penerima}. Error: {e}")
        return None

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

    history_data = load_history()

    file_excel = glob.glob('BCA *.xlsx')
    file_pdf = glob.glob('BRI *.pdf')
    semua_file = file_excel + file_pdf

    if not semua_file:
        print("--> Tidak ditemukan file yang sesuai.")
        sys.exit()

    print(f"--> Ditemukan {len(semua_file)} file. Memulai proses...")

    for path_file in semua_file:
        nama_file = os.path.basename(path_file)
        nama_tanpa_ext = os.path.splitext(nama_file)[0]
        
        try:
            parts = nama_tanpa_ext.split()
            if len(parts) < 2:
                continue

            kode_unik = parts[1]
            penanda_waktu = parts[-1]

            if kode_unik in daftar_penerima:
                email_tujuan = daftar_penerima[kode_unik]
                ekstensi = os.path.splitext(nama_file)[1]
                history_key = f"{email_tujuan}_{penanda_waktu}_{ekstensi}"
                
                user_history = history_data.get(history_key, {})
                prev_msg_id = user_history.get('last_id')
                all_thread_ids = user_history.get('thread_ids', [])
                
                subject_final = subject_template.replace('{kode}', kode_unik).replace('{bulan}', penanda_waktu)
                content_final = content_template.replace('{nama_file}', nama_file)
                
                new_msg_id = kirim_email(
                    email_pengirim, 
                    email_password, 
                    email_tujuan, 
                    subject_final, 
                    content_final, 
                    path_file, 
                    nama_file,
                    reply_to_id=prev_msg_id,
                    thread_ids=all_thread_ids
                )

                if new_msg_id:
                    all_thread_ids.append(new_msg_id)
                    history_data[history_key] = {
                        'last_id': new_msg_id,
                        'thread_ids': all_thread_ids
                    }
                    save_history(history_data)
            else:
                print(f"--> Kode {kode_unik} pada file {nama_file} tidak terdaftar.")

        except Exception as e:
            print(f"--> Error memproses file {nama_file}: {e}")

    print("--> Proses selesai.")

if __name__ == "__main__":
    main()