# Strategi Efektivitas Konverter Heuristik (Word/PDF to Moodle XML)

Dokumen ini berisi rencana strategis untuk meningkatkan kehandalan program dalam menangani keberagaman format naskah soal secara otomatis dan interaktif.

## 1. Arsitektur Pengenalan Berlapis (Layered Recognition)
Alih-alih mengandalkan satu pola tunggal, program menggunakan beberapa lapisan validasi:
- **Pola Dominan**: Program menganalisis format 5 soal pertama (misal: menggunakan 1. atau Soal 1). Jika pola ditemukan, program akan memprioritaskan pola tersebut untuk soal berikutnya.
- **Analisis Proximity (Jarak)**: Sebuah teks hanya dianggap "Soal" jika diikuti oleh daftar "Opsi" (A, B, C) dalam jarak kedekatan tertentu. Hal ini mencegah Instruksi atau Judul terdeteksi sebagai soal.

## 2. Manajemen Konteks & Teks Bersama (Context Persistence)
Untuk menangani soal bertipe "Teks bacaan untuk nomor X - Y":
- **State Machine**: Program melacak status Active_Passage. 
- **Auto-Injection**: Teks bacaan akan otomatis disisipkan ke dalam questiontext setiap soal yang berada dalam rentang nomor tersebut.
- **Visual Styling**: Menggunakan kotak abu-abu atau garis pemisah (<hr/>) pada XML Moodle agar bacaan terpisah secara visual dari pertanyaan.

## 3. Sistem Confidence Score (Tingkat Keyakinan)
Mengelompokkan hasil ekstraksi berdasarkan kejelasan format:
- **High Confidence**: Format standar (1., A., B., C., Kunci:). Langsung masuk daftar review utama.
- **Low Confidence**: Teks yang memiliki angka tapi tidak memiliki opsi, atau teks panjang di antara dua soal.
- **Safety Net**: Teks Low Confidence ditampilkan di bagian "Blok Terabaikan" pada layar Review agar pengguna bisa melakukan konfirmasi manual.

## 4. Normalisasi Artifact Dokumen
Membersihkan "sampah" digital hasil konversi Pandoc/Word:
- **Space & HTML Cleaning**: Menghapus &nbsp;, spasi ganda, tag kosong <p></p>, dan tag <span> yang tidak perlu.
- **Bullet Flattening**: Memaksa penomoran otomatis Word menjadi teks statis agar mudah dibaca oleh Regex.

## 5. Loop Feedback Interaktif (Interactive Review)
Menjadikan pengguna sebagai validator terakhir:
- **Fitur Merge/Gabung**: Memungkinkan pengguna menggabungkan dua blok teks yang terpisah (misal: soal yang terpotong ke halaman baru).
- **Quick Edit**: Mengizinkan perbaikan cepat pada teks soal atau huruf kunci jawaban langsung di layar review sebelum XML di-generate.

## 6. Penanganan Media Universal
- **Image-to-Question Mapping**: Memastikan gambar di-embed tepat di bawah teks paragraf tempat gambar tersebut muncul di Word.
- **Table Formatting**: Mengonversi tabel Word menjadi tabel HTML yang memiliki border dan responsif di layar Moodle.

---
*Strategi ini dirancang untuk evolusi berkelanjutan seiring dengan bertambahnya variasi naskah soal yang dihadapi.*
