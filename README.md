# Penata Naskah Jurnal Ilmiah PLATAX (JIP 2026)

Penulis **cukup mengunggah naskah .docx**. Aplikasi membaca bagian-bagiannya lalu menerapkan templat resmi JIP 2026.
Tidak ada formulir yang harus diketik atau disalin.

Keluaran (semua .docx):
1. Naskah blind review (berkas utama OJS)
2. Title Page (berkas tambahan terpisah)
3. Naskah lengkap dengan identitas (arsip penulis)

> Alat ini **opsional**. Templat Word manual tersedia sebagai cadangan di halaman depan dan di folder `templates/`.
> Pada tahap awal, kerapian format bukan alasan penolakan naskah.

## Menjalankan
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Berkas
| Berkas | Fungsi |
|---|---|
| `app.py` | Antarmuka Streamlit versi unggah |
| `naskah_parser.py` | Membaca .docx bebas-format (judul, penulis, abstrak, bab, tabel, gambar, rumus, pustaka, pernyataan akhir) |
| `naskah_builder.py` | Menyusun .docx dari templat resmi di `templates/` |
| `templates/` | Templat resmi (versi “bersih”) |
| `app_formulir.py` | Versi lama berbasis formulir (cadangan; memakai `platax_builder.py`) |
| `tests/` | Uji dasar (`python -m pytest tests`) |

## Yang dikenali dari naskah
Judul ID/EN, penulis (superskrip afiliasi, `*` korespondensi), afiliasi, e-mail, ORCID, Abstrak/Abstract, Kata kunci/Keywords,
bab (Pendahuluan, Metode, Hasil, Pembahasan, Simpulan) beserta subbab, tabel (sel gabung), gambar (termasuk beberapa gambar satu keterangan),
rumus (OMML Word dipertahankan), sitiran APA (diberi warna dan tautan ke pustaka), konflik kepentingan, pendanaan, ucapan terima kasih,
kontribusi CRediT, ketersediaan data, etik, penggunaan AI, dan daftar pustaka.

Pengenalan bersifat heuristik. Aplikasi menampilkan ringkasan temuan dan peringatan sebelum berkas diunduh.

## Keterbatasan yang diketahui
- Format lama `.doc`, PDF, dan gambar berformat EMF/WMF tidak didukung.
- Bab “Hasil dan Pembahasan” yang digabung dibaca sebagai bab Hasil.
- Rumus belum diuji visual di Microsoft Word; periksa hasilnya sebelum diunggah ke OJS.
- Redaksi perlu menyetujui bunyi kebijakan pada halaman depan.
