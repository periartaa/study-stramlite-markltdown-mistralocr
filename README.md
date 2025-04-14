# Simple APP markltdown-mistralocr dengan Streamlit
Streamlit adalah sebuah pustaka (library) Python open-source yang digunakan untuk membuat aplikasi web interaktif dari script Python untuk keperluan data science dan machine learning.

## Libery yang diperlukan simple app dari MarkItDown
1. streamlit - Untuk membuat aplikasi web
2. pdfplumber - Untuk membaca file PDF
3. python-docx - Untuk membaca file DOCX
4. openpyxl - Untuk membaca file XLSX
5. python-pptx - Untuk membaca file PPTX

   ``` bash
   pip install streamlit pdfplumber python-docx openpyxl python-pptx
   ```
## Libery yang diperlukan simple app dari MistralOCR
1. streamlit - Unutk membuat aplikasi web interaktif dengan Python.
2. python-dotenv - Untuk membaca file .env yang berisi variabel lingkungan (environment variables).
3. pillow (PIL) - Untuk manipulasi gambar (buka, edit, simpan format JPEG/PNG).
4. pypdf2 (PyPDF2) - Untuk Membaca dan memanipulasi file PDF.
5. python-docx - Untuk membaca/menulis file Microsoft Word (.docx).
6. python-pptx - Untuk Membaca/menulis file PowerPoint (.pptx).
7. pandas - Untuk analisis data (membaca Excel/CSV, manipulasi tabel).

   ``` bash
   pip install python-dotenv requests pillow pypdf2 pdf2image python-docx python-pptx pandas streamlit
   ```

## Langkah menjalankan simple app
1. Pastikan streamlit sudah terinstall di perangkat anda
2. Donlowad atau clone ke device
   ``` bash
   git clone https://github.com/periartaa/study-stramlite-markltdown-mistralocr.git
   ```
3. Pilih file yang akan dijalankan
4. Jika memlih file main.py
   File main.py merupakan simple app dari MarkItDown
   ``` bash
   python -m streamlit run main.py
   ```
   
   
6. Jika memlih file main-2.py
   file main.py merupakan simple app dari MistralOCR
    ``` bash
    python -m streamlit run main-2.py
    ```

## Kesimpulan
1. Dalam penerapannya MarkItDown lebih mudah digunakan dan diimplementasikan karena tidak membutuhkan API untuk mengaksesnya
2. Dalam hasilnya MarkItDown menghasilkan output yang lebih baik dari MistralOCR dalam hal teks dalam table
3. Dalam hasil MarkItDown menghasilkan output yang lebih baik dari MistralOCR dalam dokumen PPTX karena MistralOCR ada format yang tidak bisa diolah.
