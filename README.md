# Simple APP markltdown-mistralocr dengan Streamlit
Streamlit adalah sebuah pustaka (library) Python open-source yang digunakan untuk membuat aplikasi web interaktif dari script Python untuk keperluan data science dan machine learning.

## Libery ayng diperlukan
1. streamlit - Untuk membuat aplikasi web
2. pdfplumber - Untuk membaca file PDF
3. python-docx - Untuk membaca file DOCX
4. openpyxl - Untuk membaca file XLSX
5. python-pptx - Untuk membaca file PPTX

   ``` bash
   pip install streamlit pdfplumber python-docx openpyxl python-pptx
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
