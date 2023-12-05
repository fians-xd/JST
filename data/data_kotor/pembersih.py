import pandas as pd

# Membaca data dari file Excel
data = pd.read_excel("e:\Kuliah\Pemodelan_Data\Project\_Tugas_Akhir\data\data_kotor\datapanel.xlsx", engine="openpyxl")

# Mengecek tipe data dari setiap kolom
data_types = data.dtypes

# Membuat salinan kolom non-numerik
non_numeric_data = data.select_dtypes(exclude='number')

# Loop melalui setiap kolom dan mengubah yang bukan numerik menjadi numerik
for column in data.columns:
    if data_types[column] == object:  # Memeriksa apakah tipe data adalah objek (teks)
        data[column] = pd.to_numeric(data[column], errors='coerce')

# Mengisi missing values dengan nilai rata-rata dari setiap kolom
data_filled = data.fillna(data.mean())

# Menggabungkan kembali data numerik dengan data non-numerik
data = pd.concat([non_numeric_data, data_filled.select_dtypes(include='number')], axis=1)

# Menyimpan data yang telah diisi kembali ke file Excel
data_filled.to_excel("e:\Kuliah\Pemodelan_Data\Project\_Tugas_Akhir\data\data_kotor\data_filled.xlsx", index=False, engine="openpyxl")
