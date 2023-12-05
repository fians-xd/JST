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


import pandas as pd

# Membaca data dari file Excel
data = pd.read_excel('/content/drive/MyDrive/Colab Notebooks/_Tugas_Akhir/data/datapanel_bersih.xlsx') 

# Membuat kedua kolom throughput
data['Throughput_DL_rata_rata_pengguna'] = (
    0.5 * data['Tel_movil_unid'] + 0.3 * data['Sub_internet_fijo_und'] + 0.2 * data['Compus_por_muni_unid']
)

data['Throughput_DL_rata_rata_sel'] = (
    0.4 * data['Hogar_1_radio_porc'] + 0.6 * data['VAB_Tel_miles_2007base'] + 0.1 * data['Hogar_con_cable_porc']
)

# Menyimpan hasil perubahan ke file Excel
data.to_excel('/content/drive/MyDrive/Colab Notebooks/_Tugas_Akhir/data/data_panel.xlsx', index=False) 
