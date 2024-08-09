import pandas as pd
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

# Load dataset
file_path = "E:\\Kuliah Indro\\Magang\\Data Analisis\\PenjualanABC.xlsx"
df = pd.read_excel(file_path)

# Membersihkan kumpulan data dengan menghapus baris yang tidak perlu dan mengatur nama kolom yang tepat
df_cleaned = df.drop([0, 1, 2]).reset_index(drop=True)
df_cleaned.columns = df.iloc[2]
df_cleaned = df_cleaned.rename(columns=lambda x: x.strip() if isinstance(x, str) else x)
df_cleaned = df_cleaned.reset_index(drop=True)

# Mengubah kolom yang tipe datanya terdapat kesalahan menjadi tipe data yang tepat
df_cleaned['SO Date'] = pd.to_numeric(df_cleaned['SO Date'], errors='coerce')
df_cleaned['SO Date'] = pd.to_datetime(df_cleaned['SO Date'], unit='D', origin='1899-12-30')

# Mengekstrak Bulan dan Tahun dari 'SO Date'
df_cleaned['Month'] = df_cleaned['SO Date'].dt.month
df_cleaned['Year'] = df_cleaned['SO Date'].dt.year

# Analisis penjualan berdasarkan rentan waktu =
sales_over_time = df_cleaned['SO Date'].value_counts().sort_index()
print("Penjualan dari waktu ke waktu: ")
print(sales_over_time)

# Menentukan produk yang paling laku terjual
best_selling_products = df_cleaned['Product'].value_counts().head(10)
print("Produk yang Paling Banyak Terjual: ")
print(best_selling_products)

# Identifikasi produk yang paling sedikit terjual
low_selling_products = df_cleaned['Product'].value_counts().tail(10)
print("Produk dengan Penjualan Rendah: ")
print(low_selling_products)

# Identifikasi Top Customer
customer_orders = df_cleaned['Customer'].value_counts().head(20)
print("Top 10 Pelanggan Teratas Berdasarkan Jumlah Pemesanan: ")
print(customer_orders)

# Group by product, month, and year untuk menghitung total pendapatan
monthly_revenue = df_cleaned.groupby(['Product', 'Year', 'Month']).agg(
    total_revenue=pd.NamedAgg(column='Amount Total', aggfunc='sum')
).reset_index()

# Pivot data monthly_revenue untuk menghasilkan total pendapatan produk tiap bulan
monthly_revenue_pivot = monthly_revenue.pivot_table(index=['Product', 'Month'], columns='Year', values='total_revenue')
monthly_revenue_pivot = monthly_revenue_pivot.fillna(0).infer_objects(copy=False)
print(monthly_revenue_pivot)

# Menyimpan hasil Pivot Table kedalam Excel
monthly_revenue_pivot.to_excel('Pendapatan_Bulanan_by_Product.xlsx')

# Group by month dan year untuk melihat pattern musiman
monthly_sales = df_cleaned.groupby(['Year', 'Month']).size().reset_index(name='Jumlah Pesanan')

# Melakukan pivot data unuk menemukan pola penjualan musiman
monthly_sales_pivot = monthly_sales.pivot(index='Month', columns='Year', values='Jumlah Pesanan')

#Tampilkan Pola Penjualan Musiman
print('Pola Penjualan Musiman')
print(monthly_sales_pivot)

# Mendefinisikan diskon berdasarkan frekuensi pemesanan paling banyak
def determine_discount(jumlah_pesanan):
    if jumlah_pesanan > 50:
        return 0.20  # 20% discount
    elif jumlah_pesanan > 30:
        return 0.15  # 15% discount
    elif jumlah_pesanan > 20:
        return 0.10  # 10% discount
    elif jumlah_pesanan > 10:
        return 0.05  # 5% discount
    else:
        return 0.00  # No discount

# Ambil data top 20 customer teratas berdasarkan pemesanan paling banyak
top_20_customers = customer_orders.head(20)

# Tetapkan diskon ke top 20 customer
top_20_discounts = top_20_customers.apply(determine_discount)

# Memasukkan data Customer yang mendapatkan diskon kedalam DataFrame
discounts_df = pd.DataFrame({
    'Customer': top_20_customers.index,
    'Jumlah Pesanan': top_20_customers.values,
    'Diskon': top_20_discounts.values
})

print("Top 20 Pelanggan yang Mendapatkan Diskon")
print(discounts_df)

# Save DataFrame kedalam file Excel 
discounts_df.to_excel('Diskon_Pelanggan.xlsx', index=False)

# Definisikan diskon untuk produk yang paling sedikit terjual
def suggest_discount(jumlah_pesanan):
    if jumlah_pesanan <= 5:
        return 0.25  # 25% diskon
    elif jumlah_pesanan <= 10:
        return 0.20  # 20% diskon
    else:
        return 0.15  # 15% diskon

low_selling_discounts = low_selling_products.apply(suggest_discount)

# Masukkan diskon untuk produk yang paling sedikit terjual DataFrame
low_selling_discounts_df = pd.DataFrame({
    'Product': low_selling_products.index,
    'Jumlah Pesanan': low_selling_products.values,
    'Diskon yang Disarankan': low_selling_discounts.values
})

print('Rekomendasi Produk Penjualan Terendah yang Memerlukan Diskon')
print(low_selling_discounts_df)

# Save DataFrame kedalam file Excel 
low_selling_discounts_df.to_excel('Diskon_Produk_dengan_Penjualan_Rendah.xlsx', index=False)

# Mengagregasi data untuk mendapatkan total pesanan dan total pengeluaran tiap pelanggan
customer_data = df_cleaned.groupby('Customer').agg(
    total_orders=pd.NamedAgg(column='SO Number', aggfunc='count'),
    total_pengeluaran=pd.NamedAgg(column='Amount Total', aggfunc='sum')
).reset_index()

# Standarisasi Data Customer Data
scaler = StandardScaler()
customer_data_scaled = scaler.fit_transform(customer_data[['total_orders', 'total_pengeluaran']])

# Mengaplikasikan K-Means Clustering untuk Menghasilkan Segmentasi Pelanggan
kmeans = KMeans(n_clusters=4, random_state=42)
customer_data['Segment'] = kmeans.fit_predict(customer_data_scaled)
print('Segmentasi Customer')
print(customer_data)

# Menyimpan Data Segmentasi Pelanggan kedalam Excel
customer_data.to_excel('Segmentasi_Pelanggan.xlsx', index=False)

# Menganalisis Segmentasi Pelanggan berdasarkan Rata-Rata Pemesanan dan Rata-Rata Pembelian
segment_analysis = customer_data.groupby('Segment').agg(
    customer_count=pd.NamedAgg(column='Customer', aggfunc='count'),
    avg_total_orders=pd.NamedAgg(column='total_orders', aggfunc='mean'),
    avg_total_spent=pd.NamedAgg(column='total_pengeluaran', aggfunc='mean')
).reset_index()
print('Analisis Segmentasi Customer')
print(segment_analysis)

# Menyimpan Data Analisis Segmentasi kedalam Excel
segment_analysis.to_excel('Analisis_Segmentasi Pelanggan.xlsx', index=False)

# Strategi yang Diusulkan
strategies = {
    0: {
        "Deskripsi": "Pelanggan bernilai tinggi dengan pesanan yang sering",
        "Strategi Marketing": "Tawarkan hadiah loyalitas dan diskon eksklusif untuk mempertahankan pelanggan bernilai tinggi ini."
    },
    1: {
        "Deskripsi": "Pelanggan bernilai rendah dengan pesanan yang jarang",
        "Strategi Marketing": "Kirim promosi dan diskon yang ditargetkan untuk mendorong pembelian yang lebih sering."
    },
    2: {
        "Deskripsi": "Pelanggan bernilai sedang dengan frekuensi moderat",
        "Strategi Marketing": "Berikan rekomendasi yang dipersonalisasi dan penawaran bundel untuk meningkatkan nilai pesanan rata-rata."
    },
    3: {
        "Deskripsi": "Pelanggan baru atau dengan keterlibatan rendah",
        "Strategi Marketing": "Terlibat dengan kampanye penyambutan dan penawaran pengenalan untuk meningkatkan keterlibatan awal."
    }
}

for segment, strategy in strategies.items():
    print(f"Segment {segment}: {strategy['Deskripsi']}")
    print(f"Marketing Strategy: {strategy['Strategi Marketing']}\n")


# Daftar Harga berdasarkan Kategori Pelanggan
pricing_categories = df_cleaned['Price List Category'].value_counts()
print("Distribusi Pesanan Berdasarkan Kategori Harga: ")
print(pricing_categories)

# Visualisasi Segmentasi Pelanggan
plt.figure(figsize=(10, 6))
plt.scatter(customer_data['total_orders'], customer_data['total_pengeluaran'], c=customer_data['Segment'], cmap='viridis')
plt.title('Segmentasi Pelanggan')
plt.xlabel('Total Pesanan')
plt.ylabel('Total Pengeluaran')
plt.colorbar(label='Segment')
plt.grid(True)
plt.tight_layout()

# Visualisasi Harga Berdasarkan Kategori Pelanggan
plt.figure(figsize=(12, 6))
pricing_categories.plot(kind='bar', color='green')
plt.title('Kategori Harga')
plt.xlabel('Kategori Pelanggan')
plt.ylabel('Jumlah Pesanan')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.tight_layout()

# Visualisasi Produk dengan Penjualan Terendah
plt.figure(figsize=(12, 6))
low_selling_products.plot(kind='bar', color='red')
plt.title('Produk dengan Penjualan Rendah')
plt.xlabel('Produk')
plt.ylabel('Jumlah Pesanan')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.tight_layout()

# Visualisasi Produk Terlaris
plt.figure(figsize=(12, 6))
best_selling_products.plot(kind='bar', color='skyblue')
plt.title('Produk Terlaris')
plt.xlabel('Produk')
plt.ylabel('Jumlah Pesanan')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.tight_layout()

# Visualisasi Pelanggan dengan Jumlah Pemesanan Terbanyak
plt.figure(figsize=(12, 6))
customer_orders.plot(kind='bar', color='skyblue')
plt.title('Pelanggan Teratas dengan Jumlah Pemesanan Terbanyak')
plt.xlabel('Pelanggan')
plt.ylabel('Jumlah Pesanan')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.tight_layout()

# Visualisasi Top 20 Pelanggan yang Mendapatkan Diskon
plt.figure(figsize=(14, 9))
discounts_df.set_index('Customer')['Diskon'].plot(kind='bar', color='orange')
plt.title('Top 20 Pelanggan yang Mendapatkan Diskon')
plt.xlabel('Pelanggan')
plt.ylabel('Diskon')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.tight_layout()

# Visualisasi Pola Penjualan Musiman
plt.figure(figsize=(14, 8))
monthly_sales_pivot.plot(kind='line', marker='o')
plt.title('Pola Penjualan Musiman')
plt.xlabel('Bulan')
plt.ylabel('Jumlah Pesanan')
plt.grid(True)
plt.xticks(range(1, 13))
plt.tight_layout()

# Visualisasi Tren Penjualan
plt.figure(figsize=(12, 6))
sales_over_time.sort_index().plot(kind='line', marker='o')
plt.title('Tren Penjualan Seiring Waktu')
plt.xlabel('Tanggal')
plt.ylabel('Jumlah Pesanan Penjualan')
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()

# Show the plot
plt.show()
