# WITH ... END WITH
menjalankan serangkaian pernyataan pada satu objek yang sama tanpa harus menyebutkan nama objek tersebut berulang kali.
```
with [nama objek]
	.properti = nilai
	.metode
end with
```
*Setiap properti atau metode di dalam blok tersebut **wajib** diawali dengan tanda titik (.)

Perbandingan tanpa menggunakan with
```
Range("A1").Font.Bold = True

Range("A1").Font.Color = vbRed

Range("A1").Interior.Color = vbYellow

Range("A1").Value = "Data Tersimpan"
```

Jika menggunakan with
```
With Range("A1")

    .Font.Bold = True

    .Font.Color = vbRed

    .Interior.Color = vbYellow

    .Value = "Data Tersimpan"

End With
```

# IF NOT … IS NOTHING THEN

sebuah pola keamanan (guard clause) yang digunakan untuk memastikan bahwa sebuah variabel objek telah terisi atau berhasil mereferensikan sesuatu sebelum kita melakukan tindakan lebih lanjut.

Penjelasan:

1. **Nothing**: artinya "tidak ada objek". Ini adalah nilai _default_ untuk semua variabel objek sebelum mereka diberikan tugas menggunakan perintah Set.
2. **Is Nothing**: Digunakan untuk mengecek apakah sebuah objek saat ini sedang kosong/kosong.
3. **Not**: Pembalik logika. Jadi, Not ... Is Nothing berarti "Bukan tidak ada" alias **"Ada"**.

Penting:
*Jika Anda mencoba mengakses properti atau metode dari objek yang masih `Nothing`, VBA akan langsung berhenti dan memunculkan error fatal:

**"Run-time error '91': Object variable or With block variable not set"**