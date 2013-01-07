Attribute VB_Name = "mdlText"
Option Explicit

Public Const strServerSetting As String = "SQL SERVER SETTING"

Public Const strProfile As String = "REGISTRASI"

Public Const strLogin As String = "MENU LOGIN"

Public Const strWall As String = "UBAH WALLPAPER"

Public Const strAnalysis As String = "ANALISIS"
Public Const strFax As String = "FAX"
Public Const strPrinter As String = "SET PRINTER"
Public Const strChat As String = "CHAT"
Public Const strReminderList As String = "DAFTAR PENGINGAT"
Public Const strSyncFinance As String = "SYNCHRONIZE FINANCE"
Public Const strSyncAccounting As String = "SYNCHRONIZE ACCOUNTING"

Public Const strAddUser As String = "TAMBAH PEMAKAI"
Public Const strBRWTMUSER As String = "DAFTAR - PEMAKAI"

Public Const strReminder As String = "JADWAL PENGINGAT"
Public Const strPwdChange As String = "UBAH PASSWORD"
Public Const strBackup As String = "BACKUP DATABASE"
Public Const strUpdate As String = "UPDATE PROGRAM"

Public Const strAbout As String = "Tentang Program MJStone Software Production"

Public Const strTMITEM As String = "MASTER - BARANG"
Public Const strBRWTMITEM As String = "DAFTAR - BARANG"
Public Const strBRWTMITEMOPT As String = "DAFTAR - BARANG OPTIONAL"
Public Const strMISTMITEM As String = "INFORMASI - BARANG"

Public Const strTMITEMPRICE As String = "MASTER - HARGA JUAL"
Public Const strBRWTMITEMPRICE As String = "DAFTAR - HARGA JUAL"
Public Const strMISTMITEMPRICE As String = "INFORMASI - HARGA JUAL"

Public Const strTMPRICELIST As String = "MASTER - DAFTAR HARGA BARANG"
Public Const strBRWTMPRICELIST As String = "DAFTAR - DAFTAR HARGA BARANG"
Public Const strMISTMPRICELIST As String = "INFORMASI - DAFTAR HARGA BARANG"

Public Const strTMCONVERTPRICE As String = "MASTER - KONVERSI HARGA"

Public Const strTMSTOCKINIT As String = "MASTER - STOK AWAL"
Public Const strBRWTMSTOCKINIT As String = "DAFTAR - STOK AWAL"

Public Const strTMUNITY As String = "MASTER - SATUAN"
Public Const strBRWTMUNITY As String = "DAFTAR - SATUAN"
Public Const strMISTMUNITY As String = "INFORMASI - SATUAN"

Public Const strTMBRAND As String = "MASTER - MERK"
Public Const strBRWTMBRAND As String = "DAFTAR - MERK"
Public Const strMISTMBRAND As String = "INFORMASI - MERK"

Public Const strTMCATEGORY As String = "MASTER - JENIS"
Public Const strBRWTMCATEGORY As String = "DAFTAR - JENIS"
Public Const strMISTMCATEGORY As String = "INFORMASI - JENIS"

Public Const strTMGROUP As String = "MASTER - GRUP"
Public Const strBRWTMGROUP As String = "DAFTAR - GRUP"
Public Const strMISTMGROUP As String = "INFORMASI - GRUP"

Public Const strTMEMPLOYEE As String = "MASTER - KARYAWAN"
Public Const strBRWTMEMPLOYEE As String = "DAFTAR - KARYAWAN"
Public Const strMISTMEMPLOYEE As String = "INFORMASI - KARYAWAN"

Public Const strTMJOBTYPE As String = "MASTER - JABATAN"
Public Const strBRWTMJOBTYPE As String = "DAFTAR - JABATAN"
Public Const strMISTMJOBTYPE As String = "INFORMASI - JABATAN"

Public Const strTMDIVISION As String = "MASTER - DIVISI"
Public Const strBRWTMDIVISION As String = "DAFTAR - DIVISI"
Public Const strMISTMDIVISION As String = "INFORMASI - DIVISI"

Public Const strTMVENDOR As String = "MASTER - PEMASOK"
Public Const strBRWTMVENDOR As String = "DAFTAR - PEMASOK"
Public Const strMISTMVENDOR As String = "INFORMASI - PEMASOK"

Public Const strTMCONTACTVENDOR As String = "MASTER - KONTAK PEMASOK"

Public Const strTMCUSTOMER As String = "MASTER - CUSTOMER"
Public Const strBRWTMCUSTOMER As String = "DAFTAR - CUSTOMER"
Public Const strMISTMCUSTOMER As String = "INFORMASI - CUSTOMER"
Public Const strMISTMCUSTOMERTRANS As String = "INFORMASI - TRANSAKSI CUSTOMER"

Public Const strTMCONTACTCUSTOMER As String = "MASTER - KONTAK CUSTOMER"
Public Const strTMDELIVERYCUSTOMER  As String = "MASTER - ALAMAT KIRIM CUSTOMER"
Public Const strTMCUSTOMERNOTES As String = "MASTER - CATATAN CUSTOMER"
Public Const strTMCONTACTNOTES As String = "MASTER - CATATAN KONTAK"

Public Const strTMCURRENCY As String = "MASTER - MATA UANG"
Public Const strBRWTMCURRENCY As String = "DAFTAR - MATA UANG"
Public Const strMISTMCURRENCY As String = "INFORMASI - MATA UANG"

Public Const strTMCONVERTCURRENCY As String = "INFORMASI - NILAI TUKAR"

Public Const strTMWAREHOUSE As String = "MASTER - GUDANG"
Public Const strBRWTMWAREHOUSE As String = "DAFTAR - GUDANG"
Public Const strMISTMWAREHOUSE As String = "INFORMASI - GUDANG"

Public Const strTMREMINDERCUSTOMER As String = "PENGINGAT - CUSTOMER"

Public Const strTHITEMIN As String = "TRANSAKSI - PEMASUKKAN BARANG"
Public Const strTDITEMIN As String = "DATA BARANG (PEMASUKKAN BARANG)"
Public Const strBRWTHITEMIN As String = "DAFTAR - PEMASUKKAN BARANG"
Public Const strMISTHITEMIN As String = "INFORMASI - PEMASUKKAN BARANG"

Public Const strTHITEMOUT As String = "TRANSAKSI - PENGELUARAN BARANG"
Public Const strTDITEMOUT As String = "DATA BARANG (PENGELUARAN BARANG)"
Public Const strBRWTHITEMOUT As String = "DAFTAR - PENGELUARAN BARANG"
Public Const strMISTHITEMOUT As String = "INFORMASI - PENGELUARAN BARANG"

Public Const strTHMUTITEM As String = "TRANSAKSI - MUTASI BARANG"
Public Const strTDMUTITEM As String = "DATA BARANG (MUTASI BARANG)"
Public Const strBRWTHMUTITEM As String = "DAFTAR - MUTASI BARANG"
Public Const strMISTHMUTITEM As String = "INFORMASI - MUTASI BARANG"

Public Const strTHPOBUY As String = "TRANSAKSI - PURCHASE ORDER (BELI)"
Public Const strTDPOBUY As String = "DATA BARANG (PURCHASE ORDER (BELI))"
Public Const strBRWTHPOBUY As String = "DAFTAR - PURCHASE ORDER (BELI)"
Public Const strMISTHPOBUY As String = "INFORMASI - PURCHASE ORDER (BELI)"

Public Const strTHDOBUY As String = "TRANSAKSI - DELIVERY ORDER (BELI)"
Public Const strTDDOBUY As String = "DATA BARANG (DELIVERY ORDER (BELI))"
Public Const strBRWTHDOBUY As String = "DAFTAR - DELIVERY ORDER (BELI)"
Public Const strMISTHDOBUY As String = "INFORMASI - DELIVERY ORDER (BELI)"

Public Const strTHSJBUY As String = "TRANSAKSI - SURAT JALAN (BELI)"
Public Const strTDSJBUY As String = "DATA BARANG (SURAT JALAN (BELI))"
Public Const strBRWTHSJBUY As String = "DAFTAR - SURAT JALAN (BELI)"
Public Const strMISTHSJBUY As String = "INFORMASI - SURAT JALAN (BELI)"

Public Const strTHFKTBUY As String = "TRANSAKSI - FAKTUR (BELI)"
Public Const strTDFKTBUY As String = "DATA SURAT JALAN (FAKTUR (BELI))"
Public Const strBRWTHFKTBUY As String = "DAFTAR - FAKTUR (BELI)"
Public Const strMISTHFKTBUY As String = "INFORMASI - FAKTUR (BELI)"

Public Const strTHRTRBUY As String = "TRANSAKSI - RETUR (BELI)"
Public Const strTDRTRBUY As String = "DATA BARANG (RETUR (BELI))"
Public Const strBRWTHRTRBUY As String = "DAFTAR - RETUR (BELI)"
Public Const strMISTHRTRBUY As String = "INFORMASI - RETUR (BELI)"

Public Const strTHSALESSUM As String = "TRANSAKSI - SALES SUMMARY (JUAL)"
Public Const strBRWTHSALESSUM As String = "DAFTAR - SALES SUMMARY (JUAL)"
Public Const strMISTHSALESSUM As String = "INFORMASI - SALES SUMMARY (JUAL)"

Public Const strTHPOSELL As String = "TRANSAKSI - PURCHASE ORDER (JUAL)"
Public Const strTDPOSELL As String = "DATA BARANG (PURCHASE ORDER (JUAL))"
Public Const strBRWTHPOSELL As String = "DAFTAR - PURCHASE ORDER (JUAL)"
Public Const strMISTHPOSELL As String = "INFORMASI - PURCHASE ORDER (JUAL)"

Public Const strTHSOSELL As String = "TRANSAKSI - SALES ORDER (JUAL)"
Public Const strTDSOSELL As String = "DATA BARANG (SALES ORDER (JUAL))"
Public Const strBRWTHSOSELL As String = "DAFTAR - SALES ORDER (JUAL)"
Public Const strMISTHSOSELL As String = "INFORMASI - SALES ORDER (JUAL)"

Public Const strTHSJSELL As String = "TRANSAKSI - SURAT JALAN (JUAL)"
Public Const strTDSJSELL As String = "DATA BARANG (SURAT JALAN (JUAL))"
Public Const strBRWTHSJSELL As String = "DAFTAR - SURAT JALAN (JUAL)"
Public Const strMISTHSJSELL As String = "INFORMASI - SURAT JALAN (JUAL)"

Public Const strTHFKTSELL As String = "TRANSAKSI - FAKTUR (JUAL)"
Public Const strTDFKTSELL As String = "DATA SURAT JALAN (FAKTUR (JUAL))"
Public Const strBRWTHFKTSELL As String = "DAFTAR - FAKTUR (JUAL)"
Public Const strMISTHFKTSELL As String = "INFORMASI - FAKTUR (JUAL)"

Public Const strTHRTRSELL As String = "TRANSAKSI - RETUR (JUAL)"
Public Const strTDRTRSELL As String = "DATA BARANG (RETUR (JUAL))"
Public Const strBRWTHRTRSELL As String = "DAFTAR - RETUR (JUAL)"
Public Const strMISTHRTRSELL As String = "INFORMASI - RETUR (JUAL)"

Public Const strTHRECYCLE As String = "DATA - RECYCLE"

Public Const strMISTHSTOCK As String = "INFORMASI - SISA STOK"

Public Const strRPTTMSTOCKINIT As String = "LAPORAN - STOK AWAL"
Public Const strRPTTHSTOCK As String = "LAPORAN - STOK"
Public Const strRPTTHITEMIN As String = "LAPORAN - PEMASUKKAN BARANG"
Public Const strRPTTHITEMOUT As String = "LAPORAN - PENGELUARAN BARANG"
Public Const strRPTTHMUTITEM As String = "LAPORAN - MUTASI BARANG"

Public Const strRPTTHPOBUY As String = "LAPORAN - PURCHASE ORDER (BELI)"
Public Const strRPTTHDOBUY As String = "LAPORAN - DELIVERY ORDER (BELI)"
Public Const strRPTTHSJBUY As String = "LAPORAN - SURAT JALAN (BELI)"
Public Const strRPTTHFKTBUY As String = "LAPORAN - FAKTUR (BELI)"
Public Const strRPTTHRTRBUY As String = "LAPORAN - RETUR (BELI)"

Public Const strRPTTHPOSELL As String = "LAPORAN - PURCHASE ORDER (JUAL)"
Public Const strRPTTHSOSELL As String = "LAPORAN - SALES ORDER (JUAL)"
Public Const strRPTTHSJSELL As String = "LAPORAN - SURAT JALAN (JUAL)"
Public Const strRPTTHFKTSELL As String = "LAPORAN - FAKTUR (JUAL)"
Public Const strRPTTHRTRSELL As String = "LAPORAN - RETUR (JUAL)"

Public Const strRPTSUMTHPOSELL As String = "LAPORAN - TOTAL ORDER CUSTOMER"

Public Const strCUSTOMERIDINIT As String = "C"
Public Const strCUSTOMERREMINDER As String = "Customer"
Public Const strVENDORIDINIT As String = "V"
Public Const strPOIDINIT As String = "PO"
Public Const strSOIDINIT As String = "SO"
Public Const strDOIDINIT As String = "DO"
Public Const strSJIDINIT As String = "SJ"
Public Const strFKTIDINIT As String = "FK"
Public Const strRTRIDINIT As String = "RT"
Public Const strSELLINIT As String = "SELL"
Public Const strBUYINIT As String = "BUY"
Public Const strITEMINIDINIT As String = "IN"
Public Const strITEMOUTIDINIT As String = "OUT"
Public Const strMUTITEMIDINIT As String = "MUT"
Public Const strPOSELLREMINDER As String = "PO Jual"

Public Const strTownInit As String = "J"
