-- phpMyAdmin SQL Dump
-- version 5.0.3
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Waktu pembuatan: 08 Bulan Mei 2021 pada 10.10
-- Versi server: 10.4.14-MariaDB
-- Versi PHP: 7.2.34

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `form`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `peserta`
--

CREATE TABLE `peserta` (
  `id` int(30) NOT NULL,
  `tglform` varchar(200) NOT NULL,
  `jenispendaftaran` varchar(200) NOT NULL,
  `tglmasuksekolah` date NOT NULL,
  `nis` varchar(200) NOT NULL,
  `nmrpeserta` varchar(200) NOT NULL,
  `paud` varchar(200) NOT NULL,
  `tk` varchar(200) NOT NULL,
  `skhun` varchar(200) NOT NULL,
  `ijazah` varchar(200) NOT NULL,
  `hobi` varchar(200) NOT NULL,
  `cita` varchar(200) NOT NULL,
  `namalengkap` varchar(200) NOT NULL,
  `jk` varchar(200) NOT NULL,
  `nisn` varchar(200) NOT NULL,
  `nik` varchar(200) NOT NULL,
  `tempatlahir` varchar(200) NOT NULL,
  `tgllahir` date NOT NULL,
  `agama` varchar(200) NOT NULL,
  `bkpribadi` varchar(200) NOT NULL,
  `alamat` varchar(200) NOT NULL,
  `rt` varchar(200) NOT NULL,
  `rw` varchar(200) NOT NULL,
  `namadusun` varchar(200) NOT NULL,
  `namadesa` varchar(200) NOT NULL,
  `kecamatan` varchar(200) NOT NULL,
  `kdpos` varchar(200) NOT NULL,
  `tinggal` varchar(200) NOT NULL,
  `transportasi` varchar(200) NOT NULL,
  `nohp` varchar(200) NOT NULL,
  `notelp` varchar(200) NOT NULL,
  `email` varchar(200) NOT NULL,
  `penkip` varchar(200) NOT NULL,
  `nokip` varchar(200) NOT NULL,
  `kwn` varchar(200) NOT NULL,
  `namaayah` varchar(200) NOT NULL,
  `thnlahirayah` date NOT NULL,
  `pendikayah` varchar(200) NOT NULL,
  `kerjaayah` varchar(200) NOT NULL,
  `hasilayah` varchar(200) NOT NULL,
  `bkayah` varchar(200) NOT NULL,
  `namaibu` varchar(200) NOT NULL,
  `thnlahiribu` date NOT NULL,
  `pendikibu` varchar(200) NOT NULL,
  `kerjaibu` varchar(200) NOT NULL,
  `hasilibu` varchar(200) NOT NULL,
  `bkibu` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data untuk tabel `peserta`
--

INSERT INTO `peserta` (`id`, `tglform`, `jenispendaftaran`, `tglmasuksekolah`, `nis`, `nmrpeserta`, `paud`, `tk`, `skhun`, `ijazah`, `hobi`, `cita`, `namalengkap`, `jk`, `nisn`, `nik`, `tempatlahir`, `tgllahir`, `agama`, `bkpribadi`, `alamat`, `rt`, `rw`, `namadusun`, `namadesa`, `kecamatan`, `kdpos`, `tinggal`, `transportasi`, `nohp`, `notelp`, `email`, `penkip`, `nokip`, `kwn`, `namaayah`, `thnlahirayah`, `pendikayah`, `kerjaayah`, `hasilayah`, `bkayah`, `namaibu`, `thnlahiribu`, `pendikibu`, `kerjaibu`, `hasilibu`, `bkibu`) VALUES
(1, '2021-05-04', 'Siswa Baru', '2021-05-04', '2010095', '12901290129012901290', 'Tidak', 'Ya', '1920192019201920', '1920192019201920', 'Olahraga', 'PNS', 'Rifat Hafiz Aditya', 'Laki - laki', '0013234918', '36788833338110001', 'Surabaya', '2001-08-30', 'Islam', 'Tidak', 'keputih tegal ', '2', '3', 'sukolilo', 'Keputih', 'sukolilo', '6012', 'Bersama Orangtua', 'Kendaraan Pribadi', '081336234008', '081336234008', 'hafisarfutz@gmail.com', 'Tidak', '', 'Indonesia(WNI)', 'Suryo', '2021-04-25', 'SMA Sederajat', 'Wirausaha', '500,000 - 999,999', 'Tidak', 'Setio', '2021-04-25', 'D4/S1', 'PNS/TNI/POLRI', '500,000 - 999,999', 'Tidak'),
(2, '2021-05-07', 'Siswa Baru', '2021-05-07', '12345', '12345', 'Ya', 'Ya', '12345', '12345', 'Olahraga', 'TNI/POLRI', 'Rifat Hafiz Aditya', 'Laki - laki', '111122222', '12345', 'Surabaya', '2021-05-03', 'Islam', 'Tidak', 'keputih tegal ', '2', '3', 'sukolilo', 'Keputih', 'sukolilo', '6012', 'Bersama Orangtua', 'Kendaraan Pribadi', '1212323123123', '12312312321', 'hafisarfutz@gmail.com', 'Tidak', '12312312', 'Indonesia(WNI)', 'Suryo', '2021-04-25', 'SMA Sederajat', 'Wirausaha', 'Kurang dari 500,000', 'Tidak', 'Setio', '2021-04-25', 'SMA Sederajat', 'Wirausaha', 'Kurang dari 500,000', 'Tidak');

--
-- Indexes for dumped tables
--

--
-- Indeks untuk tabel `peserta`
--
ALTER TABLE `peserta`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT untuk tabel yang dibuang
--

--
-- AUTO_INCREMENT untuk tabel `peserta`
--
ALTER TABLE `peserta`
  MODIFY `id` int(30) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
