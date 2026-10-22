-- =============================================================================
-- Skema MySQL/MariaDB untuk Portal Karyawan (K.O.M.A.N.D.O)
-- Pengganti Cloud Firestore.
--
-- Dibuat otomatis oleh db_mysql.py saat aplikasi pertama kali dijalankan.
-- File ini disediakan untuk dokumentasi / import manual lewat phpMyAdmin.
--
-- Konsep:
--   1 dokumen Firestore = 1 baris
--   `id`   = ID dokumen Firestore (dipertahankan)
--   `data` = seluruh isi dokumen (JSON)
--   kolom lain = salinan field yang dipakai untuk pencarian (indeks)
-- =============================================================================

CREATE DATABASE IF NOT EXISTS `portal_karyawan`
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE `portal_karyawan`;

-- ---------------------------------------------------------------------------
-- users
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `users` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `email` VARCHAR(190) NULL,
  `role` VARCHAR(20) NULL,
  `superior_id` VARCHAR(191) NULL,
  `fullname` VARCHAR(190) NULL,
  `telegram_chat_id` VARCHAR(64) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_users_email` (`email`),
  KEY `idx_users_role` (`role`),
  KEY `idx_users_superior` (`superior_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- struktur (pohon organisasi)
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `struktur` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `user_id` VARCHAR(191) NULL,
  `parent_id` VARCHAR(191) NULL,
  `nama` VARCHAR(190) NULL,
  `urutan` INT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_struktur_user` (`user_id`),
  KEY `idx_struktur_parent` (`parent_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- tugas
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `tugas` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `user_id` VARCHAR(191) NULL,
  `assigned_by` VARCHAR(191) NULL,
  `status` VARCHAR(20) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_tugas_user` (`user_id`),
  KEY `idx_tugas_assigned` (`assigned_by`),
  KEY `idx_tugas_status` (`status`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- disposisi_surat
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `disposisi_surat` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `atasan_id` VARCHAR(191) NULL,
  `pengirim_id` VARCHAR(191) NULL,
  `asal_surat_id` VARCHAR(191) NULL,
  `status` VARCHAR(30) NULL,
  `file_nama_simpan` VARCHAR(255) NULL,
  `lampiran_nama_simpan` VARCHAR(255) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_disp_atasan` (`atasan_id`),
  KEY `idx_disp_pengirim` (`pengirim_id`),
  KEY `idx_disp_asal` (`asal_surat_id`),
  KEY `idx_disp_file` (`file_nama_simpan`),
  KEY `idx_disp_lampiran` (`lampiran_nama_simpan`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- disposisi_log
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `disposisi_log` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `surat_id` VARCHAR(191) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_displog_surat` (`surat_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- disposisi_hist (riwayat isian form disposisi)
-- id = id kolom form, mis. 'd-paraf'
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `disposisi_hist` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- bantuan_tiket
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `bantuan_tiket` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `user_id` VARCHAR(191) NULL,
  `status` VARCHAR(30) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_tiket_user` (`user_id`),
  KEY `idx_tiket_status` (`status`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- pengumuman
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `pengumuman` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- jadwal_rapat
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `jadwal_rapat` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- format_surat
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `format_surat` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------------
-- password_reset_otp (id = email lowercase)
-- ---------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `password_reset_otp` (
  `id` VARCHAR(191) NOT NULL,
  `data` LONGTEXT NOT NULL,
  `email` VARCHAR(190) NULL,
  PRIMARY KEY (`id`),
  KEY `idx_otp_email` (`email`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
