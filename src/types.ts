/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type Role = 'Siswa' | 'Guru' | 'Admin';

export interface UserSession {
  success: boolean;
  role: Role;
  name: string;
  uid: string;
  kelas: string;
  isWali: boolean;
  jabatan?: string;
  message?: string;
  roleBackup?: Role;
  pass?: string;
}

export interface Student {
  nisn: string;
  nama: string;
  jenisKelamin: 'L' | 'P';
  tempat: string;
  tgl: string;
  kelas: string;
  ayah: string;
  ibu: string;
  hp: string;
  foto?: string;
}

export type Jabatan = 'Kamad' | 'Wakamad' | 'Guru' | 'Guru Wali Kelas';

export interface Teacher {
  nip: string;
  nama: string;
  jabatan: Jabatan;
  kelas: string;
  user: string;
  pass: string;
  role: Role;
  foto?: string;
}

export interface TeachingSchedule {
  id: string;
  nip: string;
  namaGuru: string;
  kelas: string;
  targetPertemuan: number;
  mapel?: string;
  hari: string;
  jps?: number[];
}

export interface Classroom {
  nama: string;
  wali: string;
  jumlah: number;
}

export interface Attendance {
  id: string;
  nisn: string;
  nama: string;
  kelas: string;
  tanggal: string;
  jam: string;
  status: string;
  terlambat: number;
  keterangan: string;
  scanTime?: string;
}

export interface TeacherAttendance {
  id: string;
  nip: string;
  nama: string;
  kelas: string;
  tanggal: string;
  jam: string;
  terlambat: number;
  mapel?: string;
  status?: string;
  keterangan?: string;
}

export interface DashboardStats {
  siswaCount: number;
  guruCount: number;
  hadirHariIni: number;
  terlambatHariIni: number;
  chartData?: any[];
}

export interface Holiday {
  tanggal: string;
  keterangan: string;
}

export interface DaySetting {
  hari: string;
  masuk: string;
  pulang: string;
  activeJps?: number[];
  reasonInactive?: string;
  targetDate?: string;
  jpTimes?: { [jp: number]: { start: string; end: string } };
}

export interface AppConfig {
  logoUrl?: string;
  schoolName?: string;
  cardTemplateUrl?: string;
  subjects?: string[];
}
