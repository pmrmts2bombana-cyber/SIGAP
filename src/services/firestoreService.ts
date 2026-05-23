import { 
  collection, 
  doc, 
  getDoc, 
  getDocs, 
  setDoc, 
  updateDoc, 
  deleteDoc, 
  query, 
  where, 
  orderBy, 
  limit, 
  onSnapshot,
  Timestamp,
  addDoc,
  deleteField
} from 'firebase/firestore';
import { db, auth } from '../firebase';
import { 
  Student, Teacher, Classroom, Attendance, TeacherAttendance, 
  DashboardStats, Holiday, DaySetting, UserSession, Role, TeachingSchedule 
} from '../types';
import { handleFirestoreError, OperationType } from '../lib/firebaseUtils';

export const firestoreService = {
  // Bootstrap Admin if empty
  bootstrapAdmin: async () => {
    console.log("[DEBUG] bootstrapAdmin triggered");
    try {
      const q = query(collection(db, 'teachers'), limit(1));
      const snapshot = await getDocs(q);
      console.log(`[DEBUG] Teachers collection snapshot size: ${snapshot.size}`);
      if (snapshot.empty) {
        console.log("[DEBUG] Database empty, bootstrapping admin account...");
        await setDoc(doc(db, 'teachers', 'ADMIN001'), {
          nip: 'ADMIN001',
          nama: 'Administrator Utama',
          user: 'admin',
          pass: 'admin123',
          role: 'Admin',
          kelas: ''
        });
        await firestoreService.initializeSettings();
        console.log("[DEBUG] Bootstrap SUCCESS");
        return true;
      }
      return false;
    } catch (e) {
      console.error("[DEBUG] Bootstrap error:", e);
      handleFirestoreError(e, OperationType.GET, "teachers");
      return false;
    }
  },

  // Auth Bridge
  checkLogin: async (identifier: string, p: string, r: Role): Promise<UserSession> => {
    try {
      if (r === 'Siswa') {
        const q = query(collection(db, 'students'), where('nisn', '==', identifier));
        const snapshot = await getDocs(q);
        if (snapshot.empty) return { success: false, message: "NISN tidak terdaftar.", role: 'Siswa', name: '', uid: '', kelas: '', isWali: false };
        const data = snapshot.docs[0].data() as Student;
        return { success: true, name: data.nama, role: 'Siswa', uid: data.nisn, kelas: data.kelas, isWali: false };
      } else {
        const q = query(collection(db, 'teachers'), where('user', '==', identifier), where('pass', '==', p));
        const snapshot = await getDocs(q);
        if (snapshot.empty) return { success: false, message: "Username atau Password salah.", role: 'Guru', name: '', uid: '', kelas: '', isWali: false };
        const data = snapshot.docs[0].data() as Teacher;
        if (data.role !== r && data.role !== 'Admin') return { success: false, message: "Akses ditolak.", role: data.role, name: '', uid: '', kelas: '', isWali: false };
        return { 
          success: true, 
          name: data.nama, 
          role: data.role, 
          uid: data.nip,
          kelas: data.kelas,
          isWali: !!data.kelas,
          jabatan: data.jabatan
        };
      }
    } catch (e) {
      handleFirestoreError(e, OperationType.GET, 'auth');
      return { success: false, message: "Error.", role: 'Siswa', name: '', uid: '', kelas: '', isWali: false };
    }
  },

  // Students
  saveSiswa: async (data: Student) => {
    try {
      await setDoc(doc(db, 'students', data.nisn), data);
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `students/${data.nisn}`);
      throw e;
    }
  },

  hapusSiswa: async (nisn: string) => {
    try {
      await deleteDoc(doc(db, 'students', nisn));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `students/${nisn}`);
      throw e;
    }
  },

  // Teachers
  saveGuru: async (data: Teacher) => {
    try {
      if (!data.role) data.role = 'Guru';
      await setDoc(doc(db, 'teachers', data.nip), data);
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `teachers/${data.nip}`);
      throw e;
    }
  },

  hapusGuru: async (nip: string) => {
    try {
      await deleteDoc(doc(db, 'teachers', nip));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `teachers/${nip}`);
      throw e;
    }
  },

  // Classrooms
  saveKelas: async (data: Classroom) => {
    try {
      await setDoc(doc(db, 'classrooms', data.nama), data);
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `classrooms/${data.nama}`);
      throw e;
    }
  },

  hapusKelas: async (nama: string) => {
    try {
      await deleteDoc(doc(db, 'classrooms', nama));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `classrooms/${nama}`);
      throw e;
    }
  },

  // Attendance
  saveAbsensiManual: async (data: any) => {
    try {
      const id = data.id || `${data.nisn}-${data.tanggal}`;
      await setDoc(doc(db, 'attendance', id), { ...data, id });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `attendance/${data.id}`);
      throw e;
    }
  },

  updateAbsensiStatus: async (id: string, status: string, ket?: string) => {
    try {
      await updateDoc(doc(db, 'attendance', id), { status, keterangan: ket || '' });
    } catch (e) {
      handleFirestoreError(e, OperationType.UPDATE, `attendance/${id}`);
      throw e;
    }
  },

  hapusAbsensi: async (id: string) => {
    try {
      await deleteDoc(doc(db, 'attendance', id));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `attendance/${id}`);
      throw e;
    }
  },

  hapusAbsensiGuru: async (id: string) => {
    try {
      await deleteDoc(doc(db, 'teacherAttendance', id));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `teacherAttendance/${id}`);
      throw e;
    }
  },

  resetAbsensiSiswa: async (bulan?: string) => {
    try {
      let q;
      if (bulan) {
        q = query(
          collection(db, 'attendance'),
          where('tanggal', '>=', bulan),
          where('tanggal', '<=', bulan + '\uf8ff')
        );
      } else {
        q = query(collection(db, 'attendance'));
      }
      const snapshot = await getDocs(q);
      const promises = snapshot.docs.map(d => deleteDoc(doc(db, 'attendance', d.id)));
      await Promise.all(promises);
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, 'attendance-all');
      throw e;
    }
  },

  resetAbsensiGuru: async (bulan?: string) => {
    try {
      let q;
      if (bulan) {
        q = query(
          collection(db, 'teacherAttendance'),
          where('tanggal', '>=', bulan),
          where('tanggal', '<=', bulan + '\uf8ff')
        );
      } else {
        q = query(collection(db, 'teacherAttendance'));
      }
      const snapshot = await getDocs(q);
      const promises = snapshot.docs.map(d => deleteDoc(doc(db, 'teacherAttendance', d.id)));
      await Promise.all(promises);
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, 'teacherAttendance-all');
      throw e;
    }
  },

  // Scan Logic
  processScan: async (nisn: string) => {
    try {
      const studentDoc = await getDoc(doc(db, 'students', nisn));
      if (!studentDoc.exists()) return { success: false, message: "Siswa tidak terdaftar." };
      const student = studentDoc.data() as Student;
      
      const now = new Date();
      const dayMap = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
      const dayName = dayMap[now.getDay()];
      
      if (dayName === "Minggu") {
        return { success: false, message: "Hari Minggu adalah hari libur." };
      }

      // Check for holiday
      const tanggal = now.toISOString().split('T')[0];
      const holidayDoc = await getDoc(doc(db, 'holidays', tanggal));
      if (holidayDoc.exists()) {
        const holidayData = holidayDoc.data() as Holiday;
        return { success: false, message: `Hari Libur: ${holidayData.keterangan}` };
      }

      const jam = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', hour12: false });
      
      const id = `${nisn}-${tanggal}`;
      const existing = await getDoc(doc(db, 'attendance', id));
      if (existing.exists()) return { success: false, message: "Anda sudah melakukan presensi hari ini." };

      // Get settings for the day
      const settingsDoc = await getDoc(doc(db, 'settings', dayName));
      let cutoffMasuk = "07:15"; // Default
      let cutoffPulang = "14:00"; // Default
      if (settingsDoc.exists()) {
        const s = settingsDoc.data() as DaySetting;
        if (s.masuk) cutoffMasuk = s.masuk;
        if (s.pulang) cutoffPulang = s.pulang;
      }

      // Calculate total minutes for robust comparison
      const hNow = now.getHours();
      const mNow = now.getMinutes();
      const currentMinutes = hNow * 60 + mNow;
      const [hMasuk, mMasuk] = cutoffMasuk.split(':').map(Number);
      const [hPulang, mPulang] = cutoffPulang.split(':').map(Number);
      
      const masukMinutes = hMasuk * 60 + mMasuk;
      const pulangMinutes = hPulang * 60 + mPulang;

      // Check if after dismissal time
      if (pulangMinutes > 0 && currentMinutes >= pulangMinutes) {
        const record: Attendance = {
          id,
          nisn,
          nama: student.nama,
          kelas: student.kelas,
          tanggal,
          jam,
          status: 'Alfa',
          terlambat: 0,
          keterangan: 'Absensi gagal sudah jam pulang'
        };
        await setDoc(doc(db, 'attendance', id), record);
        return { 
          success: false, 
          status: 'Alfa',
          message: `Absensi Gagal: Sudah jam pulang sekolah (${cutoffPulang}). Anda dinyatakan ALFA.` 
        };
      }

      const late = currentMinutes - masukMinutes;
      const terlambat = late > 0 ? late : 0;

      const formatLateDescription = (min: number) => {
        if (min <= 0) return 'Tepat Waktu';
        const h = Math.floor(min / 60);
        const m = min % 60;
        const hStr = h > 0 ? `${h} jam ` : '';
        const mStr = `${m} menit`;
        return `Terlambat ${hStr}${mStr}`;
      };

      const record: Attendance = {
        id,
        nisn,
        nama: student.nama,
        kelas: student.kelas,
        tanggal,
        jam,
        status: 'Hadir',
        terlambat,
        keterangan: formatLateDescription(terlambat)
      };

      await setDoc(doc(db, 'attendance', id), record);
      const lateMsg = terlambat > 0 ? ` (Terlambat ${formatLateDescription(terlambat)})` : ' (Tepat Waktu)';
      return { 
        success: true, 
        status: 'Hadir',
        message: `Berhasil Absen: ${student.nama}${lateMsg}` 
      };
    } catch (e) {
      return handleFirestoreError(e, OperationType.WRITE, 'scan');
    }
  },

  processClassScan: async (nip: string, className: string) => {
    try {
      const teacherDoc = await getDoc(doc(db, 'teachers', nip));
      if (!teacherDoc.exists()) return { success: false, message: "Guru tidak terdaftar." };
      const teacher = teacherDoc.data() as Teacher;

      const now = new Date();
      const dayMap = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
      const dayName = dayMap[now.getDay()];

      if (dayName === "Minggu") {
        return { success: false, message: "Hari Minggu tidak ada jadwal mengajar." };
      }

      // Check for holiday
      const tanggal = now.toISOString().split('T')[0];
      const holidayDoc = await getDoc(doc(db, 'holidays', tanggal));
      if (holidayDoc.exists()) {
        const holidayData = holidayDoc.data() as Holiday;
        return { success: false, message: `Hari ini Libur: ${holidayData.keterangan}` };
      }

      // Check settings for dismissal time
      const settingsDoc = await getDoc(doc(db, 'settings', dayName));
      let cutoffPulang = "14:00"; // Default
      if (settingsDoc.exists()) {
        const s = settingsDoc.data() as DaySetting;
        if (s.pulang) cutoffPulang = s.pulang;
      }

      const jam = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', hour12: false });
      const [hNow, mNow] = jam.split(':').map(Number);
      const [hPulang, mPulang] = cutoffPulang.split(':').map(Number);
      
      const currentMinutes = hNow * 60 + mNow;
      const pulangMinutes = hPulang * 60 + mPulang;

      if (pulangMinutes > 0 && currentMinutes >= pulangMinutes) {
        return { success: false, message: "Gagal Scan: Sudah jam pulang sekolah." };
      }

      // Calculate lateness based on school start time
      let cutoffMasuk = "07:15";
      if (settingsDoc.exists()) {
        const s = settingsDoc.data() as DaySetting;
        if (s.masuk) cutoffMasuk = s.masuk;
      }
      const [hMasuk, mMasuk] = cutoffMasuk.split(':').map(Number);
      const masukMinutes = hMasuk * 60 + mMasuk;
      const late = currentMinutes - masukMinutes;
      const terlambat = late > 0 ? late : 0;

      // Find the subject for this teacher and class
      const schedQuery = query(
        collection(db, 'teachingSchedules'), 
        where('nip', '==', nip), 
        where('kelas', '==', className)
      );
      const schedSnap = await getDocs(schedQuery);
      let mapel = '-';
      if (!schedSnap.empty) {
        mapel = schedSnap.docs[0].data().mapel || '-';
      }

      const id = `${nip}-${className}-${now.getTime()}`;
      const record: TeacherAttendance = {
        id,
        nip,
        nama: teacher.nama,
        kelas: className,
        tanggal,
        jam,
        terlambat,
        mapel
      };

      await setDoc(doc(db, 'teacherAttendance', id), record);
      return { success: true, message: `Berhasil Scan Kelas: ${className}` };
    } catch (e) {
      return handleFirestoreError(e, OperationType.WRITE, 'class-scan');
    }
  },

  // Teaching Schedules
  saveJadwalMengajar: async (data: TeachingSchedule) => {
    try {
      const id = data.id || `${data.nip}-${data.kelas}`;
      await setDoc(doc(db, 'teachingSchedules', id), { ...data, id });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `teachingSchedules/${data.id}`);
      throw e;
    }
  },

  hapusJadwalMengajar: async (id: string) => {
    try {
      await deleteDoc(doc(db, 'teachingSchedules', id));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `teachingSchedules/${id}`);
      throw e;
    }
  },

  saveTeachingBatch: async (nip: string, namaGuru: string, schedules: {kelas: string, target: number, hari: string, jps?: number[]}[], mapel: string) => {
    try {
      for (const item of schedules) {
        const id = `${nip}-${item.kelas}-${item.hari}`;
        await setDoc(doc(db, 'teachingSchedules', id), {
          id, nip, namaGuru, kelas: item.kelas, targetPertemuan: item.target, mapel, hari: item.hari, jps: item.jps || []
        });
      }
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'teachingScheduleBatch');
      throw e;
    }
  },

  // Settings
  initializeSettings: async () => {
    try {
      const dayMap = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
      for (const hari of dayMap) {
        const d = doc(db, 'settings', hari);
        const snap = await getDoc(d);
        if (!snap.exists()) {
          await setDoc(d, { hari, masuk: "07:15", pulang: "14:00" });
        }
      }
      // Also add Sunday as a placeholder but marked as libur in UI
      const dSun = doc(db, 'settings', 'Minggu');
      const snapSun = await getDoc(dSun);
      if (!snapSun.exists()) {
        await setDoc(dSun, { hari: 'Minggu', masuk: "", pulang: "" });
      }
      return true;
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'settings-init');
      return false;
    }
  },

  savePengaturanHari: async (hari: string, masuk: string, pulang: string, activeJps?: number[], reasonInactive?: string, targetDate?: string) => {
    try {
      const updateData: any = { hari, masuk, pulang };
      if (activeJps !== undefined) {
        updateData.activeJps = activeJps;
      }
      if (reasonInactive !== undefined) {
        updateData.reasonInactive = reasonInactive;
      }
      if (targetDate !== undefined) {
        updateData.targetDate = targetDate;
      }
      await setDoc(doc(db, 'settings', hari), updateData, { merge: true });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `settings/${hari}`);
      throw e;
    }
  },

  hapusPengaturanHariKhusus: async (hari: string) => {
    try {
      await updateDoc(doc(db, 'settings', hari), {
        activeJps: deleteField(),
        reasonInactive: deleteField(),
        targetDate: deleteField()
      });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `settings/${hari}`);
      throw e;
    }
  },

  saveLiburAgenda: async (tanggal: string, keterangan: string) => {
    try {
      await setDoc(doc(db, 'holidays', tanggal), { tanggal, keterangan });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `holidays/${tanggal}`);
      throw e;
    }
  },

  hapusLiburAgenda: async (tanggal: string) => {
    try {
      await deleteDoc(doc(db, 'holidays', tanggal));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `holidays/${tanggal}`);
      throw e;
    }
  },
  saveAppConfig: async (data: any) => {
    try {
      await setDoc(doc(db, 'appConfig', 'general'), data, { merge: true });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'appConfig/general');
      throw e;
    }
  },

  generateOneTimeToken: async (session: any): Promise<string> => {
    try {
      if (!session || !session.uid) {
        throw new Error("No active session found.");
      }
      // Generate a secure, hard-to-guess 32-character token string
      const array = new Uint32Array(8);
      window.crypto.getRandomValues(array);
      const token = Array.from(array, dec => dec.toString(16).padStart(8, '0')).join('');

      const tokenRef = doc(db, 'oneTimeTokens', token);
      const now = Date.now();
      const expiresAt = now + 60 * 1000; // Token valid for 60 seconds (1 minute)

      await setDoc(tokenRef, {
        token,
        uid: session.uid,
        name: session.name || '',
        role: session.role || 'Guru',
        createdAt: now,
        expiresAt,
        used: false
      });

      return token;
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, 'oneTimeTokens');
      throw e;
    }
  },


};
