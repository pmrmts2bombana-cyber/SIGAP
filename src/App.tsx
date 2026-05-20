/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { 
  User, School, QrCode, UserPlus, GraduationCap, LayoutGrid, Settings, LogOut, 
  ChevronRight, ChevronLeft, ClipboardList, FileBarChart, Table as TableIcon, Search, Plus, 
  RefreshCcw, Printer, Download, Eye, EyeOff, Calendar, Clock, Trash2, Edit, Save,
  ArrowLeft, Upload, FileSpreadsheet, BarChart3, Info, CheckCircle2, XCircle, AlertTriangle, AlertCircle,
  Maximize2, CreditCard, Award, ExternalLink, ShieldCheck, Sparkles, Heart
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { firestoreService } from './services/firestoreService';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as htmlToImage from 'html-to-image';
import { QRCodeSVG } from 'qrcode.react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Legend, AreaChart, Area
} from 'recharts';
import { 
  UserSession, Student, Teacher, Classroom, Attendance, DashboardStats, 
  Holiday, DaySetting, TeacherAttendance, AppConfig, TeachingSchedule
} from './types';
import { auth, db } from './firebase';
import { doc, getDoc, setDoc, onSnapshot, collection, query, orderBy, where, getDocs } from 'firebase/firestore';
import { handleFirestoreError, OperationType, onQuotaExceeded } from './lib/firebaseUtils';

const formatIndoDate = (dateStr: string) => {
  if (!dateStr) return '-';
  try {
    // Handle YYYY-MM-DD to avoid timezone shifting
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const [y, m, d] = dateStr.split('-').map(Number);
      const date = new Date(y, m - 1, d);
      return new Intl.DateTimeFormat('id-ID', { 
        weekday: 'long', 
        day: 'numeric', 
        month: 'long', 
        year: 'numeric' 
      }).format(date);
    }
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return new Intl.DateTimeFormat('id-ID', { 
      weekday: 'long', 
      day: 'numeric', 
      month: 'long', 
      year: 'numeric' 
      }).format(d);
  } catch (e) {
    return dateStr;
  }
};

const countDaysInMonth = (year: number, month: number, dayName: string, holidays: Holiday[], upToDay?: number) => {
  const dayIndex = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"].indexOf(dayName);
  if (dayIndex === -1) return 0;
  
  let count = 0;
  let date = new Date(year, month, 1);
  const lastDay = upToDay || new Date(year, month + 1, 0).getDate();
  
  while (date.getMonth() === month && date.getDate() <= lastDay) {
    if (date.getDay() === dayIndex) {
      // Manual formatting YYYY-MM-DD in local time to match holiday data precisely
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      const dateStr = `${y}-${m}-${d}`;
      
      const isHoliday = (holidays || []).some(h => h.tanggal === dateStr);
      if (!isHoliday) {
        count++;
      }
    }
    date.setDate(date.getDate() + 1);
  }
  return count;
};

const getEffectiveDays = (year: number, month: number, holidays: Holiday[], settings: DaySetting[], upToDay?: number) => {
  const dayNames = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const days: string[] = [];
  const lastDay = upToDay || new Date(year, month + 1, 0).getDate();
  
  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month, d);
    const dayName = dayNames[date.getDay()];
    
    // 1. Skip Sunday
    if (dayName === "Minggu") continue;
    
    // 2. Skip Holidays
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    if (holidays.some(h => h.tanggal === dateStr)) continue;
    
    // 3. Skip if day setting is "Libur" (empty masuk or pulang)
    const setting = settings.find(s => s.hari === dayName);
    if (setting && (!setting.masuk || !setting.pulang)) continue;
    
    days.push(dateStr);
  }
  return days;
};

const AttendanceChart = ({ data }: { data: any[] }) => (
  <div className="h-64 w-full">
    <ResponsiveContainer width="100%" height="100%">
      <AreaChart data={data} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
        <defs>
          <linearGradient id="colorHadir" x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor="#16a34a" stopOpacity={0.8}/>
            <stop offset="95%" stopColor="#16a34a" stopOpacity={0}/>
          </linearGradient>
          <linearGradient id="colorAlfa" x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor="#dc2626" stopOpacity={0.8}/>
            <stop offset="95%" stopColor="#dc2626" stopOpacity={0}/>
          </linearGradient>
        </defs>
        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
        <XAxis dataKey="name" fontSize={10} axisLine={false} tickLine={false} />
        <YAxis fontSize={10} axisLine={false} tickLine={false} />
        <Tooltip 
          contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1)' }}
          labelStyle={{ fontWeight: 'bold' }}
        />
        <Legend iconType="circle" />
        <Area type="monotone" dataKey="hadir" stroke="#16a34a" fillOpacity={1} fill="url(#colorHadir)" />
        <Area type="monotone" dataKey="alfa" stroke="#dc2626" fillOpacity={1} fill="url(#colorAlfa)" />
      </AreaChart>
    </ResponsiveContainer>
  </div>
);


const formatMinutes = (minutes: number) => {
  if (minutes <= 0) return 'Tepat Waktu';
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h > 0) return `${h} jam ${m} menit`;
  return `${m} menit`;
};

export default function App() {
  const [students, setStudents] = useState<Student[]>([]);
  const [teachers, setTeachers] = useState<any[]>([]);
  const [classrooms, setClassrooms] = useState<Classroom[]>([]);
  const [attendance, setAttendance] = useState<Attendance[]>([]);
  const [teacherAttendance, setTeacherAttendance] = useState<TeacherAttendance[]>([]);
  const [settings, setSettings] = useState<DaySetting[]>([]);
  const [teachingSchedules, setTeachingSchedules] = useState<TeachingSchedule[]>([]);
  const [holidays, setHolidays] = useState<Holiday[]>([]);
  const [appConfig, setAppConfig] = useState<AppConfig>({});

  const [searchTermGuru, setSearchTermGuru] = useState('');
  const [searchTermSiswa, setSearchTermSiswa] = useState('');
  const [searchTermKelas, setSearchTermKelas] = useState('');
  const [searchTermJadwal, setSearchTermJadwal] = useState('');
  const [refreshing, setRefreshing] = useState(false);
  const [session, setSession] = useState<UserSession | null>(null);
  const [studentProfileClassFilter, setStudentProfileClassFilter] = useState('');

  const [activePanel, setActivePanel] = useState('home');

  const [analysisClass, setAnalysisClass] = useState('');
  const [appLogoInput, setAppLogoInput] = useState('');

  // Force analysisClass for Wali Kelas
  useEffect(() => {
    if (activePanel === 'analisis' && session?.role === 'Guru' && session?.isWali && session?.kelas) {
      setAnalysisClass(session.kelas);
    }
  }, [activePanel, session]);

  const effectiveDaysThisMonth = useMemo(() => {
    const now = new Date();
    return getEffectiveDays(now.getFullYear(), now.getMonth(), holidays, settings, now.getDate());
  }, [holidays, settings]);

  const DEFAULT_SUBJECTS = useMemo(() => [
    'Fiqih', "Alqur'an Hadis", 'Akida Akhlak', 'SKI', 'Bhs. Indonesia', 
    'Bhs. Arab', 'Bhs. Inggris', 'Matematika', 'IPA', 'IPS', 
    'Penjas', 'PPKN', 'Informatika', 'Koding K.A', 'Mulok'
  ], []);

  const [loading, setLoading] = useState(false);
  const [firebaseConnected, setFirebaseConnected] = useState(false);
  const [quotaExceeded, setQuotaExceeded] = useState(false);
  const [dismissQuotaNotice, setDismissQuotaNotice] = useState(false);

  const [filterClassAbsensi, setFilterClassAbsensi] = useState('');
  const [filterNamaAbsensi, setFilterNamaAbsensi] = useState('');
  const [filterAdminSiswaClass, setFilterAdminSiswaClass] = useState('');
  const [siswaDashboardDate, setSiswaDashboardDate] = useState(new Date().toISOString().split('T')[0]);

  const [pagination, setPagination] = useState({
    guru: 0,
    siswa: 0,
    kelas: 0,
    jadwal: 0,
    absensiSiswa: 0,
    absensiGuru: 0,
    rekapMapel: 0
  });
  const [pageSize, setPageSize] = useState(10);
  const getLocalISO = () => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  };

  const [currentDate, setCurrentDate] = useState(getLocalISO);

  // Timer to ensure "Today" updates if page is left open overnight
  useEffect(() => {
    const timer = setInterval(() => {
      const live = getLocalISO();
      if (live !== currentDate) setCurrentDate(live);
    }, 30000); // Check every 30 seconds
    return () => clearInterval(timer);
  }, [currentDate]);

  const [rekapMapelFilter, setRekapMapelFilter] = useState({
    nip: '',
    tanggal: getLocalISO(),
    mapel: ''
  });

  const [confirmModal, setConfirmModal] = useState<{ show: boolean, title: string, message: string, entityName?: string, onConfirm: () => void } | null>(null);

  useEffect(() => {
    onQuotaExceeded((isExceeded) => {
      setQuotaExceeded(isExceeded);
    });

    const initApp = async () => {
      setLoading(true);
      try {
        await firestoreService.bootstrapAdmin();
        setFirebaseConnected(true);
      } catch (e) {
        console.error("Firebase connection error:", e);
        setFirebaseConnected(false);
      } finally {
        setLoading(false);
      }
    };
    initApp();
  }, []);

  // Load general configuration as long as database is connected (needed for login page branding)
  useEffect(() => {
    if (!firebaseConnected) return;

    const unsubAppConfig = onSnapshot(doc(db, 'appConfig', 'general'), (snap) => {
      if (snap.exists()) {
        const data = snap.data() as AppConfig;
        setAppConfig(data);
      }
    }, (error) => handleFirestoreError(error, OperationType.GET, 'appConfig/general'));

    return () => {
      unsubAppConfig();
    };
  }, [firebaseConnected]);

  // Load transaction and master data ONLY when a user session is active (logged in)
  useEffect(() => {
    if (!firebaseConnected || !session) {
      // Clear data states when logged out to avoid stale/previous user data and stop memory/network consumption
      setAttendance([]);
      setTeacherAttendance([]);
      setStudents([]);
      setTeachers([]);
      setClassrooms([]);
      setSettings([]);
      setHolidays([]);
      setTeachingSchedules([]);
      return;
    }

    // Real-time synchronization for critical data (restricted to last 30 days)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const dateLimit = thirtyDaysAgo.toISOString().split('T')[0];

    const role = session.role;
    const uid = session.uid;
    const kelasWali = session.kelas;
    const isWali = session.isWali;
    const jabatan = session.jabatan || '';

    // Deterministic state check for school leadership
    const isLeadership = role === 'Guru' && (jabatan === 'Kamad' || jabatan === 'Wakamad');

    let unsubAttendance = () => {};
    let unsubTeacherAttendance = () => {};
    let unsubStudents = () => {};
    let unsubTeachers = () => {};
    let unsubClassrooms = () => {};
    let unsubSettings = () => {};
    let unsubHolidays = () => {};
    let unsubSchedules = () => {};

    // 1. Settings & Holidays are lightweight config models, fetch safely
    unsubSettings = onSnapshot(collection(db, 'settings'), (snap) => {
      setSettings(snap.docs.map(d => d.data() as DaySetting));
    }, (error) => handleFirestoreError(error, OperationType.GET, 'settings'));

    unsubHolidays = onSnapshot(collection(db, 'holidays'), (snap) => {
      const data = snap.docs.map(d => d.data() as Holiday);
      setHolidays(data.sort((a, b) => b.tanggal.localeCompare(a.tanggal)));
    }, (error) => handleFirestoreError(error, OperationType.GET, 'holidays'));

    // 2. Performance & Quota Optimization dynamically applied
    if (role === 'Siswa') {
      // Student accounts only fetch their specific account profile, and their specific attendance logs
      unsubStudents = onSnapshot(doc(db, 'students', uid), (snap) => {
        if (snap.exists()) {
          setStudents([{ nisn: snap.id, ...snap.data() } as Student]);
        } else {
          setStudents([]);
        }
      }, (error) => handleFirestoreError(error, OperationType.GET, `students/${uid}`));

      unsubAttendance = onSnapshot(
        query(
          collection(db, 'attendance'),
          where('nisn', '==', uid)
        ),
        (snap) => {
          setAttendance(snap.docs.map(d => d.data() as Attendance));
        },
        (error) => handleFirestoreError(error, OperationType.GET, 'attendance')
      );

      // Nullify heavy teacher/classroom lists for student sessions
      setTeachers([]);
      setClassrooms([]);
      setTeacherAttendance([]);
      setTeachingSchedules([]);

    } else if (role === 'Guru') {
      if (isLeadership) {
        // Leadership needs wider scope
        unsubAttendance = onSnapshot(
          query(
            collection(db, 'attendance'), 
            where('tanggal', '>=', dateLimit),
            orderBy('tanggal', 'desc')
          ), 
          (snap) => {
            setAttendance(snap.docs.map(d => d.data() as Attendance));
          }, 
          (error) => handleFirestoreError(error, OperationType.GET, 'attendance')
        );

        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, 'teacherAttendance'), 
            where('tanggal', '>=', dateLimit),
            orderBy('tanggal', 'desc')
          ), 
          (snap) => {
            setTeacherAttendance(snap.docs.map(d => d.data() as TeacherAttendance));
          }, 
          (error) => handleFirestoreError(error, OperationType.GET, 'teacherAttendance')
        );

        unsubStudents = onSnapshot(collection(db, 'students'), (snap) => {
          const data = snap.docs.map(d => ({ nisn: d.id, ...d.data() } as Student));
          setStudents(data.sort((a, b) => a.nama.localeCompare(b.nama)));
        }, (error) => handleFirestoreError(error, OperationType.GET, 'students'));

        unsubTeachers = onSnapshot(collection(db, 'teachers'), (snap) => {
          const data = snap.docs.map(d => ({ nip: d.id, ...d.data() } as Teacher));
          setTeachers(data.sort((a, b) => a.nama.localeCompare(b.nama)));
        }, (error) => handleFirestoreError(error, OperationType.GET, 'teachers'));

        unsubClassrooms = onSnapshot(collection(db, 'classrooms'), (snap) => {
          const data = snap.docs.map(d => d.data() as Classroom);
          setClassrooms(data.sort((a, b) => a.nama.localeCompare(b.nama)));
        }, (error) => handleFirestoreError(error, OperationType.GET, 'classrooms'));

        unsubSchedules = onSnapshot(collection(db, 'teachingSchedules'), (snap) => {
          setTeachingSchedules(snap.docs.map(d => d.data() as TeachingSchedule));
        }, (error) => handleFirestoreError(error, OperationType.GET, 'teachingSchedules'));

      } else if (isWali && kelasWali) {
        // Homeroom Teacher: scope student records and attendance within their class
        unsubStudents = onSnapshot(
          query(
            collection(db, 'students'),
            where('kelas', '==', kelasWali)
          ),
          (snap) => {
            const data = snap.docs.map(d => ({ nisn: d.id, ...d.data() } as Student));
            setStudents(data.sort((a, b) => a.nama.localeCompare(b.nama)));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'students')
        );

        unsubAttendance = onSnapshot(
          query(
            collection(db, 'attendance'),
            where('kelas', '==', kelasWali),
            where('tanggal', '>=', dateLimit)
          ),
          (snap) => {
            setAttendance(snap.docs.map(d => d.data() as Attendance));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'attendance')
        );

        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, 'teacherAttendance'),
            where('nip', '==', uid),
            where('tanggal', '>=', dateLimit)
          ),
          (snap) => {
            setTeacherAttendance(snap.docs.map(d => d.data() as TeacherAttendance));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'teacherAttendance')
        );

        unsubSchedules = onSnapshot(
          query(
            collection(db, 'teachingSchedules'),
            where('nip', '==', uid)
          ),
          (snap) => {
            setTeachingSchedules(snap.docs.map(d => d.data() as TeachingSchedule));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'teachingSchedules')
        );

        unsubTeachers = onSnapshot(doc(db, 'teachers', uid), (snap) => {
          if (snap.exists()) {
            setTeachers([{ nip: snap.id, ...snap.data() } as Teacher]);
          } else {
            setTeachers([]);
          }
        }, (error) => handleFirestoreError(error, OperationType.GET, `teachers/${uid}`));

        unsubClassrooms = onSnapshot(collection(db, 'classrooms'), (snap) => {
          const data = snap.docs.map(d => d.data() as Classroom);
          setClassrooms(data.sort((a, b) => a.nama.localeCompare(b.nama)));
        }, (error) => handleFirestoreError(error, OperationType.GET, 'classrooms'));

      } else {
        // Regular Subject Teacher: scoped resources
        unsubTeachers = onSnapshot(doc(db, 'teachers', uid), (snap) => {
          if (snap.exists()) {
            setTeachers([{ nip: snap.id, ...snap.data() } as Teacher]);
          } else {
            setTeachers([]);
          }
        }, (error) => handleFirestoreError(error, OperationType.GET, `teachers/${uid}`));

        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, 'teacherAttendance'),
            where('nip', '==', uid),
            where('tanggal', '>=', dateLimit)
          ),
          (snap) => {
            setTeacherAttendance(snap.docs.map(d => d.data() as TeacherAttendance));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'teacherAttendance')
        );

        unsubSchedules = onSnapshot(
          query(
            collection(db, 'teachingSchedules'),
            where('nip', '==', uid)
          ),
          (snap) => {
            setTeachingSchedules(snap.docs.map(d => d.data() as TeachingSchedule));
          },
          (error) => handleFirestoreError(error, OperationType.GET, 'teachingSchedules')
        );

        setStudents([]);
        setClassrooms([]);
        setAttendance([]);
      }

    } else if (role === 'Admin') {
      unsubAttendance = onSnapshot(
        query(
          collection(db, 'attendance'), 
          where('tanggal', '>=', dateLimit),
          orderBy('tanggal', 'desc')
        ), 
        (snap) => {
          setAttendance(snap.docs.map(d => d.data() as Attendance));
        }, 
        (error) => handleFirestoreError(error, OperationType.GET, 'attendance')
      );

      unsubTeacherAttendance = onSnapshot(
        query(
          collection(db, 'teacherAttendance'), 
          where('tanggal', '>=', dateLimit),
          orderBy('tanggal', 'desc')
        ), 
        (snap) => {
          setTeacherAttendance(snap.docs.map(d => d.data() as TeacherAttendance));
        }, 
        (error) => handleFirestoreError(error, OperationType.GET, 'teacherAttendance')
      );

      unsubStudents = onSnapshot(collection(db, 'students'), (snap) => {
        const data = snap.docs.map(d => ({ nisn: d.id, ...d.data() } as Student));
        setStudents(data.sort((a, b) => a.nama.localeCompare(b.nama)));
      }, (error) => handleFirestoreError(error, OperationType.GET, 'students'));

      unsubTeachers = onSnapshot(collection(db, 'teachers'), (snap) => {
        const data = snap.docs.map(d => ({ nip: d.id, ...d.data() } as Teacher));
        setTeachers(data.sort((a, b) => a.nama.localeCompare(b.nama)));
      }, (error) => handleFirestoreError(error, OperationType.GET, 'teachers'));

      unsubClassrooms = onSnapshot(collection(db, 'classrooms'), (snap) => {
        const data = snap.docs.map(d => d.data() as Classroom);
        setClassrooms(data.sort((a, b) => a.nama.localeCompare(b.nama)));
      }, (error) => handleFirestoreError(error, OperationType.GET, 'classrooms'));

      unsubSchedules = onSnapshot(collection(db, 'teachingSchedules'), (snap) => {
        setTeachingSchedules(snap.docs.map(d => d.data() as TeachingSchedule));
      }, (error) => handleFirestoreError(error, OperationType.GET, 'teachingSchedules'));
    }

    return () => {
      unsubAttendance();
      unsubTeacherAttendance();
      unsubStudents();
      unsubTeachers();
      unsubClassrooms();
      unsubSettings();
      unsubHolidays();
      unsubSchedules();
    };
  }, [firebaseConnected, session]);

  const [selectedTeacherNipForCapaian, setSelectedTeacherNipForCapaian] = useState<string | null>(null);

  const stats: DashboardStats & { siswaL?: number, siswaP?: number, sakitCount?: number, izinCount?: number, alfaCount?: number } = useMemo(() => {
    const now = new Date();
    const today = currentDate; // Use the state-synced date
    const dayMap = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const dayName = dayMap[now.getDay()];
    
    const isHoliday = holidays.some(h => h.tanggal === today) || dayName === 'Minggu';

    const todayAttendance = attendance.filter(a => a.tanggal === today);
    const hadir = todayAttendance.filter(a => a.status === 'Hadir').length;
    const lambat = todayAttendance.filter(a => a.terlambat > 0).length;
    const sakit = todayAttendance.filter(a => a.status === 'Sakit').length;
    const izin = todayAttendance.filter(a => a.status === 'Izin').length;
    
    // ALFA Logic: 
    // 1. Those recorded as 'Alfa' in DB (e.g. check-in after school hours)
    const recordedAlfa = todayAttendance.filter(a => a.status === 'Alfa').length;
    
    // 2. Those who haven't checked in at all today
    const attendedNisns = new Set(todayAttendance.map(a => a.nisn));
    const notYetCheckedIn = students.length - attendedNisns.size;

    const alfaTotal = isHoliday ? 0 : (recordedAlfa + notYetCheckedIn);

    const siswaL = students.filter(s => s.jenisKelamin === 'L' || s.jenisKelamin === 'Laki-Laki').length;
    const siswaP = students.filter(s => s.jenisKelamin === 'P' || s.jenisKelamin === 'Perempuan').length;

    // Chart data (last 7 days)
    const chartData = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const dateStr = d.toISOString().split('T')[0];
      const dayData = attendance.filter(a => a.tanggal === dateStr);
      const dayHadir = dayData.filter(a => a.status === 'Hadir').length;
      
      // Daily Alfa calculation for chart too?
      // For chart, we use the same "Missing" logic for realistic data
      const dDayMap = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
      const dDayName = dDayMap[d.getDay()];
      const dIsHoliday = holidays.some(h => h.tanggal === dateStr) || dDayName === 'Minggu';
      
      const dAttendedNisns = new Set(dayData.map(a => a.nisn));
      const dNotYet = students.length - dAttendedNisns.size;
      const dRecordedAlfa = dayData.filter(a => a.status === 'Alfa').length;
      const dAlfaTotal = dIsHoliday ? 0 : (dNotYet + dRecordedAlfa);

      chartData.push({
        name: d.toLocaleDateString('id-ID', { weekday: 'short' }),
        hadir: dayHadir,
        alfa: dAlfaTotal
      });
    }

    return {
      siswaCount: students.length,
      guruCount: teachers.length,
      hadirHariIni: hadir,
      terlambatHariIni: lambat,
      sakitCount: sakit,
      izinCount: izin,
      alfaCount: alfaTotal,
      chartData,
      siswaL,
      siswaP
    };
  }, [students, teachers, attendance, settings, holidays]);

  // Reporting states
  const [rekapFilter, setRekapFilter] = useState({ bulan: new Date().toISOString().slice(0, 7), kelas: '', type: 'Siswa' });
  
  useEffect(() => {
    if (session?.kelas && (session?.role === 'Guru Wali Kelas' || session?.isWali)) {
      setRekapFilter(prev => ({ ...prev, kelas: session.kelas }));
      setFilterClassAbsensi(session.kelas);
    }
  }, [session]);

  const [editingProfileData, setEditingProfileData] = useState<any>(null);
  const [isUpdatingProfile, setIsUpdatingProfile] = useState(false);

  // Optimized Maps for Rekap
  const monthAttendanceMap = useMemo(() => {
    const map: Record<string, Attendance[]> = {};
    attendance.forEach(a => {
      if (a.tanggal.startsWith(rekapFilter.bulan)) {
        if (!map[a.nisn]) map[a.nisn] = [];
        map[a.nisn].push(a);
      }
    });
    return map;
  }, [attendance, rekapFilter.bulan]);

  const teacherAttendanceMap = useMemo(() => {
    const map: Record<string, TeacherAttendance[]> = {};
    teacherAttendance.forEach(ta => {
      if (ta.tanggal.startsWith(rekapFilter.bulan)) {
        if (!map[ta.nip]) map[ta.nip] = [];
        map[ta.nip].push(ta);
      }
    });
    return map;
  }, [teacherAttendance, rekapFilter.bulan]);

  const handleUpdateProfile = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingProfileData) return;
    setIsUpdatingProfile(true);
    try {
      await firestoreService.saveGuru(editingProfileData);
      alert("Profil berhasil diperbarui.");
      // Refresh teacher state if needed
    } catch (err) {
      alert("Gagal memperbarui profil.");
    } finally {
      setIsUpdatingProfile(false);
    }
  };

  const [searchTeacherReport, setSearchTeacherReport] = useState('');





  const renderRoster = () => {
    // Generate URL using session data. session.uid is guaranteed after login.
    // We add a timestamp to prevent caching and ensure a fresh trigger.
    const currentUserToken = session?.uid || 'guest';
    const currentUserName = session?.name || 'User';
    const currentUserRole = session?.role || 'Guru';
    const rosterUrl = `https://criet-roster.vercel.app/?token=${currentUserToken}&role=${currentUserRole}&name=${encodeURIComponent(currentUserName)}&ts=${Date.now()}`;

    return (
      <div className="flex flex-col items-center justify-center min-h-[500px] bg-white rounded-[3rem] p-12 shadow-2xl border border-gray-100 text-center relative overflow-hidden">
        {/* Background Decorative Elements */}
        <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-green-400 via-emerald-500 to-teal-600"></div>
        <div className="absolute -top-24 -right-24 w-64 h-64 bg-green-50 rounded-full blur-3xl opacity-50"></div>
        <div className="absolute -bottom-24 -left-24 w-64 h-64 bg-blue-50 rounded-full blur-3xl opacity-50"></div>

        <div className="relative z-10 max-w-lg mx-auto">
          <motion.div 
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="w-24 h-24 bg-gradient-to-br from-green-500 to-emerald-600 rounded-[2.5rem] flex items-center justify-center mb-8 mx-auto shadow-lg shadow-green-200"
          >
            <Sparkles className="text-white" size={44} />
          </motion.div>

          <h3 className="text-3xl font-black text-gray-900 uppercase tracking-tighter mb-4">
            Criet Roster <span className="text-green-600">AI</span>
          </h3>
          
          <p className="text-sm font-bold text-gray-500 mb-10 leading-relaxed">
            Hasilkan jadwal mengajar (roster) sekolah secara otomatis dan efisien menggunakan teknologi AI. 
            Klik tombol di bawah untuk diarahkan ke platform pembuatan roster.
          </p>

          <div className="space-y-4">
            <motion.a 
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              href={rosterUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full bg-zinc-900 text-white rounded-[2rem] py-6 px-8 font-black uppercase tracking-widest shadow-xl hover:bg-zinc-800 transition-all flex items-center justify-center gap-4 group cursor-pointer no-underline"
            >
              <ExternalLink size={24} className="text-green-400 group-hover:rotate-12 transition-transform" />
              Buka Dashboard Roster
            </motion.a>

            <div className="flex items-center justify-center gap-2 py-3 px-6 bg-zinc-50 rounded-2xl border border-zinc-100">
              <ShieldCheck size={16} className="text-green-600" />
              <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                Akses Terproteksi (Sesi Terintegrasi)
              </span>
            </div>
          </div>

          <p className="mt-8 text-[11px] font-bold text-zinc-400 italic">
            *Anda akan diarahkan ke tab baru secara aman. Gunakan akun SIGAP Anda untuk sinkronisasi data.
          </p>
        </div>
      </div>
    );
  };

  const renderKBC = () => {
    return (
      <div className="flex flex-col items-center justify-center min-h-[500px] bg-white rounded-[3rem] p-12 shadow-2xl border border-rose-50 text-center relative overflow-hidden">
        {/* Background Decorative Elements */}
        <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-rose-400 via-pink-500 to-red-600"></div>
        <div className="absolute -top-24 -right-24 w-64 h-64 bg-rose-50 rounded-full blur-3xl opacity-50"></div>
        <div className="absolute -bottom-24 -left-24 w-64 h-64 bg-pink-50 rounded-full blur-3xl opacity-50"></div>

        <div className="relative z-10 max-w-lg mx-auto">
          <motion.div 
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="w-24 h-24 bg-gradient-to-br from-rose-500 to-pink-600 rounded-[2.5rem] flex items-center justify-center mb-8 mx-auto shadow-lg shadow-rose-200"
          >
            <Heart className="text-white fill-white/20" size={44} />
          </motion.div>

          <h3 className="text-3xl font-black text-gray-900 uppercase tracking-tighter mb-4">
            SIGAP <span className="text-rose-600">KBC</span>
          </h3>
          
          <p className="text-sm font-bold text-gray-500 mb-10 leading-relaxed">
            Aplikasi administrasi guru berbasis Kurikulum Berbasis Cinta (KBC). 
            Klik tombol di bawah untuk diarahkan ke platform aplikasi KBC.
          </p>

          <div className="space-y-4">
            <motion.a 
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              href="https://kbc-ecru.vercel.app/"
              target="_blank"
              rel="noopener noreferrer"
              className="w-full bg-zinc-900 text-white rounded-[2rem] py-6 px-8 font-black uppercase tracking-widest shadow-xl hover:bg-zinc-800 transition-all flex items-center justify-center gap-4 group cursor-pointer no-underline"
            >
              <ExternalLink size={24} className="text-rose-400 group-hover:rotate-12 transition-transform" />
              Buka Aplikasi KBC
            </motion.a>

            <div className="flex items-center justify-center gap-2 py-3 px-6 bg-zinc-50 rounded-2xl border border-zinc-100">
              <ShieldCheck size={16} className="text-green-600" />
              <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                Akses Terproteksi (Sesi Aktif)
              </span>
            </div>
          </div>

          <p className="mt-8 text-[11px] font-bold text-zinc-400 italic">
            *Anda akan diarahkan ke tab baru secara aman. Pastikan Anda tetap login di sistem SIGAP.
          </p>
        </div>
      </div>
    );
  };

  const renderLaporanCapaianGuru = () => {
    const selectedTeacher = teachers.find(t => t.nip === selectedTeacherNipForCapaian);

    return (
      <div className="space-y-6">
        {!selectedTeacher ? (
          <div className="bg-white rounded-[2rem] p-8 shadow-xl border border-gray-100">
            <div className="flex flex-col md:flex-row justify-between md:items-center gap-4 mb-8">
              <h3 className="text-xl font-black text-green-950 uppercase tracking-widest flex items-center gap-3">
                <User className="text-green-600" /> Pilih Guru
              </h3>
              <div className="relative w-full md:w-64">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300" size={16} />
                <input 
                  type="text" 
                  placeholder="Cari nama guru..."
                  value={searchTeacherReport}
                  onChange={e => setSearchTeacherReport(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-gray-50 border-0 rounded-xl text-sm font-bold focus:ring-2 focus:ring-green-500"
                />
              </div>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {teachers
                .filter(t => t.nama.toLowerCase().includes(searchTeacherReport.toLowerCase()))
                .map((t) => (
                <button
                  key={t.nip}
                  onClick={() => setSelectedTeacherNipForCapaian(t.nip)}
                  className="bg-gray-50 hover:bg-green-50 p-6 rounded-[2rem] border border-gray-100 text-left transition-all hover:shadow-md flex items-center gap-4 group"
                >
                  <div className="w-12 h-12 rounded-2xl bg-zinc-900 flex items-center justify-center text-white overflow-hidden shadow-sm">
                    {t.foto ? <img src={t.foto} className="w-full h-full object-cover" alt={t.nama} /> : <User size={20} />}
                  </div>
                  <div className="flex-1">
                    <p className="text-sm font-black text-zinc-900 group-hover:text-green-900 transition-colors">{t.nama}</p>
                    <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">{t.jabatan || 'Guru'}</p>
                  </div>
                  <ChevronRight size={16} className="text-zinc-300 group-hover:text-green-600" />
                </button>
              ))}
            </div>
          </div>
        ) : (
          <motion.div initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} className="space-y-6 pb-20">
            <div className="flex items-center gap-4">
               <button 
                onClick={() => setSelectedTeacherNipForCapaian(null)}
                className="bg-white p-3 rounded-2xl shadow-sm border border-gray-100 text-zinc-400 hover:text-green-900 transition-colors"
               >
                 <ArrowLeft size={20} />
               </button>
               <div>
                  <h3 className="font-black text-xl text-green-950">{selectedTeacher.nama}</h3>
                  <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">Laporan Capaian Mengajar</p>
               </div>
            </div>

            <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-xl border border-gray-100">
               <h3 className="text-lg font-black text-green-950 mb-6 flex items-center gap-3 uppercase tracking-widest">
                 <Calendar className="text-green-600" /> Jam Mengajar & Laporan Capaian
               </h3>
               <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {(() => {
                    const now = new Date();
                    const year = now.getFullYear();
                    const month = now.getMonth();
                    const currentMonthPrefix = now.toISOString().slice(0, 7);
                    
                    const mySchedules = teachingSchedules.filter(ts => ts.nip === selectedTeacher.nip);

                    // Group schedules by Class and Mapel
                    const groupedSchedules: {[key: string]: {
                      kelas: string,
                      mapel: string,
                      days: string[],
                      sessions: any[]
                    }} = {};

                    mySchedules.forEach(s => {
                      const key = `${s.kelas}-${s.mapel || 'unnamed'}`;
                      if (!groupedSchedules[key]) {
                        groupedSchedules[key] = {
                          kelas: s.kelas,
                          mapel: s.mapel || '-',
                          days: [],
                          sessions: []
                        };
                      }
                      if (!groupedSchedules[key].days.includes(s.hari)) {
                        groupedSchedules[key].days.push(s.hari);
                      }
                      groupedSchedules[key].sessions.push(s);
                    });

                    const sortedGroups = Object.values(groupedSchedules);

                    return sortedGroups.length > 0 ? sortedGroups.map((group, idx) => {
                      let totalMonthlyTarget = 0;
                      let totalActualSessions = 0;

                      group.sessions.forEach(s => {
                        const targetOccurrences = countDaysInMonth(year, month, s.hari, holidays);
                        totalMonthlyTarget += targetOccurrences * (Number(s.targetPertemuan) || 0);

                        const actual = teacherAttendance.filter(ta => 
                          ta.nip === selectedTeacher.nip && 
                          ta.kelas === s.kelas && 
                          ta.tanggal.startsWith(currentMonthPrefix) &&
                          (() => {
                            const [y, m, d] = ta.tanggal.split('-').map(Number);
                            const dObj = new Date(y, m-1, d);
                            const dayNames = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
                            return dayNames[dObj.getDay()] === s.hari;
                          })()
                        ).length;
                        totalActualSessions += actual;
                      });

                      const achievementPerc = totalMonthlyTarget > 0 ? Math.min(100, (totalActualSessions / totalMonthlyTarget) * 100) : 0;

                      return (
                        <div key={idx} className="bg-white p-6 rounded-[2rem] border border-gray-100 shadow-sm hover:shadow-md transition-shadow relative overflow-hidden group">
                           <div className="absolute top-0 right-0 p-4">
                              <div className={`w-10 h-10 rounded-2xl flex items-center justify-center ${achievementPerc >= 100 ? 'bg-green-100 text-green-600' : 'bg-orange-100 text-orange-600'}`}>
                                 <Award size={20} />
                              </div>
                           </div>
                           
                           <div className="mb-4">
                              <span className="text-[10px] font-black text-green-600 bg-green-50 px-3 py-1 rounded-full uppercase tracking-widest">
                                {group.days.join(' / ')}
                              </span>
                           </div>

                           <div className="space-y-1 mb-6">
                              <h4 className="text-xl font-black text-zinc-900">{group.kelas}</h4>
                              <p className="text-sm font-bold text-gray-400">{group.mapel}</p>
                           </div>

                           <div className="grid grid-cols-2 gap-4 pt-4 border-t border-gray-50">
                              <div>
                                 <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">Target</p>
                                 <div className="flex items-baseline gap-1">
                                    <span className="text-lg font-black text-zinc-900">{totalMonthlyTarget}</span>
                                    <span className="text-[10px] font-bold text-gray-400">Sesi</span>
                                 </div>
                              </div>
                              <div>
                                 <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">Capaian</p>
                                 <div className="flex items-baseline gap-1">
                                    <span className="text-lg font-black text-green-700">{totalActualSessions}</span>
                                    <span className="text-[10px] font-bold text-gray-400">Sesi</span>
                                 </div>
                              </div>
                           </div>

                           <div className="mt-4">
                              <div className="flex justify-between items-center mb-1.5">
                                 <span className="text-[9px] font-black text-gray-400 uppercase tracking-widest">Progress Bulanan</span>
                                 <span className="text-[10px] font-black text-green-700">{Math.round(achievementPerc)}%</span>
                              </div>
                              <div className="h-2 w-full bg-gray-100 rounded-full overflow-hidden text-[0px]">
                                 <motion.div 
                                   initial={{ width: 0 }}
                                   animate={{ width: `${achievementPerc}%` }}
                                   transition={{ duration: 1, ease: "easeOut" }}
                                   className={`h-full rounded-full ${achievementPerc >= 100 ? 'bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]' : 'bg-orange-500'}`}
                                 />
                              </div>
                           </div>
                        </div>
                      );
                    }) : (
                      <div className="col-span-full p-12 text-center bg-gray-50 rounded-[2rem] border border-dashed border-gray-200">
                         <p className="text-sm font-bold text-gray-400">Guru ini belum memiliki target jam mengajar.</p>
                      </div>
                    );
                  })()}
               </div>
            </div>
          </motion.div>
        )}
      </div>
    );
  };

  const renderProfile = () => {
    const teacher = teachers.find(t => t.nip === session?.uid);
    const data = editingProfileData || teacher;

    if (!data) return <div className="p-8 text-center text-gray-400">Memuat data profil...</div>;

    const mySchedules = teachingSchedules.filter(ts => ts.nip === data.nip);

    return (
      <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-4xl mx-auto space-y-8 pb-20">
        <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-xl border border-gray-100 overflow-hidden relative">
          <div className="absolute top-0 left-0 w-full h-32 bg-green-900/5 -z-10" />
          
          <div className="flex flex-col items-center mb-10">
            <div className="relative mb-6">
              <div className="w-32 h-32 rounded-[2.5rem] bg-green-900 flex items-center justify-center text-white overflow-hidden shadow-2xl ring-4 ring-white">
                {data.foto ? (
                   <img src={data.foto} className="w-full h-full object-cover" crossOrigin="anonymous" alt="Foto Profil" />
                ) : (
                   <User size={64} />
                )}
              </div>
              <label className="absolute -bottom-2 -right-2 bg-green-600 text-white p-2 rounded-xl shadow-lg border-2 border-white cursor-pointer hover:bg-green-700 transition-colors">
                 <Upload size={16} />
                 <input 
                    type="file" 
                    accept="image/*" 
                    className="hidden" 
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        const reader = new FileReader();
                        reader.onloadend = () => {
                          setEditingProfileData({...data, foto: reader.result as string});
                        };
                        reader.readAsDataURL(file);
                      }
                    }}
                  />
              </label>
            </div>
            <h2 className="text-2xl font-black text-green-950 mb-1">{data.nama}</h2>
            <p className="text-xs font-bold text-green-600 uppercase tracking-widest bg-green-50 px-4 py-1.5 rounded-full">{data.jabatan || 'Guru'}</p>
          </div>

          <form onSubmit={handleUpdateProfile} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">NIP / ID</label>
                <input 
                  type="text" 
                  readOnly 
                  value={data.nip} 
                  className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-400 cursor-not-allowed" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Nama Lengkap</label>
                <input 
                  type="text" 
                  required 
                  value={data.nama} 
                  onChange={e => setEditingProfileData({...data, nama: e.target.value})}
                  className="w-full bg-gray-50 focus:bg-white focus:ring-2 focus:ring-green-100 border-0 rounded-2xl px-6 py-4 text-sm font-bold transition-all" 
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 p-6 bg-zinc-50 rounded-[2rem] border border-zinc-100">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Username Login</label>
                <input 
                  type="text" 
                  required 
                  value={data.user} 
                  onChange={e => setEditingProfileData({...data, user: e.target.value})}
                  className="w-full bg-white border border-zinc-200 rounded-2xl px-6 py-4 text-sm font-bold shadow-sm" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Password Login</label>
                <div className="relative">
                  <input 
                    type={showPass ? "text" : "password"} 
                    required 
                    value={data.pass} 
                    onChange={e => setEditingProfileData({...data, pass: e.target.value})}
                    className="w-full bg-white border border-zinc-200 rounded-2xl px-6 py-4 text-sm font-bold shadow-sm" 
                  />
                  <button 
                    type="button" 
                    onClick={() => setShowPass(!showPass)}
                    className="absolute right-4 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-green-600 transition-colors"
                  >
                    {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
                  </button>
                </div>
              </div>
            </div>

            <button 
              type="submit" 
              disabled={isUpdatingProfile || !editingProfileData}
              className="w-full bg-green-900 text-white py-5 rounded-2xl font-black text-sm uppercase tracking-widest shadow-xl shadow-green-100 hover:bg-green-800 disabled:opacity-50 transition-all flex items-center justify-center gap-3 mt-4"
            >
              {isUpdatingProfile ? (
                 <div className="w-5 h-5 border-2 border-white/20 border-t-white rounded-full animate-spin" />
              ) : (
                 <Save size={20} />
              )}
              Update Data Profil
            </button>
          </form>
        </div>

        <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-xl border border-gray-100">
           <h3 className="text-lg font-black text-green-950 mb-6 flex items-center gap-3 uppercase tracking-widest">
             <Calendar className="text-green-600" /> Jam Mengajar & Laporan Capaian
           </h3>
           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {(() => {
                const now = new Date();
                const year = now.getFullYear();
                const month = now.getMonth();
                const currentMonthPrefix = now.toISOString().slice(0, 7);

                // Group schedules by Class and Mapel
                const groupedSchedules: {[key: string]: {
                  kelas: string,
                  mapel: string,
                  days: string[],
                  sessions: any[]
                }} = {};

                mySchedules.forEach(s => {
                  const key = `${s.kelas}-${s.mapel || 'unnamed'}`;
                  if (!groupedSchedules[key]) {
                    groupedSchedules[key] = {
                      kelas: s.kelas,
                      mapel: s.mapel || '-',
                      days: [],
                      sessions: []
                    };
                  }
                  if (!groupedSchedules[key].days.includes(s.hari)) {
                    groupedSchedules[key].days.push(s.hari);
                  }
                  groupedSchedules[key].sessions.push(s);
                });

                const sortedGroups = Object.values(groupedSchedules);

                return sortedGroups.length > 0 ? sortedGroups.map((group, idx) => {
                  let totalMonthlyTarget = 0;
                  let totalActualSessions = 0;

                  group.sessions.forEach(s => {
                    const targetOccurrences = countDaysInMonth(year, month, s.hari, holidays);
                    totalMonthlyTarget += targetOccurrences * (Number(s.targetPertemuan) || 0);

                    const actual = teacherAttendance.filter(ta => 
                      ta.nip === data.nip && 
                      ta.kelas === s.kelas && 
                      ta.tanggal.startsWith(currentMonthPrefix) &&
                      (() => {
                        const [y, m, d] = ta.tanggal.split('-').map(Number);
                        const dObj = new Date(y, m-1, d);
                        const dayNames = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
                        return dayNames[dObj.getDay()] === s.hari;
                      })()
                    ).length;
                    totalActualSessions += actual;
                  });

                  const achievementPerc = totalMonthlyTarget > 0 ? Math.min(100, (totalActualSessions / totalMonthlyTarget) * 100) : 0;

                  return (
                    <div key={idx} className="bg-white p-6 rounded-[2rem] border border-gray-100 shadow-sm hover:shadow-md transition-shadow relative overflow-hidden group">
                       <div className="absolute top-0 right-0 p-4">
                          <div className={`w-10 h-10 rounded-2xl flex items-center justify-center ${achievementPerc >= 100 ? 'bg-green-100 text-green-600' : 'bg-orange-100 text-orange-600'}`}>
                             <Award size={20} />
                          </div>
                       </div>
                       
                       <div className="mb-4">
                          <span className="text-[10px] font-black text-green-600 bg-green-50 px-3 py-1 rounded-full uppercase tracking-widest">
                            {group.days.join(' / ')}
                          </span>
                       </div>

                       <div className="space-y-1 mb-6">
                          <h4 className="text-xl font-black text-zinc-900">{group.kelas}</h4>
                          <p className="text-sm font-bold text-gray-400">{group.mapel}</p>
                       </div>

                       <div className="grid grid-cols-2 gap-4 pt-4 border-t border-gray-50">
                          <div>
                             <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">Target</p>
                             <div className="flex items-baseline gap-1">
                                <span className="text-lg font-black text-zinc-900">{totalMonthlyTarget}</span>
                                <span className="text-[10px] font-bold text-gray-400">Sesi</span>
                             </div>
                          </div>
                          <div>
                             <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">Capaian</p>
                             <div className="flex items-baseline gap-1">
                                <span className="text-lg font-black text-green-700">{totalActualSessions}</span>
                                <span className="text-[10px] font-bold text-gray-400">Sesi</span>
                             </div>
                          </div>
                       </div>

                       <div className="mt-4">
                          <div className="flex justify-between items-center mb-1.5">
                             <span className="text-[9px] font-black text-gray-400 uppercase tracking-widest">Progress Bulanan</span>
                             <span className="text-[10px] font-black text-green-700">{Math.round(achievementPerc)}%</span>
                          </div>
                          <div className="h-2 w-full bg-gray-100 rounded-full overflow-hidden text-[0px]">
                             <motion.div 
                               initial={{ width: 0 }}
                               animate={{ width: `${achievementPerc}%` }}
                               transition={{ duration: 1, ease: "easeOut" }}
                               className={`h-full rounded-full ${achievementPerc >= 100 ? 'bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]' : 'bg-orange-500'}`}
                             />
                          </div>
                       </div>
                    </div>
                  );
                }) : (
                  <div className="col-span-full p-12 text-center bg-gray-50 rounded-[2rem] border border-dashed border-gray-200">
                     <p className="text-sm font-bold text-gray-400">Belum ada target jam mengajar yang diatur oleh Admin.</p>
                  </div>
                );
              })()}
           </div>
           <p className="text-[10px] text-gray-400 mt-6 font-bold uppercase tracking-widest italic">*Data jam mengajar hanya dapat diubah oleh Administrator.</p>
        </div>
        
        <p className="text-center text-[10px] text-gray-400 font-bold uppercase tracking-widest">Pembaruan data profil anda akan sinkron ke dashboard admin secara otomatis.</p>
      </motion.div>
    );
  };

  // Login states
  const [loginRole, setLoginRole] = useState<'Siswa' | 'Guru' | 'Admin'>('Siswa');
  const [loginNisn, setLoginNisn] = useState('');
  const [loginUser, setLoginUser] = useState('');
  const [loginPass, setLoginPass] = useState('');
  const [showPass, setShowPass] = useState(false);

  // Scanner states
  const [scanType, setScanType] = useState<'Siswa' | 'Kelas'>('Siswa');
  const [scanResult, setScanResult] = useState<{ success: boolean, message: string } | null>(null);
  const [scanning, setScanning] = useState(false);

  useEffect(() => {
    let html5QrCode: any = null;
    
    const startScanner = async () => {
      if (activePanel === 'scanner' && scanning) {
        // Wait a small bit for DOM to settle
        await new Promise(resolve => setTimeout(resolve, 100));
        
        const element = document.getElementById("reader");
        if (!element) return;

        try {
          const { Html5Qrcode } = await import('html5-qrcode');
          html5QrCode = new Html5Qrcode("reader");
          
          const qrConfig = { fps: 10, qrbox: { width: 250, height: 250 } };
          
          await html5QrCode.start(
            { facingMode: "environment" },
            qrConfig,
            async (decodedText: string) => {
              setScanning(false);
              if (html5QrCode) {
                try {
                  await html5QrCode.stop();
                } catch (err) {
                  console.error("Stop error:", err);
                }
              }
              
              toggleLoader(true);
              try {
                let res;
                if (scanType === 'Siswa') {
                  res = await firestoreService.processScan(decodedText);
                } else {
                  res = await firestoreService.processClassScan(session?.uid || '', decodedText);
                }
                
                if (res.success) {
                  const audio = new Audio('https://assets.mixkit.co/active_storage/sfx/2568/2568-preview.mp3');
                  audio.play().catch(e => console.log('Audio play blocked'));
                  if (navigator.vibrate) navigator.vibrate(200);
                }
                
                setScanResult({ success: res.success, message: res.message, status: res.status });
              } catch (e) {
                setScanResult({ success: false, message: "Gagal memproses pindaian." });
              } finally {
                toggleLoader(false);
              }
            },
            (errorMessage: string) => {
              // Ignore frequent scan errors
            }
          );
        } catch (err) {
          console.error("Scanner startup error:", err);
          setScanning(false);
        }
      }
    };

    startScanner();

    return () => {
      if (html5QrCode && html5QrCode.isScanning) {
        html5QrCode.stop().catch((err: any) => console.error("Unmount stop error:", err));
      }
    };
  }, [activePanel, scanning, scanType, session]);

  // Modal states
  const ConfirmModal = () => {
    if (!confirmModal || !confirmModal.show) return null;
    return (
      <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
        <motion.div 
          initial={{ scale: 0.9, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          className="bg-white p-8 rounded-[2.5rem] shadow-2xl max-w-sm w-full border border-gray-100"
        >
          <div className="flex flex-col items-center text-center">
            <div className="bg-red-50 p-4 rounded-3xl text-red-600 mb-6 scale-110">
              <Trash2 size={32} />
            </div>
            <h3 className="text-xl font-black text-gray-900 mb-2">{confirmModal.title}</h3>
            <p className="text-sm text-gray-500 font-medium mb-8 leading-relaxed">
              {confirmModal.entityName ? (
                <>Data <strong>{confirmModal.entityName}</strong> akan dihapus permanen dari sistem.</>
              ) : (
                confirmModal.message
              )}
            </p>
            <div className="flex gap-3 w-full">
              <button 
                onClick={() => setConfirmModal(null)} 
                className="flex-1 bg-gray-100 text-gray-400 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-gray-200 transition-all"
              >
                Batal
              </button>
              <button 
                onClick={() => { confirmModal.onConfirm(); setConfirmModal(null); }} 
                className="flex-1 bg-red-600 text-white py-4 rounded-2xl font-black text-xs uppercase tracking-widest shadow-lg shadow-red-200 hover:bg-red-700 transition-all scale-105"
              >
                Hapus Sekarang
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    );
  };

  const [showSiswaModal, setShowSiswaModal] = useState(false);
  const [editingSiswa, setEditingSiswa] = useState<Student | null>(null);
  const [showGuruModal, setShowGuruModal] = useState(false);
  const [editingGuru, setEditingGuru] = useState<any | null>(null);
  const [showKelasModal, setShowKelasModal] = useState(false);
  const [editingKelas, setEditingKelas] = useState<any | null>(null);
  const [showSuccessToast, setShowSuccessToast] = useState<{title: string, message: string} | null>(null);
  const [newLibur, setNewLibur] = useState({ tanggal: '', keterangan: '' });

  const [showCelebration, setShowCelebration] = useState(false);

  const triggerSuccess = (title: string, message: string) => {
    setShowSuccessToast({ title, message });
    setShowCelebration(true);
    setSearchTermGuru('');
    setSearchTermSiswa('');
    setSearchTermKelas('');
    setSearchTermJadwal('');
    setPagination({ guru: 0, siswa: 0, kelas: 0, jadwal: 0, absensiSiswa: 0, absensiGuru: 0, rekapMapel: 0 });
    setTimeout(() => {
      setShowSuccessToast(null);
      setShowCelebration(false);
    }, 3000);
  };

  const triggerError = (title: string, message: string) => {
    alert(`${title}: ${message}`);
  };
  const [mobileExtraMenuOpen, setMobileExtraMenuOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [showJadwalModal, setShowJadwalModal] = useState(false);
  const [editingJadwal, setEditingJadwal] = useState<TeachingSchedule | null>(null);
  const [showSessionDetail, setShowSessionDetail] = useState<{name: string, mapping: any[]} | null>(null);

  const printBarcode = (className: string) => {
    const win = window.open('', '_blank');
    if (win) {
      win.document.write(`
        <html>
          <head><title>Print Barcode ${className}</title></head>
          <body style="display:flex; flex-direction:column; align-items:center; justify-content:center; height:100vh; font-family:sans-serif;">
            <div style="padding:40px; border:4px solid #166534; border-radius:24px; text-align:center;">
              <h3 style="text-align:center; font-size:24px; color:#166534; margin-bottom:20px;">SCAN ABSENSI MENGAJAR<br/>KELAS ${className}</h3>
              <img src="https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${className}" />
              <p style="font-weight:900; font-size:20px; color:#166534; margin-top:20px;">SIGAP MTsN 2 BOMBANA</p>
            </div>
            <script>window.onload = () => { window.print(); window.close(); }</script>
          </body>
        </html>
      `);
      win.document.close();
    }
  };

  const captureElement = async (elementId: string) => {
    const el = document.getElementById(elementId);
    if (!el) {
      console.error(`Capture Error: Element #${elementId} not found`);
      alert("Elemen kartu tidak ditemukan. Segarkan halaman.");
      return null;
    }

    // Capture original state
    const originalStyle = el.style.cssText;
    
    // Temporarily make it visible and independent for capture
    // We force it to be at the top level and fully visible without any transitions
    el.style.cssText += `
      position: fixed !important;
      top: 0 !important;
      left: 0 !important;
      z-index: 9999999 !important;
      display: flex !important;
      visibility: visible !important;
      opacity: 1 !important;
      transform: none !important;
      transition: none !important;
      pointer-events: none !important;
      background: white !important;
      width: fit-content !important;
      height: fit-content !important;
      box-shadow: none !important;
      padding: 0 !important;
      margin: 0 !important;
    `;

    try {
      console.log(`Starting capture for #${elementId}...`);
      
      // Delay to ensure the DOM has rendered the new styles and images
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Force reload images with crossOrigin for CORS compatibility
      const images = Array.from(el.querySelectorAll('img'));
      await Promise.all(images.map(img => {
        if (img.complete && img.naturalWidth !== 0) return Promise.resolve();
        return new Promise(resolve => {
          img.crossOrigin = "anonymous";
          const src = img.src;
          if (src) {
            img.src = ""; // Reset
            img.src = src; // Re-trigger
          }
          img.onload = () => resolve(true);
          img.onerror = () => resolve(false);
          // Safety timeout for image loading
          setTimeout(() => resolve(false), 3000);
        });
      }));

      // Additional small delay for layout stabilization
      await new Promise(resolve => setTimeout(resolve, 300));

      // Try html-to-image first
      try {
        const dataUrl = await htmlToImage.toPng(el, {
          pixelRatio: 2,
          backgroundColor: '#ffffff',
          cacheBust: true,
          style: {
            transform: 'none',
            visibility: 'visible',
            display: 'flex',
            opacity: '1',
            margin: '0',
            padding: '0'
          }
        });
        
        if (dataUrl && dataUrl.length > 500) { 
          console.log("Capture with html-to-image successful.");
          return dataUrl;
        }
      } catch (innerErr) {
        console.warn("html-to-image failed, trying html2canvas...", innerErr);
      }

      // html2canvas fallback
      const html2canvas = (await import('html2canvas')).default;
      const canvas = await html2canvas(el, {
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        scale: 2,
        logging: false,
        onclone: (clonedDoc) => {
          const win = clonedDoc.defaultView || window;
          const clonedEl = clonedDoc.getElementById(elementId);
          if (clonedEl) {
             clonedEl.style.cssText += "transform: none !important; visibility: visible !important; display: flex !important; opacity: 1 !important; position: static !important; background: white !important;";
             
             // Recursively remove oklch/oklab from ALL stylesheets in the clone
             // html2canvas parser crashes when it encounters these keywords anywhere
             Array.from(clonedDoc.styleSheets).forEach(sheet => {
               try {
                 const rules = sheet.cssRules;
                 for (let i = rules.length - 1; i >= 0; i--) {
                   if (rules[i].cssText.includes('oklch') || rules[i].cssText.includes('oklab')) {
                     sheet.deleteRule(i);
                   }
                 }
               } catch (e) {
                 // Cross-origin sheets might throw, so we fallback to fixing style tags
               }
             });

             const styleTags = clonedDoc.querySelectorAll('style');
             styleTags.forEach(tag => {
               try {
                 tag.innerHTML = tag.innerHTML.replace(/(oklch|oklab)\s*\([^)]+\)/g, '#333333');
               } catch (e) {}
             });

             // Recursively find and fix oklch colors in inline/computed styles for elements
             const allElements = clonedEl.querySelectorAll('*');
             const canvasHelper = clonedDoc.createElement('canvas'); 
             const ctx = canvasHelper.getContext('2d');
             
             allElements.forEach((child) => {
               const element = child as HTMLElement;
               
               // Remove filter/backdropFilter as they cause issues
               element.style.filter = 'none';
               element.style.backdropFilter = 'none';

               // Properties to check for oklch colors
               const props = ['color', 'backgroundColor', 'borderColor', 'outlineColor', 'fill', 'stroke', 'stopColor'];
               
               props.forEach(prop => {
                 // Try to get value carefully to avoid triggering crash if possible
                 let val = '';
                 try {
                    val = element.style.getPropertyValue(prop);
                    if (!val || val.includes('oklch') || val.includes('oklab')) {
                      const computed = win.getComputedStyle(element);
                      val = computed.getPropertyValue(prop);
                    }
                 } catch (e) {}

                 if (val && (val.includes('oklch') || val.includes('oklab'))) {
                   let converted = false;
                   if (ctx) {
                     try {
                       ctx.fillStyle = val;
                       const rgbVal = ctx.fillStyle; 
                       if (rgbVal && !rgbVal.includes('oklch') && !rgbVal.includes('oklab') && rgbVal !== '#000000') {
                         element.style.setProperty(prop, rgbVal, 'important');
                         converted = true;
                       }
                     } catch (e) {}
                   }
                   
                   if (!converted) {
                      if (prop === 'color' || prop === 'stroke') element.style.setProperty(prop, '#000000', 'important');
                      else if (prop === 'backgroundColor' || prop === 'fill') element.style.setProperty(prop, '#ffffff', 'important');
                      else element.style.setProperty(prop, 'transparent', 'important');
                   }
                 }
               });
             });
          }
        }
      });
      
      const canvasDataUrl = canvas.toDataURL('image/png');
      console.log("Capture with html2canvas successful.");
      return canvasDataUrl;
    } catch (err) {
      console.error("All capture methods failed:", err);
      try {
        // Absolute last resort: simplest possible capture
        return await htmlToImage.toPng(el);
      } catch (finalErr) {
        alert(`Gagal mengambil gambar: ${err instanceof Error ? err.message : 'Unknown Error'}`);
        return null;
      }
    } finally {
      // Revert to original state
      el.style.cssText = originalStyle;
    }
  };

  const downloadAsPDF = async (elementId: string, fileName: string) => {
    toggleLoader(true);
    try {
      console.log(`Preparing PDF download for ${elementId}...`);
      const dataUrl = await captureElement(elementId);
      if (!dataUrl) {
        toggleLoader(false);
        return;
      }
      
      // Create image object to get real dimensions
      const img = new Image();
      img.src = dataUrl;
      await new Promise(resolve => img.onload = resolve);
      
      // Use logical dimensions (divided by 2 because we used pixelRatio: 2)
      const logicalWidth = img.width / 2;
      const logicalHeight = img.height / 2;

      const pdf = new jsPDF({
        orientation: logicalWidth > logicalHeight ? 'l' : 'p',
        unit: 'px',
        format: [logicalWidth, logicalHeight]
      });

      pdf.addImage(dataUrl, 'PNG', 0, 0, logicalWidth, logicalHeight);
      
      pdf.save(fileName.replace('.png', '.pdf'));
      console.log("PDF download triggered successfully.");
    } catch (err) {
      console.error("PDF download failed:", err);
      alert("Gagal mengunduh PDF. Silakan coba lagi.");
    } finally {
      toggleLoader(false);
    }
  };

  const downloadClassPDF = (className: string) => {
    downloadAsPDF(`card-kelas-${className}`, `Barcode-Kelas-${className}.pdf`);
  };

  useEffect(() => {
    setLoginNisn('');
    setLoginUser('');
    setLoginPass('');
  }, [loginRole]);

  const [siswaAbsensiManual, setSiswaAbsensiManual] = useState({ nisn: '', status: 'Hadir', ket: '', tanggal: new Date().toISOString().split('T')[0] });
  const [selectedStudentCard, setSelectedStudentCard] = useState<Student | null>(null);

  const exportExcel = () => {
    const isSiswa = rekapFilter.type === 'Siswa';
    const data = isSiswa 
      ? students
        .filter(s => (rekapFilter.kelas ? s.kelas === rekapFilter.kelas : true))
        .map(s => {
          const monthAttendance = attendance.filter(a => a.nisn === s.nisn && a.tanggal.startsWith(rekapFilter.bulan));
          const hadir = monthAttendance.filter(a => a.status === 'Hadir').length;
          const izin = monthAttendance.filter(a => a.status === 'Izin').length;
          const sakit = monthAttendance.filter(a => a.status === 'Sakit').length;
          const alfa = monthAttendance.filter(a => a.status === 'Alfa').length;
          const totalLambat = monthAttendance.reduce((sum, a) => sum + (a.terlambat || 0), 0);
          const [y, m] = rekapFilter.bulan.split('-').map(Number);
          const now = new Date();
          const limit = (now.getFullYear() === y && (now.getMonth() + 1) === m) ? now.getDate() : undefined;
          const effCount = getEffectiveDays(y, m - 1, holidays, settings, limit).length;
          const perc = effCount > 0 ? Math.round(((hadir + izin + sakit) / effCount) * 100) : 0;
          return {
            "Nama Siswa": s.nama,
            "Kelas": s.kelas,
            "Hadir": hadir,
            "Izin": izin,
            "Sakit": sakit,
            "Alfa": alfa,
            "Terlambat (Menit)": totalLambat,
            "Jam Terlambat": formatMinutes(totalLambat),
            "% Kehadiran": `${perc}%`
          };
        })
      : teachers
        .filter(t => {
           if (!rekapFilter.kelas) return true;
           const scheds = teachingSchedules.filter(ts => ts.nip === t.nip);
           return scheds.some(sc => sc.kelas === rekapFilter.kelas);
        })
        .map(t => {
          const schedules = teachingSchedules.filter(ts => ts.nip === t.nip && (rekapFilter.kelas ? ts.kelas === rekapFilter.kelas : true));
          const [y, m] = rekapFilter.bulan.split('-').map(Number);
          const totalTarget = schedules.reduce((sum, sch) => {
            const occ = countDaysInMonth(y, m - 1, sch.hari, holidays);
            return sum + (occ * (Number(sch.targetPertemuan) || 0));
          }, 0);
          const monthAtts = teacherAttendance.filter(ta => ta.nip === t.nip && (rekapFilter.kelas ? ta.kelas === rekapFilter.kelas : true) && ta.tanggal.startsWith(rekapFilter.bulan));
          const actual = monthAtts.length;
          const totalLambat = monthAtts.reduce((sum, a) => sum + (a.terlambat || 0), 0);
          const perc = totalTarget > 0 ? Math.round((actual / totalTarget) * 100) : 0;
          return {
            "Nama Guru": t.nama,
            "Jabatan": t.jabatan || 'Guru',
            "Kelas": rekapFilter.kelas || schedules.map(s => s.kelas).join(', ') || '-',
            "Target Sesi": totalTarget,
            "Aktual Sesi": actual,
            "Terlambat (Menit)": totalLambat,
            "Jam Terlambat": formatMinutes(totalLambat),
            "% Kinerja": `${perc}%`
          };
        });
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Kehadiran");
    XLSX.writeFile(wb, `Rekap-${rekapFilter.type}-${rekapFilter.bulan}.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF() as any;
    const isSiswa = rekapFilter.type === 'Siswa';
    const data = isSiswa 
      ? students
        .filter(s => (rekapFilter.kelas ? s.kelas === rekapFilter.kelas : true))
        .map(s => {
          const m = attendance.filter(a => a.nisn === s.nisn && a.tanggal.startsWith(rekapFilter.bulan));
          const h = m.filter(a => a.status === 'Hadir').length;
          const i = m.filter(a => a.status === 'Izin').length;
          const sk = m.filter(a => a.status === 'Sakit').length;
          const f = m.filter(a => a.status === 'Alfa').length;
          const lbt = m.reduce((sum, a) => sum + (a.terlambat || 0), 0);
          const [y, mm] = rekapFilter.bulan.split('-').map(Number);
          const now = new Date();
          const limit = (now.getFullYear() === y && (now.getMonth() + 1) === mm) ? now.getDate() : undefined;
          const effCount = getEffectiveDays(y, mm - 1, holidays, settings, limit).length;
          const p = effCount > 0 ? Math.round(((h + i + sk) / effCount) * 100) : 0;
          return [s.nama, s.kelas, h, i, sk, f, lbt, `${p}%`];
        })
      : teachers
        .filter(t => {
          if (!rekapFilter.kelas) return true;
          const scheds = teachingSchedules.filter(ts => ts.nip === t.nip);
          return scheds.some(sc => sc.kelas === rekapFilter.kelas);
        })
        .map(t => {
          const schs = teachingSchedules.filter(ts => ts.nip === t.nip && (rekapFilter.kelas ? ts.kelas === rekapFilter.kelas : true));
          const [y, m] = rekapFilter.bulan.split('-').map(Number);
          const tgt = schs.reduce((sum, sch) => {
            const occ = countDaysInMonth(y, m - 1, sch.hari, holidays);
            return sum + (occ * (Number(sch.targetPertemuan) || 0));
          }, 0);
          const atts = teacherAttendance.filter(ta => ta.nip === t.nip && (rekapFilter.kelas ? ta.kelas === rekapFilter.kelas : true) && ta.tanggal.startsWith(rekapFilter.bulan));
          const act = atts.length;
          const lbt = atts.reduce((sum, a) => sum + (a.terlambat || 0), 0);
          const p = tgt > 0 ? Math.round((act / tgt) * 100) : 0;
          return [t.nama, t.jabatan || 'Guru', rekapFilter.kelas || schs.map(s => s.kelas).join(', ') || '-', tgt, act, lbt, `${p}%`];
        });
    
    doc.text(`Laporan Rekap Kehadiran ${rekapFilter.type}`, 14, 15);
    doc.text(`Bulan: ${rekapFilter.bulan}`, 14, 25);
    
    const columns = isSiswa 
      ? ["Nama", "Kelas", "H", "I", "S", "A", "Lbt(m)", "%"]
      : ["Nama Guru", "Jabatan", "Kelas", "Target", "Aktual", "Lbt(m)", "%"];
      
    autoTable(doc, {
      startY: 30,
      head: [columns],
      body: data,
    });
    doc.save(`Rekap-${rekapFilter.type}-${rekapFilter.bulan}.pdf`);
  };

  const toggleLoader = (val: boolean) => setLoading(val);
  const triggerCurrentPanelRefresh = async () => {
    setRefreshing(true);
    // Reset all filters and pagination to ensure user sees latest data
    setSearchTermGuru('');
    setSearchTermSiswa('');
    setSearchTermKelas('');
    setSearchTermJadwal('');
    setFilterClassAbsensi('');
    setFilterNamaAbsensi('');
    setFilterAdminSiswaClass('');
    setStudentProfileClassFilter('');
    setAnalysisClass('');
    
    // Reset pagination
    setPagination({ 
      guru: 0, 
      siswa: 0, 
      kelas: 0, 
      jadwal: 0, 
      absensiSiswa: 0, 
      absensiGuru: 0, 
      rekapMapel: 0 
    });

    // Real-time data is synced via onSnapshot, but we can simulate a fresh state
    setLoading(true);
    setTimeout(() => {
      setRefreshing(false);
      setLoading(false);
      triggerSuccess("DIPERBARUI", "Sinkronisasi data terbaru berhasil dilakukan. Semua filter telah direset.");
    }, 1000);
  };

  const handleLogin = async () => {
    const u = loginRole === 'Siswa' ? loginNisn : loginUser;
    const p = loginRole === 'Siswa' ? '' : loginPass;
    
    if (!u || (loginRole !== 'Siswa' && !p)) {
      alert("Harap isi semua kolom login.");
      return;
    }

    toggleLoader(true);
    try {
      const res = await firestoreService.checkLogin(u, p, loginRole);
      toggleLoader(false);
      if (res.success) {
        setSession(res);
        if (res.role === 'Siswa') setActivePanel('siswa-personal');
      } else {
        alert(res.message);
      }
    } catch (e) {
      toggleLoader(false);
      alert("Gagal terhubung ke database.");
    }
  };

  const defaultLogo = "/logo.png";
  const fallbackLogo = "https://www.kemenag.go.id/assets/images/logo-kemenag.png";

  const appLogo = useMemo(() => {
    return appConfig.logoUrl || defaultLogo;
  }, [appConfig.logoUrl]);

  const dayOrder: any = { "Senin": 1, "Selasa": 2, "Rabu": 3, "Kamis": 4, "Jumat": 5, "Sabtu": 6, "Minggu": 7 };

  const renderQuotaNoticeModal = () => {
    if (!quotaExceeded || dismissQuotaNotice) return null;

    return (
      <div className="fixed inset-0 bg-zinc-950/80 backdrop-blur-md z-[10000] flex items-center justify-center p-4 overflow-y-auto">
        <motion.div 
          initial={{ scale: 0.95, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          className="bg-white rounded-[2.5rem] w-full max-w-lg shadow-2xl overflow-hidden border border-red-100 flex flex-col relative"
        >
          {/* Top Warning Banner Strip */}
          <div className="absolute top-0 left-0 w-full h-3 bg-gradient-to-r from-red-500 via-orange-500 to-amber-500"></div>
          
          <div className="p-8 md:p-10 flex flex-col items-center text-center">
            <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mb-6 text-red-600 shadow-sm shadow-red-100 animate-pulse">
              <AlertTriangle size={32} />
            </div>

            <h2 className="text-2xl font-black text-gray-900 uppercase tracking-tighter mb-3 leading-tight">
              Kapasitas Database Harian Terlampaui
            </h2>
            <div className="bg-red-50 text-[10px] font-black text-red-700 uppercase tracking-widest px-3 py-1 rounded-full border border-red-100 mb-6">
              Quota Limit Exceeded
            </div>

            <div className="space-y-4 text-left bg-zinc-50 p-6 rounded-2xl border border-zinc-100 text-xs text-zinc-600 leading-relaxed font-medium">
              <p>
                Sistem mendeteksi bahwa batas kuota gratis harian (<strong>Firestore Free Tier Read units</strong>) untuk proyek database Firebase ini telah terlampaui.
              </p>
              <div className="flex gap-2.5 items-start">
                <div className="w-2 h-2 rounded-full bg-red-500 mt-1.5 shrink-0" />
                <p>
                  Batas harian Spark Plan adalah <strong>50.000 pembacaan</strong> per hari. Layanan akan dipulihkan secara otomatis setelah kuota di-reset kembali keesokan harinya oleh Google Cloud.
                </p>
              </div>
              <div className="flex gap-2.5 items-start">
                <div className="w-2 h-2 rounded-full bg-zinc-900 mt-1.5 shrink-0" />
                <p>
                  <strong>Untuk Pemilik Proyek:</strong> Anda dapat mengaktifkan billing (upgrade ke Blaze/Pay-as-you-go plan) pada konsol Firebase untuk menghindari batas harian ini sama sekali.
                </p>
              </div>
            </div>

            <div className="mt-8 w-full space-y-3">
              <a 
                href="https://console.firebase.google.com/project/gen-lang-client-0541836694/firestore/databases/ai-studio-8c6b2059-8af6-4d53-bb89-22a78bab9d06/data?openUpgradeDialog=true"
                target="_blank"
                rel="noopener noreferrer"
                className="w-full bg-zinc-950 text-white rounded-2xl py-4.5 px-6 font-black uppercase tracking-widest text-xs hover:bg-zinc-800 transition-all flex items-center justify-center gap-2.5 group cursor-pointer shadow-lg shadow-zinc-200 no-underline"
              >
                <ExternalLink size={16} className="text-amber-400 group-hover:rotate-12 transition-transform" />
                Buka Konsol Firebase & Upgrade
              </a>

              <div className="flex gap-3 w-full">
                <a 
                  href="https://firebase.google.com/pricing#cloud-firestore"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="flex-1 bg-white hover:bg-zinc-50 text-zinc-700 border border-zinc-200 rounded-2xl py-3.5 px-4 font-bold text-[11px] uppercase tracking-wider transition-all text-center no-underline"
                >
                  Informasi Harga (Spark Plan)
                </a>

                <button 
                  onClick={() => setDismissQuotaNotice(true)}
                  className="flex-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 rounded-2xl py-3.5 px-4 font-bold text-[11px] uppercase tracking-wider transition-all animate-fade-in"
                >
                  Sembunyikan
                </button>
              </div>
            </div>

            <p className="mt-6 text-[10px] font-bold text-zinc-400 italic">
              *Proyek ID: gen-lang-client-0541836694 • DB: ai-studio-8c6b2059-8af6-4d53-bb89-22a78bab9d06
            </p>
          </div>
        </motion.div>
      </div>
    );
  };

  const renderLogin = () => (
    <div className="min-h-screen bg-green-950 flex items-center justify-center p-4">
      {renderQuotaNoticeModal()}
      {loading && (
        <div className="fixed inset-0 bg-black/20 backdrop-blur-[2px] z-[9999] flex items-center justify-center">
          <div className="bg-white p-6 rounded-3xl shadow-2xl flex flex-col items-center">
            <RefreshCcw className="animate-spin text-green-700 mb-2" size={32} />
            <p className="font-bold text-green-900 text-xs uppercase tracking-widest">Sinkronisasi...</p>
          </div>
        </div>
      )}
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="bg-white rounded-3xl shadow-2xl w-full max-w-md overflow-hidden"
      >
        <div className="bg-green-800 p-8 text-center text-white">
          <div className="bg-white w-24 h-24 rounded-full flex items-center justify-center mx-auto mb-4 backdrop-blur-sm shadow-xl overflow-hidden p-3 border-4 border-green-700/20 relative">
            <img 
              src={appLogo} 
              alt="Logo" 
              crossOrigin="anonymous"
              className="w-full h-full object-contain relative z-10" 
              onError={(e) => {
                const target = e.currentTarget;
                if (target.src === defaultLogo) {
                  target.src = fallbackLogo;
                } else if (target.src === fallbackLogo) {
                  target.style.opacity = '0';
                  const parent = target.parentElement;
                  if (parent) {
                    const icon = parent.querySelector('.fallback-icon');
                    if (icon) (icon as HTMLElement).style.opacity = '1';
                  }
                } else {
                   target.src = defaultLogo;
                }
              }}
            />
            <School size={48} className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity" />
          </div>
          <h1 className="text-4xl font-black tracking-tighter">SIGAP</h1>
          <p className="text-green-100 text-[10px] font-bold uppercase tracking-widest mt-1 opacity-80">Sistem Informasi Guru dan Absensi Pelajar</p>
          <div className="mt-3 inline-block">
            <p className="text-white font-black text-xs bg-green-900/40 py-1.5 rounded-full border border-white/20 px-5 shadow-lg backdrop-blur-sm uppercase tracking-wider">
              MTsN 2 BOMBANA
            </p>
          </div>
          
          <div className="mt-4 flex flex-col items-center justify-center gap-1">
            <div className="flex items-center gap-2">
              <div className={`w-2 h-2 rounded-full animate-pulse ${firebaseConnected ? 'bg-green-400 shadow-[0_0_8px_#4ade80]' : 'bg-red-500'}`} />
              <span className="text-[9px] font-bold text-white/60 uppercase tracking-widest">
                Database: {firebaseConnected ? 'Terhubung (Online)' : 'Terputus (Offline)'}
              </span>
            </div>
            {quotaExceeded && (
              <div className="bg-red-500/20 px-3 py-1.5 rounded-lg border border-red-500/30 mt-2">
                <p className="text-[9px] font-black text-red-200 uppercase tracking-widest animate-pulse flex items-center gap-2">
                  <AlertCircle size={10} /> Quota Firebase Habis (Limit Tercapai)
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="p-8">
          {(() => {
            const appRole = import.meta.env.VITE_APP_ROLE || 'staff';
            const roles = (['Siswa', 'Guru', 'Admin'] as const).filter(r => {
              if (appRole === 'student') return r === 'Siswa';
              if (appRole === 'teacher') return r === 'Siswa' || r === 'Guru';
              return true;
            });

            if (roles.length <= 1) return null;

            return (
              <div className="flex bg-gray-100 p-1 rounded-xl mb-6">
                {roles.map(role => (
                  <button
                    key={role}
                    disabled={loading}
                    onClick={() => setLoginRole(role)}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${loginRole === role ? 'bg-white text-green-800 shadow-sm' : 'text-gray-500'}`}
                  >
                    {role}
                  </button>
                ))}
              </div>
            );
          })()}

          <div className="space-y-4">
            {loginRole === 'Siswa' ? (
              <div>
                <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">Nomor Induk Siswa Nasional (NISN)</label>
                <input 
                  type="number" 
                  disabled={loading}
                  value={loginNisn}
                  onChange={e => setLoginNisn(e.target.value)}
                  className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50" 
                  placeholder="Masukkan NISN"
                />
              </div>
            ) : (
              <>
                <div>
                  <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">ID Pengguna / NIP</label>
                  <input 
                    type="text" 
                    disabled={loading}
                    value={loginUser}
                    onChange={e => setLoginUser(e.target.value)}
                    className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50" 
                    placeholder="Username"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">Kata Sandi</label>
                  <div className="relative">
                    <input 
                      type={showPass ? "text" : "password"} 
                      disabled={loading}
                      value={loginPass}
                      onChange={e => setLoginPass(e.target.value)}
                      className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50" 
                      placeholder="••••••••"
                    />
                    <button onClick={() => setShowPass(!showPass)} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-green-800 transition-colors">
                      {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
                    </button>
                  </div>
                </div>
              </>
            )}

            <button 
              onClick={handleLogin}
              disabled={loading}
              className={`w-full bg-green-800 text-white rounded-xl py-4 font-bold shadow-lg hover:bg-green-700 active:scale-95 transition-all mt-4 flex items-center justify-center gap-2 ${loading ? 'opacity-70 cursor-wait' : ''}`}
            >
              {loading ? (
                <>
                  <RefreshCcw className="animate-spin" size={20} />
                  Memproses...
                </>
              ) : "Masuk Aplikasi"}
            </button>
          </div>
        </div>
      </motion.div>
    </div>
  );

  const [scanInput, setScanInput] = useState('');

  const handleManualScan = async () => {
    if (!scanInput) return;
    toggleLoader(true);
    try {
      let res;
      if (scanType === 'Siswa') {
        res = await firestoreService.processScan(scanInput);
      } else {
        res = await firestoreService.processClassScan(session?.uid || '', scanInput);
      }
      
      if (res.success) {
        const audio = new Audio('https://assets.mixkit.co/active_storage/sfx/2568/2568-preview.mp3');
        audio.play().catch(e => console.log('Audio play blocked'));
        if (navigator.vibrate) navigator.vibrate(200);
      }
      
      setScanResult({ success: res.success, message: res.message, status: res.status });
      setScanInput('');
    } catch (e) {
      setScanResult({ success: false, message: "Gagal memproses pindaian." });
    } finally {
      toggleLoader(false);
    }
  };

  const handleLogout = () => {
    setSession(null);
    setLoginNisn('');
    setLoginUser('');
    setLoginPass('');
    setActivePanel('home');
  };

  const handleImportExcel = (type: 'Siswa' | 'Guru') => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx, .xls';
    input.onchange = (e: any) => {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = async (evt: any) => {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        
        toggleLoader(true);
        try {
          const promises = (data as any[]).map(item => {
            if (type === 'Siswa') return firestoreService.saveSiswa(item);
            return firestoreService.saveGuru({
              ...item,
              role: (item.role || 'Guru') as any
            });
          });
          await Promise.all(promises);
          triggerSuccess("BERHASIL", `Sukses mengimpor ${data.length} data ${type}. Data akan muncul otomatis di tabel.`);
        } catch (err) {
          triggerError("GAGAL", "Gagal mengimpor data. Pastikan format template sesuai dan data tidak duplikat.");
        } finally {
          toggleLoader(false);
        }
      };
      reader.readAsBinaryString(file);
    };
    input.click();
  };

  const downloadTemplate = (type: 'Siswa' | 'Guru') => {
    const headers = type === 'Siswa' 
      ? [["nisn", "nama", "tempat", "tgl", "kelas", "ayah", "ibu", "hp"]]
      : [["nip", "nama", "kelas", "user", "pass"]];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, `Template_Import_${type}.xlsx`);
  };

  const getSiswaPercentage = useMemo(() => {
    if (session?.role !== 'Siswa') return 0;
    const today = new Date();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();
    const currentMonthAbsen = attendance.filter(a => {
      const [aYear, aMonth] = a.tanggal.split('-').map(Number);
      return a.nisn === session?.uid && aMonth === month && aYear === year;
    });
    
    if (currentMonthAbsen.length === 0 && effectiveDaysThisMonth.length === 0) return 0;
    const hadir = currentMonthAbsen.filter(a => a.status === 'Hadir' || a.status === 'Izin' || a.status === 'Sakit').length;
    const effCount = effectiveDaysThisMonth.length;
    return effCount > 0 ? Math.round((hadir / effCount) * 100) : 0;
  }, [attendance, session, effectiveDaysThisMonth]);

  const menuItems = useMemo(() => {
    if (!session) return [];
    if (session?.role === 'Admin') {
      return [
        { id: 'home', label: 'Dashboard Utama', icon: LayoutGrid },
        { id: 'scanner', label: 'Scanner Presensi', icon: QrCode },
        { id: 'guru', label: 'Data Guru', icon: UserPlus },
        { id: 'siswa', label: 'Data Siswa', icon: GraduationCap },
        { id: 'kelas', label: 'Data Kelas', icon: TableIcon },
        { id: 'jadwal-mengajar', label: 'Jam Mengajar', icon: Calendar },
        { id: 'absensi-umum', label: 'Absensi Siswa', icon: ClipboardList },
        { id: 'absensi-guru', label: 'Absensi Guru', icon: User },
        { id: 'rekap', label: 'Laporan Rekap', icon: FileBarChart },
        { id: 'rekap-mapel', label: 'Rekap Mapel', icon: ClipboardList },
        { id: 'analisis', label: 'Analisis Kehadiran', icon: BarChart3 },
        { id: 'pengaturan', label: 'Pengaturan', icon: Settings }
      ];
    } else if (session?.role === 'Guru') {
      const teacher = teachers.find(t => t.nip === session?.uid);
      const jabatan = teacher?.jabatan || 'Guru';
      
      const items = [
        { id: 'home', label: 'Dashboard Utama', icon: LayoutGrid },
        { id: 'scanner', label: 'Scanner Presensi', icon: QrCode }
      ];

      items.push({ id: 'profil', label: 'Profil Saya', icon: User });

      if (jabatan === 'Kamad') {
        items.push({ id: 'absensi-umum', label: 'Absensi Siswa', icon: ClipboardList });
        items.push({ id: 'absensi-guru', label: 'Absensi Guru', icon: User });
        items.push({ id: 'capaian-guru', label: 'Laporan Capaian Guru', icon: Award });
        items.push({ id: 'analisis', label: 'Analisis Kehadiran', icon: BarChart3 });
      } else if (jabatan === 'Wakamad') {
        items.push({ id: 'absensi-umum', label: 'Absensi Siswa', icon: ClipboardList });
        items.push({ id: 'absensi-guru', label: 'Absensi Guru', icon: User });
        items.push({ id: 'roster', label: 'Roster Mengajar', icon: TableIcon });
        items.push({ id: 'profil-siswa', label: 'Profil Siswa', icon: GraduationCap });
        items.push({ id: 'rekap', label: 'Laporan Rekap', icon: FileBarChart });
        items.push({ id: 'rekap-mapel', label: 'Rekap Mapel', icon: ClipboardList });
        items.push({ id: 'analisis', label: 'Analisis Kehadiran', icon: BarChart3 });
      } else if (jabatan === 'Guru Wali Kelas' || session?.isWali) {
        items.push({ id: 'profil-siswa', label: 'Profil Siswa', icon: GraduationCap });
        items.push({ id: 'absensi-wali', label: 'Riwayat Kehadiran', icon: ClipboardList });
        items.push({ id: 'rekap', label: 'Laporan Rekap', icon: FileBarChart });
        items.push({ id: 'rekap-mapel', label: 'Rekap Mapel', icon: ClipboardList });
        items.push({ id: 'analisis', label: 'Analisis Kehadiran', icon: BarChart3 });
      } else {
        items.push({ id: 'rekap-mapel', label: 'Rekap Mapel', icon: ClipboardList });
      }

      items.push({ id: 'kbc', label: 'KBC', icon: Heart });

      return items;
    } else {
      return [
        { id: 'siswa-personal', label: 'Dashboard Siswa', icon: User },
        { id: 'siswa-data', label: 'Data Siswa', icon: GraduationCap }
      ];
    }
  }, [session, teachers]);

  const renderSidebar = () => (
    <div className="w-64 bg-zinc-950 min-h-screen flex flex-col text-white sticky top-0 hidden md:flex">
      <div className="p-6 bg-green-900 flex flex-col items-center border-b border-green-800">
        <div className="bg-white w-16 h-16 rounded-full mb-3 overflow-hidden p-2 shadow-xl flex items-center justify-center relative border-4 border-green-700/20 backdrop-blur-sm">
          <img 
            src={appLogo} 
            alt="Logo" 
            className="w-full h-full object-contain relative z-10" 
            onError={(e) => {
              const target = e.currentTarget;
              if (target.src === defaultLogo) {
                target.src = fallbackLogo;
              } else if (target.src === fallbackLogo) {
                target.style.opacity = '0';
                const parent = target.parentElement;
                if (parent) {
                  const icon = parent.querySelector('.fallback-icon');
                  if (icon) (icon as HTMLElement).style.opacity = '1';
                }
              } else {
                 target.src = defaultLogo;
              }
            }}
          />
          <School size={28} className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity" />
        </div>
        <h2 className="font-black text-xs text-center text-white uppercase tracking-widest">SIGAP MTsN 2</h2>
        <span className="bg-white/10 text-white text-[9px] px-2 py-0.5 rounded-full font-bold mt-2 uppercase border border-white/10">
          {session?.role} {session?.isWali && `(Wali ${session.kelas})`}
        </span>
      </div>

      <nav className="flex-grow py-4 overflow-y-auto">
        {menuItems.map(item => (
          <button
            key={item.id}
            onClick={() => setActivePanel(item.id)}
            className={`w-full flex items-center px-6 py-4 transition-all text-sm font-medium border-l-4 ${activePanel === item.id ? 'bg-green-900/20 text-white border-green-600' : 'text-zinc-400 border-transparent hover:text-white hover:bg-zinc-900'}`}
          >
            <item.icon size={18} className="mr-3" />
            {item.label}
          </button>
        ))}
      </nav>

      <div className="p-4 border-t border-zinc-800 space-y-3">
        <div className="flex items-center justify-center gap-2 p-2 bg-zinc-900 rounded-lg border border-zinc-800">
          <div className={`w-2 h-2 rounded-full ${firebaseConnected ? 'bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.6)]' : 'bg-red-500'}`} />
          <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-widest">
            Firebase: {firebaseConnected ? 'Connected' : 'Disconnected'}
          </span>
        </div>
        <button 
          onClick={handleLogout}
          className="w-full flex items-center justify-center p-3 text-red-400 hover:bg-red-950/30 rounded-xl transition-all font-bold text-sm border border-red-900/30"
        >
          <LogOut size={16} className="mr-2" />
          Keluar
        </button>
      </div>
    </div>
  );

  const handleSaveSiswa = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingSiswa) return;
    toggleLoader(true);
    try {
      await firestoreService.saveSiswa(editingSiswa);
      setShowSiswaModal(false);
      setEditingSiswa(null);
      setSearchTermSiswa(''); // Reset search to show newcomer
      triggerSuccess("BERHASIL", "Data siswa telah berhasil disimpan. Data akan muncul otomatis di daftar.");
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteSiswa = (nisn: string, nama: string) => {
    setConfirmModal({
      show: true,
      title: 'Hapus data siswa?',
      message: 'Data siswa dan riwayat absensinya akan dihapus permanen dari sistem.',
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusSiswa(nisn);
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      }
    });
  };

  const handleSaveGuru = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingGuru) return;
    toggleLoader(true);
    // Ensure kelas is only for Wali Kelas
    const cleanedGuru = {
      ...editingGuru,
      jabatan: editingGuru.jabatan || 'Guru',
      role: editingGuru.role || 'Guru',
      kelas: editingGuru.jabatan === 'Guru Wali Kelas' ? editingGuru.kelas : ''
    };
    try {
      await firestoreService.saveGuru(cleanedGuru);
      setShowGuruModal(false);
      setEditingGuru(null);
      setSearchTermGuru(''); // Reset search to show latest data
      triggerSuccess("BERHASIL", "Data guru dan akun akses telah diperbarui. Perubahan akan muncul otomatis.");
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteGuru = (nip: string, nama: string) => {
    setConfirmModal({
      show: true,
      title: 'Hapus guru?',
      message: 'Data guru dan akses loginnya akan dihapus permanen.',
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusGuru(nip);
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      }
    });
  };

  const handleSaveKelas = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingKelas) return;
    toggleLoader(true);
    try {
      await firestoreService.saveKelas(editingKelas);
      setShowKelasModal(false);
      setEditingKelas(null);
      setSearchTermKelas(''); // Reset search
      triggerSuccess("BERHASIL", `Kelas ${editingKelas.nama} berhasil disimpan dan sinkron.`);
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteKelas = (nama: string) => {
    setConfirmModal({
      show: true,
      title: 'Hapus kelas?',
      message: 'Data kelas akan dihapus. Siswa yang terdaftar di kelas ini mungkin perlu dipindahkan.',
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusKelas(nama);
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      }
    });
  };

  const renderSiswaModal = () => (
    <AnimatePresence>
      {showSiswaModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setShowSiswaModal(false)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden">
            <div className="bg-green-800 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Form Data Siswa</h2>
            </div>
            <form onSubmit={handleSaveSiswa} className="p-8 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">NISN</label>
                  <input 
                    type="text" 
                    required 
                    disabled={!!(editingSiswa && students.some(s => s.nisn === editingSiswa.nisn && s.nisn !== ''))}
                    value={editingSiswa?.nisn || ''} 
                    onChange={e => setEditingSiswa({...(editingSiswa as Student), nisn: e.target.value})} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed" 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Nama Lengkap</label>
                  <input type="text" required value={editingSiswa?.nama || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), nama: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Jenis Kelamin</label>
                  <select required value={editingSiswa?.jenisKelamin || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), jenisKelamin: e.target.value as 'L' | 'P'})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold">
                    <option value="">Pilih Gender</option>
                    <option value="L">Laki-Laki (L)</option>
                    <option value="P">Perempuan (P)</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Kelas</label>
                  <select required value={editingSiswa?.kelas || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), kelas: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold">
                    <option value="">Pilih</option>
                    {classrooms.map(k => <option key={k.nama} value={k.nama}>{k.nama}</option>)}
                  </select>
                </div>
                <div className="md:grid md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Ayah</label>
                    <input type="text" value={editingSiswa?.ayah || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), ayah: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" placeholder="Nama Ayah" />
                  </div>
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Ibu</label>
                    <input type="text" value={editingSiswa?.ibu || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), ibu: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" placeholder="Nama Ibu" />
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">HP Orang Tua</label>
                  <input type="text" value={editingSiswa?.hp || ''} onChange={e => setEditingSiswa({...(editingSiswa as Student), hp: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Foto Siswa (Data URL)</label>
                  <div className="flex items-center gap-4">
                    {editingSiswa?.foto && <img src={editingSiswa.foto} className="w-12 h-12 rounded-lg object-cover border" alt="Profile" />}
                    <input 
                      type="file" 
                      accept="image/*"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) {
                          const reader = new FileReader();
                          reader.onloadend = () => {
                            setEditingSiswa({...(editingSiswa as Student), foto: reader.result as string});
                          };
                          reader.readAsDataURL(file);
                        }
                      }}
                      className="w-full text-xs" 
                    />
                  </div>
                </div>
              </div>
              <div className="pt-4 flex gap-3">
                <button type="button" onClick={() => setShowSiswaModal(false)} className="flex-1 py-4 font-bold text-zinc-400">Batal</button>
                <button type="submit" className="flex-1 bg-green-800 text-white rounded-2xl py-4 font-bold shadow-xl">Simpan Data</button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const renderGuruModal = () => (
    <AnimatePresence>
      {showGuruModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setShowGuruModal(false)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden">
            <div className="bg-zinc-900 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Form Data Guru</h2>
            </div>
            <form onSubmit={handleSaveGuru} className="p-8 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">NIP</label>
                  <input 
                    type="text" 
                    required 
                    disabled={!!(editingGuru && teachers.some(t => t.nip === editingGuru.nip && t.nip !== ''))}
                    value={editingGuru?.nip || ''} 
                    onChange={e => setEditingGuru({...editingGuru, nip: e.target.value})} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed" 
                  />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Nama Guru</label>
                  <input type="text" required value={editingGuru?.nama || ''} onChange={e => setEditingGuru({...editingGuru, nama: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Jabatan</label>
                  <select 
                    required 
                    value={editingGuru?.jabatan || ''} 
                    onChange={e => {
                      const jabatan = e.target.value as any;
                      setEditingGuru({...editingGuru, jabatan, kelas: jabatan === 'Guru Wali Kelas' ? editingGuru.kelas : ''});
                    }} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  >
                    <option value="">Pilih Jabatan</option>
                    <option value="Kamad">Kamad</option>
                    <option value="Wakamad">Wakamad</option>
                    <option value="Guru">Guru</option>
                    <option value="Guru Wali Kelas">Guru Wali Kelas</option>
                  </select>
                </div>
                {editingGuru?.jabatan === 'Guru Wali Kelas' && (
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Kelas Wali</label>
                    <select 
                      required
                      value={editingGuru?.kelas || ''} 
                      onChange={e => setEditingGuru({...editingGuru, kelas: e.target.value})} 
                      className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                    >
                      <option value="">-- Pilih Kelas --</option>
                      {classrooms.map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                    </select>
                  </div>
                )}
                <div className="col-span-2">
                   <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Username Login</label>
                        <input type="text" required value={editingGuru?.user || ''} onChange={e => setEditingGuru({...editingGuru, user: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Password</label>
                        <input type="text" required value={editingGuru?.pass || ''} onChange={e => setEditingGuru({...editingGuru, pass: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
                      </div>
                   </div>
                </div>
              </div>
              <div className="pt-4 flex gap-3">
                <button type="button" onClick={() => setShowGuruModal(false)} className="flex-1 py-4 font-bold text-zinc-400">Batal</button>
                <button type="submit" className="flex-1 bg-zinc-900 text-white rounded-2xl py-4 font-bold shadow-xl">Simpan Akun</button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const renderKelasModal = () => (
    <AnimatePresence>
      {showKelasModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setShowKelasModal(false)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-md rounded-3xl shadow-2xl overflow-hidden">
            <div className="bg-blue-900 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Konfigurasi Ruang Kelas</h2>
            </div>
            <form onSubmit={handleSaveKelas} className="p-8 space-y-4">
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Nama Kelas</label>
                <input 
                  type="text" 
                  required 
                  disabled={!!(editingKelas && classrooms.some(c => c.nama === (editingKelas.oldNama || editingKelas.nama) && c.nama !== ''))}
                  value={editingKelas?.nama || ''} 
                  onChange={e => setEditingKelas({...editingKelas, nama: e.target.value})} 
                  className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed" 
                  placeholder="Contoh: VII-A" 
                />
              </div>
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Wali Kelas</label>
                <input type="text" value={editingKelas?.wali || ''} onChange={e => setEditingKelas({...editingKelas, wali: e.target.value})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
              </div>
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Jumlah Siswa</label>
                <input type="number" value={editingKelas?.jumlah || ''} onChange={e => setEditingKelas({...editingKelas, jumlah: parseInt(e.target.value)})} className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold" />
              </div>
              <div className="pt-4 flex gap-3">
                <button type="button" onClick={() => setShowKelasModal(false)} className="flex-1 py-4 font-bold text-zinc-400">Tutup</button>
                <button type="submit" className="flex-1 bg-blue-900 text-white rounded-2xl py-4 font-bold shadow-xl">Terapkan</button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const [teachingSessions, setTeachingSessions] = useState<{kelas: string, hari: string, target: number}[]>([]);

  const handleSaveJadwal = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingJadwal?.nip) return;
    toggleLoader(true);
    try {
      // Cleanup existing schedules for this specific teacher, mapel, and class
      // to avoid duplicates if days are changed.
      const teacherSchedules = teachingSchedules.filter(ts => 
        ts.nip === editingJadwal.nip && 
        ts.mapel === editingJadwal.mapel && 
        ts.kelas === editingJadwal.kelas
      );
      for (const old of teacherSchedules) {
        await firestoreService.hapusJadwalMengajar(old.id);
      }

      const g = teachers.find(t => t.nip === editingJadwal.nip);
      // Ensure all sessions have the correct class from editingJadwal
      const finalSessions = teachingSessions.map(s => ({ ...s, kelas: editingJadwal.kelas || '' }));
      
      await firestoreService.saveTeachingBatch(
        editingJadwal.nip, 
        g?.nama || '', 
        finalSessions, 
        editingJadwal.mapel || ''
      );
      setShowJadwalModal(false);
      setTeachingSessions([]);
      triggerSuccess("BERHASIL", "Konfigurasi jam mengajar telah diperbarui.");
    } catch (e) {
      alert("Gagal menyimpan jadwal.");
    } finally {
      toggleLoader(false);
    }
  };

  const downloadPhoto = (dataUrl: string, fileName: string) => {
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = fileName;
    link.click();
  };

  const renderProfilSiswa = () => {
    const teacher = teachers.find(t => t.nip === session?.uid);
    const jabatan = teacher?.jabatan || (session?.role === 'Admin' ? 'Wakamad' : 'Guru');
    const isKamadWakamad = ['Kamad', 'Wakamad'].includes(jabatan) || session?.role === 'Admin';
    
    const waliKelas = session?.kelas;
    
    // Determine which students to show
    let filteredStudents = [];
    let title = "";
    let description = "";

    if (isKamadWakamad) {
      filteredStudents = students.filter(s => studentProfileClassFilter ? s.kelas === studentProfileClassFilter : true);
      title = studentProfileClassFilter ? `Profil Siswa Kelas ${studentProfileClassFilter}` : "Seluruh Profil Siswa";
      description = "Melihat profil siswa untuk seluruh tingkatan/kelas.";
    } else if (waliKelas) {
      filteredStudents = students.filter(s => s.kelas === waliKelas);
      title = `Profil Siswa Kelas ${waliKelas}`;
      description = "Daftar siswa dalam perwakilan kelas Anda.";
    } else {
      return <div className="p-8 text-center text-gray-400">Akses Dibatasi. Anda bukan Wali Kelas atau Wakamad.</div>;
    }

    return (
      <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-6 pb-20">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 md:p-8 rounded-[2.5rem] shadow-sm border border-gray-100 gap-4">
           <div>
              <h2 className="text-xl font-black flex items-center gap-2 uppercase tracking-tight">
                 <GraduationCap className="text-green-700" size={24} /> {title}
              </h2>
              <p className="text-xs text-gray-400 font-bold mt-1 uppercase tracking-widest">{description}</p>
           </div>
           
           <div className="flex flex-wrap items-center gap-4 w-full md:w-auto">
             {isKamadWakamad && (
               <div className="flex items-center gap-2 bg-zinc-50 px-4 py-2 rounded-2xl border border-zinc-100">
                  <TableIcon size={14} className="text-zinc-400" />
                  <select 
                    value={studentProfileClassFilter}
                    onChange={e => setStudentProfileClassFilter(e.target.value)}
                    className="bg-transparent border-0 text-xs font-black uppercase tracking-widest focus:ring-0 text-zinc-600"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                  </select>
               </div>
             )}
             <div className="text-right bg-green-50 px-6 py-2 rounded-2xl border border-green-100">
                <p className="text-[9px] font-black text-green-600 uppercase tracking-widest mb-1">Total Tampil</p>
                <p className="text-xl font-black text-green-800">{filteredStudents.length}</p>
             </div>
           </div>
        </div>

        <div className="bg-white rounded-[2.5rem] shadow-xl border border-gray-100 overflow-hidden">
           <div className="overflow-x-auto">
             <table className="w-full text-left text-sm">
                <thead className="bg-zinc-50 text-zinc-400 uppercase text-[10px] font-black tracking-widest border-b border-zinc-100">
                   <tr>
                      <th className="px-8 py-5">Nama Siswa</th>
                      <th className="px-8 py-5">NISN / Kelas</th>
                      <th className="px-8 py-5 text-center">Aksi Pelayanan</th>
                   </tr>
                </thead>
                <tbody className="divide-y divide-zinc-50">
                   {filteredStudents.map((s, idx) => (
                      <tr key={idx} className="hover:bg-zinc-50 transition-colors group">
                         <td className="px-8 py-5">
                            <div className="flex items-center gap-4">
                               <div className="w-10 h-10 rounded-2xl bg-zinc-100 flex items-center justify-center overflow-hidden border border-zinc-200 group-hover:border-green-300 transition-colors">
                                  {s.foto ? <img src={s.foto} className="w-full h-full object-cover" /> : <User size={18} className="text-zinc-300" />}
                               </div>
                               <div>
                                 <span className="font-black text-zinc-900 block leading-tight">{s.nama}</span>
                                 <span className="text-[10px] font-bold text-zinc-400 uppercase tracking-tighter">Status Aktif</span>
                               </div>
                            </div>
                         </td>
                         <td className="px-8 py-5">
                            <span className="font-bold text-zinc-500 block">{s.nisn}</span>
                            <span className="text-[10px] font-black text-green-700 bg-green-50 px-2 py-0.5 rounded-lg uppercase tracking-widest">{s.kelas}</span>
                         </td>
                         <td className="px-8 py-5 text-center">
                            <button 
                               onClick={() => setSelectedStudentCard(s)}
                               className="bg-zinc-900 text-white px-6 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg hover:shadow-green-100 shadow-zinc-100"
                            >
                               Lihat Detail Profil
                            </button>
                         </td>
                      </tr>
                   ))}
                   {filteredStudents.length === 0 && (
                      <tr>
                         <td colSpan={3} className="p-20 text-center text-zinc-300">
                           <div className="max-w-xs mx-auto">
                              <GraduationCap size={48} className="mx-auto mb-4 opacity-20" />
                              <p className="font-bold uppercase tracking-widest text-[10px]">Data tidak ditemukan</p>
                              <p className="text-xs font-medium mt-1">Belum ada siswa yang terdaftar di pilihan ini.</p>
                           </div>
                         </td>
                      </tr>
                   )}
                </tbody>
             </table>
           </div>
        </div>
      </motion.div>
    );
  };

  const renderBiodataSiswaModal = () => (
    <AnimatePresence>
      {selectedStudentCard && activePanel === 'profil-siswa' && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setSelectedStudentCard(null)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden max-h-[90vh] flex flex-col">
             <div className="p-8 border-b flex justify-between items-center bg-gray-50">
                <h2 className="text-xl font-black text-green-950 uppercase tracking-widest">Detail Profil Siswa</h2>
                <button onClick={() => setSelectedStudentCard(null)} className="text-gray-400 hover:text-red-500 transition-colors">
                   <XCircle size={28} />
                </button>
             </div>
             
             <div className="p-8 overflow-y-auto custom-scrollbar flex-1">
                <div className="flex flex-col md:flex-row gap-8 mb-10 items-center md:items-start">
                   <div className="relative group">
                      <div className="w-40 h-40 rounded-[2.5rem] bg-gray-100 overflow-hidden border-4 border-white shadow-xl relative">
                         {selectedStudentCard.foto ? (
                            <img src={selectedStudentCard.foto} className="w-full h-full object-cover" crossOrigin="anonymous" />
                         ) : (
                            <div className="w-full h-full flex items-center justify-center text-gray-300">
                               <User size={64} />
                            </div>
                         )}
                      </div>
                      {selectedStudentCard.foto && (
                        <button 
                           onClick={() => downloadPhoto(selectedStudentCard.foto!, `Foto-${selectedStudentCard.nama}.png`)}
                           className="absolute -bottom-2 -right-2 bg-green-600 text-white p-3 rounded-2xl shadow-lg border-4 border-white hover:bg-green-700 transition-all scale-90 group-hover:scale-100"
                           title="Download Foto"
                        >
                           <Download size={18} />
                        </button>
                      )}
                   </div>
                   
                   <div className="flex-1 text-center md:text-left">
                      <h3 className="text-2xl font-black text-green-900 mb-2">{selectedStudentCard.nama}</h3>
                      <p className="text-xs font-bold text-green-600 uppercase tracking-widest bg-green-50 px-4 py-1.5 rounded-full inline-block">
                         Kelas {selectedStudentCard.kelas}
                      </p>
                      <div className="mt-6 grid grid-cols-2 gap-4">
                         <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                            <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Jenis Kelamin</p>
                            <p className="font-bold text-sm">{selectedStudentCard.jenisKelamin === 'L' ? 'Laki-Laki' : 'Perempuan'}</p>
                         </div>
                         <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                            <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Tempat Lahir</p>
                            <p className="font-bold text-sm">{selectedStudentCard.tempat || '-'}</p>
                         </div>
                         <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                            <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Tanggal Lahir</p>
                            <p className="font-bold text-sm">{selectedStudentCard.tgl ? formatIndoDate(selectedStudentCard.tgl) : '-'}</p>
                         </div>
                      </div>
                   </div>
                </div>

                <div className="space-y-6">
                   <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                         <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Nama Ayah</label>
                         <input type="text" readOnly value={selectedStudentCard.ayah || '-'} className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed" />
                      </div>
                      <div className="space-y-2">
                         <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Nama Ibu</label>
                         <input type="text" readOnly value={selectedStudentCard.ibu || '-'} className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed" />
                      </div>
                   </div>
                   <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">WhatsApp Orang Tua</label>
                      <input type="text" readOnly value={selectedStudentCard.hp || '-'} className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed" />
                   </div>
                </div>
             </div>
             
             <div className="p-8 border-t bg-gray-50 flex justify-center">
                <button onClick={() => setSelectedStudentCard(null)} className="bg-zinc-900 text-white px-12 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-zinc-800 transition-all shadow-xl">
                   Tutup Profil
                </button>
             </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const renderJadwalModal = () => (
    <AnimatePresence>
      {showJadwalModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setShowJadwalModal(false)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden">
            <div className="bg-green-950 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Konfigurasi Jam Mengajar</h2>
            </div>
            <div className="p-8 space-y-6 max-h-[75vh] overflow-y-auto custom-scrollbar">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Nama Guru</label>
                  <select 
                    required 
                    value={editingJadwal?.nip || ''} 
                    onChange={e => {
                      const t = teachers.find(tg => tg.nip === e.target.value);
                      setEditingJadwal({...(editingJadwal as any), nip: e.target.value, namaGuru: t?.nama || ''});
                    }} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Guru --</option>
                    {teachers.map(t => <option key={t.nip} value={t.nip}>{t.nama}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Mata Pelajaran</label>
                  <select 
                    required 
                    value={editingJadwal?.mapel || ''} 
                    onChange={e => setEditingJadwal({...(editingJadwal as any), mapel: e.target.value})} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Mapel --</option>
                    {(appConfig.subjects || []).map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Kelas</label>
                  <select 
                    required 
                    value={editingJadwal?.kelas || ''} 
                    onChange={e => setEditingJadwal({...(editingJadwal as any), kelas: e.target.value})} 
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Kelas --</option>
                    {classrooms.map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                  </select>
                </div>
              </div>

              <div>
                <div className="flex justify-between items-center mb-3">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">Sesi Mengajar (Hari & Target)</label>
                  <button 
                    onClick={() => setTeachingSessions([...teachingSessions, { kelas: editingJadwal?.kelas || '', hari: 'Senin', target: 2 }])}
                    className="text-[10px] font-black text-blue-600 uppercase flex items-center gap-1 hover:text-blue-800 transition-colors"
                  >
                    <Plus size={10} /> Tambah Hari
                  </button>
                </div>
                
                <div className="space-y-3">
                  {teachingSessions.length === 0 && (
                    <div className="text-center py-6 border-2 border-dashed border-gray-100 rounded-2xl text-gray-400 text-xs italic">
                      Belum ada hari mengajar ditambahkan.
                    </div>
                  )}
                  {teachingSessions.map((session, idx) => (
                    <div key={idx} className="bg-gray-50 p-4 rounded-2xl border border-gray-100 flex items-center gap-4 relative group">
                      <div className="flex-1">
                        <label className="text-[9px] font-black text-gray-400 uppercase mb-1 ml-1 block">Hari</label>
                        <select 
                          value={session.hari} 
                          onChange={e => {
                            const newSess = [...teachingSessions];
                            newSess[idx].hari = e.target.value;
                            setTeachingSessions(newSess);
                          }}
                          className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold focus:ring-1 focus:ring-green-500"
                        >
                          {["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"].map(h => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </div>
                      <div className="w-32">
                        <label className="text-[9px] font-black text-gray-400 uppercase mb-1 ml-1 block">Sesi/Jam</label>
                        <input 
                          type="number" 
                          min={1}
                          value={session.target}
                          onChange={e => {
                            const newSess = [...teachingSessions];
                            newSess[idx].target = parseInt(e.target.value);
                            setTeachingSessions(newSess);
                          }}
                          className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold focus:ring-1 focus:ring-green-500"
                        />
                      </div>
                      <button 
                        onClick={() => setTeachingSessions(teachingSessions.filter((_, i) => i !== idx))}
                        className="text-red-400 hover:text-red-600 transition-colors pt-4"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="pt-4 flex gap-3 sticky bottom-0 bg-white py-4 border-t">
              <button type="button" onClick={() => setShowJadwalModal(false)} className="flex-1 py-4 font-bold text-zinc-400">Batal</button>
              <button 
                onClick={handleSaveJadwal}
                disabled={!editingJadwal?.nip || !editingJadwal?.mapel || !editingJadwal?.kelas || teachingSessions.length === 0}
                className="flex-1 bg-green-900 text-white rounded-2xl py-4 font-bold shadow-xl disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Simpan Konfigurasi
              </button>
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const renderSessionDetailModal = () => (
    <AnimatePresence>
      {showSessionDetail && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
          <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setShowSessionDetail(null)} />
          <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }} className="relative bg-white w-full max-w-sm rounded-[2.5rem] shadow-2xl overflow-hidden border border-gray-100">
            <div className="bg-green-100 p-8 text-center">
              <div className="bg-white w-16 h-16 rounded-3xl flex items-center justify-center mx-auto mb-4 shadow-sm">
                <School className="text-green-800" size={32} />
              </div>
              <h2 className="text-lg font-black text-green-950 uppercase tracking-tight">{showSessionDetail.name}</h2>
              <p className="text-xs font-bold text-green-700/60 uppercase tracking-widest mt-1">Detail Jadwal Mengajar</p>
            </div>
            <div className="p-8">
              <div className="space-y-4">
                {showSessionDetail.mapping.map((s, idx) => (
                  <div key={idx} className="bg-gray-50 p-4 rounded-2xl border border-gray-100 flex justify-between items-center">
                    <div>
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">{s.kelas}</p>
                      <p className="font-bold text-sm text-gray-900">{s.hari}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Target</p>
                      <p className="font-black text-green-800 text-sm">{s.targetPertemuan} Sesi</p>
                    </div>
                  </div>
                ))}
              </div>
              <button 
                onClick={() => setShowSessionDetail(null)}
                className="w-full bg-zinc-900 text-white rounded-2xl py-4 font-black text-xs uppercase tracking-widest mt-8 shadow-xl shadow-zinc-200 hover:bg-black transition-all"
              >
                Tutup Detail
              </button>
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const currentSiswaData = students.find(s => s.nisn === session?.uid);

  const [biodataForm, setBiodataForm] = useState({
    nama: '',
    jenisKelamin: '',
    tempat: '',
    tgl: '',
    ayah: '',
    ibu: '',
    hp: ''
  });

  useEffect(() => {
    if (currentSiswaData) {
      setBiodataForm({
        nama: currentSiswaData.nama || '',
        jenisKelamin: currentSiswaData.jenisKelamin || 'L',
        tempat: currentSiswaData.tempat || '',
        tgl: currentSiswaData.tgl || '',
        ayah: currentSiswaData.ayah || '',
        ibu: currentSiswaData.ibu || '',
        hp: currentSiswaData.hp || ''
      });
    }
  }, [currentSiswaData]);

  // Filtered lists for panels
  const filteredGuru = useMemo(() => {
    if (!searchTermGuru) return teachers;
    const s = searchTermGuru.toLowerCase();
    return teachers.filter(t => t.nama.toLowerCase().includes(s) || t.nip.toLowerCase().includes(s));
  }, [teachers, searchTermGuru]);

  const filteredSiswa = useMemo(() => {
    if (!searchTermSiswa) return students;
    const s = searchTermSiswa.toLowerCase();
    return students.filter(st => st.nama.toLowerCase().includes(s) || st.nisn.toLowerCase().includes(s) || st.kelas.toLowerCase().includes(s));
  }, [students, searchTermSiswa]);

  const filteredKelas = useMemo(() => {
    if (!searchTermKelas) return classrooms;
    const s = searchTermKelas.toLowerCase();
    return classrooms.filter(c => c.nama.toLowerCase().includes(s) || c.wali.toLowerCase().includes(s));
  }, [classrooms, searchTermKelas]);

  const filteredJadwal = useMemo(() => {
    if (!searchTermJadwal) return teachingSchedules;
    const s = searchTermJadwal.toLowerCase();
    return teachingSchedules.filter(ts => ts.namaGuru.toLowerCase().includes(s) || ts.mapel.toLowerCase().includes(s) || ts.kelas.toLowerCase().includes(s));
  }, [teachingSchedules, searchTermJadwal]);

  const groupedJadwal = useMemo(() => {
    const groups: any = {};
    filteredJadwal.forEach(s => {
      const key = `${s.nip}-${s.mapel || 'N/A'}-${s.kelas}`;
      if (!groups[key]) {
        groups[key] = {
          nip: s.nip,
          namaGuru: s.namaGuru,
          mapel: s.mapel,
          kelas: s.kelas,
          sessions: []
        };
      }
      groups[key].sessions.push(s);
    });
    return Object.values(groups);
  }, [filteredJadwal]);

  const renderSuccessToast = () => (
    <AnimatePresence>
      {showSuccessToast && (
        <motion.div 
          initial={{ opacity: 0, y: 100, scale: 0.8 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          exit={{ opacity: 0, y: 20, scale: 0.8 }}
          className="fixed bottom-10 left-1/2 -translate-x-1/2 z-[300] bg-zinc-900 border border-zinc-800 text-white px-8 py-4 rounded-[2rem] shadow-2xl flex items-center gap-4 min-w-[300px]"
        >
          <div className="bg-green-500 p-2 rounded-xl text-white">
            <CheckCircle2 size={24} />
          </div>
          <div>
            <p className="text-[10px] font-black uppercase tracking-widest text-zinc-400 leading-none mb-1">{showSuccessToast.title}</p>
            <p className="text-sm font-bold">{showSuccessToast.message}</p>
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 overflow-hidden rounded-b-[2rem]">
             <motion.div 
               initial={{ width: "100%" }}
               animate={{ width: "0%" }}
               transition={{ duration: 3, ease: "linear" }}
               className="h-full bg-green-500"
             />
          </div>
        </motion.div>
      )}
    </AnimatePresence>
  );

  if (!session) return renderLogin();

  return (
    <div className="flex min-h-screen bg-gray-50 text-zinc-900 font-sans pb-20 md:pb-0">
      {renderQuotaNoticeModal()}
      {loading && (
        <div className="fixed inset-0 bg-white/80 backdrop-blur-sm z-[9999] flex items-center justify-center">
          <div className="flex flex-col items-center">
            <RefreshCcw className="animate-spin text-green-700 mb-4" size={48} />
            <p className="font-bold text-green-900">Sinkronisasi Data...</p>
          </div>
        </div>
      )}

      {/* Mobile Sidebar Drawer */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <div className="fixed inset-0 z-[110] md:hidden">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsMobileMenuOpen(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ x: '-100%' }}
              animate={{ x: 0 }}
              exit={{ x: '-100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="absolute top-0 bottom-0 left-0 w-[280px] bg-zinc-950 flex flex-col shadow-2xl"
            >
              <div className="p-6 bg-green-900 flex flex-col items-center border-b border-green-800 relative">
                <button 
                  onClick={() => setIsMobileMenuOpen(false)}
                  className="absolute top-4 right-4 text-white/50 hover:text-white"
                >
                  <XCircle size={24} />
                </button>
                <div className="bg-white w-20 h-20 rounded-full flex items-center justify-center mb-3 overflow-hidden p-3 relative border-4 border-green-700/20 shadow-xl backdrop-blur-sm">
                  <img 
                    src={appLogo} 
                    alt="Logo" 
                    className="w-full h-full object-contain relative z-10" 
                    onError={(e) => {
                      const target = e.currentTarget;
                      if (target.src === defaultLogo) {
                        target.src = fallbackLogo;
                      } else if (target.src === fallbackLogo) {
                        target.style.opacity = '0';
                        const parent = target.parentElement;
                        if (parent) {
                          const icon = parent.querySelector('.fallback-icon');
                          if (icon) (icon as HTMLElement).style.opacity = '1';
                        }
                      } else {
                         target.src = defaultLogo;
                      }
                    }}
                  />
                  <School size={36} className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity" />
                </div>
                <h2 className="font-black text-lg text-white">SIGAP MTsN 2</h2>
                <span className="bg-orange-500 text-black text-[10px] px-3 py-1 rounded-full font-black mt-2 uppercase tracking-widest">
                  {session?.role}
                </span>
              </div>

              <nav className="flex-grow py-6 overflow-y-auto">
                {menuItems.map(item => (
                  <button
                    key={item.id}
                    onClick={() => { setActivePanel(item.id); setIsMobileMenuOpen(false); }}
                    className={`w-full flex items-center px-8 py-4 transition-all text-sm font-bold border-l-4 ${activePanel === item.id ? 'bg-green-900/20 text-green-400 border-green-500' : 'text-zinc-500 border-transparent'}`}
                  >
                    <item.icon size={20} className="mr-4" />
                    {item.label}
                  </button>
                ))}
              </nav>

              <div className="p-6 border-t border-zinc-900">
                <button 
                  onClick={handleLogout}
                  className="w-full flex items-center justify-center p-4 bg-red-950/20 text-red-500 rounded-2xl font-black text-sm uppercase tracking-widest border border-red-900/30"
                >
                  <LogOut size={18} className="mr-3" />
                  Logout
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {renderSidebar()}
      
      {/* Mobile Bottom Navigation */}
      <div className="fixed bottom-0 left-0 right-0 bg-zinc-950 text-white z-50 flex justify-around items-center p-2 border-t border-zinc-800 md:hidden">
        {menuItems.slice(0, 4).map(item => (
          <button
            key={item.id}
            onClick={() => setActivePanel(item.id)}
            className={`flex flex-col items-center p-2 rounded-xl transition-all ${activePanel === item.id ? 'text-green-500' : 'text-zinc-500'}`}
          >
            <item.icon size={20} />
            <span className="text-[8px] font-bold mt-1 uppercase tracking-tighter truncate w-12 text-center">
              {item.id === 'guru' ? 'Guru' : item.id === 'siswa' ? 'Siswa' : item.label.split(' ')[0]}
            </span>
          </button>
        ))}
        {menuItems.length > 4 && (
          <div className="relative">
            <button 
              onClick={() => setMobileExtraMenuOpen(!mobileExtraMenuOpen)}
              className={`flex flex-col items-center p-2 transition-all ${mobileExtraMenuOpen ? 'text-green-500' : 'text-zinc-500'}`}
            >
              <TableIcon size={20} />
              <span className="text-[8px] font-bold mt-1 uppercase tracking-tighter">Lainnya</span>
            </button>
            
            <AnimatePresence>
              {mobileExtraMenuOpen && (
                <>
                  <motion.div 
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    onClick={() => setMobileExtraMenuOpen(false)}
                    className="fixed inset-0 bg-black/40 z-[-1]"
                  />
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.9, y: 20 }}
                    animate={{ opacity: 1, scale: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 0.9, y: 20 }}
                    className="absolute bottom-full right-0 mb-4 bg-zinc-900 rounded-2xl shadow-2xl border border-zinc-800 p-2 w-56 overflow-hidden"
                  >
                    {menuItems.slice(4).map(item => (
                      <button
                        key={item.id}
                        onClick={() => { setActivePanel(item.id); setMobileExtraMenuOpen(false); }}
                        className={`w-full flex items-center px-4 py-4 text-xs font-bold rounded-xl transition-all ${activePanel === item.id ? 'bg-green-900 text-white' : 'text-zinc-400 hover:bg-zinc-800'}`}
                      >
                        <item.icon size={16} className="mr-3 text-green-500" />
                        {item.label}
                      </button>
                    ))}
                    <div className="border-t border-zinc-800 my-2 pt-2">
                      <button 
                        onClick={() => { handleLogout(); setMobileExtraMenuOpen(false); }}
                        className="w-full flex items-center px-4 py-4 text-xs font-bold text-red-400 hover:bg-red-950/30 rounded-xl transition-all"
                      >
                        <LogOut size={16} className="mr-3" />
                        Keluar Aplikasi
                      </button>
                    </div>
                  </motion.div>
                </>
              )}
            </AnimatePresence>
          </div>
        )}
      </div>

      {renderSiswaModal()}
      {renderGuruModal()}
      {renderKelasModal()}
      {renderJadwalModal()}
      {renderBiodataSiswaModal()}

      <main className="flex-grow overflow-x-hidden p-6">
        <header className="flex justify-between items-center mb-8 bg-white p-4 rounded-2xl shadow-sm border border-gray-100">
          <div className="flex items-center gap-2">
            <button 
              onClick={() => setIsMobileMenuOpen(true)}
              className="p-2 bg-green-50 text-green-700 rounded-xl md:hidden"
            >
              <LayoutGrid size={20} />
            </button>
            <div className="flex items-center gap-3">
              <h1 className="text-xl font-extrabold text-green-900">{menuItems.find(m => m.id === activePanel)?.label || 'Dashboard'}</h1>
              <button 
                className={`p-2 hover:bg-gray-100 rounded-xl transition-all ${refreshing ? 'text-green-600' : 'text-gray-400'}`} 
                onClick={() => triggerCurrentPanelRefresh()}
              >
                <RefreshCcw size={16} className={`${refreshing ? 'animate-spin' : ''}`} />
              </button>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">Selamat Datang</p>
              <p className="font-bold text-green-800">{session?.name || 'User'}</p>
            </div>
            <div className="w-10 h-10 bg-green-100 rounded-full flex items-center justify-center text-green-700 font-bold border-2 border-green-200">
              {(session?.name || 'U')[0]}
            </div>
          </div>
        </header>

        <AnimatePresence mode="wait">
          {activePanel === 'home' && (
            <motion.div key="home" initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-8">
              {/* Stats Grid */}
              <div className="grid grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-4">
                {[
                  { label: 'Total Siswa', value: stats.siswaCount, icon: GraduationCap, color: 'bg-blue-50 text-blue-600', sub: `L: ${stats.siswaL} / P: ${stats.siswaP}` },
                  { label: 'Total Guru', value: stats.guruCount, icon: UserPlus, color: 'bg-zinc-100 text-zinc-600' },
                  { label: 'Hadir', value: stats.hadirHariIni, icon: ClipboardList, color: 'bg-green-50 text-green-600' },
                  { label: 'Terlambat', value: stats.terlambatHariIni, icon: Clock, color: 'bg-red-50 text-red-600' },
                  { label: 'Sakit', value: stats.sakitCount, icon: Info, color: 'bg-amber-50 text-amber-600' },
                  { label: 'Izin', value: stats.izinCount, icon: Info, color: 'bg-indigo-50 text-indigo-600' },
                  { label: 'Alfa', value: stats.alfaCount, icon: XCircle, color: 'bg-rose-50 text-rose-600' }
                ].map((s, i) => (
                  <div key={i} className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 relative overflow-hidden group">
                    <div className={`${s.color} w-10 h-10 rounded-xl flex items-center justify-center mb-3 group-hover:scale-110 transition-transform`}>
                      <s.icon size={20} />
                    </div>
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">{s.label}</p>
                    <div className="flex items-baseline gap-2">
                       <h3 className="text-2xl font-black mt-1">{s.value}</h3>
                       {s.sub && <span className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter">{s.sub}</span>}
                    </div>
                  </div>
                ))}
              </div>

              {/* Quick Menu for Mobile */}
              <div className="md:hidden space-y-4">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-black text-green-900 uppercase tracking-widest">Menu Utama</h3>
                </div>
                <div className="grid grid-cols-3 gap-3">
                  {menuItems.map(item => (
                    <button
                      key={item.id}
                      onClick={() => setActivePanel(item.id)}
                      className={`bg-white p-4 rounded-2xl border border-gray-100 shadow-sm flex flex-col items-center justify-center transition-all ${activePanel === item.id ? 'ring-2 ring-green-600' : ''}`}
                    >
                      <div className="bg-green-50 p-2 rounded-xl text-green-700 mb-2">
                        <item.icon size={20} />
                      </div>
                      <span className="text-[9px] font-bold text-gray-600 text-center leading-tight">{item.label}</span>
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                 <div className="lg:col-span-2 bg-white p-8 rounded-[2rem] shadow-sm border border-gray-100">
                    <div className="flex justify-between items-center mb-6">
                       <div>
                          <h3 className="text-sm font-black text-green-900 uppercase tracking-widest">Grafik Kehadiran</h3>
                          <p className="text-[10px] text-gray-400 font-bold">Statistik kehadiran 7 hari terakhir</p>
                       </div>
                       <BarChart3 className="text-green-600" size={20} />
                    </div>
                    {stats.chartData ? (
                      <AttendanceChart data={stats.chartData} />
                    ) : (
                      <div className="h-64 flex items-center justify-center text-gray-300 italic text-xs">Memuat data grafik...</div>
                    )}
                 </div>
                 
                 <div className="bg-green-950 p-8 rounded-[2rem] text-white flex flex-col justify-between overflow-hidden relative group">
                    <div className="relative z-10">
                       <h3 className="text-lg font-black leading-tight mb-2">Selamat Datang,<br/>{session?.name}!</h3>
                       <p className="text-green-300 text-[10px] font-bold uppercase tracking-widest">Akses Panel {session?.role}</p>
                    </div>
                    
                    <div className="mt-8 space-y-4 relative z-10">
                       <div className="bg-white/10 p-4 rounded-2xl backdrop-blur-sm border border-white/10">
                          <p className="text-[10px] font-black text-green-200 uppercase tracking-widest mb-1">Status Sistem</p>
                          <div className="flex items-center gap-2">
                             <div className="w-2 h-2 rounded-full bg-green-400 animate-pulse"></div>
                             <span className="text-xs font-bold">Online / Real-time</span>
                          </div>
                       </div>
                    </div>
                    
                    <School size={120} className="absolute -bottom-6 -right-6 text-white/5 group-hover:scale-110 transition-transform duration-500" />
                 </div>
              </div>
            </motion.div>
          )}

          {activePanel === 'scanner' && (
            <motion.div key="scanner" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col items-center">
              <div className="text-center mb-6">
                <div className="inline-block px-4 py-2 bg-white rounded-full shadow-sm border border-zinc-100 mb-2">
                  <span className="text-xs font-black text-green-900 uppercase tracking-widest flex items-center gap-2">
                    <Calendar size={14} />
                    {new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                  </span>
                </div>
                <h2 className="text-2xl font-black text-green-950">Presensi Digital</h2>
                <p className="text-zinc-500 text-sm font-medium">Arahkan barcode kartu ke kamera</p>
              </div>

              <div className="bg-white p-6 rounded-3xl shadow-xl w-full max-w-lg">
                <div className="flex gap-2 p-1 bg-gray-100 rounded-xl mb-4">
                  <button 
                    onClick={() => { setScanType('Siswa'); setScanResult(null); setScanning(false); }}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${scanType === 'Siswa' ? 'bg-green-800 text-white shadow' : 'text-gray-500'}`}
                  >
                    Scan Siswa
                  </button>
                  <button 
                    onClick={() => { setScanType('Kelas'); setScanResult(null); setScanning(false); }}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${scanType === 'Kelas' ? 'bg-blue-800 text-white shadow' : 'text-gray-500'}`}
                  >
                    Scan Kelas
                  </button>
                </div>

                <div className="aspect-square bg-zinc-900 rounded-2xl flex flex-col items-center justify-center relative overflow-hidden mb-4 group ring-8 ring-gray-50">
                  {scanning ? (
                    <div id="reader" className="w-full h-full"></div>
                  ) : (
                    <>
                      <QrCode size={120} className="text-white/20 group-hover:scale-110 transition-transform duration-500" />
                      <div className="absolute inset-0 border-2 border-dashed border-green-500/50 m-8 rounded-2xl animate-pulse"></div>
                      <div className="absolute top-4 left-1/2 -translate-x-1/2 bg-white/10 backdrop-blur-md px-3 py-1 rounded-full text-[10px] text-white font-bold uppercase tracking-widest border border-white/10">
                        Mode: {scanType === 'Siswa' ? 'Absensi Siswa' : 'Absensi Guru (Mengajar)'}
                      </div>
                      <button 
                        onClick={() => setScanning(true)}
                        className="absolute bottom-6 bg-green-500 px-6 py-2 rounded-xl text-xs text-white font-black uppercase tracking-widest shadow-lg animate-bounce flex items-center gap-2 hover:bg-green-600 transition-all"
                      >
                        <Maximize2 size={16} /> Aktifkan Kamera
                      </button>
                    </>
                  )}
                </div>

                <div className="mb-4">
                  <div className="flex items-center justify-between mb-1 ml-1">
                    <label className="block text-[10px] font-black text-gray-400 uppercase">Input Manual</label>
                    {scanning && (
                      <button onClick={() => setScanning(false)} className="text-[10px] font-black text-red-500 uppercase">Matikan Kamera</button>
                    )}
                  </div>
                  <div className="flex gap-2">
                    <input 
                      type="text" 
                      value={scanInput}
                      onChange={e => setScanInput(e.target.value)}
                      placeholder={scanType === 'Siswa' ? 'Contoh NISN: 12345' : 'Masukkan Nama Kelas'}
                      className="flex-grow bg-gray-100 border-0 rounded-xl px-4 py-3 text-sm font-bold focus:ring-2 focus:ring-green-600 transition-all"
                    />
                    <button 
                      onClick={handleManualScan}
                      className="bg-green-800 text-white px-6 py-3 rounded-xl font-bold text-xs hover:bg-green-700 transition-all"
                    >
                      Kirim
                    </button>
                  </div>
                </div>

                {scanResult && (
                  <motion.div 
                    initial={{ scale: 0.95, opacity: 0 }} 
                    animate={{ scale: 1, opacity: 1 }}
                    className={`p-4 rounded-2xl border-2 flex flex-col gap-2`} style={{
                      ...(scanResult.status === 'Alfa' ? { backgroundColor: '#fff1f2', borderColor: '#fecdd3', color: '#9f1239' } :
                          scanResult.success ? { backgroundColor: '#f0fdf4', borderColor: '#bbf7d0', color: '#166534' } : 
                          { backgroundColor: '#fef2f2', borderColor: '#fecaca', color: '#991b1b' })
                    }}
                  >
                    <div className="flex items-center gap-3">
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center ${scanResult.status === 'Alfa' ? 'bg-rose-500' : scanResult.success ? 'bg-green-500' : 'bg-red-500'} text-white`}>
                        {scanResult.success ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
                      </div>
                      <div className="flex-1">
                        <p className="font-bold">
                          {scanResult.status === 'Alfa' ? 'Gagal: Sudah Jam Pulang' : scanResult.success ? 'Berhasil Terdaftar' : 'Gagal Memproses'}
                        </p>
                        <p className="text-xs opacity-75">{scanResult.message}</p>
                      </div>
                    </div>
                    {scanResult.success && scanResult.message.includes('Terlambat') && (
                       <div className="bg-orange-500 text-white p-3 rounded-xl flex items-center gap-3 animate-bounce">
                          <Clock size={20} className="animate-spin-slow" />
                          <div>
                             <p className="text-xs font-black uppercase tracking-widest">Peringatan Terlambat!</p>
                             <p className="text-[10px] font-bold">Waktu Anda akan tercatat dalam akumulasi keterlambatan.</p>
                          </div>
                       </div>
                    )}
                  </motion.div>
                )}
              </div>
              <p className="mt-4 text-xs text-zinc-400 font-medium">Pastikan QR Code berada di tengah kotak pemindai.</p>
            </motion.div>
          )}

          {activePanel === 'siswa' && (
            <motion.div key="siswa" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="mb-4 flex flex-wrap gap-3 justify-between items-center">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input 
                      type="text" 
                      placeholder="Cari Siswa..." 
                      value={searchTermSiswa}
                      onChange={e => { setSearchTermSiswa(e.target.value); setPagination({...pagination, siswa: 0}); }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500" 
                    />
                  </div>
                  <select 
                    value={filterAdminSiswaClass} 
                    onChange={e => setFilterAdminSiswaClass(e.target.value)}
                    className="bg-white border border-gray-100 rounded-xl px-4 py-2 text-sm font-bold text-gray-700"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                  </select>
                </div>
                <div className="flex gap-2">
                  <select 
                    value={pageSize} 
                    onChange={e => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                  </select>
                  <button 
                    onClick={() => downloadTemplate('Siswa')}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-zinc-200 transition-all"
                  >
                    <Download size={14} /> Template
                  </button>
                  <button 
                    onClick={() => handleImportExcel('Siswa')}
                    className="bg-blue-100 text-blue-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-blue-200 transition-all"
                  >
                    <Upload size={14} /> Import
                  </button>
                  <button 
                    onClick={() => { setEditingSiswa({ nisn: '', nama: '', jenisKelamin: 'L', tempat: '', tgl: '', kelas: '', ayah: '', ibu: '', hp: '' }); setShowSiswaModal(true); }} 
                    className="bg-green-800 text-white px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-green-700 transition-all shadow-sm"
                  >
                    <Plus size={14} /> Tambah
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden text-zinc-900">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm min-w-[800px]">
                    <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                      <tr>
                        <th className="px-6 py-4 text-center w-12">No</th>
                        <th className="px-6 py-4">Foto</th>
                        <th className="px-6 py-4">NISN</th>
                        <th className="px-6 py-4">Nama</th>
                        <th className="px-6 py-4">Kelas</th>
                        <th className="px-6 py-4">Orang Tua</th>
                        <th className="px-6 py-4">Kontak</th>
                        <th className="px-6 py-4">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {filteredSiswa.filter(s => filterAdminSiswaClass ? s.kelas === filterAdminSiswaClass : true).slice(pagination.siswa, pagination.siswa + pageSize).map((s, i) => (
                        <tr key={s.nisn} className="hover:bg-gray-50 transition-colors group">
                          <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{pagination.siswa + i + 1}</td>
                          <td className="px-6 py-4">
                          {s.foto ? (
                            <img src={s.foto} className="w-8 h-8 rounded-lg object-cover ring-2 ring-gray-100" alt={s.nama} />
                          ) : (
                            <div className="w-8 h-8 rounded-lg bg-gray-100 flex items-center justify-center text-gray-400">
                              <User size={14} />
                            </div>
                          )}
                        </td>
                        <td className="px-6 py-4 font-mono font-bold text-gray-400">{s.nisn}</td>
                        <td className="px-6 py-4 font-bold">{s.nama}</td>
                        <td className="px-6 py-4"><span className="text-[10px] font-bold">{s.kelas}</span></td>
                        <td className="px-6 py-4">
                           <p className="text-[10px] font-medium text-gray-500">A: {s.ayah || '-'}</p>
                           <p className="text-[10px] font-medium text-gray-500">I: {s.ibu || '-'}</p>
                        </td>
                        <td className="px-6 py-4">{s.hp}</td>
                        <td className="px-6 py-4">
                          <div className="flex gap-2">
                             <button onClick={() => setSelectedStudentCard(s)} className="text-green-600 hover:text-green-800"><QrCode size={16} /></button>
                             <button onClick={() => { setEditingSiswa(s); setShowSiswaModal(true); }} className="text-blue-500 hover:text-blue-800"><Edit size={16} /></button>
                             <button onClick={() => handleDeleteSiswa(s.nisn, s.nama)} className="text-red-500 hover:text-red-800"><Trash2 size={16} /></button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
                  {filteredSiswa.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.siswa / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.siswa === 0} onClick={() => setPagination({...pagination, siswa: pagination.siswa - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.siswa + pageSize >= filteredSiswa.length} onClick={() => setPagination({...pagination, siswa: pagination.siswa + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'analisis' && (
            <motion.div key="analisis" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-8">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-[2rem] shadow-sm border border-zinc-100 mb-8 text-zinc-900">
                  <div className="flex flex-col">
                    <h1 className="text-2xl font-black text-green-950">Analisis Kehadiran</h1>
                    <p className="text-zinc-500 font-medium">Monitoring pergerakan data presensi sekolah</p>
                  </div>
                  <div className="flex items-center gap-3">
                    <select 
                      value={pageSize} 
                      onChange={e => setPageSize(parseInt(e.target.value))}
                      className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                    >
                      {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                    </select>
                    <div className="relative flex items-center gap-2">
                       {session?.role === 'Guru' && session?.isWali && (
                         <span className="absolute -top-3 right-0 bg-red-100 text-red-600 text-[8px] font-black px-2 py-0.5 rounded-full uppercase tracking-widest border border-red-200">Terkunci</span>
                       )}
                       <select 
                         value={analysisClass} 
                         onChange={e => setAnalysisClass(e.target.value)}
                         disabled={session?.role === 'Guru' && session?.isWali}
                         className="bg-zinc-50 border-0 rounded-xl px-4 py-2 font-bold text-sm disabled:opacity-50 disabled:cursor-not-allowed min-w-[120px]"
                       >
                         {!(session?.role === 'Guru' && session?.isWali) && <option value="">Semua Kelas</option>}
                         {(classrooms || [])
                           .filter(c => (session?.role === 'Guru' && session?.isWali) ? c.nama === session?.kelas : true)
                           .map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                       </select>
                    </div>
                  </div>
                </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Siswa Attendance Analysis */}
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 min-h-[400px]">
                   <h3 className="text-sm font-black text-green-900 uppercase tracking-widest mb-6">
                     {session?.role === 'Guru' && session?.isWali ? `Analisis Kehadiran Siswa Kelas ${session.kelas}` : 'Analisis Kehadiran Siswa Per Kelas'}
                   </h3>
                   <div className="h-64 relative">
                      {classrooms.length > 0 ? (
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={classrooms.filter(c => analysisClass ? c.nama === analysisClass : true).map(c => {
                            const currentMonthPrefix = currentDate.slice(0, 7);
                            const classAttendance = (attendance || []).filter(a => a.kelas === c.nama && a.tanggal.startsWith(currentMonthPrefix));
                            const totalSiswa = (students || []).filter(s => s.kelas === c.nama).length;
                            
                            const daysPassed = effectiveDaysThisMonth.length;
                            const totalPeluang = totalSiswa * daysPassed;
                            const hadirCount = classAttendance.filter(a => a.status === 'Hadir' || a.status === 'Izin' || a.status === 'Sakit').length;
                            return {
                              name: c.nama || 'Kelas',
                              percentage: totalPeluang > 0 ? Math.round((hadirCount / totalPeluang) * 100) : 0
                            };
                          })}>
                             <CartesianGrid strokeDasharray="3 3" vertical={false} />
                             <XAxis dataKey="name" fontSize={10} axisLine={false} tickLine={false} />
                             <YAxis domain={[0, 100]} fontSize={10} axisLine={false} tickLine={false} />
                             <Tooltip />
                             <Bar dataKey="percentage" fill="#16a34a" radius={[4, 4, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      ) : (
                        <div className="h-full flex items-center justify-center text-zinc-400 text-xs italic">Memuat data kelas...</div>
                      )}
                   </div>
                   <div className="mt-6">
                      <h4 className="text-[10px] font-black text-red-600 uppercase tracking-widest mb-3">Siswa Kehadiran &lt; 80% ({new Date().toLocaleDateString('id-ID', { month: 'long', year: 'numeric' })})</h4>
                      <div className="max-h-48 overflow-y-auto space-y-2">
                        {(() => {
                           const daysCount = effectiveDaysThisMonth.length;
                           if (daysCount === 0) return (
                             <div className="text-center py-4 text-gray-400 text-[10px] font-bold uppercase italic">Belum ada hari sekolah bulan ini</div>
                           );

                           const filteredStudents = (students || []).filter(s => analysisClass ? s.kelas === analysisClass : true);
                           const lowAttendanceSiswa = filteredStudents.filter(s => {
                             const currentMonthString = new Date().toISOString().slice(0, 7);
                             const studentAttendance = (attendance || []).filter(a => a.nisn === s.nisn && a.tanggal.startsWith(currentMonthString));
                             const hadir = studentAttendance.filter(a => ['Hadir', 'Izin', 'Sakit'].includes(a.status)).length;
                             const perc = (hadir / daysCount) * 100;
                             return perc < 80;
                           });

                           if (lowAttendanceSiswa.length === 0) return (
                             <div className="text-center py-4 text-gray-400 text-[10px] font-bold uppercase italic">Semua siswa tertib presensi ✨</div>
                           );

                           return lowAttendanceSiswa.map(s => {
                             const currentMonthString = new Date().toISOString().slice(0, 7);
                             const studentAttendance = (attendance || []).filter(a => a.nisn === s.nisn && a.tanggal.startsWith(currentMonthString));
                             const hadir = studentAttendance.filter(a => ['Hadir', 'Izin', 'Sakit'].includes(a.status)).length;
                             const perc = Math.round((hadir / daysCount) * 100);
                             
                             return (
                               <div key={s.nisn} className="flex justify-between items-center p-2 bg-red-50 rounded-lg border border-red-100">
                                 <div className="flex flex-col">
                                   <span className="text-xs font-bold">{s.nama}</span>
                                   <span className="text-[10px] font-medium text-gray-500">{s.kelas} • {hadir} dari {daysCount} hari sekolah</span>
                                 </div>
                                 <span className="text-xs font-black text-red-600">{perc}%</span>
                               </div>
                             );
                           });
                        })()}
                      </div>
                   </div>
                </div>

                {/* Teacher Attendance Analysis */}
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 min-h-[400px]">
                   <h3 className="text-sm font-black text-blue-900 uppercase tracking-widest mb-6">
                     {session?.role === 'Guru' && session?.isWali ? `Analisis Kehadiran Guru Pengajar di ${session.kelas}` : 'Analisis Kehadiran Guru'}
                   </h3>
                   <div className="h-64 relative">
                      {teachers.length > 0 ? (
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={teachers.filter(t => {
                             if (!analysisClass) return true;
                             const scheds = (teachingSchedules || []).filter(ts => ts.nip === t.nip);
                             return scheds.some(s => s.kelas === analysisClass);
                           }).map(t => {
                             const schedules = (teachingSchedules || []).filter(ts => ts.nip === t.nip && (analysisClass ? ts.kelas === analysisClass : true));
                             const now = new Date();
                             const year = now.getFullYear();
                             const month = now.getMonth();
                             const totalTarget = (schedules || []).reduce((sum, s) => {
                               const occ = countDaysInMonth(year, month, s.hari, holidays);
                               return sum + (occ * (Number(s.targetPertemuan) || 0));
                             }, 0);
                             const currentMonth = now.toISOString().slice(0, 7);
                             const actual = (teacherAttendance || []).filter(ta => ta.nip === t.nip && (analysisClass ? ta.kelas === analysisClass : true) && ta.tanggal.startsWith(currentMonth)).length;
                             return {
                               name: (t.nama || 'Guru').split(' ')[0],
                               percentage: totalTarget > 0 ? Math.round((actual / totalTarget) * 100) : 0
                             };
                           })}>
                             <CartesianGrid strokeDasharray="3 3" vertical={false} />
                             <XAxis dataKey="name" fontSize={10} axisLine={false} tickLine={false} />
                             <YAxis domain={[0, 100]} fontSize={10} axisLine={false} tickLine={false} />
                             <Tooltip />
                             <Bar dataKey="percentage" fill="#2563eb" radius={[4, 4, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      ) : (
                        <div className="h-full flex items-center justify-center text-zinc-400 text-xs italic">Memuat data guru...</div>
                      )}
                   </div>
                   <div className="mt-6">
                      <h4 className="text-[10px] font-black text-red-600 uppercase tracking-widest mb-3">
                         {session?.role === 'Guru' && session?.isWali ? `Performa Guru di ${session.kelas} < 50%` : 'Guru Performa < 50% (Bulan Ini)'}
                       </h4>
                      <div className="max-h-48 overflow-y-auto space-y-2">
                        {(teachers || []).filter(t => {
                           if (!analysisClass) return true;
                           const scheds = (teachingSchedules || []).filter(ts => ts.nip === t.nip);
                           return scheds.some(s => s.kelas === analysisClass);
                        }).map(t => {
                          const schedules = (teachingSchedules || []).filter(ts => ts.nip === t.nip && (analysisClass ? ts.kelas === analysisClass : true));
                          const now = new Date();
                          const year = now.getFullYear();
                          const month = now.getMonth();
                          const totalTarget = (schedules || []).reduce((sum, s) => {
                            const occ = countDaysInMonth(year, month, s.hari, holidays, now.getDate());
                            return sum + (occ * (Number(s.targetPertemuan) || 1));
                          }, 0);
                          const currentMonth = now.toISOString().slice(0, 7);
                          const actual = (teacherAttendance || []).filter(ta => ta.nip === t.nip && (analysisClass ? ta.kelas === analysisClass : true) && ta.tanggal.startsWith(currentMonth)).length;
                          const perc = totalTarget > 0 ? (actual / totalTarget) * 100 : 0;
                          if (perc < 50 && totalTarget > 0) {
                            return (
                              <div key={t.nip} className="flex justify-between items-center p-2 bg-red-50 rounded-lg border border-red-100">
                                <span className="text-xs font-bold">{t.nama}</span>
                                <span className="text-xs font-black text-red-600">{Math.round(perc)}%</span>
                              </div>
                            );
                          }
                          return null;
                        })}
                      </div>
                   </div>
                </div>
              </div>
            </motion.div>
          )}

          {activePanel === 'jadwal-mengajar' && (
            <motion.div key="jadwal-mengajar" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input 
                      type="text" 
                      placeholder="Cari Jadwal..." 
                      value={searchTermJadwal}
                      onChange={e => { setSearchTermJadwal(e.target.value); setPagination({...pagination, jadwal: 0}); }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500" 
                    />
                  </div>
                </div>
                <div className="flex gap-2">
                  <select 
                    value={pageSize} 
                    onChange={e => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                  </select>
                  <button 
                    onClick={() => { setEditingJadwal({ id: '', nip: '', namaGuru: '', kelas: '', hari: '', targetPertemuan: 2 } as any); setTeachingSessions([]); setShowJadwalModal(true); }}
                    className="bg-green-800 text-white px-4 py-2 rounded-xl text-xs font-bold flex items-center gap-2 shadow-sm"
                  >
                    <Plus size={14} /> Tambah Jam Mengajar
                  </button>
                </div>
              </div>
              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                    <tr>
                      <th className="px-6 py-4 text-center w-12 text-nowrap">No</th>
                      <th className="px-6 py-4">Guru</th>
                      <th className="px-6 py-4">Mapel</th>
                      <th className="px-6 py-4">Kelas</th>
                      <th className="px-6 py-4">Total Target / Bln</th>
                      <th className="px-6 py-4 text-center">Aksi</th>
                    </tr>
                  </thead>
                    <tbody className="divide-y divide-gray-50">
                    {(() => {
                      const now = new Date();
                      const year = now.getFullYear();
                      const month = now.getMonth();
                      
                      return groupedJadwal.slice(pagination.jadwal, pagination.jadwal + pageSize).map((g: any, i) => {
                        const totalTarget = g.sessions.reduce((sum: number, s: any) => {
                          const occ = countDaysInMonth(year, month, s.hari, holidays);
                          return sum + (occ * (Number(s.targetPertemuan) || 0));
                        }, 0);

                        return (
                          <tr key={i}>
                            <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{pagination.jadwal + i + 1}</td>
                            <td className="px-6 py-4">
                               <p className="font-bold">{g.namaGuru}</p>
                               <p className="text-[9px] font-mono font-bold text-gray-400 uppercase">{g.nip}</p>
                            </td>
                            <td className="px-6 py-4 font-bold text-green-950 text-xs">{g.mapel || '-'}</td>
                            <td className="px-6 py-4">
                               <div className="flex flex-col gap-1">
                                  <span className="font-bold">{g.kelas}</span>
                                  <button 
                                    onClick={() => setShowSessionDetail({ name: g.namaGuru, mapping: g.sessions })}
                                    className="text-[9px] font-black text-blue-600 underline uppercase tracking-tighter text-left"
                                  >
                                    Mengajar di
                                  </button>
                               </div>
                            </td>
                            <td className="px-6 py-4">
                               <div className="flex flex-col">
                                  <span className="font-black text-green-900">{totalTarget} Pertemuan</span>
                                  <span className="text-[9px] font-bold text-gray-400 italic">Bulan Ini ({new Date().toLocaleDateString('id-ID', { month: 'long' })})</span>
                               </div>
                            </td>
                            <td className="px-6 py-4 text-center text-nowrap">
                              <div className="flex justify-center gap-3">
                                 <button onClick={() => { 
                                   setEditingJadwal({ nip: g.nip, namaGuru: g.namaGuru, mapel: g.mapel } as any); 
                                   setTeachingSessions(g.sessions.map((s: any) => ({ kelas: s.kelas, hari: s.hari, target: s.targetPertemuan })));
                                   setShowJadwalModal(true); 
                                 }} className="text-blue-500 hover:bg-blue-50 p-2 rounded-lg transition-all"><Edit size={16} /></button>
                                 <button onClick={() => { 
                                   setConfirmModal({
                                     show: true,
                                     title: 'Hapus Jam Mengajar?',
                                     message: 'Seluruh jam mengajar guru untuk mapel ini di kelas ini akan dihapus.',
                                     entityName: `${g.namaGuru} - ${g.kelas} (${g.mapel})`,
                                     onConfirm: async () => {
                                       for (const s of g.sessions) {
                                         await firestoreService.hapusJadwalMengajar(s.id);
                                       }
                                     }
                                   });
                                 }} className="text-red-500 hover:bg-red-50 p-2 rounded-lg transition-all"><Trash2 size={16} /></button>
                              </div>
                            </td>
                          </tr>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </div>
                  {groupedJadwal.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.jadwal / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.jadwal === 0} onClick={() => setPagination({...pagination, jadwal: pagination.jadwal - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.jadwal + pageSize >= groupedJadwal.length} onClick={() => setPagination({...pagination, jadwal: pagination.jadwal + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'guru' && (
            <motion.div key="guru" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="relative w-72">
                  <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                  <input 
                    type="text" 
                    placeholder="Cari Guru..." 
                    value={searchTermGuru}
                    onChange={e => { setSearchTermGuru(e.target.value); setPagination({...pagination, guru: 0}); }}
                    className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500" 
                  />
                </div>
                <div className="flex gap-2 text-zinc-900">
                  <select 
                    value={pageSize} 
                    onChange={e => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                  </select>
                  <button 
                    onClick={() => downloadTemplate('Guru')}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-zinc-200 transition-all"
                  >
                    <Download size={14} /> Template
                  </button>
                  <button 
                    onClick={() => handleImportExcel('Guru')}
                    className="bg-blue-100 text-blue-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-blue-200 transition-all"
                  >
                    <Upload size={14} /> Import
                  </button>
                  <button 
                    onClick={() => { setEditingGuru({ nip: '', nama: '', jabatan: 'Guru', kelas: '', user: '', pass: '', role: 'Guru' }); setShowGuruModal(true); }} 
                    className="bg-zinc-900 text-white px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-zinc-800 transition-all shadow-sm"
                  >
                    <Plus size={14} /> Tambah
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden text-zinc-900">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm min-w-[800px]">
                    <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                      <tr>
                        <th className="px-6 py-4 text-center w-12 text-nowrap">No</th>
                        <th className="px-6 py-4">Foto / NIP</th>
                        <th className="px-6 py-4">Nama Guru</th>
                        <th className="px-6 py-4">Jabatan</th>
                        <th className="px-6 py-4">Wali Kelas</th>
                        <th className="px-6 py-4">Username</th>
                        <th className="px-6 py-4">Aksi</th>
                      </tr>
                    </thead>
                      <tbody className="divide-y divide-gray-50">
                        {filteredGuru.slice(pagination.guru, pagination.guru + pageSize).map((g, i) => (
                          <tr key={g.nip} className="hover:bg-gray-50 transition-colors">
                            <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{pagination.guru + i + 1}</td>
                            <td className="px-6 py-4">
                              <div className="w-10 h-10 rounded-lg bg-gray-100 flex items-center justify-center overflow-hidden border">
                                {g.foto ? <img src={g.foto} className="w-full h-full object-cover" /> : <User size={16} className="text-gray-300" />}
                              </div>
                            </td>
                          <td className="px-6 py-4">
                             <p className="font-bold">{g.nama}</p>
                             <p className="text-[10px] font-mono font-bold text-gray-400">NIP: {g.nip}</p>
                          </td>
                          <td className="px-6 py-4">
                            <span className="bg-zinc-100 text-zinc-600 px-2 py-0.5 rounded text-[9px] font-black uppercase">{g.jabatan || 'Guru'}</span>
                          </td>
                          <td className="px-6 py-4"><span className="text-green-800 text-[10px] font-bold">{g.kelas || '-'}</span></td>
                          <td className="px-6 py-4 text-gray-600 font-medium">{g.user}</td>
                          <td className="px-6 py-4">
                            <div className="flex gap-2">
                               <button onClick={() => { setEditingGuru({ nip: g.nip, nama: g.nama, jabatan: g.jabatan, kelas: g.kelas, user: g.user, pass: g.pass, role: g.role, foto: g.foto }); setShowGuruModal(true); }} className="text-blue-500 hover:text-blue-800"><Edit size={16} /></button>
                               <button onClick={() => handleDeleteGuru(g.nip, g.nama)} className="text-red-500 hover:text-red-800"><Trash2 size={16} /></button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
                    {filteredGuru.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.guru / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.guru === 0} onClick={() => setPagination({...pagination, guru: pagination.guru - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.guru + pageSize >= filteredGuru.length} onClick={() => setPagination({...pagination, guru: pagination.guru + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'kelas' && (
            <motion.div key="kelas" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input 
                      type="text" 
                      placeholder="Cari Kelas..." 
                      value={searchTermKelas}
                      onChange={e => { setSearchTermKelas(e.target.value); setPagination({...pagination, kelas: 0}); }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500" 
                    />
                  </div>
                  <select 
                    value={pageSize} 
                    onChange={e => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                  </select>
                </div>
                <button 
                  onClick={() => { setEditingKelas({ nama: '', wali: '', jumlah: 0 }); setShowKelasModal(true); }}
                  className="bg-green-800 text-white px-4 py-2 rounded-xl text-sm font-bold flex items-center gap-2"
                >
                  <Plus size={16} /> Tambah Kelas
                </button>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-zinc-900">
                {filteredKelas.slice(pagination.kelas, pagination.kelas + pageSize).map((k, i) => (
                  <div key={i} className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex flex-col justify-between">
                    <div>
                      <div className="flex justify-between items-start mb-4">
                        <div className="flex flex-col">
                          <span className="text-[10px] font-black text-gray-400">#{pagination.kelas + i + 1}</span>
                          <h3 className="text-2xl font-black text-green-900">{k.nama}</h3>
                        </div>
                        <div className="bg-green-100 p-2 rounded-xl text-green-700">
                          <TableIcon size={20} />
                        </div>
                      </div>
                      <p className="text-sm font-bold text-gray-400 uppercase tracking-tight">Wali Kelas</p>
                      <p className="font-bold text-gray-800 mb-4">{k.wali}</p>
                      
                      <div className="flex items-center justify-between p-3">
                        <span className="text-xs font-bold text-gray-400">KUOTA</span>
                        <span className="text-sm font-black">{k.jumlah} Siswa</span>
                      </div>
                      
                      <div className="mt-4 flex flex-col items-center p-3 rounded-2xl">
                        <p className="text-[9px] font-black text-gray-400 uppercase mb-2">Barcode Absensi Guru</p>
                        <div className="bg-white p-2 rounded-lg shadow-sm border border-gray-50 flex items-center justify-center">
                          <QRCodeSVG 
                            value={k.nama} 
                            size={80}
                            level="H"
                            includeMargin={false}
                          />
                        </div>
                      </div>
                    </div>

                    <div className="mt-6 flex gap-2">
                      <div 
                        id={`card-kelas-${k.nama}`} 
                        className="opacity-0 pointer-events-none fixed"
                        style={{ left: '-5000px', top: '-5000px', zIndex: -1000 }}
                      >
                        <div style={{ padding:'50px', backgroundColor:'#ffffff', textAlign:'center', width:'600px', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center' }}>
                          <h2 style={{ color:'#14532d', fontSize:'36px', marginBottom:'30px', fontWeight:'900', lineHeight:'1.2', textTransform:'uppercase' }}>ABSENSI MENGAJAR<br/>KELAS {k.nama}</h2>
                          <div style={{ backgroundColor: '#ffffff', padding: '20px', borderRadius: '30px', boxShadow: '0 10px 25px rgba(0,0,0,0.1)', border: '1px solid #f0f0f0', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                            <QRCodeSVG 
                              value={k.nama} 
                              size={400}
                              level="H"
                              includeMargin={false}
                            />
                          </div>
                          <p style={{ color:'#14532d', fontSize:'28px', marginTop:'40px', fontWeight:'900', letterSpacing:'0.1em' }}>SIGAP MTsN 2 BOMBANA</p>
                        </div>
                      </div>
                      <button 
                        onClick={() => printBarcode(k.nama)}
                        className="flex-1 bg-zinc-900 text-white text-[10px] font-bold py-2 rounded-lg flex items-center justify-center gap-1 hover:bg-zinc-800 transition-all text-nowrap"
                      >
                        <Printer size={12} /> Cetak
                      </button>
                      <button 
                        onClick={() => downloadClassPDF(k.nama)}
                        className="flex-1 bg-blue-600 text-white text-[10px] font-bold py-2 rounded-lg flex items-center justify-center gap-1 hover:bg-blue-700 transition-all text-nowrap"
                      >
                        <Download size={12} /> PDF
                      </button>
                      <button onClick={() => { setEditingKelas({ ...k, oldNama: k.nama }); setShowKelasModal(true); }} className="bg-gray-100 text-gray-600 p-2 rounded-lg hover:bg-gray-200">
                         <Edit size={14} />
                      </button>
                      <button onClick={() => handleDeleteKelas(k.nama)} className="bg-red-50 text-red-600 p-2 rounded-lg hover:bg-red-100">
                         <Trash2 size={14} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
              {filteredKelas.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.kelas / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.kelas === 0} onClick={() => setPagination({...pagination, kelas: pagination.kelas - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.kelas + pageSize >= filteredKelas.length} onClick={() => setPagination({...pagination, kelas: pagination.kelas + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'absensi-wali' && (
            <motion.div key="absensi-wali" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
              <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
                <div className="xl:col-span-2 space-y-6">
                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                    <div className="p-6 bg-green-50 border-b border-green-100 flex justify-between items-center">
                       <div>
                         <h2 className="text-sm font-black text-green-900 uppercase tracking-widest">Riwayat Kehadiran Siswa ({session?.kelas}) Hari Ini</h2>
                         <p className="text-[10px] text-green-700 font-bold mt-0.5">Hasil absensi scanner atau manual hari ini.</p>
                       </div>
                  <div className="flex items-center gap-2">
                    <select 
                      value={pageSize} 
                      onChange={e => setPageSize(parseInt(e.target.value))}
                      className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                    >
                      {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                    </select>
                    <div className="bg-white px-3 py-1 rounded-full border border-green-200 text-[10px] font-black">{new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long' })}</div>
                  </div>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-sm min-w-[600px]">
                        <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                          <tr>
                            <th className="px-6 py-4 text-center w-12 text-nowrap">No</th>
                            <th className="px-6 py-4">Siswa</th>
                            <th className="px-6 py-4">Status</th>
                            <th className="px-6 py-4">Keterangan</th>
                            <th className="px-6 py-4 text-center">Update</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {students.filter(s => s.kelas === session?.kelas).map((s, i) => {
                            const today = new Date().toISOString().split('T')[0];
                            const current = attendance.find(a => a.nisn === s.nisn && a.tanggal === today);
                            return (
                              <tr key={s.nisn} className="hover:bg-gray-50 transition-colors">
                                <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{i + 1}</td>
                                <td className="px-6 py-4">
                                  <p className="font-bold">{s.nama}</p>
                                  <p className="text-[10px] text-gray-400 font-mono">{s.nisn}</p>
                                </td>
                                <td className="px-6 py-4">
                                  <span className={`px-3 py-1 rounded-full text-[10px] font-black ${current?.status === 'Hadir' ? 'bg-green-100 text-green-700' : current ? 'bg-red-100 text-red-700' : 'bg-gray-100 text-gray-400'}`}>
                                    {current?.status || 'BELUM ABSEN'}
                                  </span>
                                </td>
                                <td className="px-6 py-4 text-[11px] text-gray-500 italic max-w-[200px] truncate">
                                  {current?.keterangan || '-'}
                                </td>
                                <td className="px-6 py-4 text-center">
                                  <button 
                                    onClick={() => setSiswaAbsensiManual({ 
                                      nisn: s.nisn, 
                                      status: current?.status || 'Hadir', 
                                      ket: current?.keterangan || '', 
                                      tanggal: today 
                                    })}
                                    className="p-2 border border-zinc-100 rounded-lg hover:bg-zinc-50 text-green-600"
                                  >
                                    <Edit size={14} />
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                <div className="space-y-6">
                  <div className="bg-white rounded-3xl shadow-sm border border-gray-100 p-8">
                    <h3 className="text-sm font-black text-zinc-900 uppercase tracking-widest mb-6 flex items-center gap-2">
                       <Plus size={16} className="text-green-600" />
                       Input Kehadiran Manual
                    </h3>
                    <div className="space-y-4">
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Pilih Tanggal</label>
                        <input 
                          type="date" 
                          value={siswaAbsensiManual.tanggal}
                          onChange={e => setSiswaAbsensiManual({...siswaAbsensiManual, tanggal: e.target.value})}
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm" 
                        />
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Pilih Siswa</label>
                        <select 
                          value={siswaAbsensiManual.nisn}
                          onChange={e => {
                            const nisn = e.target.value;
                            const current = attendance.find(a => a.nisn === nisn && a.tanggal === siswaAbsensiManual.tanggal);
                            setSiswaAbsensiManual({ ...siswaAbsensiManual, nisn, status: current?.status || 'Hadir', ket: current?.keterangan || '' });
                          }}
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm"
                        >
                          <option value="">-- Pilih Siswa --</option>
                          {students.filter(s => s.kelas === session?.kelas).map(s => (
                            <option key={s.nisn} value={s.nisn}>{s.nama}</option>
                          ))}
                        </select>
                      </div>
                      <div className="grid grid-cols-2 gap-2">
                         {['Hadir', 'Sakit', 'Izin', 'Alfa'].map(stat => (
                           <button
                             key={stat}
                             onClick={() => setSiswaAbsensiManual({...siswaAbsensiManual, status: stat})}
                             className={`py-3 rounded-xl text-[10px] font-black uppercase transition-all ${siswaAbsensiManual.status === stat ? 'bg-green-800 text-white shadow-lg scale-105' : 'bg-zinc-100 text-zinc-400 hover:bg-zinc-200'}`}
                           >
                             {stat}
                           </button>
                         ))}
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Keterangan / Alasan</label>
                        <textarea 
                          rows={3}
                          value={siswaAbsensiManual.ket}
                          onChange={e => setSiswaAbsensiManual({...siswaAbsensiManual, ket: e.target.value})}
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm"
                          placeholder="Penyebab tidak hadir..."
                        />
                      </div>
                      <button 
                        disabled={!siswaAbsensiManual.nisn}
                        onClick={async () => {
                          const student = students.find(s => s.nisn === siswaAbsensiManual.nisn);
                          if (!student) return;
                          toggleLoader(true);
                          try {
                            await firestoreService.saveAbsensiManual({
                              nisn: student.nisn,
                              nama: student.nama,
                              kelas: student.kelas,
                              status: siswaAbsensiManual.status,
                              keterangan: siswaAbsensiManual.ket,
                              tanggal: siswaAbsensiManual.tanggal,
                              jam: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', hour12: false })
                            });
                            alert("Absensi berhasil disimpan!");
                          } catch(e) {
                            alert("Gagal menyimpan data.");
                          } finally {
                            toggleLoader(false);
                          }
                        }}
                        className="w-full bg-green-800 text-white py-4 rounded-2xl font-black uppercase tracking-widest text-xs hover:bg-green-700 transition-all shadow-xl disabled:opacity-30"
                      >
                        SIMPAN ABSENSI
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activePanel === 'absensi-umum' && (
            <motion.div key="absensi-umum" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="mb-6 flex flex-col md:flex-row gap-4 bg-white p-6 rounded-3xl border border-gray-100 shadow-sm">
                  <div className="flex-1">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Filter Kelas</label>
                    <select 
                      value={filterClassAbsensi} 
                      onChange={e => setFilterClassAbsensi(e.target.value)}
                      className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-2 font-bold text-sm mt-1"
                    >
                      <option value="">Semua Kelas</option>
                      {classrooms.map(c => <option key={c.nama} value={c.nama}>{c.nama}</option>)}
                    </select>
                  </div>
                  <div className="flex-1">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Cari Nama</label>
                    <div className="relative mt-1">
                      <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300" />
                      <input 
                        type="text" 
                        placeholder="Ketik nama siswa..." 
                        value={filterNamaAbsensi}
                        onChange={e => setFilterNamaAbsensi(e.target.value)}
                        className="w-full bg-zinc-50 border-0 rounded-xl px-10 py-2 font-bold text-sm"
                      />
                    </div>
                  </div>
                  <div className="flex items-end gap-2">
                    <select 
                      value={pageSize} 
                      onChange={e => setPageSize(parseInt(e.target.value))}
                      className="bg-zinc-100 text-zinc-500 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0 h-[40px]"
                    >
                      {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                    </select>
                    <button 
                      onClick={() => { setFilterClassAbsensi(''); setFilterNamaAbsensi(''); }}
                      className="bg-zinc-100 text-zinc-500 px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-zinc-200"
                    >
                      Reset
                    </button>
                  </div>
               </div>

               <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-x-auto text-zinc-900">
                <table className="w-full text-left text-sm min-w-[800px]">
                  <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                    <tr>
                      <th className="px-6 py-4 text-center w-12">No</th>
                      <th className="px-6 py-4">Tanggal</th>
                      <th className="px-6 py-4">Nama Siswa</th>
                      <th className="px-6 py-4">Kelas</th>
                      <th className="px-6 py-4 text-center">Status</th>
                      <th className="px-6 py-4">Keterangan</th>
                      <th className="px-6 py-4 text-center">Aksi</th>
                      {session?.role === 'Admin' && <th className="px-6 py-4 text-center">Hapus</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {attendance
                      .filter(a => {
                        const matchesClass = filterClassAbsensi ? a.kelas === filterClassAbsensi : true;
                        const matchesName = filterNamaAbsensi ? a.nama.toLowerCase().includes(filterNamaAbsensi.toLowerCase()) : true;
                        return matchesClass && matchesName;
                      })
                      .sort((a, b) => b.tanggal.localeCompare(a.tanggal) || b.jam.localeCompare(a.jam))
                      .slice(pagination.absensiSiswa, pagination.absensiSiswa + pageSize).map((a, i) => (
                      <tr key={a.id} className="hover:bg-gray-50 transition-colors">
                        <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{pagination.absensiSiswa + i + 1}</td>
                        <td className="px-6 py-4 font-medium">{a.tanggal}</td>
                        <td className="px-6 py-4">
                          <p className="font-bold">{a.nama}</p>
                          <p className="text-[10px] text-gray-400 font-mono">{a.nisn}</p>
                        </td>
                        <td className="px-6 py-4 font-bold">{a.kelas}</td>
                         <td className="px-6 py-4 text-center">
                          <span className={`px-2 py-1 rounded-lg text-[10px] font-black border-0 ${
                              a.status === 'Hadir' ? 'bg-green-100 text-green-700' : 
                              a.status === 'Alfa' ? 'bg-red-100 text-red-700' : 
                              'bg-orange-100 text-orange-700'
                            }`}
                          >
                            {a.status}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                           <p className="text-[10px] font-medium text-gray-500 italic max-w-[150px] truncate" title={a.keterangan || '-'}>
                              {a.keterangan || '-'}
                           </p>
                        </td>
                        <td className="px-6 py-4 text-center">
                          { (session?.role === 'Admin' || session?.isWali) ? (
                            <div className="flex justify-center gap-1">
                               {['Hadir', 'Sakit', 'Izin', 'Alfa'].map(stat => (
                                 <button 
                                   key={stat}
                                   title={stat}
                                   onClick={async () => {
                                      toggleLoader(true);
                                      await firestoreService.updateAbsensiStatus(a.id, stat);
                                      toggleLoader(false);
                                   }}
                                   className={`w-6 h-6 rounded flex items-center justify-center text-[8px] font-bold ${a.status === stat ? 'bg-green-800 text-white shadow' : 'bg-gray-100 text-gray-400 hover:bg-gray-200'}`}
                                 >
                                   {stat[0]}
                                 </button>
                               ))}
                            </div>
                          ) : (
                            <span className="text-gray-300 italic text-[10px]">Read Only</span>
                          )}
                        </td>
                        <td className="px-6 py-4 text-center">
                           {session?.role === 'Admin' && (
                             <button onClick={async () => { 
                             setConfirmModal({
                               show: true,
                               title: 'Hapus Log Absensi?',
                               message: 'Rekaman presensi siswa ini akan dihapus dari riwayat hari ini.',
                               onConfirm: async () => {
                                 toggleLoader(true); 
                                 await firestoreService.hapusAbsensi(a.id); 
                                 toggleLoader(false);
                               }
                             });
                           }} className="text-red-300 hover:text-red-600">
                                <Trash2 size={16}/>
                             </button>
                           )}
                        </td>
                      </tr>
                    ))}
                    {attendance.length === 0 && (
                      <tr><td colSpan={7} className="text-center py-10 text-gray-400 font-bold uppercase tracking-widest text-xs">Belum ada rekaman presensi</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              {attendance.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.absensiSiswa / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.absensiSiswa === 0} onClick={() => setPagination({...pagination, absensiSiswa: pagination.absensiSiswa - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.absensiSiswa + pageSize >= attendance.length} onClick={() => setPagination({...pagination, absensiSiswa: pagination.absensiSiswa + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'absensi-guru' && (
            <motion.div key="absensi-guru" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="flex flex-wrap gap-4 justify-between items-center mb-4">
                 <h2 className="text-sm font-black text-green-900 uppercase tracking-widest flex items-center gap-2">
                   <Clock size={16} /> Data Mengajar Guru - {formatIndoDate(new Date().toISOString().split('T')[0])}
                 </h2>
                 <select 
                    value={pageSize} 
                    onChange={e => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                  </select>
               </div>
               <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm min-w-[700px]">
                  <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                    <tr>
                      <th className="px-6 py-4 text-center w-12">No</th>
                      <th className="px-6 py-4">NIP</th>
                      <th className="px-6 py-4">Nama Guru</th>
                      <th className="px-6 py-4">Mata Pelajaran</th>
                      <th className="px-6 py-4">Kelas Mengajar</th>
                      <th className="px-6 py-4">Hari/Tanggal</th>
                      <th className="px-6 py-4">Jam Scan</th>
                      {session?.role === 'Admin' && <th className="px-6 py-4 text-center">Opsi</th>}
                    </tr>
                  </thead>
                    <tbody className="divide-y divide-gray-50">
                    {teacherAttendance.slice(pagination.absensiGuru, pagination.absensiGuru + pageSize).map((a, i) => (
                      <tr key={a.id} className="hover:bg-gray-50 transition-colors">
                        <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">{pagination.absensiGuru + i + 1}</td>
                        <td className="px-6 py-4 text-gray-400 font-mono font-bold">{a.nip}</td>
                        <td className="px-6 py-4 font-bold">{a.nama}</td>
                        <td className="px-6 py-4"><span className="text-zinc-600 font-bold">{a.mapel || '-'}</span></td>
                        <td className="px-6 py-4"><span className="bg-green-100 text-green-800 px-2 py-0.5 rounded text-[10px] font-bold">{a.kelas}</span></td>
                        <td className="px-6 py-4 font-medium text-gray-600">{formatIndoDate(a.tanggal)}</td>
                        <td className="px-6 py-4 font-bold">{a.jam}</td>
                        {session?.role === 'Admin' && (
                          <td className="px-6 py-4 text-center">
                            <button 
                              onClick={async () => {
                                if (window.confirm(`Hapus data absensi guru ${a.nama}?`)) {
                                  toggleLoader(true);
                                  try {
                                    await firestoreService.hapusAbsensiGuru(a.id);
                                  } catch (err) {
                                    alert("Gagal menghapus data.");
                                  } finally {
                                    toggleLoader(false);
                                  }
                                }
                              }}
                              className="text-red-500 hover:bg-red-50 p-2 rounded-lg transition-all"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        )}
                      </tr>
                    ))}
                    {teacherAttendance.length === 0 && (
                      <tr><td colSpan={session?.role === 'Admin' ? 8 : 7} className="text-center py-8 text-gray-400 font-medium italic">Belum ada data mengajar hari ini.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
              {teacherAttendance.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Halaman {Math.floor(pagination.absensiGuru / pageSize) + 1}</span>
                  <div className="flex gap-2">
                    <button disabled={pagination.absensiGuru === 0} onClick={() => setPagination({...pagination, absensiGuru: pagination.absensiGuru - pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronLeft size={16} /></button>
                    <button disabled={pagination.absensiGuru + pageSize >= teacherAttendance.length} onClick={() => setPagination({...pagination, absensiGuru: pagination.absensiGuru + pageSize})} className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"><ChevronRight size={16} /></button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === 'rekap-mapel' && (
            <motion.div key="rekap-mapel" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
               <div className="bg-white rounded-3xl shadow-sm border border-gray-100 p-8 mb-6">
                 <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                   <div className="flex-1">
                     <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                       <FileBarChart className="text-green-700" /> Rekap Presensi Mapel
                     </h2>
                     <p className="text-xs text-gray-500 font-medium mt-1 italic">Laporan performa pengajar dan kehadiran siswa di kelas.</p>
                   </div>
                   <div className="flex items-center gap-2">
                      <select 
                        value={pageSize} 
                        onChange={e => setPageSize(parseInt(e.target.value))}
                        className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                      >
                        {[10, 20, 50, 100].map(v => <option key={v} value={v}>Tampil {v}</option>)}
                      </select>
                   </div>
                   <div className="flex items-center gap-3 w-full md:w-auto">
                     <button 
                        onClick={() => {
                          const dataForExcel = teacherAttendance
                            .filter(a => {
                              const matchesNip = rekapMapelFilter.nip ? a.nip === rekapMapelFilter.nip : (session?.role !== 'Admin' ? a.nip === session?.uid : true);
                              const matchesDate = rekapMapelFilter.tanggal ? a.tanggal === rekapMapelFilter.tanggal : true;
                              const matchesMapel = rekapMapelFilter.mapel ? a.mapel === rekapMapelFilter.mapel : true;
                              return matchesNip && matchesDate && matchesMapel;
                            })
                            .map(a => {
                              const classAttendance = attendance.filter(sa => sa.kelas === a.kelas && sa.tanggal === a.tanggal);
                              return {
                                'Nama Guru': a.nama,
                                'Mapel': a.mapel || '-',
                                'Kelas': a.kelas,
                                'Tanggal': a.tanggal,
                                'Hadir': classAttendance.filter(sa => sa.status === 'Hadir').length,
                                'Sakit': classAttendance.filter(sa => sa.status === 'Sakit').length,
                                'Izin': classAttendance.filter(sa => sa.status === 'Izin').length,
                                'Alfa': classAttendance.filter(sa => sa.status === 'Alfa').length,
                                'Total Siswa': classAttendance.length,
                                'Status': 'Mengajar'
                              };
                            });
                          const ws = XLSX.utils.json_to_sheet(dataForExcel);
                          const wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, ws, "Rekap Mapel");
                          XLSX.writeFile(wb, `Rekap_Mapel_${rekapMapelFilter.tanggal}.xlsx`);
                        }}
                        className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-green-50 text-green-700 px-5 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-green-100 transition-all border border-green-100"
                     >
                       <FileSpreadsheet size={16} /> Excel
                     </button>
                     <button 
                        onClick={() => {
                          const doc = new jsPDF('l', 'mm', 'a4');
                          doc.setFontSize(16);
                          doc.text('REKAPITULASI MENGAJAR DAN PRESENSI MAPEL', 14, 20);
                          doc.setFontSize(10);
                          doc.text(`Tanggal: ${rekapMapelFilter.tanggal}`, 14, 28);
                          
                          const filtered = teacherAttendance.filter(a => {
                            const matchesNip = rekapMapelFilter.nip ? a.nip === rekapMapelFilter.nip : (session?.role !== 'Admin' ? a.nip === session?.uid : true);
                            const matchesDate = rekapMapelFilter.tanggal ? a.tanggal === rekapMapelFilter.tanggal : true;
                            const matchesMapel = rekapMapelFilter.mapel ? a.mapel === rekapMapelFilter.mapel : true;
                            return matchesNip && matchesDate && matchesMapel;
                          });

                          const tableData = filtered.map((a, idx) => {
                            const classAttendance = attendance.filter(sa => sa.kelas === a.kelas && sa.tanggal === a.tanggal);
                            return [
                              idx + 1,
                              a.nama,
                              a.mapel || '-',
                              a.kelas,
                              classAttendance.filter(sa => sa.status === 'Hadir').length,
                              classAttendance.filter(sa => sa.status === 'Sakit').length,
                              classAttendance.filter(sa => sa.status === 'Izin').length,
                              classAttendance.filter(sa => sa.status === 'Alfa').length,
                              classAttendance.length,
                              'Mengajar'
                            ];
                          });

                          autoTable(doc, {
                            startY: 35,
                            head: [['No', 'Nama Guru', 'Mapel', 'Kelas', 'H', 'S', 'I', 'A', 'Total', 'Status']],
                            body: tableData,
                            theme: 'striped',
                            headStyles: { fillColor: [22, 101, 52] }
                          });
                          doc.save(`Rekap_Mapel_${rekapMapelFilter.tanggal}.pdf`);
                        }}
                        className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-red-50 text-red-700 px-5 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-red-100 transition-all border border-red-100"
                     >
                       <Download size={16} /> PDF
                     </button>
                   </div>
                 </div>

                 <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-8">
                    <div>
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Filter Guru</label>
                      <select 
                        disabled={session?.role !== 'Admin'}
                        value={session?.role !== 'Admin' ? session?.uid : rekapMapelFilter.nip}
                        onChange={e => setRekapMapelFilter({...rekapMapelFilter, nip: e.target.value, mapel: ''})}
                        className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1 disabled:opacity-50"
                      >
                        {session?.role === 'Admin' ? (
                          <>
                            <option value="">Semua Guru</option>
                            {teachers.filter(t => t.role !== 'Siswa').map(t => <option key={t.nip} value={t.nip}>{t.nama}</option>)}
                          </>
                        ) : (
                          <option value={session?.uid}>{session?.name}</option>
                        )}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Filter Mapel</label>
                      <select 
                        value={rekapMapelFilter.mapel}
                        onChange={e => setRekapMapelFilter({...rekapMapelFilter, mapel: e.target.value})}
                        className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1"
                      >
                        <option value="">Semua Mapel</option>
                        {Array.from(new Set(teachingSchedules
                          .filter(s => {
                             const targetNip = rekapMapelFilter.nip || (session?.role !== 'Admin' ? session?.uid : '');
                             return targetNip ? s.nip === targetNip : true;
                          })
                          .map(s => s.mapel)
                          .filter(Boolean)
                        )).map(m => <option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Pilih Tanggal</label>
                      <input 
                        type="date"
                        value={rekapMapelFilter.tanggal}
                        onChange={e => setRekapMapelFilter({...rekapMapelFilter, tanggal: e.target.value})}
                        className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1"
                      />
                    </div>
                 </div>
               </div>

               <div className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden">
                 <div className="overflow-x-auto">
                   <table className="w-full text-left text-sm min-w-[1000px]">
                     <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                       <tr>
                         <th className="px-6 py-4 text-center">No</th>
                         <th className="px-6 py-4 text-nowrap">Nama Guru</th>
                         <th className="px-6 py-4 text-nowrap">Mapel</th>
                         <th className="px-6 py-4 text-nowrap">Kelas</th>
                         <th className="px-6 py-4 text-center">H</th>
                         <th className="px-6 py-4 text-center">S</th>
                         <th className="px-6 py-4 text-center">I</th>
                         <th className="px-6 py-4 text-center">A</th>
                         <th className="px-6 py-4 text-center">Total</th>
                         <th className="px-6 py-4 text-nowrap">Status</th>
                       </tr>
                     </thead>
                     <tbody className="divide-y divide-gray-50">
                       {teacherAttendance
                        .filter(a => {
                          const matchesNip = rekapMapelFilter.nip ? a.nip === rekapMapelFilter.nip : (session?.role !== 'Admin' ? a.nip === session?.uid : true);
                          const matchesDate = rekapMapelFilter.tanggal ? a.tanggal === rekapMapelFilter.tanggal : true;
                          const matchesMapel = rekapMapelFilter.mapel ? a.mapel === rekapMapelFilter.mapel : true;
                          return matchesNip && matchesDate && matchesMapel;
                        })
                        .map((a, i) => {
                          const classAttendance = attendance.filter(sa => sa.kelas === a.kelas && sa.tanggal === a.tanggal);
                          const h = classAttendance.filter(sa => sa.status === 'Hadir').length;
                          const s = classAttendance.filter(sa => sa.status === 'Sakit').length;
                          const iz = classAttendance.filter(sa => sa.status === 'Izin').length;
                          const al = classAttendance.filter(sa => sa.status === 'Alfa').length;
                          return (
                            <tr key={a.id} className="hover:bg-gray-50">
                              <td className="px-6 py-4 text-center text-gray-400 font-bold">{i + 1}</td>
                              <td className="px-6 py-4 font-bold">{a.nama}</td>
                              <td className="px-6 py-4 text-green-700 font-black text-xs uppercase">{a.mapel || '-'}</td>
                              <td className="px-6 py-4 font-bold">{a.kelas}</td>
                              <td className="px-6 py-4 text-center">
                                <span className="bg-green-100 text-green-700 px-2 py-0.5 rounded-full text-[10px] font-black">{h}</span>
                              </td>
                              <td className="px-6 py-4 text-center">
                                <span className="bg-orange-100 text-orange-700 px-2 py-0.5 rounded-full text-[10px] font-black">{s}</span>
                              </td>
                              <td className="px-6 py-4 text-center">
                                <span className="bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full text-[10px] font-black">{iz}</span>
                              </td>
                              <td className="px-6 py-4 text-center">
                                <span className="bg-red-100 text-red-700 px-2 py-0.5 rounded-full text-[10px] font-black">{al}</span>
                              </td>
                              <td className="px-6 py-4 text-center font-black">{classAttendance.length}</td>
                              <td className="px-6 py-4 italic text-gray-500 font-medium">Mengajar</td>
                            </tr>
                          );
                        })}
                     </tbody>
                   </table>
                 </div>
               </div>
            </motion.div>
          )}

          {activePanel === 'siswa-data' && (
            <motion.div key="siswa-data" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
              <div className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden max-w-2xl mx-auto">
                <div className="bg-green-800 p-6 text-white flex justify-between items-center">
                   <div>
                     <h2 className="text-lg font-bold">Tentang Data Siswa</h2>
                     <p className="text-green-200 text-xs mt-1">Status data biodata terakhir yang tersimpan di sistem.</p>
                   </div>
                   <div className="bg-white/20 p-2 rounded-xl">
                     <GraduationCap size={24} />
                   </div>
                </div>
                <div className="p-8">
                  <div className="flex flex-col items-center mb-8">
                    <div className="relative group">
                      <div className="w-32 h-32 rounded-[2rem] bg-zinc-100 border-4 border-white shadow-lg overflow-hidden flex items-center justify-center">
                        {currentSiswaData?.foto ? (
                          <img src={currentSiswaData.foto} className="w-full h-full object-cover" alt="Profile" />
                        ) : (
                          <User size={64} className="text-zinc-300" />
                        )}
                      </div>
                      <label className="absolute bottom-0 right-0 p-2 bg-green-800 text-white rounded-xl shadow-lg cursor-pointer hover:bg-green-700 transition-all border-4 border-white">
                        <Upload size={16} />
                        <input 
                          type="file" 
                          className="hidden" 
                          accept="image/*"
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (file && currentSiswaData) {
                              const reader = new FileReader();
                              reader.onloadend = async () => {
                                toggleLoader(true);
                                try {
                                  await firestoreService.saveSiswa({
                                    ...currentSiswaData,
                                    foto: reader.result as string
                                  });
                                  alert("Foto profile berhasil diunggah!");
                                } catch (err) {
                                  alert("Gagal mengunggah foto.");
                                } finally {
                                  toggleLoader(false);
                                }
                              };
                              reader.readAsDataURL(file);
                            }
                          }}
                        />
                      </label>
                    </div>
                    <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mt-3">Upload Foto Profil Anda</p>
                  </div>
                  
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">NISN (Kunci)</label>
                        <input type="text" readOnly value={currentSiswaData?.nisn || ''} className="w-full bg-zinc-50 border-0 rounded-2xl py-3 px-4 font-bold text-zinc-400 cursor-not-allowed" />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Nama Lengkap</label>
                        <input 
                          type="text" 
                          value={biodataForm.nama} 
                          onChange={e => setBiodataForm({...biodataForm, nama: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500" 
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Jenis Kelamin</label>
                        <select 
                          value={biodataForm.jenisKelamin} 
                          onChange={e => setBiodataForm({...biodataForm, jenisKelamin: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500 border-0"
                        >
                          <option value="L">Laki-Laki</option>
                          <option value="P">Perempuan</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Tempat Lahir</label>
                        <input 
                          type="text" 
                          placeholder="Contoh: Bombana" 
                          value={biodataForm.tempat} 
                          onChange={e => setBiodataForm({...biodataForm, tempat: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500" 
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Tanggal Lahir</label>
                        <input 
                          type="date" 
                          value={biodataForm.tgl} 
                          onChange={e => setBiodataForm({...biodataForm, tgl: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500" 
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Nama Ayah Kandung</label>
                        <input 
                          type="text" 
                          value={biodataForm.ayah} 
                          onChange={e => setBiodataForm({...biodataForm, ayah: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500" 
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Nama Ibu Kandung</label>
                        <input 
                          type="text" 
                          value={biodataForm.ibu} 
                          onChange={e => setBiodataForm({...biodataForm, ibu: e.target.value})}
                          className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500" 
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">Kelas Aktif (Kunci)</label>
                        <input type="text" readOnly value={currentSiswaData?.kelas || ''} className="w-full bg-zinc-50 border-0 rounded-2xl py-3 px-4 font-bold text-zinc-400 cursor-not-allowed" />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">WhatsApp Ortu (Kunci)</label>
                        <input 
                          type="text" 
                          readOnly
                          value={currentSiswaData?.hp || ''} 
                          className="w-full bg-zinc-50 border-0 rounded-2xl py-3 px-4 font-bold text-zinc-400 cursor-not-allowed" 
                        />
                      </div>
                    </div>
                    <div className="mt-8 pt-8 border-t border-zinc-50 flex justify-end">
                      <button 
                        onClick={async () => {
                          if (!currentSiswaData) return;
                          toggleLoader(true);
                          try {
                            await firestoreService.saveSiswa({
                              ...currentSiswaData,
                              ...biodataForm
                            });
                            alert("Biodata berhasil diperbarui.");
                          } catch (e) {
                            alert("Gagal memperbarui biodata.");
                          } finally {
                            toggleLoader(false);
                          }
                        }}
                        className="bg-green-800 text-white px-8 py-3 rounded-2xl font-bold shadow-lg hover:bg-green-700 transition-all flex items-center gap-2"
                      >
                         <RefreshCcw size={18} /> Update Data Siswa
                      </button>
                    </div>
                </div>
              </div>
            </motion.div>
          )}

          {activePanel === 'profil' && renderProfile()}

          {activePanel === 'capaian-guru' && renderLaporanCapaianGuru()}

          {activePanel === 'roster' && renderRoster()}

          {activePanel === 'kbc' && renderKBC()}

          {activePanel === 'profil-siswa' && renderProfilSiswa()}

          {activePanel === 'rekap' && (
            <motion.div key="rekap" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 mb-6">
                <div className="flex gap-2 p-1 bg-gray-100 rounded-xl mb-6 w-fit">
                   <button 
                     onClick={() => setRekapFilter({...rekapFilter, type: 'Siswa'})}
                     className={`px-6 py-2 text-xs font-black rounded-lg transition-all ${rekapFilter.type === 'Siswa' ? 'bg-green-800 text-white shadow' : 'text-gray-400 uppercase tracking-widest'}`}
                   >
                     REKAP SISWA
                   </button>
                   <button 
                     onClick={() => setRekapFilter({...rekapFilter, type: 'Guru'})}
                     className={`px-6 py-2 text-xs font-black rounded-lg transition-all ${rekapFilter.type === 'Guru' ? 'bg-green-800 text-white shadow' : 'text-gray-400 uppercase tracking-widest'}`}
                   >
                     REKAP GURU
                   </button>
                </div>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
                    <div>
                      <label className="block text-[10px] font-black text-gray-400 uppercase mb-1 ml-1">Pilih Bulan</label>
                      <input 
                        type="month" 
                        value={rekapFilter.bulan}
                        onChange={e => setRekapFilter({...rekapFilter, bulan: e.target.value})}
                        className="w-full bg-gray-50 border-0 rounded-xl px-4 py-2 text-sm font-bold" 
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-black text-gray-400 uppercase mb-1 ml-1">Filter Kelas</label>
                      <select 
                        disabled={(() => {
                           const teacher = teachers.find(t => t.nip === session?.uid);
                           const isKamadWakamad = teacher && ['Kamad', 'Wakamad'].includes(teacher.jabatan);
                           return !isKamadWakamad && (session?.role === 'Guru Wali Kelas' || session?.isWali);
                        })()}
                        value={rekapFilter.kelas}
                        onChange={e => setRekapFilter({...rekapFilter, kelas: e.target.value})}
                        className="w-full bg-gray-50 border-0 rounded-xl px-4 py-2 text-sm font-bold disabled:opacity-50"
                      >
                        <option value="">Semua Kelas</option>
                        {classrooms.map(k => <option key={k.nama} value={k.nama}>{k.nama}</option>)}
                      </select>
                    </div>
                    <div className="md:col-span-2 flex gap-2">
                      <button 
                        onClick={exportExcel}
                        className="flex-1 bg-green-800 text-white px-4 py-2 rounded-xl text-xs font-bold flex items-center justify-center gap-2 hover:bg-green-700 shadow-md"
                      >
                        <Download size={14} /> Download Excel
                      </button>
                      <button 
                        onClick={exportPDF}
                        className="flex-1 bg-red-600 text-white px-4 py-2 rounded-xl text-xs font-bold flex items-center justify-center gap-2 hover:bg-red-500 shadow-md"
                      >
                        <FileBarChart size={14} /> Download PDF
                      </button>
                    </div>
                </div>
              </div>

              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                      {rekapFilter.type === 'Siswa' ? (
                        <tr>
                          <th className="px-6 py-4">Nama Siswa</th>
                          <th className="px-6 py-4 text-center">Kelas</th>
                          <th className="px-6 py-4 text-center">Hadir</th>
                          <th className="px-6 py-4 text-center">Izin</th>
                          <th className="px-6 py-4 text-center">Sakit</th>
                          <th className="px-6 py-4 text-center">Alfa</th>
                          <th className="px-6 py-4 text-center">% Kehadiran</th>
                        </tr>
                      ) : (
                        <tr>
                          <th className="px-6 py-4">Nama Guru</th>
                          <th className="px-6 py-4 text-center">Jabatan</th>
                          <th className="px-6 py-4 text-center">Kelas Mengajar</th>
                          <th className="px-6 py-4 text-center">Target Sesi</th>
                          <th className="px-6 py-4 text-center">Aktual</th>
                          <th className="px-6 py-4 text-center">% Kinerja</th>
                        </tr>
                      )}
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {rekapFilter.type === 'Siswa' ? (
                        students
                          .filter(s => (rekapFilter.kelas ? s.kelas === rekapFilter.kelas : true))
                          .map((s, i) => {
                            const monthAttendance = monthAttendanceMap[s.nisn] || [];
                            const hadir = monthAttendance.filter(a => a.status === 'Hadir').length;
                            const izin = monthAttendance.filter(a => a.status === 'Izin').length;
                            const sakit = monthAttendance.filter(a => a.status === 'Sakit').length;
                            const alfa = monthAttendance.filter(a => a.status === 'Alfa').length;
                            const [y, m] = rekapFilter.bulan.split('-').map(Number);
                            const now = new Date();
                            const limit = (now.getFullYear() === y && (now.getMonth() + 1) === m) ? now.getDate() : undefined;
                            const effCount = getEffectiveDays(y, m - 1, holidays, settings, limit).length;
                            const perc = effCount > 0 ? Math.round(((hadir + izin + sakit) / effCount) * 100) : 0;
                            
                            return (
                              <tr key={i} className="hover:bg-gray-50 transition-colors">
                                <td className="px-6 py-4 font-bold">{s.nama}</td>
                                <td className="px-6 py-4 text-center font-medium">{s.kelas}</td>
                                <td className="px-6 py-4 text-center">{hadir}</td>
                                <td className="px-6 py-4 text-center">{izin}</td>
                                <td className="px-6 py-4 text-center">{sakit}</td>
                                <td className="px-6 py-4 text-center text-red-600 font-bold">{alfa}</td>
                                <td className="px-6 py-4 text-center">
                                  <span className={`px-2 py-0.5 rounded-full text-[10px] font-black ${perc >= 80 ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}`}>
                                    {perc}%
                                  </span>
                                </td>
                              </tr>
                            );
                          })
                      ) : (
                        teachers
                          .filter(t => {
                            if (rekapFilter.kelas) {
                              return teachingSchedules.some(ts => ts.nip === t.nip && ts.kelas === rekapFilter.kelas);
                            }
                            return true;
                          })
                          .map((t, i) => {
                            const schedules = teachingSchedules.filter(ts => ts.nip === t.nip && (rekapFilter.kelas ? ts.kelas === rekapFilter.kelas : true));
                            const [y, m] = rekapFilter.bulan.split('-').map(Number);
                            const totalTarget = schedules.reduce((sum, sch) => {
                              const occ = countDaysInMonth(y, m - 1, sch.hari, holidays);
                              return sum + (occ * (Number(sch.targetPertemuan) || 0));
                            }, 0);
                            const actual = (teacherAttendanceMap[t.nip] || [])
                              .filter(ta => rekapFilter.kelas ? ta.kelas === rekapFilter.kelas : true)
                              .length;
                            const perc = totalTarget > 0 ? Math.round((actual / totalTarget) * 100) : 0;
                            
                            return (
                              <tr key={i} className="hover:bg-gray-50 transition-colors">
                                <td className="px-6 py-4 font-bold">{t.nama}</td>
                                <td className="px-6 py-4 text-center text-[10px] font-black uppercase text-gray-400">{t.jabatan || 'Guru'}</td>
                                <td className="px-6 py-4 text-center text-[10px] font-black uppercase text-gray-400">
                                  {rekapFilter.kelas ? rekapFilter.kelas : (schedules.map(s => s.kelas).join(', ') || '-')}
                                </td>
                                <td className="px-6 py-4 text-center font-bold">{totalTarget}</td>
                                <td className="px-6 py-4 text-center font-bold text-blue-600">{actual}</td>
                                <td className="px-6 py-4 text-center">
                                  <span className={`px-2 py-0.5 rounded-full text-[10px] font-black ${perc >= 50 ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                                    {perc}%
                                  </span>
                                </td>
                              </tr>
                            );
                          })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </motion.div>
          )}

          {activePanel === 'pengaturan' && (
            <motion.div key="pengaturan" initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100">
                  <div className="flex items-center justify-between mb-6">
                    <div className="flex items-center gap-3">
                      <div className="bg-green-100 p-2 rounded-xl text-green-700">
                        <Clock size={20} />
                      </div>
                      <h3 className="text-lg font-bold">Jam Masuk & Pulang</h3>
                    </div>
                    {settings.length === 0 && (
                      <button 
                        onClick={async () => {
                          toggleLoader(true);
                          await firestoreService.initializeSettings();
                          toggleLoader(false);
                        }}
                        className="bg-green-600 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-green-700"
                      >
                        Inisialisasi
                      </button>
                    )}
                  </div>
                  
                  <div className="space-y-4">
                    {[...settings].sort((a,b) => dayOrder[a.hari] - dayOrder[b.hari]).map((s, idx) => {
                      const today = new Date();
                      // Day number: Minggu=0, Senin=1, ...
                      const currentDay = today.getDay();
                      // targetDay from dayOrder: Senin=1, ... Minggu=7 -> map to 0-6
                      const targetDayNum = dayOrder[s.hari] === 7 ? 0 : dayOrder[s.hari];
                      
                      // Calculate date for this specific day of THIS current week
                      const diff = targetDayNum - currentDay;
                      const dateAtDay = new Date();
                      dateAtDay.setDate(today.getDate() + diff);
                      
                      // Format precisely as YYYY-MM-DD in LOCAL time
                      const dateStr = dateAtDay.toLocaleDateString('en-CA'); 
                      const holiday = holidays.find(h => h.tanggal === dateStr);
                      
                      return (
                        <div key={idx} className={`flex items-center justify-between p-3 rounded-2xl ${holiday ? 'bg-red-50 border border-red-100' : 'bg-gray-50'}`}>
                          <div className="flex flex-col">
                            <span className={`font-bold text-sm ${holiday ? 'text-red-900' : ''}`}>{s.hari}</span>
                            <span className="text-[8px] text-gray-400 font-black uppercase tracking-widest">{dateAtDay.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' })}</span>
                            {holiday && <span className="text-[8px] text-red-600 font-bold uppercase mt-1">LIBUR: {holiday.keterangan}</span>}
                          </div>
                          <div className="flex items-center gap-4">
                            <div className="flex flex-col items-center">
                              <span className="text-[8px] font-black text-gray-400 uppercase mb-1">Masuk</span>
                              <input 
                                type="time" 
                                defaultValue={s.masuk} 
                                disabled={s.hari === 'Minggu'}
                                onBlur={async (e) => { await firestoreService.savePengaturanHari(s.hari, e.target.value, s.pulang); }}
                                className="bg-white border border-gray-100 rounded-xl px-3 py-1.5 text-xs font-bold shadow-sm focus:ring-2 focus:ring-green-500 outline-none disabled:opacity-30" 
                              />
                            </div>
                            <div className="flex flex-col items-center">
                              <span className="text-[8px] font-black text-gray-400 uppercase mb-1">Pulang</span>
                              <input 
                                type="time" 
                                defaultValue={s.pulang} 
                                disabled={s.hari === 'Minggu'}
                                onBlur={async (e) => { await firestoreService.savePengaturanHari(s.hari, s.masuk, e.target.value); }}
                                className="bg-white border border-gray-100 rounded-xl px-3 py-1.5 text-xs font-bold shadow-sm focus:ring-2 focus:ring-green-500 outline-none disabled:opacity-30" 
                              />
                            </div>
                          </div>
                        </div>
                      );
                    })}
                    <div className="bg-red-50 p-4 rounded-2xl border border-red-100 mt-4 flex items-center justify-center gap-2">
                       <Calendar size={14} className="text-red-600" />
                       <span className="text-[10px] font-black text-red-700 uppercase tracking-widest">HARI MINGGU: LIBUR RUTIN (TANPA SETTING JAM)</span>
                    </div>
                  </div>
                  <p className="mt-4 text-[10px] text-gray-400 font-bold uppercase tracking-widest text-center italic">* Simpan otomatis saat kursor keluar</p>
                </div>

                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100">
                  <div className="flex items-center gap-3 mb-6">
                    <div className="bg-red-100 p-2 rounded-xl text-red-600">
                      <Calendar size={20} />
                    </div>
                    <h3 className="text-lg font-bold">Hari Libur</h3>
                  </div>

                  <div className="space-y-3 mb-6">
                    {holidays.map((h, idx) => (
                      <div key={idx} className="flex items-center justify-between p-4 bg-red-50/50 rounded-2xl border border-red-100">
                        <div>
                          <p className="font-black text-red-900 text-sm">{h.keterangan}</p>
                          <p className="text-xs text-red-400 font-bold">{h.tanggal}</p>
                        </div>
                        <button 
                          onClick={async () => { await firestoreService.hapusLiburAgenda(h.tanggal); }}
                          className="text-red-300 hover:text-red-600 p-2"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    ))}
                  </div>

                  <div className="bg-gray-50 p-4 rounded-2xl border-2 border-dashed border-gray-200">
                    <div className="space-y-3">
                      <input 
                        type="date" 
                        value={newLibur.tanggal}
                        onChange={e => setNewLibur({...newLibur, tanggal: e.target.value})}
                        className="w-full bg-white border-0 rounded-xl px-4 py-2 text-sm font-bold shadow-sm" 
                      />
                      <input 
                        type="text" 
                        placeholder="Keterangan Libur" 
                        value={newLibur.keterangan}
                        onChange={e => setNewLibur({...newLibur, keterangan: e.target.value})}
                        className="w-full bg-white border-0 rounded-xl px-4 py-2 text-sm font-bold shadow-sm" 
                      />
                      <button 
                        onClick={async () => {
                          if (!newLibur.tanggal || !newLibur.keterangan) return;
                          toggleLoader(true);
                          try {
                            await firestoreService.saveLiburAgenda(newLibur.tanggal, newLibur.keterangan);
                            setNewLibur({ tanggal: '', keterangan: '' });
                            triggerSuccess("BERHASIL", "Hari libur ditambahkan.");
                          } catch (e) {
                            alert("Gagal menyimpan.");
                          } finally {
                            toggleLoader(false);
                          }
                        }}
                        className="w-full bg-red-600 text-white rounded-xl py-2 font-bold hover:bg-red-500 transition-all text-xs"
                      >
                        Tambah Hari Libur
                      </button>
                    </div>
                  </div>
                </div>

                {/* Logo & Identity Configuration */}
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 lg:col-span-2 space-y-8">
                   <div className="flex items-center gap-3 mb-2">
                    <div className="bg-blue-100 p-2 rounded-xl text-blue-600">
                      <LayoutGrid size={20} />
                    </div>
                    <h3 className="text-lg font-bold">Identitas & Desain Aplikasi</h3>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-start">
                    <div className="space-y-6">
                      {/* Logo Section */}
                      <div>
                        <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">Logo Aplikasi (Sekolah)</label>
                        
                        <div className="flex flex-col gap-3 mb-4">
                          <input 
                            type="text" 
                            placeholder="Link Logo (URL)..."
                            defaultValue={appConfig.logoUrl || ''}
                            onBlur={async (e) => {
                               toggleLoader(true);
                               await setDoc(doc(db, 'appConfig', 'general'), { logoUrl: e.target.value }, { merge: true });
                               toggleLoader(false);
                            }}
                            className="bg-gray-50 border border-gray-100 rounded-xl px-4 py-2 text-xs font-bold"
                          />
                        </div>

                        <div className="flex items-center gap-6 p-6 bg-gray-50 rounded-[2rem] border border-gray-100 group relative">
                          <div className="w-24 h-24 bg-white rounded-[1.5rem] flex items-center justify-center shadow-lg overflow-hidden p-2 border-4 border-white transition-transform group-hover:scale-105 relative">
                            <img 
                              src={appLogo} 
                              alt="Preview" 
                              className="w-full h-full object-contain relative z-10" 
                              onError={(e) => {
                                console.error("[DEBUG] Logo load error (Settings Preview):", e.currentTarget.src);
                                if (e.currentTarget.src !== fallbackLogo && !e.currentTarget.src.includes('/logo.png')) {
                                  e.currentTarget.src = fallbackLogo;
                                } else {
                                  e.currentTarget.style.opacity = '0';
                                  const parent = e.currentTarget.parentElement;
                                  if (parent) {
                                    const icon = parent.querySelector('.fallback-icon');
                                    if (icon) (icon as HTMLElement).style.opacity = '1';
                                  }
                                }
                              }}
                            />
                            <School size={36} className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity" />
                          </div>
                          <div className="flex-1 space-y-3">
                            <p className="text-[10px] text-gray-400 font-medium">Gunakan gambar transparan (PNG) untuk hasil terbaik.</p>
                            <label className="inline-block bg-green-800 text-white px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg hover:bg-green-700 cursor-pointer transition-all">
                              Pilih Gambar
                              <input 
                                type="file" 
                                className="hidden" 
                                accept="image/*"
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) {
                                    const reader = new FileReader();
                                    reader.onloadend = async () => {
                                      toggleLoader(true);
                                      try {
                                        const result = reader.result as string;
                                        if (result.length > 1048000) {
                                          throw new Error("Ukuran gambar terlalu besar. Gunakan gambar di bawah 750KB.");
                                        }
                                        await setDoc(doc(db, 'appConfig', 'general'), {
                                          logoUrl: result
                                        }, { merge: true });
                                        console.log("Logo updated on server");
                                        alert("Logo aplikasi berhasil diperbarui!");
                                      } catch (err: any) {
                                        console.error("Logo upload error:", err);
                                        alert("Gagal memperbarui logo: " + (err.message || "Unknown error"));
                                      } finally {
                                        toggleLoader(false);
                                      }
                                    };
                                    reader.readAsDataURL(file);
                                  }
                                }}
                              />
                            </label>
                            {appConfig.logoUrl && (
                              <button 
                                onClick={async () => {
                                  setConfirmModal({
                                    show: true,
                                    title: 'Hapus Logo?',
                                    message: 'Hapus logo kustom dan kembali ke default?',
                                    onConfirm: async () => {
                                      toggleLoader(true);
                                      await setDoc(doc(db, 'appConfig', 'general'), { logoUrl: null }, { merge: true });
                                      toggleLoader(false);
                                    }
                                  });
                                }}
                                className="ml-2 inline-block bg-zinc-200 text-zinc-600 px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-zinc-300"
                              >
                                Reset Logo
                              </button>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* Data Mata Pelajaran Madrasah */}
                      <div className="bg-gray-50 p-6 rounded-[2rem] border border-gray-100">
                        <div className="flex items-center justify-between mb-4">
                          <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">Data Mata Pelajaran Madrasah</label>
                          <button 
                            onClick={async () => {
                              const currentSubs = appConfig.subjects || [];
                              const allSelected = DEFAULT_SUBJECTS.every(s => currentSubs.includes(s));
                              const newSubs = allSelected ? [] : [...DEFAULT_SUBJECTS];
                              await setDoc(doc(db, 'appConfig', 'general'), { subjects: newSubs }, { merge: true });
                            }}
                            className="bg-green-100 text-green-700 px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-wider hover:bg-green-200 transition-all"
                          >
                            {DEFAULT_SUBJECTS.every(s => (appConfig.subjects || []).includes(s)) ? 'Hapus Semua' : 'Ceklis Semua'}
                          </button>
                        </div>
                        
                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                          {DEFAULT_SUBJECTS.map(subject => {
                            const isChecked = (appConfig.subjects || []).includes(subject);
                            return (
                              <label key={subject} className={`flex items-center gap-2 p-3 rounded-xl border transition-all cursor-pointer ${isChecked ? 'bg-green-50 border-green-200 text-green-700' : 'bg-white border-gray-100 text-gray-400 hover:border-gray-300'}`}>
                                <input 
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={async (e) => {
                                    const currentSubs = appConfig.subjects || [];
                                    let newSubs;
                                    if (e.target.checked) {
                                      newSubs = [...currentSubs, subject];
                                    } else {
                                      newSubs = currentSubs.filter(s => s !== subject);
                                    }
                                    await setDoc(doc(db, 'appConfig', 'general'), { subjects: newSubs }, { merge: true });
                                  }}
                                  className="w-4 h-4 rounded border-gray-300 text-green-600 focus:ring-green-500"
                                />
                                <span className="text-[11px] font-bold truncate">{subject}</span>
                              </label>
                            );
                          })}
                        </div>
                        <p className="mt-4 text-[9px] text-gray-400 italic">Mata pelajaran yang diceklis akan muncul pada menu jam mengajar.</p>
                      </div>
                    </div>


                    <div className="bg-zinc-950 p-6 rounded-[2rem] text-white">
                      <h4 className="text-xs font-black uppercase tracking-[0.2em] text-green-500 mb-4">Informasi Sistem</h4>
                      <div className="space-y-4">
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                           <p className="text-[9px] font-black text-gray-500 uppercase mb-1">Nama Aplikasi</p>
                           <p className="font-bold text-sm">SIGAP MTsN 2 BOMBANA</p>
                        </div>
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                           <p className="text-[9px] font-black text-gray-500 uppercase mb-1">Versi</p>
                           <p className="font-bold text-sm">Production v1.2.0-PRO</p>
                        </div>
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                           <p className="text-[9px] font-black text-gray-500 uppercase mb-1">URL Logo Aplikasi</p>
                           <input 
                             type="text" 
                             value={appLogoInput} 
                             onChange={e => setAppLogoInput(e.target.value)}
                             placeholder="https://example.com/logo.png"
                             className="w-full bg-white/10 border-0 rounded-xl py-2 px-3 text-xs font-bold text-white mt-1 focus:ring-1 focus:ring-green-500"
                           />
                        </div>
                        <button 
                          onClick={() => {
                            if (!appLogoInput) return;
                            toggleLoader(true);
                            firestoreService.saveAppConfig({ logoUrl: appLogoInput })
                              .then(() => alert("Logo berhasil diterapkan."))
                              .finally(() => toggleLoader(false));
                          }}
                          className="bg-green-800 text-white px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest mt-2 hover:bg-green-700 transition-all"
                        >
                          Terapkan Logo
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {selectedStudentCard && (session?.role === 'Siswa' || session?.role === 'Admin') && (
            <div className="fixed inset-0 z-[200] flex items-center justify-center p-2 sm:p-4 overflow-y-auto" style={{ backgroundColor: 'rgba(0, 0, 0, 0.6)', backdropFilter: 'blur(4px)' }}>
              <div className="absolute inset-0" onClick={() => setSelectedStudentCard(null)} />
              <motion.div 
                initial={{ scale: 0.9, opacity: 0 }} 
                animate={{ scale: 1, opacity: 1 }} 
                className="relative p-4 md:p-8 rounded-[2rem] md:rounded-[3rem] shadow-2xl flex flex-col items-center border w-fit max-w-[95vw] sm:max-w-none max-h-[95vh] overflow-y-auto scale-[0.75] sm:scale-[0.85] md:scale-90 lg:scale-100 my-auto origin-top"
                style={{ backgroundColor: 'rgba(255, 255, 255, 0.1)', backdropFilter: 'blur(16px)', borderColor: 'rgba(255, 255, 255, 0.2)' }}
              >
                  <div id="print-student-card" className="flex flex-col md:flex-row gap-4 p-4 w-full" style={{ width: 'fit-content', backgroundColor: 'rgba(0,0,0,0.05)', borderRadius: '2rem' }}>
                    {/* FRONT SIDE */}
                    <div style={{ 
                      background: appConfig.cardTemplateUrl ? 'none' : 'linear-gradient(135deg, #14532d 0%, #052e16 100%)', 
                      color: '#ffffff', 
                      width: '320px', 
                      height: '480px', 
                      borderRadius: '32px', 
                      position: 'relative', 
                      overflow: 'hidden', 
                      boxSizing: 'border-box',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      padding: '40px 24px',
                      fontFamily: 'Inter, sans-serif'
                    }}>
                      {appConfig.cardTemplateUrl && (
                        <img 
                           src={appConfig.cardTemplateUrl} 
                           style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover', zIndex: 0 }} 
                           crossOrigin="anonymous" 
                           alt="Background"
                        />
                      )}
                      {!appConfig.cardTemplateUrl && (
                        <div style={{ width: '80px', height: '80px', backgroundColor: 'white', borderRadius: '24px', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '10px', marginBottom: '20px', boxShadow: '0 8px 24px rgba(0,0,0,0.2)', zIndex: 10 }}>
                          <img src={appLogo} style={{ width: '100%', height: '100%', objectFit: 'contain' }} crossOrigin="anonymous" />
                        </div>
                      )}

                      <div style={{ textAlign: 'center', marginBottom: '30px', zIndex: 10 }}>
                        <h3 style={{ margin: 0, fontSize: '18px', fontWeight: '900', letterSpacing: '0.05em', textTransform: 'uppercase', lineHeight: '1.2' }}>MTS NEGERI 2 BOMBANA</h3>
                        <p style={{ margin: 0, fontSize: '10px', fontWeight: '700', color: 'rgba(255,255,255,0.6)', letterSpacing: '0.2em', marginTop: '4px' }}>SIGAP PRESENSI DIGITAL</p>
                      </div>

                      <div style={{ width: '110px', height: '145px', borderRadius: '24px', backgroundColor: 'rgba(255,255,255,0.1)', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '25px', border: '4px solid rgba(255,255,255,0.1)', zIndex: 20 }}>
                         {selectedStudentCard.foto ? (
                           <img src={selectedStudentCard.foto} style={{ width: '100%', height: '100%', objectFit: 'cover', display: 'block' }} crossOrigin="anonymous" />
                         ) : (
                           <User size={64} style={{ color: 'rgba(255,255,255,0.1)' }} />
                         )}
                      </div>

                      <div style={{ width: '100%', zIndex: 30, textAlign: 'center' }}>
                        <div style={{ marginBottom: '12px' }}>
                          <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Nama Lengkap</span>
                          <h4 style={{ margin: 0, fontWeight: '900', fontSize: '16px', lineHeight: '1.2', textShadow: '0 2px 4px rgba(0,0,0,0.3)', wordBreak: 'break-word' }}>{selectedStudentCard.nama}</h4>
                        </div>
                        
                        <div style={{ display: 'flex', justifyContent: 'center', gap: '20px' }}>
                          <div>
                            <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Kelas</span>
                            <p style={{ margin: 0, fontSize: '14px', fontWeight: '900' }}>{selectedStudentCard.kelas}</p>
                          </div>
                          <div>
                            <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>NISN</span>
                            <p style={{ margin: 0, fontSize: '14px', fontWeight: '900' }}>{selectedStudentCard.nisn}</p>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* BACK SIDE */}
                    <div style={{ 
                      background: appConfig.cardTemplateUrl ? 'none' : 'linear-gradient(135deg, #052e16 0%, #14532d 100%)', 
                      color: '#ffffff', 
                      width: '320px', 
                      height: '480px', 
                      borderRadius: '32px', 
                      position: 'relative', 
                      overflow: 'hidden', 
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      padding: '40px 24px',
                      fontFamily: 'Inter, sans-serif',
                      justifyContent: 'center'
                    }}>
                      {appConfig.cardTemplateUrl && (
                        <img 
                           src={appConfig.cardTemplateUrl} 
                           style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover', zIndex: 0 }} 
                           crossOrigin="anonymous" 
                           alt="Background"
                        />
                      )}
                      <div style={{ marginBottom: '30px', textAlign: 'center', zIndex: 10 }}>
                        <h4 style={{ fontSize: '14px', fontWeight: '900', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: '8px' }}>Kartu Presensi Digital</h4>
                        <div style={{ height: '2px', width: '40px', background: '#16a34a', margin: '0 auto' }}></div>
                      </div>

                      <div style={{ backgroundColor: '#ffffff', padding: '15px', borderRadius: '24px', boxShadow: '0 8px 24px rgba(0,0,0,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        <QRCodeSVG 
                          value={selectedStudentCard.nisn} 
                          size={150}
                          level="H"
                          includeMargin={false}
                        />
                      </div>
                      <div style={{ marginTop: '40px', textAlign: 'center', padding: '0 20px' }}>
                        <p style={{ fontSize: '10px', fontWeight: '500', opacity: 0.8, lineHeight: '1.6' }}>
                          Simpan kartu ini dengan baik. Gunakan barcode di atas untuk melakukan presensi pada mesin yang tersedia di sekolah.
                        </p>
                        <p style={{ fontSize: '9px', fontWeight: '900', marginTop: '20px', opacity: 0.4, letterSpacing: '0.2em' }}>MTS NEGERI 2 BOMBANA</p>
                      </div>
                    </div>
                  </div>

                  <div className="mt-8 flex gap-3 w-full max-w-md">
                    <button 
                      onClick={() => downloadAsPDF('print-student-card', `Kartu-Siswa-${selectedStudentCard.nisn}.pdf`)}
                      className="flex-1 bg-green-700 text-white rounded-2xl py-4 md:py-5 font-black text-xs md:text-sm uppercase tracking-widest shadow-xl hover:bg-green-600 active:scale-95 transition-all flex items-center justify-center gap-2"
                    >
                      <Download size={18} /> Simpan Kartu (PDF)
                    </button>
                    <button 
                      onClick={() => setSelectedStudentCard(null)} 
                      className="flex-1 bg-white/10 text-white hover:bg-white/20 rounded-2xl py-4 md:py-5 font-black text-xs md:text-sm uppercase tracking-widest transition-all"
                    >
                      Tutup
                    </button>
                  </div>
              </motion.div>
            </div>
          )}

          {activePanel === 'siswa-personal' && (() => {
            const selectedDateAttendance = attendance.find(a => a.nisn === session?.uid && a.tanggal === siswaDashboardDate);
            const dayMapLong = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
            const dateObj = new Date(siswaDashboardDate);
            const dayNameLong = dayMapLong[dateObj.getDay()];
            const dateSetting = settings.find(s => s.hari === dayNameLong);

            return (
              <motion.div key="siswa-personal" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-8 pb-32">
                <div className="flex flex-wrap gap-4 items-center justify-between p-4 rounded-3xl border" style={{ backgroundColor: 'rgba(255, 255, 255, 0.5)', backdropFilter: 'blur(4px)', borderColor: 'rgba(255, 255, 255, 0.2)' }}>
                  <div className="flex items-center gap-3">
                    <div className="bg-green-100 p-3 rounded-2xl text-green-700">
                      <Calendar size={24} />
                    </div>
                    <div>
                      <h2 className="text-xl font-black text-green-950 uppercase tracking-tighter">Ringkasan Harian</h2>
                      <p className="text-[10px] font-bold uppercase tracking-widest" style={{ color: 'rgba(22, 163, 74, 0.6)' }}>{formatIndoDate(siswaDashboardDate)}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2 bg-white p-2 rounded-2xl shadow-sm border">
                    <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest px-2">Pilih Tanggal:</span>
                    <input 
                      type="date" 
                      value={siswaDashboardDate}
                      onChange={e => setSiswaDashboardDate(e.target.value)}
                      className="border-0 focus:ring-0 text-sm font-black text-green-900 bg-transparent"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 xl:grid-cols-3 gap-8">
                  <div className="xl:col-span-1 space-y-6">
                    <div className="bg-white rounded-[2.5rem] p-8 border border-gray-100 shadow-xl relative overflow-hidden group">
                      <div className="absolute top-0 right-0 w-32 h-32 bg-green-50 rounded-bl-[4rem] -z-10 transition-transform group-hover:scale-110" />
                      <h3 className="text-sm font-black text-green-900 uppercase tracking-[0.2em] mb-8">Status Kehadiran</h3>
                      
                      <div className="space-y-8">
                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-4">
                            <div className="w-12 h-12 rounded-2xl bg-zinc-50 border border-zinc-100 flex items-center justify-center text-zinc-400">
                              <Clock size={20} />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Jam Masuk</p>
                              <p className="text-lg font-black text-zinc-900">{dateSetting?.masuk || '--:--'}</p>
                            </div>
                          </div>
                          {selectedDateAttendance?.jam && (
                            <div className="text-right">
                              <p className="text-[10px] font-black text-green-600 uppercase tracking-widest">Absen Masuk</p>
                              <p className="text-lg font-black text-green-700">{selectedDateAttendance.jam}</p>
                            </div>
                          )}
                        </div>

                        <div className="flex items-center justify-between">
                          <div className="flex items-center gap-4">
                            <div className="w-12 h-12 rounded-2xl bg-zinc-50 border border-zinc-100 flex items-center justify-center text-zinc-400">
                              <LogOut size={20} className="rotate-180" />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Jam Pulang</p>
                              <p className="text-lg font-black text-zinc-900">{dateSetting?.pulang || '--:--'}</p>
                            </div>
                          </div>
                          {selectedDateAttendance?.jamPulang && (
                            <div className="text-right">
                              <p className="text-[10px] font-black text-blue-600 uppercase tracking-widest">Absen Pulang</p>
                              <p className="text-lg font-black text-blue-700">{selectedDateAttendance.jamPulang}</p>
                            </div>
                          )}
                        </div>

                        <div className="pt-6 border-t border-gray-50">
                          <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-3">Status Akhir</p>
                          <div className={`py-4 rounded-3xl text-center border-2 transition-all`} style={{
                            ...(selectedDateAttendance?.status === 'Hadir' ? { backgroundColor: '#f0fdf4', borderColor: '#bbf7d0', color: '#166534' } :
                                selectedDateAttendance?.status === 'Izin' ? { backgroundColor: '#eff6ff', borderColor: '#bfdbfe', color: '#1e40af' } :
                                selectedDateAttendance?.status === 'Sakit' ? { backgroundColor: '#fffbeb', borderColor: '#fde68a', color: '#92400e' } :
                                selectedDateAttendance?.status === 'Alfa' ? { backgroundColor: '#fef2f2', borderColor: '#fecaca', color: '#991b1b' } :
                                { backgroundColor: '#f9fafb', borderColor: '#e5e7eb', color: '#9ca3af' })
                          }}>
                            <span className="text-2xl font-black uppercase tracking-widest block">
                              {selectedDateAttendance?.status || 'Belum Absen'}
                            </span>
                            {selectedDateAttendance?.status === 'Hadir' && (
                              <p className="text-[10px] font-black mt-1 uppercase tracking-widest" style={{ color: 'rgba(22, 163, 74, 0.6)' }}>Selamat Belajar!</p>
                            )}
                            {selectedDateAttendance?.terlambat && selectedDateAttendance.terlambat > 0 ? (
                               <p className="text-[10px] font-black mt-2 px-3 py-1 rounded-full inline-block" style={{ backgroundColor: '#fee2e2', color: '#b91c1c' }}>Terlambat {selectedDateAttendance.terlambat} Menit</p>
                            ) : null}
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="bg-white rounded-[2rem] p-8 border border-gray-100 shadow-sm">
                      <h3 className="text-sm font-black text-green-900 uppercase tracking-widest mb-6 flex items-center gap-2">
                        <BarChart3 size={16} className="text-green-600" />
                        Statistik Kehadiran Bulan Ini
                      </h3>
                      <div className="flex flex-col sm:flex-row items-center gap-6">
                        <div className="relative w-32 h-32 flex items-center justify-center shrink-0">
                          <svg className="w-full h-full transform -rotate-90">
                            <circle cx="64" cy="64" r="56" stroke="currentColor" strokeWidth="12" fill="transparent" className="text-gray-100" />
                            <circle cx="64" cy="64" r="56" stroke="currentColor" strokeWidth="12" fill="transparent" 
                              strokeDasharray={351.8} 
                              strokeDashoffset={351.8 - (351.8 * getSiswaPercentage / 100)} 
                              strokeLinecap="round"
                              className="text-green-600 transition-all duration-1000" 
                            />
                          </svg>
                          <div className="absolute flex flex-col items-center">
                            <span className="text-2xl font-black text-green-900">{getSiswaPercentage}%</span>
                            <span className="text-[8px] font-bold text-gray-400 uppercase">Hadir</span>
                          </div>
                        </div>
                        <div className="flex-1 w-full space-y-3">
                          <div className="p-3 bg-gray-50 rounded-2xl flex items-center justify-between">
                            <span className="text-xs font-bold text-gray-500">Total Absen</span>
                            <span className="text-sm font-black text-green-900">
                                {attendance.filter(a => a.nisn === session?.uid && a.tanggal.startsWith(new Date().toISOString().slice(0, 7))).length}
                            </span>
                          </div>
                          <div className="p-3 bg-green-50 rounded-2xl flex items-center justify-between">
                            <span className="text-xs font-bold text-green-700">Hadir</span>
                            <span className="text-sm font-black text-green-800">
                                {attendance.filter(a => a.nisn === session?.uid && a.tanggal.startsWith(new Date().toISOString().slice(0, 7)) && a.status === 'Hadir').length}
                            </span>
                          </div>
                        </div>
                      </div>
                      <p className="mt-6 text-[10px] text-gray-400 font-medium italic text-center">"Kedisiplinan adalah modal utama meraih cita-cita."</p>
                    </div>
                  </div>

                  <div className="xl:col-span-2 space-y-8">
                    <div 
                      className="flex flex-col items-center w-full p-8 rounded-[3rem] border shadow-sm relative group"
                      style={{ backgroundColor: 'rgba(255, 255, 255, 0.4)', backdropFilter: 'blur(12px)', borderColor: 'rgba(255, 255, 255, 0.5)' }}
                    >
                      <div className="absolute top-6 right-6 opacity-20 group-hover:opacity-100 transition-opacity">
                         <CreditCard className="text-green-900" size={32} />
                      </div>
                      <div id="personal-student-card" className="flex flex-col lg:flex-row gap-4 p-4 w-fit max-w-full overflow-auto scale-75 md:scale-95 lg:scale-100 origin-top" style={{ backgroundColor: 'rgba(0,0,0,0.05)', borderRadius: '2.5rem', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)' }}>
                        {/* FRONT SIDE */}
                        <div style={{ 
                          background: appConfig.cardTemplateUrl ? 'none' : 'linear-gradient(135deg, #14532d 0%, #052e16 100%)', 
                          color: '#ffffff', 
                          width: '320px', 
                          height: '480px', 
                          borderRadius: '32px', 
                          position: 'relative', 
                          overflow: 'hidden', 
                          boxSizing: 'border-box',
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: 'center',
                          padding: '40px 24px',
                          fontFamily: 'Inter, sans-serif'
                        }}>
                          {appConfig.cardTemplateUrl && (
                            <img 
                               src={appConfig.cardTemplateUrl} 
                               style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover', zIndex: 0 }} 
                               crossOrigin="anonymous" 
                               alt="Background"
                            />
                          )}
                          {!appConfig.cardTemplateUrl && (
                            <div style={{ width: '80px', height: '80px', backgroundColor: 'white', borderRadius: '24px', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '10px', marginBottom: '20px', boxShadow: '0 8px 24px rgba(0,0,0,0.2)', zIndex: 10 }}>
                              <img src={appLogo} style={{ width: '100%', height: '100%', objectFit: 'contain' }} crossOrigin="anonymous" />
                            </div>
                          )}

                          <div style={{ textAlign: 'center', marginBottom: '30px', zIndex: 10 }}>
                            <h3 style={{ margin: 0, fontSize: '18px', fontWeight: '900', letterSpacing: '0.05em', textTransform: 'uppercase', lineHeight: '1.2' }}>MTS NEGERI 2 BOMBANA</h3>
                            <p style={{ margin: 0, fontSize: '10px', fontWeight: '700', color: 'rgba(255,255,255,0.6)', letterSpacing: '0.2em', marginTop: '4px' }}>SIGAP PRESENSI DIGITAL</p>
                          </div>

                          <div style={{ width: '110px', height: '145px', borderRadius: '24px', backgroundColor: 'rgba(255,255,255,0.1)', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '25px', border: '4px solid rgba(255,255,255,0.1)', zIndex: 20 }}>
                            {currentSiswaData?.foto ? (
                              <img src={currentSiswaData.foto} style={{ width: '100%', height: '100%', objectFit: 'cover', display: 'block' }} crossOrigin="anonymous" />
                            ) : (
                              <User size={64} style={{ color: 'rgba(255,255,255,0.1)' }} />
                            )}
                          </div>

                          <div style={{ width: '100%', zIndex: 30, textAlign: 'center' }}>
                            <div style={{ marginBottom: '12px' }}>
                              <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Nama Lengkap</span>
                              <h4 style={{ margin: 0, fontWeight: '900', fontSize: '16px', lineHeight: '1.2', textShadow: '0 2px 4px rgba(0,0,0,0.3)', wordBreak: 'break-word' }}>{session?.name || 'Siswa'}</h4>
                            </div>
                            
                            <div style={{ display: 'flex', justifyContent: 'center', gap: '20px' }}>
                              <div>
                                <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>Kelas</span>
                                <p style={{ margin: 0, fontSize: '14px', fontWeight: '900' }}>{session?.kelas}</p>
                              </div>
                              <div>
                                <span style={{ fontSize: '9px', fontWeight: '900', color: 'rgba(255,255,255,0.4)', textTransform: 'uppercase', letterSpacing: '0.1em' }}>NISN</span>
                                <p style={{ margin: 0, fontSize: '14px', fontWeight: '900' }}>{session?.uid}</p>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* BACK SIDE */}
                        <div style={{ 
                          background: appConfig.cardTemplateUrl ? 'none' : 'linear-gradient(135deg, #052e16 0%, #14532d 100%)', 
                          color: '#ffffff', 
                          width: '320px', 
                          height: '480px', 
                          borderRadius: '32px', 
                          position: 'relative', 
                          overflow: 'hidden', 
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: 'center',
                          padding: '40px 24px',
                          fontFamily: 'Inter, sans-serif',
                          justifyContent: 'center'
                        }}>
                          {appConfig.cardTemplateUrl && (
                            <img 
                               src={appConfig.cardTemplateUrl} 
                               style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover', zIndex: 0 }} 
                               crossOrigin="anonymous" 
                               alt="Background"
                            />
                          )}
                          <div style={{ marginBottom: '30px', textAlign: 'center', zIndex: 10 }}>
                            <h4 style={{ fontSize: '14px', fontWeight: '900', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: '8px' }}>Kartu Presensi Digital</h4>
                            <div style={{ height: '2px', width: '40px', background: '#16a34a', margin: '0 auto' }}></div>
                          </div>

                          <div style={{ backgroundColor: '#ffffff', padding: '15px', borderRadius: '24px', boxShadow: '0 8px 24px rgba(0,0,0,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                            <QRCodeSVG 
                              value={session?.uid || ''} 
                              size={150}
                              level="H"
                              includeMargin={false}
                            />
                          </div>

                          <div style={{ marginTop: '40px', textAlign: 'center', padding: '0 20px' }}>
                            <p style={{ fontSize: '10px', fontWeight: '500', opacity: 0.8, lineHeight: '1.6' }}>
                              Simpan kartu ini dengan baik. Gunakan barcode di atas untuk melakukan presensi pada mesin yang tersedia di sekolah.
                            </p>
                            <p style={{ fontSize: '9px', fontWeight: '900', marginTop: '20px', opacity: 0.4, letterSpacing: '0.2em' }}>MTS NEGERI 2 BOMBANA</p>
                          </div>
                        </div>
                      </div>

                      <div className="mt-8 flex gap-4 w-full justify-center">
                        <button 
                          onClick={() => downloadAsPDF('personal-student-card', `Kartu-${session?.uid}.pdf`)}
                          className="bg-green-800 text-white px-8 py-4 rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl flex items-center justify-center gap-2 hover:bg-green-700 active:scale-95 transition-all"
                        >
                          <Download size={18} /> Simpan Kartu (PDF)
                        </button>
                      </div>
                    </div>

                    <div className="bg-white rounded-[2rem] shadow-sm border border-gray-100 overflow-hidden">
                      <div className="p-6 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
                        <h3 className="text-sm font-black text-green-900 uppercase tracking-widest">Riwayat Kehadiran Terakhir</h3>
                        <div className="bg-white px-3 py-1 rounded-full border border-gray-200 text-[9px] font-black uppercase text-gray-400">7 Data Terakhir</div>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs text-left">
                          <thead>
                            <tr style={{ backgroundColor: 'rgba(249, 250, 251, 0.5)' }}>
                              <th className="px-6 py-4 font-black">TANGGAL</th>
                              <th className="px-6 py-4 font-black">JAM MASUK</th>
                              <th className="px-6 py-4 font-black">JAM PULANG</th>
                              <th className="px-6 py-4 font-black">STATUS</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-50">
                            {attendance
                              .filter(a => a.nisn === session?.uid)
                              .sort((a, b) => b.tanggal.localeCompare(a.tanggal))
                              .slice(0, 7)
                              .map((a, i) => (
                                <tr key={i} className="hover:bg-gray-50 transition-colors">
                                  <td className="px-6 py-4 font-bold text-gray-500 text-nowrap">{formatIndoDate(a.tanggal)}</td>
                                  <td className="px-6 py-4 font-bold text-green-600">{a.jam || '-'}</td>
                                  <td className="px-6 py-4 font-bold text-blue-600">{a.jamPulang || '-'}</td>
                                  <td className="px-6 py-4">
                                    <span className={`px-2 py-0.5 rounded-full text-[10px] font-black ${
                                      a.status === 'Hadir' ? 'bg-green-100 text-green-700' : 
                                      a.status === 'Alfa' ? 'bg-red-100 text-red-700' :
                                      'bg-amber-100 text-amber-700'
                                    }`}>
                                      {a.status}
                                    </span>
                                  </td>
                                </tr>
                            ))}
                            {attendance.filter(a => a.nisn === session?.uid).length === 0 && (
                              <tr>
                                <td colSpan={4} className="px-6 py-12 text-center text-gray-400 font-bold uppercase tracking-widest text-[10px]">Belum ada riwayat kehadiran.</td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </motion.div>
            );
          })()}

        </AnimatePresence>
        {renderSessionDetailModal()}
        {renderSuccessToast()}
        <AnimatePresence>
          {showCelebration && (
             <motion.div 
               initial={{ opacity: 0 }}
               animate={{ opacity: 1 }}
               exit={{ opacity: 0 }}
               className="fixed inset-0 z-[400] flex items-center justify-center pointer-events-none px-4"
             >
                <motion.div 
                  initial={{ scale: 0, rotate: -20 }}
                  animate={{ scale: [0, 1.2, 1], rotate: [0, 5, -5, 0] }}
                  className="bg-white/95 backdrop-blur-xl p-10 md:p-14 rounded-[3rem] md:rounded-[4rem] shadow-[0_0_100px_rgba(0,0,0,0.1)] border-4 border-green-500 flex flex-col items-center text-center"
                >
                   <div className="bg-green-100 p-6 rounded-full mb-6 shadow-inner">
                      <Sparkles className="text-green-600" size={64} />
                   </div>
                   <h2 className="text-2xl md:text-4xl font-black text-green-950 uppercase tracking-tighter">DATA BERHASIL DISIMPAN</h2>
                   <p className="text-xs md:text-sm font-bold text-zinc-500 mt-2 uppercase tracking-widest">Sinkronisasi Real-time Berhasil</p>
                </motion.div>
             </motion.div>
          )}
        </AnimatePresence>
        <ConfirmModal />
      </main>
    </div>
  );
}

